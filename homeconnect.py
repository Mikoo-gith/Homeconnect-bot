"""
Bot Tiket HOME SURVEY - Telegram + Google Sheets (v3)
Jalankan: python homeconnect.py (konfigurasi di file .env atau Streamlit Secrets)
"""
import asyncio
import html
import logging
import math
import os
import re
import statistics
from datetime import datetime
from urllib.parse import quote, unquote
from zoneinfo import ZoneInfo

import gspread
import httpx
from dotenv import load_dotenv
from gspread.utils import rowcol_to_a1
from telegram import (
    InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, Update,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application, CallbackQueryHandler, CommandHandler, ContextTypes,
    ConversationHandler, MessageHandler, filters,
)

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("tiketbot")


def env(key, default=""):
    return os.getenv(key, default).strip()


BOT_TOKEN = env("BOT_TOKEN")
SHEET_ID = env("SHEET_ID")
CREDS_FILE = env("GOOGLE_CREDS_FILE", "credentials.json")
START_LINK = env("START_LINK")
START_COORD = env("START_COORD")  # opsional: "lat,lng" (menggantikan START_LINK)
MAX_TICKETS = int(env("MAX_TICKETS", "40"))
STOPS_PER_LINK = min(int(env("STOPS_PER_LINK", "10")), 10)  # 9 waypoint + 1 tujuan
TZ = ZoneInfo(env("TIMEZONE", "Asia/Jakarta"))
ALLOWED = [int(x) for x in env("ALLOWED_USER_IDS").split(",") if x.strip()]
PETUGAS_NAMA = env("PETUGAS_NAMA", "Jatmiko Satrio Wibowo")
PETUGAS_NIK = env("PETUGAS_NIK")
PERUSAHAAN = env("PERUSAHAAN", "Telkomsel Indihome")
CLOSE_KETERANGAN = env("CLOSE_KETERANGAN", "DONE DAN SCC")

if not BOT_TOKEN or not SHEET_ID or not ALLOWED:
    raise SystemExit("BOT_TOKEN, SHEET_ID, dan ALLOWED_USER_IDS wajib diisi di .env / Secrets")

UA = {"User-Agent": "Mozilla/5.0 (Linux; Android 13) AppleWebKit/537.36 Chrome/120 Mobile Safari/537.36"}
GEO_UA = {"User-Agent": "HomeSurveyTiketBot/1.0"}

BASE_HEADERS = [
    "received_at", "ticket_date", "inc", "jenis", "sa", "pic", "odp", "nama",
    "telepon", "lat", "lng", "alamat", "link_maps", "times_received", "last_seen_at",
]
HEADERS_AKTIF = BASE_HEADERS + ["layanan"]
HEADERS_SELESAI = BASE_HEADERS + ["closed_at", "layanan"]
HEADERS_LOG = ["logged_at", "inc", "event", "ticket_date"]
IDX = {h: i for i, h in enumerate(HEADERS_AKTIF)}


# ================================================================ Google Sheets
class Store:
    def __init__(self):
        sh = gspread.service_account(filename=CREDS_FILE).open_by_key(SHEET_ID)
        self.aktif = self._ws(sh, "Aktif", HEADERS_AKTIF)
        self.selesai = self._ws(sh, "Selesai", HEADERS_SELESAI)
        self.log = self._ws(sh, "Log", HEADERS_LOG)

    @staticmethod
    def _ws(sh, name, headers):
        try:
            ws = sh.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=name, rows=1000, cols=len(headers))
        if ws.row_values(1) != headers:
            ws.update(range_name="A1", values=[headers])
        return ws

    def load_aktif(self):
        out = {}
        for n, r in enumerate(self.aktif.get_all_values()[1:], start=2):
            r = r + [""] * (len(HEADERS_AKTIF) - len(r))
            d = dict(zip(HEADERS_AKTIF, r))
            d["row"] = n
            if d["inc"]:
                out[d["inc"]] = d
        return out

    def load_selesai_inc(self):
        return set(self.selesai.col_values(IDX["inc"] + 1)[1:])

    def save(self, new_rows, updates, logs):
        if new_rows:
            self.aktif.append_rows(new_rows, value_input_option="RAW")
        if updates:
            self.aktif.batch_update(updates, value_input_option="RAW")
        if logs:
            self.log.append_rows(logs, value_input_option="RAW")

    def close(self, incs, now, addr_overrides=None):
        """Pindahkan dari Aktif ke Selesai. Return (data tiket, INC tidak ditemukan, sisa aktif dict)."""
        aktif = self.load_aktif()
        addr_overrides = addr_overrides or {}
        found = []
        for i in incs:
            if i in aktif:
                d = dict(aktif[i])
                if i in addr_overrides and addr_overrides[i]:
                    d["alamat"] = addr_overrides[i]
                found.append(d)
        missing = [i for i in incs if i not in aktif]
        if found:
            rows = [[dict(d, closed_at=now).get(h, "") for h in HEADERS_SELESAI] for d in found]
            self.selesai.append_rows(rows, value_input_option="RAW")
            for d in sorted(found, key=lambda x: -x["row"]):
                self.aktif.delete_rows(d["row"])
                aktif.pop(d["inc"], None)
            self.log.append_rows(
                [[now, d["inc"], "SELESAI", d["ticket_date"]] for d in found],
                value_input_option="RAW",
            )
        return found, missing, aktif


store: Store = None


# ================================================================ parsing & geocoding
COORD_PATTERNS = [
    r"[?&](?:q|ll|query|destination|center)=(-?\d+\.\d+)(?:,|%2C)(-?\d+\.\d+)",
    r"@(-?\d+\.\d+),(-?\d+\.\d+)",
    r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)",
]


def extract_coord(text):
    text = unquote(text or "")
    for p in COORD_PATTERNS:
        m = re.search(p, text)
        if m:
            return float(m.group(1)), float(m.group(2))
    return None


def extract_gmaps_place_name(url):
    """Ambil nama jalan/perumahan dari URL Google Maps /maps/place/<nama>/ jika tersedia."""
    if not url:
        return ""
    u = unquote(url)
    m = re.search(r"/maps/place/([^/@?]+)", u)
    if not m:
        return ""
    raw = m.group(1).replace("+", " ").strip()
    if re.match(r"^-?\d+\.\d+\s*,\s*-?\d+\.\d+$", raw):
        return ""
    if re.search(r"\d+°\d+'\d+", raw):
        return ""
    return raw


async def resolve_coord(client, url):
    c = extract_coord(url)
    if c:
        return c
    try:
        r = await client.get(url, follow_redirects=True, timeout=15)
        return extract_coord(str(r.url)) or extract_coord(r.text)
    except Exception:
        return None


def clean_text(text):
    """Bersihkan hasil copy-paste: link markdown [teks](tel:..), karakter tak terlihat, dll."""
    def md(m):
        label, target = m.group(1), m.group(2).strip()
        if target.lower().startswith("http") and not label.lower().startswith("http"):
            return f"{label} {target}"
        return label

    text = re.sub(r"\[([^\]]*)\]\(([^)]*)\)", md, text)
    text = re.sub(r"[\u200b\u200c\u200d\u2060\ufeff]", "", text)
    return text.replace("\xa0", " ").replace("\r", "")


def parse_input(text, default_date):
    """Kembalikan list tiket. Header (tanggal/SA/PIC/jenis) diambil yang terdekat di atas tiket."""
    text = clean_text(text)

    def finds(rx):
        return [(m.start(), m) for m in re.finditer(rx, text, re.M | re.I)]

    dates = finds(r"^[ \t]*(\d{1,2})[/-](\d{1,2})[/-](\d{4})[ \t]*$")
    sas = finds(r"^[ \t]*SA[ \t]*:[ \t]*(.+?)[ \t]*$")
    pics = finds(r"^[ \t]*PIC[ \t]*:[ \t]*(\S+)")
    jenis = finds(r"^[ \t]*TIKET[ \t]+(.+?)[ \t]*$")

    def last(lst, pos, fmt=lambda m: m.group(1), default="-"):
        val = default
        for p, m in lst:
            if p < pos:
                val = fmt(m)
        return val

    marks = [m.start() for m in re.finditer("🏠", text)]
    tickets = []
    for k, pos in enumerate(marks):
        end = marks[k + 1] if k + 1 < len(marks) else len(text)
        b = text[pos:end]
        inc = re.search(r"INC\d+", b)
        url = re.search(r"https?://[^\s\"”]+", b)
        if not (inc and url):
            continue
        before_cp = re.split(r"CP[ \t]*:", b, maxsplit=1)[0]
        layanan = re.search(r"(?<![\w])(\d{8,16})(?![\w])", before_cp)
        odp = re.search(r"ODP-[\w\-/]+", b)
        name = re.search(r"CP[ \t]*:[ \t]*([^\n(]+)", b)
        phone = re.search(r"CP[ \t]*:[^\n]*?\([ \t]*\+?([\d \-]+?)[ \t]*\)", b)
        tickets.append({
            "inc": inc.group(0),
            "layanan": layanan.group(1) if layanan else "-",
            "odp": odp.group(0) if odp else "-",
            "name": name.group(1).strip() if name else "-",
            "phone": phone.group(1).replace(" ", "") if phone else "-",
            "url": url.group(0),
            "ticket_date": last(
                dates, pos,
                lambda m: f"{int(m.group(1)):02d}-{int(m.group(2)):02d}-{m.group(3)}",
                default_date,
            ),
            "jenis": last(jenis, pos),
            "sa": last(sas, pos),
            "pic": last(pics, pos),
            "coord": None, "address": "", "status": None, "since": "",
        })
    return tickets


def needs_full_address(addr):
    """True jika alamat kosong atau tidak mencantumkan nama jalan / nama perumahan."""
    a = (addr or "").strip().lower()
    if not a or a in ("-", "(alamat tidak ditemukan)"):
        return True
    has_detail = bool(re.search(
        r"\b(jl\.?|jalan|perum\.?|perumahan|gg\.?|gang|komp\.?|komplek|blok|cluster|raya|kav\.?|residences?|villa|griya|bukit|taman|pesona|pondok)\b",
        a,
    ))
    return not has_detail


def _format_nominatim(data, url_place=""):
    addr = data.get("address") or {}
    place = (data.get("name") or "").strip()
    if re.match(r"^(rw|rt)\b", place, re.I):
        place = ""

    road = (
        addr.get("road")
        or addr.get("pedestrian")
        or addr.get("footway")
        or addr.get("path")
        or addr.get("street")
        or ""
    ).strip()
    house_no = (addr.get("house_number") or "").strip()
    if road and house_no and house_no not in road:
        road = f"{road} No. {house_no}"

    residential = (
        addr.get("residential")
        or addr.get("allotments")
        or addr.get("estate")
        or addr.get("complex")
        or ""
    ).strip()

    neighbourhood = (
        addr.get("neighbourhood")
        or addr.get("quarter")
        or addr.get("hamlet")
        or ""
    ).strip()
    village = (addr.get("suburb") or addr.get("village") or addr.get("town_district") or "").strip()
    district = (addr.get("city_district") or addr.get("district") or "").strip()
    city = (
        addr.get("city")
        or addr.get("town")
        or addr.get("municipality")
        or addr.get("county")
        or ""
    ).strip()
    state = (addr.get("state") or "").strip()
    postcode = (addr.get("postcode") or "").strip()

    parts = []
    for item in (url_place, place, residential, road, neighbourhood, village, district, city, state, postcode):
        if not item:
            continue
        low = item.lower()
        if any(low == p.lower() or low in p.lower() for p in parts):
            continue
        parts.append(item)

    if parts:
        return ", ".join(parts)
    return (data.get("display_name") or "").strip()


async def reverse_geocode(client, lat, lng, maps_link=""):
    """Alamat lengkap beserta nama jalan / perumahan via OpenStreetMap Nominatim."""
    url_place = extract_gmaps_place_name(maps_link)
    if not url_place and maps_link and "maps.google" not in maps_link and "google.com/maps" not in maps_link:
        try:
            r_redir = await client.get(maps_link, follow_redirects=True, timeout=15)
            url_place = extract_gmaps_place_name(str(r_redir.url))
        except Exception:
            pass

    best_addr = ""
    for zoom in (18, 17, 16):
        try:
            r = await client.get(
                "https://nominatim.openstreetmap.org/reverse",
                headers=GEO_UA,
                timeout=20,
                params={
                    "format": "jsonv2",
                    "lat": lat,
                    "lon": lng,
                    "zoom": zoom,
                    "addressdetails": 1,
                    "accept-language": "id",
                },
            )
            data = r.json()
            cand = _format_nominatim(data, url_place=url_place)
            if cand and (not best_addr or len(cand) > len(best_addr)):
                best_addr = cand
            await asyncio.sleep(1.1)
            if best_addr and not needs_full_address(best_addr):
                break
        except Exception:
            await asyncio.sleep(1.1)

    return best_addr or url_place or "(alamat tidak ditemukan)"


_start_cache = None


async def get_start(client):
    global _start_cache
    if _start_cache:
        return _start_cache
    if START_COORD:
        lat, lng = (float(x) for x in START_COORD.split(","))
        _start_cache = (lat, lng)
    else:
        c = await resolve_coord(client, START_LINK) if START_LINK else None
        if not c:
            raise RuntimeError("Titik awal gagal dibaca. Isi START_COORD (lat,lng) di .env")
        _start_cache = c
    return _start_cache


# ================================================================ WhatsApp & template
def norm_phone(p):
    """Ubah nomor ke format internasional untuk wa.me (62xxxxxxxx)."""
    d = re.sub(r"\D", "", p or "")
    if not d:
        return ""
    if d.startswith("0"):
        d = "62" + d[1:]
    elif d.startswith("8"):
        d = "62" + d
    return d


def wa_text(t):
    nama = PETUGAS_NAMA or "Jatmiko Satrio Wibowo"
    perusahaan = PERUSAHAAN or "Telkomsel Indihome"
    return (
        f"Semangat pagi Bpk/Ibu {t['name']}, Untuk nomor layanan {t.get('layanan') or '-'} berikut. "
        f"Perkenalkan saya {nama} dari {perusahaan}, saya dapat tugas untuk "
        f"pengecekan layanan dan pemetaan pelanggan ke lokasi Bpk/Ibu hari ini dengan tiket "
        f"{t['inc']}. Terimakasih"
    )


def wa_link(t):
    num = norm_phone(t["phone"])
    return f"https://wa.me/{num}?text={quote(wa_text(t), safe='')}" if num else ""


def close_template(d):
    loc = d.get("link_maps")
    if not loc:
        try:
            loc = maps_url((to_float(d["lat"]), to_float(d["lng"])))
        except ValueError:
            loc = "-"
    return (
        f"{d['inc']} | {d.get('layanan') or '-'} | {d['odp']} |\n"
        f"CP : {d['nama']} ({d['telepon']})\n"
        f"Lokasi : {loc}\n\n"
        f"PIC : {PETUGAS_NAMA.upper()}\n"
        f"NIK : {PETUGAS_NIK}\n"
        f"KETERANGAN : {CLOSE_KETERANGAN}"
    )


# ================================================================ routing & clustering
def hav(a, b):
    R = 6371.0
    la1, lo1, la2, lo2 = map(math.radians, (*a, *b))
    d = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * R * math.asin(math.sqrt(d))


def path_len(start, order):
    pts = [start] + [t["coord"] for t in order]
    return sum(hav(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def optimize_chain(start, items):
    """
    Mulai dari titik paling dekat dengan start, lalu lanjut ke titik paling dekat
    dengan titik sebelumnya, dan haluskan jalur dengan 2-opt agar menjadi 1 jalur
    tanpa tumpang tindih atau bolak-balik.
    """
    if len(items) <= 1:
        return items[:]
    first = min(items, key=lambda t: hav(start, t["coord"]))
    pool = [t for t in items if t is not first]
    order, cur = [first], first["coord"]
    while pool:
        nxt = min(pool, key=lambda t: hav(cur, t["coord"]))
        order.append(nxt)
        pool.remove(nxt)
        cur = nxt["coord"]
    improved = True
    while improved:
        improved = False
        for i in range(1, len(order) - 1):
            for j in range(i + 1, len(order)):
                new = order[:i] + order[i:j + 1][::-1] + order[j + 1:]
                if path_len(start, new) + 1e-9 < path_len(start, order):
                    order, improved = new, True
    return order


def cluster_and_route(start, items):
    """
    1. Scan semua koordinat tiket.
    2. Kelompokkan tiket yang saling berdekatan vs yang lokasinya mental jauh banget.
    3. Dari kelompok berdekatan, titik paling dekat dengan start menjadi titik #1 setelah rumah,
       lalu lanjut titik paling dekat dari titik sebelumnya (1 jalur tanpa tumpang tindih/bolak-balik).
    4. Tiket yang lokasinya mental jauh banget ditaruh di paling bawah.
    """
    if len(items) <= 2:
        return optimize_chain(start, items), []

    n = len(items)
    nn_dists = [
        min(hav(items[i]["coord"], items[j]["coord"]) for j in range(n) if i != j)
        for i in range(n)
    ]
    med_nn = statistics.median(nn_dists)
    radius = min(5.5, max(3.5, med_nn * 2.5))

    visited = [False] * n
    clusters = []
    for i in range(n):
        if visited[i]:
            continue
        comp = []
        stack = [i]
        visited[i] = True
        while stack:
            u = stack.pop()
            comp.append(items[u])
            for v in range(n):
                if not visited[v] and hav(items[u]["coord"], items[v]["coord"]) <= radius:
                    visited[v] = True
                    stack.append(v)
        clusters.append(comp)

    clusters.sort(key=lambda c: (-len(c), min(hav(start, t["coord"]) for t in c)))
    main_cluster = clusters[0]
    far_items = [t for c in clusters[1:] for t in c]

    if len(main_cluster) == 1 and n >= 3:
        d_starts = sorted(hav(start, t["coord"]) for t in items)
        med_start = statistics.median(d_starts)
        far_thresh = max(8.0, med_start * 2.2)
        near = [t for t in items if hav(start, t["coord"]) <= far_thresh]
        far = [t for t in items if hav(start, t["coord"]) > far_thresh]
        if near:
            main_order = optimize_chain(start, near)
            tail = main_order[-1]["coord"]
            return main_order, optimize_chain(tail, far)
        return optimize_chain(start, items), []

    main_order = optimize_chain(start, main_cluster)
    tail = main_order[-1]["coord"] if main_order else start
    far_order = optimize_chain(tail, far_items)
    return main_order, far_order


def maps_url(c):
    return f"https://www.google.com/maps?q={c[0]:.6f},{c[1]:.6f}"


def gmaps_links(start, order):
    links, origin, i = [], start, 0
    while i < len(order):
        chunk = order[i:i + STOPS_PER_LINK]
        dest = chunk[-1]["coord"]
        wps = "|".join(f"{t['coord'][0]:.6f},{t['coord'][1]:.6f}" for t in chunk[:-1])
        u = (
            f"https://www.google.com/maps/dir/?api=1&origin={origin[0]:.6f},{origin[1]:.6f}"
            f"&destination={dest[0]:.6f},{dest[1]:.6f}&travelmode=driving"
        )
        if wps:
            u += f"&waypoints={wps}"
        links.append((i + 1, i + len(chunk), u))
        origin, i = dest, i + len(chunk)
    return links


def mark_of(t):
    s = t.get("status")
    far_tag = "📌 JAUH " if t.get("is_far") else ""
    if s == "BARU":
        return f"{far_tag}🆕 BARU "
    if s == "LAMA":
        return f"{far_tag}♻️ LAMA (sejak {t['since']}) "
    if s == "AKTIF":
        return f"{far_tag}⏳ (sejak {t['since']}) "
    return far_tag


def date_key(x):
    try:
        return datetime.strptime(x, "%d-%m-%Y")
    except ValueError:
        return datetime.max


def build_report(items, start):
    """Laporan per tanggal saat add_tiket (HTML Telegram)."""
    e = html.escape
    groups = {}
    for t in items:
        groups.setdefault(t["ticket_date"], []).append(t)
    blocks = []
    for d in sorted(groups, key=date_key):
        g = groups[d]
        main_order, far_order = cluster_and_route(start, g)
        far_incs = {t["inc"] for t in far_order}
        order = main_order + far_order
        paras = ["\n".join([
            f"📅 <b>{e(d)}</b> | {e(g[0]['jenis'])}",
            f"SA: {e(g[0]['sa'])} | PIC: {e(g[0]['pic'])}",
            f"Total {len(order)} tiket | ±{path_len(start, order):.1f} km (garis lurus)",
        ])]
        for n, t in enumerate(order, 1):
            t["is_far"] = t["inc"] in far_incs
            murl = e(t.get("url") or maps_url(t["coord"]))
            wa = wa_link(t)
            links = f'🔗 <a href="{murl}">Buka Maps</a>'
            if wa:
                links += f' | 💬 <a href="{e(wa)}">Chat WA</a>'
            paras.append("\n".join([
                f"{n}. {e(mark_of(t))}<b>{e(t['inc'])}</b> | {e(t['odp'])}",
                f"   🧾 Layanan: {e(t.get('layanan') or '-')}",
                f"   👤 {e(t['name'])} ({e(t['phone'])})",
                f"   📍 {e(t['address'])}",
                f"   {links}",
            ]))
        route = ["🗺️ <b>Rute sekali jalan</b> (mulai dari titik awal):"]
        for a_, b_, u in gmaps_links(start, order):
            route.append(f'• <a href="{e(u)}">Stop {a_}-{b_}</a>')
        paras.append("\n".join(route))
        blocks.append("\n\n".join(paras))
    return blocks


def build_aktif_report(items, start):
    """Satu laporan gabungan seluruh tiket aktif: kelompok berdekatan mulai dari titik awal
    dalam 1 jalur tanpa tumpang tindih, dan tiket mental jauh di paling bawah."""
    e = html.escape
    main_order, far_order = cluster_and_route(start, items)
    far_incs = {t["inc"] for t in far_order}
    order = main_order + far_order

    info_far = f" ({len(main_order)} berdekatan, {len(far_order)} jauh)" if far_order else ""
    paras = ["\n".join([
        f"📋 <b>Rute Tiket Aktif</b>{e(info_far)}",
        f"Total {len(order)} tiket | ±{path_len(start, order):.1f} km (garis lurus)",
    ])]
    for n, t in enumerate(order, 1):
        t["is_far"] = t["inc"] in far_incs
        murl = e(t.get("url") or maps_url(t["coord"]))
        wa = wa_link(t)
        links = f'🔗 <a href="{murl}">Buka Maps</a>'
        if wa:
            links += f' | 💬 <a href="{e(wa)}">Chat WA</a>'
        paras.append("\n".join([
            f"{n}. {e(mark_of(t))}<b>{e(t['inc'])}</b> | {e(t['odp'])}",
            f"   📅 {e(t['ticket_date'])} | {e(t['jenis'])}",
            f"   🧾 Layanan: {e(t.get('layanan') or '-')}",
            f"   👤 {e(t['name'])} ({e(t['phone'])})",
            f"   📍 {e(t['address'])}",
            f"   {links}",
        ]))
    route = ["🗺️ <b>Rute sekali jalan</b> (mulai dari titik awal):"]
    for a_, b_, u in gmaps_links(start, order):
        route.append(f'• <a href="{e(u)}">Stop {a_}-{b_}</a>')
    paras.append("\n".join(route))
    return "\n\n".join(paras)


# ================================================================ konversi data sheet
def to_float(x):
    return float(str(x).strip().replace(",", "."))


def norm_date(s, fallback=""):
    """Terima dd-mm-yyyy, dd/mm/yyyy, yyyy-mm-dd (+jam) -> dd-mm-yyyy."""
    s = (s or "").strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        for cand in (s, s[:10]):
            try:
                return datetime.strptime(cand, fmt).strftime("%d-%m-%Y")
            except ValueError:
                pass
    return fallback


def row_to_item(a):
    try:
        coord = (to_float(a["lat"]), to_float(a["lng"]))
    except ValueError:
        coord = extract_coord(a.get("link_maps", ""))
    if not coord:
        raise ValueError(f"koordinat {a['inc']} tidak terbaca")
    since = norm_date(a["received_at"], "-")
    return {
        "inc": a["inc"], "layanan": a.get("layanan") or "-", "odp": a["odp"] or "-",
        "name": a["nama"] or "-", "phone": a["telepon"] or "-",
        "ticket_date": norm_date(a["ticket_date"], since),
        "jenis": a["jenis"] or "-", "sa": a["sa"] or "-", "pic": a["pic"] or "-",
        "address": a["alamat"] or "-", "url": (a.get("link_maps") or "").strip(),
        "since": since, "coord": coord,
    }


def get_ordered_aktif_items(aktif_dict, start):
    """Kembalikan list item terurut (1..N) sesuai rute berdekatan + outlier di bawah."""
    items = []
    for a in aktif_dict.values():
        try:
            it = row_to_item(a)
            it["status"] = "AKTIF"
            items.append(it)
        except Exception:
            continue
    if not items:
        return []
    main_order, far_order = cluster_and_route(start, items)
    far_incs = {t["inc"] for t in far_order}
    order = main_order + far_order
    for t in order:
        t["is_far"] = t["inc"] in far_incs
    return order


# ================================================================ pipeline
async def process(tickets):
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    aktif = await asyncio.to_thread(store.load_aktif)
    selesai = await asyncio.to_thread(store.load_selesai_inc)

    skipped, failed, new, items = [], [], [], []
    async with httpx.AsyncClient(headers=UA) as client:
        start = await get_start(client)
        for t in tickets:
            if t["inc"] in selesai:
                t["status"] = "SELESAI"
                skipped.append(t)
            elif t["inc"] in aktif:
                try:
                    base = row_to_item(aktif[t["inc"]])
                except ValueError:
                    new.append(t)
                    continue
                t.update(
                    coord=base["coord"],
                    address=base["address"],
                    since=base["since"],
                    status="LAMA",
                    url=t.get("url") or base.get("url", ""),
                )
                items.append(t)
            else:
                new.append(t)

        coords = await asyncio.gather(*(resolve_coord(client, t["url"]) for t in new))
        for t, c in zip(new, coords):
            t["coord"] = c
            if not c:
                t["status"] = "GAGAL"
                failed.append(t)
        for t in new:
            if t["coord"]:
                t["status"] = "BARU"
                t["since"] = datetime.now(TZ).strftime("%d-%m-%Y")
                t["address"] = await reverse_geocode(client, *t["coord"], t.get("url", ""))
                items.append(t)

    new_rows, updates, logs = [], [], []
    c1, c2 = IDX["times_received"] + 1, IDX["last_seen_at"] + 1
    for t in items:
        if t["status"] == "BARU" and t["inc"] not in aktif:
            new_rows.append([
                now, t["ticket_date"], t["inc"], t["jenis"], t["sa"], t["pic"], t["odp"],
                t["name"], t["phone"], f"{t['coord'][0]:.6f}", f"{t['coord'][1]:.6f}",
                t["address"], t["url"], 1, now, t["layanan"],
            ])
            logs.append([now, t["inc"], "BARU", t["ticket_date"]])
        elif t["status"] == "LAMA":
            a = aktif[t["inc"]]
            n = int(a["times_received"] or 1) + 1
            updates.append({
                "range": f"{rowcol_to_a1(a['row'], c1)}:{rowcol_to_a1(a['row'], c2)}",
                "values": [[n, now]],
            })
            if not a.get("layanan") and t["layanan"] != "-":
                updates.append({
                    "range": rowcol_to_a1(a["row"], IDX["layanan"] + 1),
                    "values": [[t["layanan"]]],
                })
            if t.get("url") and a.get("link_maps") != t["url"]:
                updates.append({
                    "range": rowcol_to_a1(a["row"], IDX["link_maps"] + 1),
                    "values": [[t["url"]]],
                })
            logs.append([now, t["inc"], "LAMA", t["ticket_date"]])
    logs += [[now, t["inc"], "DITOLAK_SUDAH_SELESAI", t["ticket_date"]] for t in skipped]
    await asyncio.to_thread(store.save, new_rows, updates, logs)

    return build_report(items, start), items, skipped, failed


# ================================================================ Telegram UI
MENU_ADD, MENU_CLOSE, MENU_AKTIF = "➕ Add Tiket", "✅ Close Ticket", "📋 Tiket Aktif"
MENU_RE = "^(" + "|".join(re.escape(x) for x in (MENU_ADD, MENU_CLOSE, MENU_AKTIF)) + ")$"
MAIN_KB = ReplyKeyboardMarkup([[MENU_ADD, MENU_CLOSE], [MENU_AKTIF]], resize_keyboard=True)
U = filters.User(user_id=ALLOWED)
ADD_CONFIRM, ADD_WAIT, CLOSE_WAIT = range(3)
INC_CMD_RE = r"(?i)^/INC\d+"


def split_text(text, limit=3900, html_mode=False, max_entities=80):
    def size(x):
        return len(html.unescape(re.sub(r"<[^>]+>", "", x))) if html_mode else len(x)

    def ents(x):
        return len(re.findall(r"<(?:a|b)[ >]", x)) if html_mode else 0

    def fits(x):
        return size(x) <= limit and ents(x) <= max_entities

    chunks, buf = [], ""
    for para in (p for p in text.split("\n\n") if p.strip()):
        cand = f"{buf}\n\n{para}" if buf else para
        if fits(cand):
            buf = cand
            continue
        if buf:
            chunks.append(buf)
        buf = ""
        if fits(para):
            buf = para
            continue
        for line in para.split("\n"):
            cand = f"{buf}\n{line}" if buf else line
            if not fits(cand) and buf:
                chunks.append(buf)
                buf = line
            else:
                buf = cand
    if buf.strip():
        chunks.append(buf)
    return chunks


async def send_long(msg, text, html_mode=False, **kw):
    pm = ParseMode.HTML if html_mode else None
    chunks = split_text(text, html_mode=html_mode)
    for i, c in enumerate(chunks):
        await msg.reply_text(
            c, parse_mode=pm, disable_web_page_preview=True,
            **(kw if i == len(chunks) - 1 else {}),
        )


async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 Bot Tiket Home Survey\n\n"
        "/add_tiket - input tiket baru\n\n"
        "/close_ticket - tutup tiket (tap INC)\n\n"
        "/aktif - daftar tiket aktif + rute",
        reply_markup=MAIN_KB,
    )


# ---- add tiket
async def add_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["buf"] = []
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Ya, lanjut", callback_data="add_yes"),
        InlineKeyboardButton("❌ Batal", callback_data="cancel"),
    ]])
    await update.message.reply_text(f"Input tiket baru (maks. {MAX_TICKETS} tiket). Lanjutkan?", reply_markup=kb)
    return ADD_CONFIRM


async def add_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        f"Paste tiketnya sekarang (maks. {MAX_TICKETS}).\n\n"
        "Telegram membatasi 4096 karakter/pesan, jadi kirim bertahap jika banyak, "
        "lalu tekan ✅ Proses."
    )
    return ADD_WAIT


async def add_receive(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    today = datetime.now(TZ).strftime("%d-%m-%Y")
    tickets = parse_input(update.message.text, today)
    if not tickets:
        await update.message.reply_text("Tiket tidak terbaca. Pastikan ada baris 🏠 INC... dan link lokasi.")
        return ADD_WAIT
    buf = ctx.user_data.setdefault("buf", [])
    seen, added = {t["inc"] for t in buf}, 0
    for t in tickets:
        if t["inc"] in seen:
            continue
        if len(buf) >= MAX_TICKETS:
            break
        buf.append(t)
        seen.add(t["inc"])
        added += 1
    full = " Batas maksimal tercapai." if len(buf) >= MAX_TICKETS else ""
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Proses", callback_data="add_process"),
        InlineKeyboardButton("❌ Batal", callback_data="cancel"),
    ]])
    await update.message.reply_text(
        f"Diterima {added} tiket (total {len(buf)}/{MAX_TICKETS}).{full}\n\nKirim lagi atau tekan Proses.",
        reply_markup=kb,
    )
    return ADD_WAIT


async def add_process(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    buf = ctx.user_data.pop("buf", [])
    if not buf:
        await q.edit_message_text("Belum ada tiket. Gunakan /add_tiket lagi.")
        return ConversationHandler.END
    await q.edit_message_text(f"⏳ Memproses {len(buf)} tiket (pencarian alamat ±1 detik/tiket)...")
    try:
        blocks, items, skipped, failed = await process(buf)
    except Exception as e:
        log.exception("process gagal")
        await q.message.reply_text(f"❌ Error: {e}", reply_markup=MAIN_KB)
        return ConversationHandler.END

    n_new = sum(t["status"] == "BARU" for t in items)
    n_old = sum(t["status"] == "LAMA" for t in items)
    await q.message.reply_text(
        "✅ Selesai diproses\n\n"
        f"🆕 Baru: {n_new}\n"
        f"♻️ Lama: {n_old}\n"
        f"⛔ Sudah selesai (dilewati): {len(skipped)}\n"
        f"⚠️ Lokasi gagal dibaca: {len(failed)}"
    )
    for b in blocks:
        await send_long(q.message, b, html_mode=True)
    extra = []
    if skipped:
        extra.append("\n".join(["⛔ Sudah pernah selesai (tidak diproses):"] + [f"- {t['inc']}" for t in skipped]))
    if failed:
        extra.append("⚠️ Koordinat gagal dibaca (cek manual):\n\n" + "\n\n".join(
            f"{t['inc']} | {t['name']}\n{t['url']}" for t in failed))
    if extra:
        await send_long(q.message, "\n\n".join(extra))
    await q.message.reply_text("Data tersimpan di spreadsheet.", reply_markup=MAIN_KB)
    return ConversationHandler.END


# ---- close ticket
def active_list_text(aktif, start=None):
    if start:
        order = get_ordered_aktif_items(aktif, start)
        if order:
            parts = []
            for n, t in enumerate(order, 1):
                tag = "📌 JAUH | " if t.get("is_far") else ""
                parts.append(
                    f"{n}. {tag}/{t['inc'].strip()}\n"
                    f"👤 {t['name']} | {t['odp']}\n"
                    f"📍 {t['address']}"
                )
            return "\n\n".join(parts)

    groups = {}
    for a in aktif.values():
        d = norm_date(a["ticket_date"], norm_date(a["received_at"], "-"))
        groups.setdefault(d, []).append(a)
    parts = []
    for d in sorted(groups, key=date_key):
        parts.append(f"📅 {d}")
        for a in groups[d]:
            parts.append(f"/{a['inc'].strip()}\n👤 {a['nama'] or '-'} | {a['odp'] or '-'}")
    return "\n\n".join(parts)


async def close_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        aktif = await asyncio.to_thread(store.load_aktif)
    except Exception as e:
        log.exception("load aktif gagal")
        await update.message.reply_text(f"❌ Error: {e}", reply_markup=MAIN_KB)
        return ConversationHandler.END
    if not aktif:
        await update.message.reply_text("Tidak ada tiket aktif.", reply_markup=MAIN_KB)
        return ConversationHandler.END
    start = None
    try:
        async with httpx.AsyncClient(headers=UA) as client:
            start = await get_start(client)
    except Exception:
        pass
    await update.message.reply_text(
        f"📋 {len(aktif)} tiket aktif (urut rute dari rumah)\n\n"
        "Tap INC berwarna biru untuk menutup tiket, atau ketik nomor INC manual "
        "(boleh banyak).\n\nBatal: /cancel"
    )
    await send_long(update.message, active_list_text(aktif, start))
    return CLOSE_WAIT


async def close_receive(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    incs = list(dict.fromkeys(re.findall(r"INC\d+", update.message.text.upper())))
    if not incs:
        await update.message.reply_text("Format INC tidak dikenali. Contoh: INC53101831")
        return CLOSE_WAIT
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    try:
        aktif_now = await asyncio.to_thread(store.load_aktif)
        addr_overrides = {}
        start = None
        async with httpx.AsyncClient(headers=UA) as client:
            try:
                start = await get_start(client)
            except Exception:
                pass
            for inc in incs:
                if inc not in aktif_now:
                    continue
                row = aktif_now[inc]
                if needs_full_address(row.get("alamat", "")):
                    coord = None
                    try:
                        coord = (to_float(row["lat"]), to_float(row["lng"]))
                    except ValueError:
                        coord = await resolve_coord(client, row.get("link_maps", ""))
                    if coord:
                        full_addr = await reverse_geocode(client, *coord, row.get("link_maps", ""))
                        if full_addr and full_addr != "(alamat tidak ditemukan)":
                            addr_overrides[inc] = full_addr

        found, missing, sisa_aktif = await asyncio.to_thread(store.close, incs, now, addr_overrides)
    except Exception as e:
        log.exception("close gagal")
        await update.message.reply_text(f"❌ Error: {e}", reply_markup=MAIN_KB)
        return ConversationHandler.END

    status = []
    if found:
        msg_done = f"✅ {', '.join(d['inc'] for d in found)} dipindah ke tab Selesai.\nSisa tiket aktif: {len(sisa_aktif)}"
        if sisa_aktif and start:
            next_order = get_ordered_aktif_items(sisa_aktif, start)
            if next_order:
                nxt = next_order[0]
                msg_done += f"\n\n➡️ Lanjut #1 berikutnya: /{nxt['inc']} — {nxt['name']}\n📍 {nxt['address']}"
        status.append(msg_done)
    if missing:
        status.append("⚠️ Tidak ada di tiket aktif: " + ", ".join(missing))
    await update.message.reply_text("\n\n".join(status), reply_markup=MAIN_KB)

    for d in found:
        await update.message.reply_text(d["alamat"] or "-")
        await update.message.reply_text(close_template(d))
    return ConversationHandler.END


# ---- batal
async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.pop("buf", None)
    await update.effective_message.reply_text("Dibatalkan.", reply_markup=MAIN_KB)
    return ConversationHandler.END


async def cancel_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data.pop("buf", None)
    await q.edit_message_text("Dibatalkan.")
    return ConversationHandler.END


# ---- tiket aktif
async def show_aktif(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    try:
        aktif = await asyncio.to_thread(store.load_aktif)
        if not aktif:
            await msg.reply_text("Tidak ada tiket aktif.")
            return
        items, bad = [], []
        for a in aktif.values():
            try:
                it = row_to_item(a)
                it["status"] = "AKTIF"
                items.append(it)
            except Exception:
                log.exception("baris tidak terbaca: %s", a.get("inc"))
                bad.append(a.get("inc", "?"))
        head = f"📋 {len(items)} tiket aktif"
        if bad:
            head += f" ({len(bad)} gagal dibaca: {', '.join(bad)})"
        await msg.reply_text(head + ":")
        if not items:
            return
        async with httpx.AsyncClient(headers=UA) as client:
            start = await get_start(client)
        await send_long(msg, build_aktif_report(items, start), html_mode=True)
    except Exception as e:
        log.exception("aktif gagal")
        await msg.reply_text(f"❌ Error /aktif: {e}")


async def on_error(update, ctx: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled error", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(f"❌ Error: {ctx.error}")
        except Exception:
            pass


def main():
    global store
    store = Store()
    app = Application.builder().token(BOT_TOKEN).build()

    non_menu_text = U & filters.TEXT & ~filters.COMMAND & ~filters.Regex(MENU_RE)
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("add_tiket", add_start, filters=U),
            MessageHandler(U & filters.Regex(f"^{re.escape(MENU_ADD)}$"), add_start),
            CommandHandler("close_ticket", close_start, filters=U),
            MessageHandler(U & filters.Regex(f"^{re.escape(MENU_CLOSE)}$"), close_start),
        ],
        states={
            ADD_CONFIRM: [
                CallbackQueryHandler(add_yes, pattern="^add_yes$"),
                CallbackQueryHandler(cancel_cb, pattern="^cancel$"),
            ],
            ADD_WAIT: [
                MessageHandler(non_menu_text, add_receive),
                CallbackQueryHandler(add_process, pattern="^add_process$"),
                CallbackQueryHandler(cancel_cb, pattern="^cancel$"),
            ],
            CLOSE_WAIT: [
                MessageHandler(U & filters.Regex(INC_CMD_RE), close_receive),
                MessageHandler(non_menu_text, close_receive),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel, filters=U)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(MessageHandler(U & filters.Regex(INC_CMD_RE), close_receive))
    app.add_handler(CommandHandler("start", cmd_start, filters=U))
    app.add_handler(CommandHandler("aktif", show_aktif, filters=U))
    app.add_handler(MessageHandler(U & filters.Regex(f"^{re.escape(MENU_AKTIF)}$"), show_aktif))
    app.add_error_handler(on_error)
    app.run_polling()


if __name__ == "__main__":
    main()