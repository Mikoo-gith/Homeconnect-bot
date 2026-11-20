"""
Bot Tiket HOME SURVEY - Telegram + Google Sheets (v4)
Jalankan: python homeconnect.py (konfigurasi di file .env atau Streamlit Secrets)
"""
import asyncio
import html
import json
import logging
import math
import os
import re
import statistics
import threading
from datetime import datetime
from urllib.parse import quote, unquote
from zoneinfo import ZoneInfo

import gspread
import httpx
from dotenv import load_dotenv
from gspread.utils import rowcol_to_a1
from telegram import (
    BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, Update,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application, CallbackQueryHandler, CommandHandler, ContextTypes,
    ConversationHandler, MessageHandler, filters,
)

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("tiketbot")


def _load_streamlit_secrets():
    """Sinkronkan st.secrets ke os.environ & tulis credentials.json jika berjalan di Streamlit Cloud."""
    try:
        import streamlit as st
        if not hasattr(st, "secrets"):
            return
        for k, v in st.secrets.items():
            if isinstance(v, (str, int, float, bool)):
                os.environ.setdefault(str(k), str(v))
        creds_path = os.getenv("GOOGLE_CREDS_FILE", "credentials.json").strip() or "credentials.json"
        if not os.path.exists(creds_path):
            if "GOOGLE_CREDS_JSON" in st.secrets:
                raw = st.secrets["GOOGLE_CREDS_JSON"]
                data = json.loads(raw) if isinstance(raw, str) else dict(raw)
                with open(creds_path, "w", encoding="utf-8") as f:
                    json.dump(data, f)
            elif "gcp_service_account" in st.secrets:
                with open(creds_path, "w", encoding="utf-8") as f:
                    json.dump(dict(st.secrets["gcp_service_account"]), f)
    except Exception:
        pass


_load_streamlit_secrets()


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
OSRM_URL = env("OSRM_URL", "https://router.project-osrm.org").rstrip("/")
CLUSTER_MAX_M = float(env("CLUSTER_MAX_M", "2500"))

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
HEADERS_KENDALA = [
    "logged_at", "ticket_date", "inc", "jenis_kendala", "detail",
    "odp", "nama", "telepon", "layanan", "link_maps", "status", "resolved_at",
]
IDX = {h: i for i, h in enumerate(HEADERS_AKTIF)}
IDX_KENDALA = {h: i for i, h in enumerate(HEADERS_KENDALA)}


# ================================================================ Google Sheets
class Store:
    def __init__(self):
        sh = gspread.service_account(filename=CREDS_FILE).open_by_key(SHEET_ID)
        self.aktif = self._ws(sh, "Aktif", HEADERS_AKTIF)
        self.selesai = self._ws(sh, "Selesai", HEADERS_SELESAI)
        self.log = self._ws(sh, "Log", HEADERS_LOG)
        self.kendala = self._ws(sh, "Kendala", HEADERS_KENDALA)

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

    def add_kendala(self, rows):
        """Simpan atau perbarui kendala ke worksheet Kendala tanpa duplikasi status OPEN."""
        if not rows:
            return
        try:
            existing_vals = self.kendala.get_all_values()[1:]
            open_map = {}
            for n, r in enumerate(existing_vals, start=2):
                r = r + [""] * (len(HEADERS_KENDALA) - len(r))
                d = dict(zip(HEADERS_KENDALA, r))
                d["row"] = n
                if d.get("status") == "OPEN":
                    key = (
                        (d["inc"], d["jenis_kendala"], d["detail"])
                        if d["inc"] == "-"
                        else (d["inc"], d["jenis_kendala"])
                    )
                    open_map[key] = d

            updates = []
            to_append = []
            for raw in rows:
                if isinstance(raw, dict):
                    item = {h: str(raw.get(h, "")) for h in HEADERS_KENDALA}
                else:
                    padded = list(raw) + [""] * (len(HEADERS_KENDALA) - len(raw))
                    item = {h: str(padded[i]) for i, h in enumerate(HEADERS_KENDALA)}
                if not item.get("inc"):
                    item["inc"] = "-"
                if not item.get("status"):
                    item["status"] = "OPEN"

                if item["status"] == "OPEN":
                    key = (
                        (item["inc"], item["jenis_kendala"], item["detail"])
                        if item["inc"] == "-"
                        else (item["inc"], item["jenis_kendala"])
                    )
                    if key in open_map:
                        ex = open_map[key]
                        row_idx = ex["row"]
                        merged = dict(ex)
                        merged["logged_at"] = item["logged_at"] or ex["logged_at"]
                        if item["detail"]:
                            merged["detail"] = item["detail"]
                        for fld in ("ticket_date", "odp", "nama", "telepon", "layanan", "link_maps"):
                            if item.get(fld) and item[fld] != "-" and merged.get(fld) in ("", "-"):
                                merged[fld] = item[fld]
                        row_vals = [merged.get(h, "") for h in HEADERS_KENDALA]
                        updates.append({
                            "range": f"A{row_idx}:{rowcol_to_a1(row_idx, len(HEADERS_KENDALA))}",
                            "values": [row_vals],
                        })
                        open_map[key].update(merged)
                        continue

                row_vals = [item.get(h, "") for h in HEADERS_KENDALA]
                to_append.append(row_vals)
                if item["status"] == "OPEN":
                    key = (
                        (item["inc"], item["jenis_kendala"], item["detail"])
                        if item["inc"] == "-"
                        else (item["inc"], item["jenis_kendala"])
                    )
                    stub = dict(item)
                    stub["row"] = len(existing_vals) + len(to_append) + 1
                    open_map[key] = stub

            if updates:
                self.kendala.batch_update(updates, value_input_option="RAW")
            if to_append:
                self.kendala.append_rows(to_append, value_input_option="RAW")
        except Exception:
            log.exception("Gagal menyimpan data ke sheet Kendala")

    def load_kendala(self, status_filter=None):
        """Muat daftar baris kendala dari worksheet Kendala (opsional filter status)."""
        try:
            out = []
            for n, r in enumerate(self.kendala.get_all_values()[1:], start=2):
                r = r + [""] * (len(HEADERS_KENDALA) - len(r))
                d = dict(zip(HEADERS_KENDALA, r))
                d["row"] = n
                if status_filter and d.get("status") != status_filter:
                    continue
                out.append(d)
            return out
        except Exception:
            log.exception("Gagal membaca data dari sheet Kendala")
            return []

    def resolve_kendala(self, inc, jenis=None, now=None):
        """Tandai kendala berstatus OPEN milik INC tertentu menjadi TERATASI."""
        if not inc or inc == "-":
            return
        now = now or datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        try:
            updates = []
            c_status = IDX_KENDALA["status"] + 1
            c_res = IDX_KENDALA["resolved_at"] + 1
            for n, r in enumerate(self.kendala.get_all_values()[1:], start=2):
                r = r + [""] * (len(HEADERS_KENDALA) - len(r))
                d = dict(zip(HEADERS_KENDALA, r))
                if d.get("inc") != inc or d.get("status") != "OPEN":
                    continue
                jk = d.get("jenis_kendala", "")
                if jenis is not None:
                    if isinstance(jenis, str):
                        if jenis.endswith("*"):
                            if not jk.startswith(jenis[:-1]):
                                continue
                        elif jk != jenis:
                            continue
                    elif jk not in jenis:
                        continue
                updates.append({
                    "range": f"{rowcol_to_a1(n, c_status)}:{rowcol_to_a1(n, c_res)}",
                    "values": [["TERATASI", now]],
                })
            if updates:
                self.kendala.batch_update(updates, value_input_option="RAW")
        except Exception:
            log.exception("Gagal memperbarui status kendala untuk %s", inc)

    def close(self, incs, now, addr_overrides=None):
        """Pindahkan dari Aktif ke Selesai dan tutup kendala OPEN. Return (found, missing, sisa aktif)."""
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
            for d in found:
                self.resolve_kendala(d["inc"], jenis=None, now=now)
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
    if not url:
        return None
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
    """Kembalikan list tiket. Blok tanpa URL tetap diproses; blok tanpa INC dicatat rusak."""
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
    if not marks:
        marks = [m.start() for m in re.finditer(r"(?m)^[ \t]*(?:[-•*]\s*)?INC\d+", text, re.I)]

    tickets = []
    for k, pos in enumerate(marks):
        end = marks[k + 1] if k + 1 < len(marks) else len(text)
        b = text[pos:end]
        inc = re.search(r"INC\d+", b, re.I)
        url = re.search(r"https?://[^\s\"”]+", b)
        before_cp = re.split(r"CP[ \t]*:", b, maxsplit=1)[0]
        layanan = re.search(r"(?<![\w])(\d{8,16})(?![\w])", before_cp)
        odp = re.search(r"ODP-[\w\-/]+", b, re.I)
        name = re.search(r"CP[ \t]*:[ \t]*([^\n(]+)", b)
        phone = re.search(r"CP[ \t]*:[^\n]*?\([ \t]*\+?([\d \-]+?)[ \t]*\)", b)
        t_date = last(
            dates, pos,
            lambda m: f"{int(m.group(1)):02d}-{int(m.group(2)):02d}-{m.group(3)}",
            default_date,
        )
        t_jenis = last(jenis, pos)
        t_sa = last(sas, pos)
        t_pic = last(pics, pos)

        if not inc:
            tickets.append({
                "inc": "-",
                "layanan": layanan.group(1) if layanan else "-",
                "odp": odp.group(0).upper() if odp else "-",
                "name": name.group(1).strip() if name and name.group(1).strip() else "-",
                "phone": phone.group(1).replace(" ", "") if phone and phone.group(1).strip() else "-",
                "url": url.group(0) if url else "",
                "ticket_date": t_date,
                "jenis": t_jenis,
                "sa": t_sa,
                "pic": t_pic,
                "coord": None,
                "address": "-",
                "status": "FORMAT_TIKET_RUSAK",
                "since": "",
                "raw_text": b.strip(),
            })
            continue

        tickets.append({
            "inc": inc.group(0).upper(),
            "layanan": layanan.group(1) if layanan else "-",
            "odp": odp.group(0).upper() if odp else "-",
            "name": name.group(1).strip() if name and name.group(1).strip() else "-",
            "phone": phone.group(1).replace(" ", "") if phone and phone.group(1).strip() else "-",
            "url": url.group(0) if url else "",
            "ticket_date": t_date,
            "jenis": t_jenis,
            "sa": t_sa,
            "pic": t_pic,
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
    """Hitung jarak garis lurus Haversine dalam kilometer antara dua titik (lat, lng)."""
    R = 6371.0
    la1, lo1, la2, lo2 = map(math.radians, (*a, *b))
    d = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * R * math.asin(math.sqrt(d))


def fallback_matrix(points):
    """Bangun matriks jarak tempuh estimasi (meter) memakai Haversine x 1.3."""
    n = len(points)
    return [
        [0.0 if i == j else hav(points[i], points[j]) * 1000.0 * 1.3 for j in range(n)]
        for i in range(n)
    ]


_matrix_cache = {}


async def road_matrix(client, points):
    """Ambil matriks jarak tempuh asimetris (meter) via OSRM Table API, fallback Haversine x 1.3."""
    n = len(points)
    if n == 0:
        return [], False
    if n == 1:
        return [[0.0]], False

    cache_key = tuple((round(p[0], 6), round(p[1], 6)) for p in points)
    if cache_key in _matrix_cache:
        return [row[:] for row in _matrix_cache[cache_key]], False

    try:
        dist = [[0.0] * n for _ in range(n)]
        if n <= 90:
            coords_str = ";".join(f"{p[1]:.6f},{p[0]:.6f}" for p in points)
            url = f"{OSRM_URL}/table/v1/driving/{coords_str}?annotations=distance"
            r = await client.get(url, timeout=15)
            r.raise_for_status()
            data = r.json()
            if data.get("code") != "Ok" or "distances" not in data:
                raise ValueError(f"OSRM invalid response: {data.get('code')}")
            raw_d = data["distances"]
            for i in range(n):
                for j in range(n):
                    if i == j:
                        dist[i][j] = 0.0
                    elif raw_d[i][j] is not None:
                        dist[i][j] = float(raw_d[i][j])
                    else:
                        dist[i][j] = hav(points[i], points[j]) * 1000.0 * 1.3
        else:
            chunk = 45
            blocks = [list(range(i, min(i + chunk, n))) for i in range(0, n, chunk)]
            for b_src in blocks:
                for b_dst in blocks:
                    if b_src == b_dst:
                        combined = b_src
                        src_idx = list(range(len(b_src)))
                        dst_idx = list(range(len(b_dst)))
                    else:
                        combined = b_src + b_dst
                        src_idx = list(range(len(b_src)))
                        dst_idx = list(range(len(b_src), len(combined)))
                    coords_str = ";".join(f"{points[k][1]:.6f},{points[k][0]:.6f}" for k in combined)
                    src_str = ";".join(str(x) for x in src_idx)
                    dst_str = ";".join(str(x) for x in dst_idx)
                    url = (
                        f"{OSRM_URL}/table/v1/driving/{coords_str}"
                        f"?annotations=distance&sources={src_str}&destinations={dst_str}"
                    )
                    r = await client.get(url, timeout=20)
                    r.raise_for_status()
                    data = r.json()
                    if data.get("code") != "Ok" or "distances" not in data:
                        raise ValueError(f"OSRM chunk invalid response: {data.get('code')}")
                    sub = data["distances"]
                    for ii, u in enumerate(b_src):
                        for jj, v in enumerate(b_dst):
                            if u == v:
                                dist[u][v] = 0.0
                            elif sub[ii][jj] is not None:
                                dist[u][v] = float(sub[ii][jj])
                            else:
                                dist[u][v] = hav(points[u], points[v]) * 1000.0 * 1.3

        _matrix_cache[cache_key] = [row[:] for row in dist]
        return dist, False
    except Exception as exc:
        log.warning("OSRM gagal/timeout (%s), memakai fallback Haversine x 1.3", exc)
        return fallback_matrix(points), True


def path_len(start, order):
    """Hitung panjang total jalur (km) dari start melewati urutan tiket."""
    if not order:
        return 0.0
    pts = [start] + [t["coord"] for t in order if t.get("coord") is not None]
    if len(pts) <= 1:
        return 0.0
    return sum(hav(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def _open_path_cost(entry_idx, seq, dist):
    """Hitung total jarak tempuh (meter) jalur terbuka dari entry_idx melewati seq."""
    if not seq:
        return 0.0
    total = dist[entry_idx][seq[0]]
    for i in range(len(seq) - 1):
        total += dist[seq[i]][seq[i + 1]]
    return total


def _group_colocated(indices, items, dist):
    """Gabungkan tiket berjarak < 30 m (satu ODP/gang) menjadi super-node berurutan."""
    if len(indices) <= 1:
        return [[u] for u in indices]
    sorted_idx = sorted(indices, key=lambda u: (items[u - 1].get("received_at", ""), items[u - 1].get("inc", "")))
    visited = set()
    groups = []
    for u in sorted_idx:
        if u in visited:
            continue
        comp = []
        stack = [u]
        visited.add(u)
        while stack:
            curr = stack.pop()
            comp.append(curr)
            for v in sorted_idx:
                if v not in visited:
                    d_m = min(dist[curr][v], dist[v][curr])
                    if d_m < 30.0:
                        visited.add(v)
                        stack.append(v)
        comp.sort(key=lambda idx: (items[idx - 1].get("received_at", ""), items[idx - 1].get("inc", "")))
        groups.append(comp)
    return groups


def _held_karp_open(entry_idx, groups, dist):
    """Cari urutan jalur terbuka optimal memakai DP Held-Karp untuk <= 9 super-node."""
    k = len(groups)
    if k == 0:
        return []
    if k == 1:
        return list(groups[0])

    first_g = min(range(k), key=lambda g: dist[entry_idx][groups[g][0]])
    rem = [g for g in range(k) if g != first_g]
    m = len(rem)

    def step_cost(g_from, g_to):
        return dist[groups[g_from][-1]][groups[g_to][0]]

    if m == 0:
        return list(groups[first_g])

    dp = {}
    parent = {}
    for pos, g in enumerate(rem):
        mask = 1 << pos
        dp[(mask, pos)] = step_cost(first_g, g)
        parent[(mask, pos)] = -1

    for mask in range(1, 1 << m):
        for last_pos in range(m):
            if not (mask & (1 << last_pos)):
                continue
            cur_cost = dp.get((mask, last_pos))
            if cur_cost is None:
                continue
            g_u = rem[last_pos]
            for nxt_pos in range(m):
                if mask & (1 << nxt_pos):
                    continue
                g_v = rem[nxt_pos]
                nxt_mask = mask | (1 << nxt_pos)
                cand = cur_cost + step_cost(g_u, g_v)
                if cand < dp.get((nxt_mask, nxt_pos), float("inf")):
                    dp[(nxt_mask, nxt_pos)] = cand
                    parent[(nxt_mask, nxt_pos)] = last_pos

    full_mask = (1 << m) - 1
    best_last = min(range(m), key=lambda pos: dp[(full_mask, pos)])
    seq_pos = []
    cur_mask, cur_pos = full_mask, best_last
    while cur_pos != -1:
        seq_pos.append(cur_pos)
        prev_pos = parent[(cur_mask, cur_pos)]
        cur_mask ^= (1 << cur_pos)
        cur_pos = prev_pos
    seq_pos.reverse()

    ordered_groups = [first_g] + [rem[pos] for pos in seq_pos]
    out = []
    for g in ordered_groups:
        out.extend(groups[g])
    return out


def _optimize_cluster_indices(entry_idx, indices, items, dist):
    """Urutkan tiket dalam satu klaster sebagai open path minimum tanpa memisah grup < 30 m."""
    if len(indices) <= 1:
        return list(indices)

    groups = _group_colocated(indices, items, dist)
    k = len(groups)
    if k <= 9:
        return _held_karp_open(entry_idx, groups, dist)

    def grp_cost(seq_g):
        total = dist[entry_idx][groups[seq_g[0]][0]]
        for i in range(len(seq_g) - 1):
            total += dist[groups[seq_g[i]][-1]][groups[seq_g[i + 1]][0]]
        return total

    first_g = min(range(k), key=lambda g: dist[entry_idx][groups[g][0]])
    pool = [g for g in range(k) if g != first_g]
    order_g = [first_g]
    cur_tail = groups[first_g][-1]
    while pool:
        nxt_g = min(pool, key=lambda g: dist[cur_tail][groups[g][0]])
        order_g.append(nxt_g)
        pool.remove(nxt_g)
        cur_tail = groups[nxt_g][-1]

    improved = True
    while improved:
        improved = False
        best_cost = grp_cost(order_g)
        for i in range(1, k - 1):
            for j in range(i + 1, k):
                cand = order_g[:i] + order_g[i:j + 1][::-1] + order_g[j + 1:]
                c = grp_cost(cand)
                if c + 1e-6 < best_cost:
                    order_g, best_cost, improved = cand, c, True
        for seg_len in (1, 2, 3):
            if seg_len >= k - 1:
                continue
            for i in range(1, k - seg_len + 1):
                seg = order_g[i:i + seg_len]
                rest = order_g[:i] + order_g[i + seg_len:]
                for ins in range(1, len(rest) + 1):
                    if ins == i:
                        continue
                    cand = rest[:ins] + seg + rest[ins:]
                    c = grp_cost(cand)
                    if c + 1e-6 < best_cost:
                        order_g, best_cost, improved = cand, c, True

    out = []
    for g in order_g:
        out.extend(groups[g])
    return out


def optimize_chain(start, items, dist_matrix=None):
    """Urutkan daftar tiket berkoordinat menjadi satu jalur terbuka terpendek dari start."""
    if len(items) <= 1:
        return items[:]
    pts = [start] + [t["coord"] for t in items]
    dist = dist_matrix if dist_matrix is not None else fallback_matrix(pts)
    indices = list(range(1, len(items) + 1))
    ordered_idx = _optimize_cluster_indices(0, indices, items, dist)
    return [items[i - 1] for i in ordered_idx]


def route_with_clusters(start, items, dist_matrix=None):
    """Kelompokkan tiket memakai single-linkage jarak tempuh lalu urutkan antar & dalam klaster."""
    if not items:
        return [], [], []
    pts = [start] + [t["coord"] for t in items]
    dist = dist_matrix if dist_matrix is not None else fallback_matrix(pts)
    m = len(items)

    if m == 1:
        return [items[0]], [], [[items[0]]]

    nn_dists = [
        min(min(dist[i][j], dist[j][i]) for j in range(1, m + 1) if j != i)
        for i in range(1, m + 1)
    ]
    med_nn = statistics.median(nn_dists)
    threshold = min(CLUSTER_MAX_M, max(800.0, 2.0 * med_nn))

    visited = [False] * (m + 1)
    raw_clusters = []
    for i in range(1, m + 1):
        if visited[i]:
            continue
        comp = []
        stack = [i]
        visited[i] = True
        while stack:
            u = stack.pop()
            comp.append(u)
            for v in range(1, m + 1):
                if not visited[v] and min(dist[u][v], dist[v][u]) <= threshold:
                    visited[v] = True
                    stack.append(v)
        raw_clusters.append(comp)

    med_start = statistics.median(dist[0][i] for i in range(1, m + 1))
    far_start_thresh = max(8000.0, 2.2 * med_start)

    normal_clusters = []
    far_clusters = []
    for comp in raw_clusters:
        if m >= 3 and len(comp) == 1 and dist[0][comp[0]] > far_start_thresh:
            far_clusters.append(comp)
        else:
            normal_clusters.append(comp)

    if not normal_clusters:
        normal_clusters = far_clusters
        far_clusters = []

    def order_cluster_list(c_list, cur_pos):
        rem = [list(c) for c in c_list]
        ordered_all = []
        grouped_items = []
        pos = cur_pos
        while rem:
            nxt_c = min(rem, key=lambda c: min(dist[pos][u] for u in c))
            rem.remove(nxt_c)
            seq_idx = _optimize_cluster_indices(pos, nxt_c, items, dist)
            ordered_all.extend(seq_idx)
            grouped_items.append([items[u - 1] for u in seq_idx])
            pos = seq_idx[-1]
        return ordered_all, grouped_items, pos

    main_idx, main_groups, tail_pos = order_cluster_list(normal_clusters, 0)
    far_idx, far_groups, _ = order_cluster_list(far_clusters, tail_pos)

    main_order = [items[u - 1] for u in main_idx]
    far_order = [items[u - 1] for u in far_idx]
    return main_order, far_order, main_groups + far_groups


def cluster_and_route_sync(start, items, dist_matrix=None):
    """Versi sinkron cluster_and_route untuk kebutuhan internal dan pengujian murni."""
    main_order, far_order, _ = route_with_clusters(start, items, dist_matrix=dist_matrix)
    return main_order, far_order


async def cluster_and_route(start, items, dist_matrix=None, client=None):
    """Kelompokkan dan urutkan tiket secara asinkron menggunakan matriks jarak OSRM/fallback."""
    if not items:
        return [], []
    if dist_matrix is None:
        pts = [start] + [t["coord"] for t in items]
        if client is not None:
            dist_matrix, _ = await road_matrix(client, pts)
        else:
            async with httpx.AsyncClient(headers=UA) as temp_client:
                dist_matrix, _ = await road_matrix(temp_client, pts)
    return cluster_and_route_sync(start, items, dist_matrix=dist_matrix)


def maps_url(c):
    return f"https://www.google.com/maps?q={c[0]:.6f},{c[1]:.6f}"


def gmaps_links(start, order, offset=0):
    """Bangun tautan Google Maps berantai (maks STOPS_PER_LINK per link) dengan offset nomor daftar."""
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
        links.append((offset + i + 1, offset + i + len(chunk), u))
        origin, i = dest, i + len(chunk)
    return links


def mark_of(t):
    s = t.get("status")
    no_loc_tag = "❓ LOKASI KOSONG " if t.get("coord") is None else ""
    far_tag = "📌 JAUH " if t.get("is_far") else ""
    prefix = f"{no_loc_tag}{far_tag}"
    if s == "BARU":
        return f"{prefix}🆕 BARU "
    if s == "LAMA":
        return f"{prefix}♻️ LAMA (sejak {t['since']}) "
    if s == "AKTIF":
        return f"{prefix}⏳ (sejak {t['since']}) "
    return prefix


def date_key(x):
    try:
        return datetime.strptime(x, "%d-%m-%Y")
    except ValueError:
        return datetime.max


def _no_loc_sort_key(t):
    return (t.get("received_at") or "9999-99-99", t.get("row", 10**9), t.get("inc", ""))


def _route_km_from_matrix(start, order, pts_map, dist):
    if not order:
        return 0.0
    total_m = 0.0
    prev_idx = 0
    for t in order:
        idx = pts_map[id(t)]
        total_m += dist[prev_idx][idx]
        prev_idx = idx
    return total_m / 1000.0


def _format_ticket_entry(n, t, include_date=False):
    e = html.escape
    wa = wa_link(t)
    link_parts = []
    if t.get("coord") is not None:
        murl = e(t.get("url") or maps_url(t["coord"]))
        link_parts.append(f'🔗 <a href="{murl}">Buka Maps</a>')
    elif t.get("url"):
        link_parts.append(f'🔗 <a href="{e(t["url"])}">Buka Maps</a>')
    if wa:
        link_parts.append(f'💬 <a href="{e(wa)}">Chat WA</a>')
    links = " | ".join(link_parts) if link_parts else "⚠️ Belum ada link lokasi"

    lines = [f"{n}. {e(mark_of(t))}<b>{e(t['inc'])}</b> | {e(t['odp'])}"]
    if include_date:
        lines.append(f"   📅 {e(t['ticket_date'])} | {e(t['jenis'])}")
    lines.extend([
        f"   🧾 Layanan: {e(t.get('layanan') or '-')}",
        f"   👤 {e(t['name'])} ({e(t['phone'])})",
        f"   📍 {e(t['address'])}",
        f"   {links}",
    ])
    return "\n".join(lines)


def build_report(items, start, dist_matrices=None, is_fallback=False):
    """Laporan per tanggal saat add_tiket (HTML Telegram), tiket tanpa lokasi selalu di atas."""
    e = html.escape
    dist_matrices = dist_matrices or {}
    groups = {}
    for t in items:
        groups.setdefault(t["ticket_date"], []).append(t)

    blocks = []
    for d in sorted(groups, key=date_key):
        g = groups[d]
        no_loc = sorted([t for t in g if t.get("coord") is None], key=_no_loc_sort_key)
        located = [t for t in g if t.get("coord") is not None]

        dist = dist_matrices.get(d)
        if located and dist is None:
            dist = fallback_matrix([start] + [t["coord"] for t in located])
        pts_map = {id(t): i + 1 for i, t in enumerate(located)}

        main_order, far_order, clusters_meta = route_with_clusters(start, located, dist_matrix=dist)
        far_incs = {t["inc"] for t in far_order}
        routed_all = main_order + far_order
        total_count = len(no_loc) + len(routed_all)
        total_km = _route_km_from_matrix(start, routed_all, pts_map, dist) if routed_all else 0.0
        dist_label = "(estimasi)" if is_fallback else "(jarak tempuh)"

        head_lines = [
            f"📅 <b>{e(d)}</b> | {e(g[0]['jenis'])}",
            f"SA: {e(g[0]['sa'])} | PIC: {e(g[0]['pic'])}",
            f"Total {total_count} tiket | ±{total_km:.1f} km {dist_label}",
        ]
        if is_fallback and located:
            head_lines.append("⚠️ jarak estimasi (OSRM gagal)")
        paras = ["\n".join(head_lines)]

        num = 1
        for t in no_loc:
            t["is_far"] = False
            paras.append(_format_ticket_entry(num, t, include_date=False))
            num += 1

        prev_idx = 0
        for c_idx, cl in enumerate(clusters_meta, 1):
            cl_m = 0.0
            for t in cl:
                cur_idx = pts_map[id(t)]
                cl_m += dist[prev_idx][cur_idx]
                prev_idx = cur_idx
            paras.append(f"📍 <b>Klaster {c_idx} - {len(cl)} tiket - ±{cl_m / 1000.0:.1f} km</b>")
            for t in cl:
                t["is_far"] = t["inc"] in far_incs
                paras.append(_format_ticket_entry(num, t, include_date=False))
                num += 1

        if routed_all:
            route = ["🗺️ <b>Rute sekali jalan</b> (mulai dari titik awal):"]
            for a_, b_, u in gmaps_links(start, routed_all, offset=len(no_loc)):
                route.append(f'• <a href="{e(u)}">Stop {a_}-{b_}</a>')
            paras.append("\n".join(route))

        blocks.append("\n\n".join(paras))
    return blocks


def build_aktif_report(items, start, dist_matrix=None, is_fallback=False):
    """Laporan gabungan seluruh tiket aktif: tiket tanpa lokasi paling atas, lalu rute per klaster."""
    e = html.escape
    no_loc = sorted([t for t in items if t.get("coord") is None], key=_no_loc_sort_key)
    located = [t for t in items if t.get("coord") is not None]

    if located and dist_matrix is None:
        dist_matrix = fallback_matrix([start] + [t["coord"] for t in located])
    pts_map = {id(t): i + 1 for i, t in enumerate(located)}

    main_order, far_order, clusters_meta = route_with_clusters(start, located, dist_matrix=dist_matrix)
    far_incs = {t["inc"] for t in far_order}
    routed_all = main_order + far_order
    total_count = len(no_loc) + len(routed_all)
    total_km = _route_km_from_matrix(start, routed_all, pts_map, dist_matrix) if routed_all else 0.0
    dist_label = "(estimasi)" if is_fallback else "(jarak tempuh)"

    info_far = f" ({len(main_order)} berdekatan, {len(far_order)} jauh)" if far_order else ""
    head_lines = [
        f"📋 <b>Rute Tiket Aktif</b>{e(info_far)}",
        f"Total {total_count} tiket | ±{total_km:.1f} km {dist_label}",
    ]
    if is_fallback and located:
        head_lines.append("⚠️ jarak estimasi (OSRM gagal)")
    paras = ["\n".join(head_lines)]

    num = 1
    for t in no_loc:
        t["is_far"] = False
        paras.append(_format_ticket_entry(num, t, include_date=True))
        num += 1

    prev_idx = 0
    for c_idx, cl in enumerate(clusters_meta, 1):
        cl_m = 0.0
        for t in cl:
            cur_idx = pts_map[id(t)]
            cl_m += dist_matrix[prev_idx][cur_idx]
            prev_idx = cur_idx
        paras.append(f"📍 <b>Klaster {c_idx} - {len(cl)} tiket - ±{cl_m / 1000.0:.1f} km</b>")
        for t in cl:
            t["is_far"] = t["inc"] in far_incs
            paras.append(_format_ticket_entry(num, t, include_date=True))
            num += 1

    if routed_all:
        route = ["🗺️ <b>Rute sekali jalan</b> (mulai dari titik awal):"]
        for a_, b_, u in gmaps_links(start, routed_all, offset=len(no_loc)):
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
    """Konversi baris sheet Aktif menjadi dict item; kembalikan coord=None jika koordinat kosong."""
    coord = None
    try:
        if str(a.get("lat", "")).strip() and str(a.get("lng", "")).strip():
            coord = (to_float(a["lat"]), to_float(a["lng"]))
    except ValueError:
        coord = None
    if not coord:
        coord = extract_coord(a.get("link_maps", ""))
    since = norm_date(a.get("received_at", ""), "-")
    return {
        "inc": a["inc"], "layanan": a.get("layanan") or "-", "odp": a.get("odp") or "-",
        "name": a.get("nama") or "-", "phone": a.get("telepon") or "-",
        "ticket_date": norm_date(a.get("ticket_date", ""), since),
        "jenis": a.get("jenis") or "-", "sa": a.get("sa") or "-", "pic": a.get("pic") or "-",
        "address": a.get("alamat") or "-", "url": (a.get("link_maps") or "").strip(),
        "since": since, "coord": coord,
        "received_at": a.get("received_at") or "", "row": a.get("row", 10**9),
    }


async def get_ordered_aktif_items(aktif_dict, start, client=None):
    """Kembalikan list item terurut (1..N): tiket tanpa lokasi di atas, lalu klaster terdekat."""
    items = []
    for a in aktif_dict.values():
        try:
            it = row_to_item(a)
            it["status"] = "AKTIF"
            items.append(it)
        except Exception:
            log.exception("Baris aktif gagal dibaca: %s", a.get("inc"))
    if not items:
        return []

    no_loc = sorted([t for t in items if t.get("coord") is None], key=_no_loc_sort_key)
    located = [t for t in items if t.get("coord") is not None]
    for t in no_loc:
        t["is_far"] = False

    if not located:
        return no_loc

    main_order, far_order = await cluster_and_route(start, located, client=client)
    far_incs = {t["inc"] for t in far_order}
    routed = main_order + far_order
    for t in routed:
        t["is_far"] = t["inc"] in far_incs
    return no_loc + routed


# ================================================================ pipeline
def _is_missing_contact(t):
    name = (t.get("name") or "").strip()
    phone = (t.get("phone") or "").strip()
    no_name = not name or name == "-"
    no_phone = not phone or phone == "-" or not norm_phone(phone)
    return no_name, no_phone


async def process(tickets):
    """Proses batch tiket: tangani duplikat, tiket tanpa lokasi, kendala, dan hitung rute."""
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    today_str = datetime.now(TZ).strftime("%d-%m-%Y")
    aktif = await asyncio.to_thread(store.load_aktif)
    selesai = await asyncio.to_thread(store.load_selesai_inc)

    skipped, failed, items = [], [], []
    new_rows, updates, logs, kendala_rows = [], [], [], []
    resolve_ops = []
    seen_batch = {}
    new_row_idx_by_inc = {}
    c1, c2 = IDX["times_received"] + 1, IDX["last_seen_at"] + 1

    def make_kendala(t, jenis_kendala, detail, status="OPEN"):
        return {
            "logged_at": now,
            "ticket_date": t.get("ticket_date") or today_str,
            "inc": t.get("inc") or "-",
            "jenis_kendala": jenis_kendala,
            "detail": detail,
            "odp": t.get("odp") or "-",
            "nama": t.get("name") or "-",
            "telepon": t.get("phone") or "-",
            "layanan": t.get("layanan") or "-",
            "link_maps": t.get("url") or "-",
            "status": status,
            "resolved_at": "",
        }

    async with httpx.AsyncClient(headers=UA) as client:
        start = await get_start(client)

        for raw_t in tickets:
            t = dict(raw_t)
            try:
                if t.get("status") == "FORMAT_TIKET_RUSAK" or not t.get("inc") or t.get("inc") == "-":
                    t["inc"] = "-"
                    kendala_rows.append(
                        make_kendala(
                            t,
                            "FORMAT_TIKET_RUSAK",
                            t.get("raw_text") or "Blok tiket tanpa nomor INC",
                            status="OPEN",
                        )
                    )
                    continue

                if t["inc"] in selesai:
                    t["status"] = "SELESAI"
                    skipped.append(t)
                    logs.append([now, t["inc"], "DITOLAK_SUDAH_SELESAI", t["ticket_date"]])
                    kendala_rows.append(
                        make_kendala(
                            t,
                            "SUDAH_SELESAI",
                            "INC sudah ada di sheet Selesai, ditolak",
                            status="INFO",
                        )
                    )
                    continue

                no_name, no_phone = _is_missing_contact(t)
                if no_name or no_phone:
                    if no_name and no_phone:
                        ket_kontak = "Nama dan nomor telepon kosong"
                    elif no_name:
                        ket_kontak = "Nama pelanggan kosong"
                    else:
                        ket_kontak = "Nomor telepon kosong"
                    kendala_rows.append(make_kendala(t, "KONTAK_KOSONG", ket_kontak, status="OPEN"))

                # Duplikat di dalam satu batch input
                if t["inc"] in seen_batch:
                    prev_t = seen_batch[t["inc"]]
                    if prev_t.get("coord") is None and t.get("url"):
                        c_upg = await resolve_coord(client, t["url"])
                        if c_upg:
                            addr_upg = await reverse_geocode(client, *c_upg, t["url"])
                            prev_t["coord"] = c_upg
                            prev_t["url"] = t["url"]
                            prev_t["address"] = addr_upg or "-"
                            if not addr_upg or addr_upg == "(alamat tidak ditemukan)":
                                kendala_rows.append(
                                    make_kendala(
                                        prev_t,
                                        "ALAMAT_TIDAK_DITEMUKAN",
                                        f"Reverse geocode gagal untuk {c_upg[0]:.6f},{c_upg[1]:.6f}",
                                        status="OPEN",
                                    )
                                )
                            kendala_rows = [
                                k for k in kendala_rows
                                if not (
                                    k["inc"] == t["inc"]
                                    and k["jenis_kendala"] in ("LOKASI_KOSONG", "LOKASI_GAGAL_DIBACA")
                                )
                            ]
                            failed = [f for f in failed if f["inc"] != t["inc"]]
                            if t["inc"] in new_row_idx_by_inc:
                                r_pos = new_row_idx_by_inc[t["inc"]]
                                new_rows[r_pos][IDX["lat"]] = f"{c_upg[0]:.6f}"
                                new_rows[r_pos][IDX["lng"]] = f"{c_upg[1]:.6f}"
                                new_rows[r_pos][IDX["alamat"]] = prev_t["address"]
                                new_rows[r_pos][IDX["link_maps"]] = t["url"]

                    if t["inc"] in new_row_idx_by_inc:
                        r_pos = new_row_idx_by_inc[t["inc"]]
                        new_rows[r_pos][IDX["times_received"]] = int(new_rows[r_pos][IDX["times_received"]]) + 1
                        new_rows[r_pos][IDX["last_seen_at"]] = now
                    logs.append([now, t["inc"], "DUPLIKAT_DIABAIKAN", t["ticket_date"]])
                    continue

                # Tiket sudah ada di sheet Aktif
                if t["inc"] in aktif:
                    a = aktif[t["inc"]]
                    base = row_to_item(a)
                    upgraded_coord = None
                    if base["coord"] is None and t.get("url"):
                        upgraded_coord = await resolve_coord(client, t["url"])

                    if base["coord"] is None and upgraded_coord is not None:
                        new_addr = await reverse_geocode(client, *upgraded_coord, t["url"])
                        new_addr = new_addr or "-"
                        if new_addr == "(alamat tidak ditemukan)":
                            kendala_rows.append(
                                make_kendala(
                                    dict(base, url=t["url"]),
                                    "ALAMAT_TIDAK_DITEMUKAN",
                                    f"Reverse geocode gagal untuk {upgraded_coord[0]:.6f},{upgraded_coord[1]:.6f}",
                                    status="OPEN",
                                )
                            )
                        n_recv = int(a.get("times_received") or 1) + 1
                        updates.append({
                            "range": f"{rowcol_to_a1(a['row'], IDX['lat'] + 1)}:{rowcol_to_a1(a['row'], IDX['last_seen_at'] + 1)}",
                            "values": [[
                                f"{upgraded_coord[0]:.6f}",
                                f"{upgraded_coord[1]:.6f}",
                                new_addr,
                                t["url"],
                                n_recv,
                                now,
                            ]],
                        })
                        resolve_ops.append((t["inc"], {"LOKASI_KOSONG", "LOKASI_GAGAL_DIBACA"}))
                        logs.append([now, t["inc"], "LOKASI_DIPERBARUI", t["ticket_date"]])
                        disp = dict(
                            base,
                            coord=upgraded_coord,
                            address=new_addr,
                            url=t["url"],
                            status="LAMA",
                        )
                        items.append(disp)
                        seen_batch[t["inc"]] = disp
                    else:
                        n_recv = int(a.get("times_received") or 1) + 1
                        updates.append({
                            "range": f"{rowcol_to_a1(a['row'], c1)}:{rowcol_to_a1(a['row'], c2)}",
                            "values": [[n_recv, now]],
                        })
                        logs.append([now, t["inc"], "DUPLIKAT_DIABAIKAN", t["ticket_date"]])
                        disp = dict(base, status="LAMA")
                        if disp["coord"] is None:
                            jk = "LOKASI_GAGAL_DIBACA" if disp.get("url") else "LOKASI_KOSONG"
                            det = (
                                f"Koordinat tidak dapat dibaca dari link: {disp['url']}"
                                if disp.get("url")
                                else "Tiket tidak memiliki link lokasi"
                            )
                            kendala_rows.append(make_kendala(disp, jk, det, status="OPEN"))
                            failed.append(disp)
                        items.append(disp)
                        seen_batch[t["inc"]] = disp
                    continue

                # Tiket baru
                t["status"] = "BARU"
                t["since"] = today_str
                t["received_at"] = now
                if not t.get("url"):
                    t["coord"] = None
                    t["address"] = "-"
                    kendala_rows.append(
                        make_kendala(t, "LOKASI_KOSONG", "Tiket tidak memiliki link lokasi", status="OPEN")
                    )
                    failed.append(t)
                else:
                    c = await resolve_coord(client, t["url"])
                    t["coord"] = c
                    if not c:
                        t["address"] = "-"
                        kendala_rows.append(
                            make_kendala(
                                t,
                                "LOKASI_GAGAL_DIBACA",
                                f"Koordinat tidak dapat dibaca dari link: {t['url']}",
                                status="OPEN",
                            )
                        )
                        failed.append(t)
                    else:
                        addr = await reverse_geocode(client, *c, t["url"])
                        t["address"] = addr or "-"
                        if not addr or addr == "(alamat tidak ditemukan)":
                            kendala_rows.append(
                                make_kendala(
                                    t,
                                    "ALAMAT_TIDAK_DITEMUKAN",
                                    f"Reverse geocode gagal untuk {c[0]:.6f},{c[1]:.6f}",
                                    status="OPEN",
                                )
                            )

                lat_str = f"{t['coord'][0]:.6f}" if t["coord"] else ""
                lng_str = f"{t['coord'][1]:.6f}" if t["coord"] else ""
                new_row_idx_by_inc[t["inc"]] = len(new_rows)
                new_rows.append([
                    now, t["ticket_date"], t["inc"], t["jenis"], t["sa"], t["pic"], t["odp"],
                    t["name"], t["phone"], lat_str, lng_str,
                    t["address"] or "-", t["url"], 1, now, t["layanan"],
                ])
                logs.append([now, t["inc"], "BARU", t["ticket_date"]])
                items.append(t)
                seen_batch[t["inc"]] = t

            except Exception as exc:
                log.exception("Exception saat memproses tiket %s", t.get("inc"))
                kendala_rows.append(
                    make_kendala(t, "ERROR_LAINNYA", f"{type(exc).__name__}: {exc}", status="OPEN")
                )
                inc_safe = t.get("inc") or "-"
                if inc_safe != "-" and inc_safe not in aktif and inc_safe not in selesai and inc_safe not in seen_batch:
                    t["status"] = "BARU"
                    t["since"] = t.get("since") or today_str
                    t["received_at"] = now
                    t["coord"] = t.get("coord")
                    t["address"] = t.get("address") or "-"
                    lat_str = f"{t['coord'][0]:.6f}" if t["coord"] else ""
                    lng_str = f"{t['coord'][1]:.6f}" if t["coord"] else ""
                    new_row_idx_by_inc[inc_safe] = len(new_rows)
                    new_rows.append([
                        now, t.get("ticket_date") or today_str, inc_safe, t.get("jenis") or "-",
                        t.get("sa") or "-", t.get("pic") or "-", t.get("odp") or "-",
                        t.get("name") or "-", t.get("phone") or "-", lat_str, lng_str,
                        t["address"], t.get("url") or "", 1, now, t.get("layanan") or "-",
                    ])
                    logs.append([now, inc_safe, "BARU", t.get("ticket_date") or today_str])
                    items.append(t)
                    seen_batch[inc_safe] = t

        groups = {}
        for t in items:
            groups.setdefault(t["ticket_date"], []).append(t)
        dist_matrices = {}
        any_fallback = False
        for d, g in groups.items():
            located_g = [t for t in g if t.get("coord") is not None]
            if located_g:
                pts = [start] + [t["coord"] for t in located_g]
                dm, fb = await road_matrix(client, pts)
                dist_matrices[d] = dm
                if fb:
                    any_fallback = True

    await asyncio.to_thread(store.save, new_rows, updates, logs)
    await asyncio.to_thread(store.add_kendala, kendala_rows)
    for r_inc, r_jenis in resolve_ops:
        await asyncio.to_thread(store.resolve_kendala, r_inc, r_jenis, now)

    open_kendala = [k for k in kendala_rows if k.get("status") == "OPEN"]
    blocks = build_report(items, start, dist_matrices=dist_matrices, is_fallback=any_fallback)
    return blocks, items, skipped, failed, open_kendala


# ================================================================ Telegram UI
MENU_ADD = "➕ Add Tiket"
MENU_CLOSE = "✅ Close Ticket"
MENU_AKTIF = "📋 Tiket Aktif"
MENU_KENDALA = "⚠️ Kendala"
MENU_RE = "^(" + "|".join(re.escape(x) for x in (MENU_ADD, MENU_CLOSE, MENU_AKTIF, MENU_KENDALA)) + ")$"
MAIN_KB = ReplyKeyboardMarkup(
    [[MENU_ADD, MENU_CLOSE], [MENU_AKTIF, MENU_KENDALA]],
    resize_keyboard=True,
)
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
        "/add_tiket - input tiket baru\n"
        "/aktif - daftar tiket aktif + rute\n"
        "/kendala - daftar tiket kendala lokasi/data\n"
        "/close_ticket - tutup tiket (tap INC)",
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
        await update.message.reply_text("Tiket tidak terbaca. Pastikan ada blok tiket (🏠 / INC...).")
        return ADD_WAIT

    buf = ctx.user_data.setdefault("buf", [])
    seen = {t["inc"] for t in buf if t.get("inc") and t["inc"] != "-"}
    added = 0
    for t in tickets:
        inc = t.get("inc") or "-"
        if inc != "-" and inc in seen:
            continue
        if len(buf) >= MAX_TICKETS:
            break
        buf.append(t)
        if inc != "-":
            seen.add(inc)
        added += 1

    full = " Batas maksimal tercapai." if len(buf) >= MAX_TICKETS else ""
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Proses", callback_data="add_process"),
        InlineKeyboardButton("❌ Batal", callback_data="cancel"),
    ]])
    await update.message.reply_text(
        f"Diterima {added} tiket (total {len(buf)}/{MAX_TICKETS}).{full}\n\n"
        "Kirim lagi atau tekan Proses.",
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
    await q.edit_message_text(f"⏳ Memproses {len(buf)} tiket (pencarian alamat & rute OSRM)...")
    try:
        blocks, items, skipped, failed, open_kendala = await process(buf)
    except Exception as e:
        log.exception("process gagal")
        await q.message.reply_text(f"❌ Error: {e}", reply_markup=MAIN_KB)
        return ConversationHandler.END

    n_baru = sum(t["status"] == "BARU" for t in items)
    n_lama = sum(t["status"] == "LAMA" for t in items)
    n_noloc = sum(t.get("coord") is None for t in items)
    await q.message.reply_text(
        "✅ Selesai diproses\n\n"
        f"🆕 Baru: {n_baru} | ♻️ Sudah aktif (duplikat diabaikan): {n_lama}\n"
        f"❓ Tanpa lokasi (di urutan atas): {n_noloc}\n"
        f"⚠️ Kendala dicatat: {len(open_kendala)} (cek /kendala)\n"
        f"⛔ Sudah selesai (dilewati): {len(skipped)}"
    )
    for b in blocks:
        await send_long(q.message, b, html_mode=True)

    extra = []
    if skipped:
        extra.append("\n".join(
            ["⛔ Sudah pernah selesai (tidak diproses):"] + [f"- {t['inc']}" for t in skipped]
        ))
    if failed:
        extra.append("\n".join(
            ["❓ Tiket tanpa koordinat (tetap masuk Aktif di urutan atas & dicatat di Kendala):"]
            + [f"- {t['inc']} ({t['url'] or 'tanpa link'})" for t in failed]
        ))
    if open_kendala:
        extra.append("\n".join(
            ["⚠️ Rincian kendala OPEN (ketik /kendala untuk daftar lengkap):"]
            + [f"- {k['inc']} [{k['jenis_kendala']}]: {k['detail']}" for k in open_kendala]
        ))
    if extra:
        await send_long(q.message, "\n\n".join(extra))
    await q.message.reply_text("Data tersimpan di spreadsheet.", reply_markup=MAIN_KB)
    return ConversationHandler.END


# ---- close ticket
def _format_close_list(order):
    parts = []
    for n, t in enumerate(order, 1):
        tags = []
        if t.get("coord") is None:
            tags.append("❓ LOKASI KOSONG")
        if t.get("is_far"):
            tags.append("📌 JAUH")
        tag_str = f"{' | '.join(tags)} | " if tags else ""
        parts.append(
            f"{n}. {tag_str}/{t['inc'].strip()}\n"
            f"👤 {t['name']} | {t['odp']}\n"
            f"📍 {t['address']}"
        )
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

    order = []
    try:
        async with httpx.AsyncClient(headers=UA) as client:
            start = await get_start(client)
            order = await get_ordered_aktif_items(aktif, start, client=client)
    except Exception:
        log.exception("gagal mengurutkan daftar close_ticket, fallback tanggal")

    await update.message.reply_text(
        f"📋 {len(aktif)} tiket aktif (tanpa lokasi paling atas, dilanjut rute OSRM)\n\n"
        "Tap INC berwarna biru untuk menutup tiket, atau ketik nomor INC manual "
        "(boleh banyak).\n\nBatal: /cancel"
    )
    if order:
        await send_long(update.message, _format_close_list(order))
    else:
        groups = {}
        for a in aktif.values():
            d = norm_date(a["ticket_date"], norm_date(a["received_at"], "-"))
            groups.setdefault(d, []).append(a)
        parts = []
        for d in sorted(groups, key=date_key):
            parts.append(f"📅 {d}")
            for a in groups[d]:
                parts.append(f"/{a['inc'].strip()}\n👤 {a['nama'] or '-'} | {a['odp'] or '-'}")
        await send_long(update.message, "\n\n".join(parts))
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
        next_order = []
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
                        if str(row.get("lat", "")).strip() and str(row.get("lng", "")).strip():
                            coord = (to_float(row["lat"]), to_float(row["lng"]))
                    except ValueError:
                        coord = None
                    if not coord:
                        coord = await resolve_coord(client, row.get("link_maps", ""))
                    if coord:
                        full_addr = await reverse_geocode(client, *coord, row.get("link_maps", ""))
                        if full_addr and full_addr != "(alamat tidak ditemukan)":
                            addr_overrides[inc] = full_addr

            found, missing, sisa_aktif = await asyncio.to_thread(store.close, incs, now, addr_overrides)
            if sisa_aktif and start:
                next_order = await get_ordered_aktif_items(sisa_aktif, start, client=client)
    except Exception as e:
        log.exception("close gagal")
        await update.message.reply_text(f"❌ Error: {e}", reply_markup=MAIN_KB)
        return ConversationHandler.END

    status = []
    if found:
        msg_done = f"✅ {', '.join(d['inc'] for d in found)} dipindah ke tab Selesai.\nSisa tiket aktif: {len(sisa_aktif)}"
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
            await msg.reply_text("Tidak ada tiket aktif.", reply_markup=MAIN_KB)
            return
        items = []
        for a in aktif.values():
            it = row_to_item(a)
            it["status"] = "AKTIF"
            items.append(it)
        await msg.reply_text(f"📋 {len(items)} tiket aktif:")
        async with httpx.AsyncClient(headers=UA) as client:
            start = await get_start(client)
            located = [t for t in items if t.get("coord") is not None]
            dm, fb = None, False
            if located:
                pts = [start] + [t["coord"] for t in located]
                dm, fb = await road_matrix(client, pts)
        await send_long(msg, build_aktif_report(items, start, dist_matrix=dm, is_fallback=fb), html_mode=True)
    except Exception as e:
        log.exception("aktif gagal")
        await msg.reply_text(f"❌ Error /aktif: {e}")


# ---- kendala
async def show_kendala(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    e = html.escape
    try:
        open_rows = await asyncio.to_thread(store.load_kendala, "OPEN")
        aktif_map = await asyncio.to_thread(store.load_aktif)

        # Pastikan tiket aktif yang lokasinya masih kosong selalu tampil meskipun belum ada di sheet Kendala
        seen_loc_incs = {
            r.get("inc") for r in open_rows
            if r.get("jenis_kendala") in ("LOKASI_KOSONG", "LOKASI_GAGAL_DIBACA")
        }
        extra_rows = []
        for inc, a in aktif_map.items():
            it = row_to_item(a)
            if it.get("coord") is None and inc not in seen_loc_incs:
                has_url = bool(it.get("url"))
                extra_rows.append({
                    "logged_at": a.get("received_at") or "",
                    "ticket_date": norm_date(a.get("ticket_date", ""), "-"),
                    "inc": inc,
                    "jenis_kendala": "LOKASI_GAGAL_DIBACA" if has_url else "LOKASI_KOSONG",
                    "detail": f"Koordinat gagal dibaca: {it['url']}" if has_url else "Tiket tidak memiliki link lokasi",
                    "odp": it.get("odp") or "-",
                    "nama": it.get("name") or "-",
                    "telepon": it.get("phone") or "-",
                    "layanan": it.get("layanan") or "-",
                    "link_maps": it.get("url") or "-",
                    "status": "OPEN",
                })

        all_open = open_rows + extra_rows
        if not all_open:
            await msg.reply_text("✅ Tidak ada tiket kendala (OPEN) saat ini.", reply_markup=MAIN_KB)
            return

        priority = {
            "LOKASI_KOSONG": 0,
            "LOKASI_GAGAL_DIBACA": 1,
            "ALAMAT_TIDAK_DITEMUKAN": 2,
            "KONTAK_KOSONG": 3,
            "FORMAT_TIKET_RUSAK": 4,
            "ERROR_LAINNYA": 5,
        }
        all_open.sort(
            key=lambda d: (
                priority.get(d.get("jenis_kendala", ""), 9),
                date_key(norm_date(d.get("ticket_date", ""), "")),
                d.get("inc", ""),
            )
        )

        paras = [f"⚠️ <b>Daftar Tiket Kendala OPEN ({len(all_open)} item)</b>"]
        for n, d in enumerate(all_open, 1):
            inc = (d.get("inc") or "-").strip()
            jk = d.get("jenis_kendala") or "-"
            det = d.get("detail") or "-"
            link = (d.get("link_maps") or "").strip()
            tap_inc = f"/{inc}" if inc != "-" and inc in aktif_map else inc
            link_line = (
                f'   🔗 <a href="{e(link)}">Cek Link Mentah</a>'
                if link and link != "-"
                else "   🔗 Link lokasi: <i>tidak ada</i>"
            )
            paras.append("\n".join([
                f"{n}. <b>{e(tap_inc)}</b> | {e(d.get('odp') or '-')}",
                f"   ⚠️ <b>{e(jk)}</b>: {e(det)}",
                f"   📅 {e(norm_date(d.get('ticket_date', ''), '-'))} | Layanan: {e(d.get('layanan') or '-')}",
                f"   👤 {e(d.get('nama') or '-')} ({e(d.get('telepon') or '-')})",
                link_line,
            ]))

        await send_long(msg, "\n\n".join(paras), html_mode=True)
    except Exception as err:
        log.exception("kendala gagal")
        await msg.reply_text(f"❌ Error /kendala: {err}")


async def on_error(update, ctx: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled error", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(f"❌ Error: {ctx.error}")
        except Exception:
            pass


async def post_init(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start", "Menu utama bot"),
        BotCommand("add_tiket", "Input tiket baru"),
        BotCommand("aktif", "Daftar tiket aktif & rute OSRM"),
        BotCommand("kendala", "Daftar tiket kendala lokasi & data"),
        BotCommand("close_ticket", "Tutup tiket (tap INC)"),
        BotCommand("cancel", "Batalkan proses saat ini"),
    ])


def _build_application():
    global store
    store = Store()
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

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
    app.add_handler(CommandHandler("kendala", show_kendala, filters=U))
    app.add_handler(MessageHandler(U & filters.Regex(f"^{re.escape(MENU_KENDALA)}$"), show_kendala))
    app.add_error_handler(on_error)
    return app


def main():
    app = _build_application()
    app.run_polling()


def _is_streamlit_runtime():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__":
    if _is_streamlit_runtime():
        import streamlit as st

        @st.cache_resource
        def _start_bot_daemon():
            def _runner():
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                app = _build_application()
                app.run_polling(stop_signals=None)

            t = threading.Thread(target=_runner, daemon=True, name="telegram-bot-polling")
            t.start()
            return t

        _start_bot_daemon()
        st.set_page_config(page_title="Homeconnect Bot", page_icon="🤖")
        st.title("🤖 Homeconnect Telegram Bot")
        st.success("Status: Aktif & berjalan di latar belakang.")
    else:
        main()