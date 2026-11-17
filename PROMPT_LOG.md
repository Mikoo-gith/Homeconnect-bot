# Log Prompt ke Claude - Proyek Bot Tiket Home Survey

Cara pakai: setiap kali kirim permintaan baru, tambah satu baris di tabel dan rincian di bawahnya.
Status: `Selesai` | `Diuji` | `Revisi` | `Pending`

| No | Tanggal | Ringkasan Prompt | Output | Status |
|----|---------|------------------|--------|--------|
| 1 | 17-11-2026 | Bot Telegram terima tiket HOME SURVEY, urut rute (terjauh terakhir), langkah membuat bot | `bot.py` v1 | Selesai |
| 2 | 17-11-2026 | File `.env` agar token mudah diganti | `.env`, `requirements.txt`, `bot.py` v1.1 | Selesai |
| 3 | 17-11-2026 | `/add_tiket` (maks 40) + konfirmasi, detail + alamat + link maps, rute sekali jalan dari titik awal, per tanggal, Google Sheets time series, `/close_ticket`, log prompt | `bot.py` v2, `.env`, `requirements.txt`, `PROMPT_LOG.md` | Selesai (uji lapangan: Pending) |
| 4 | 17-11-2026 | Minta langkah setup runtut dari awal (apa yang diubah selain bot) | Jawaban teks (tanpa file) | Selesai |
| 5 | 17-11-2026 | Tanya: AI apa yang dipakai / kenapa tidak ada API AI | Jawaban teks (bot tanpa AI, rule-based) | Selesai |
| 6 | 17-11-2026 | Perbaiki `/aktif` tidak muncul; tambah wa.me + template; `/close_ticket` jadi 2 chat (alamat + template) | `bot.py` v3, `.env`, `PROMPT_LOG.md` | Selesai (uji lapangan: Pending) |
| 7 | 17-11-2026 | Nomor layanan belum terdeteksi (contoh tiket berbentuk `[142410352780](tel:142410352780)`); minta deteksi otomatis | `bot.py` v3.1 | Selesai (uji lapangan: Pending) |
| 8 | 17-11-2026 | `/close_ticket` menampilkan daftar tiket aktif dari database sebagai teks biru yang bisa di-tap untuk input INC | `bot.py` v3.2 | Selesai (uji lapangan: Pending) |
| 9 | 17-11-2026 | Rapikan format teks semua command, jarak kosong 1 baris | `bot.py` v3.3 | Selesai (uji lapangan: Pending) |
| 10 | 17-11-2026 | Pertahankan link maps asli murni dari tiket tanpa diubah menjadi link koordinat (`maps?q=`), alamat tetap tampil terpisah, tanpa mengubah bagian lain | `homeconnect.py` (v3.4), `PROMPT_LOG.md` | Selesai (uji lapangan: Pending) |

## Rincian Prompt #6 (checklist)

- [x] `/aktif`: parsing data sheet dibuat toleran (koma/titik desimal, format tanggal, fallback koordinat dari `link_maps`)
- [x] `/aktif`: error ditampilkan ke chat + dicatat di terminal (sebelumnya gagal diam-diam)
- [x] Link WhatsApp `wa.me` per tiket, langsung buka chat dengan template pesan
- [x] Nomor telepon dinormalisasi ke 62xxxx (contoh `82138279896` -> `6282138279896`)
- [x] Kolom baru `layanan` (nomor layanan) di tab Aktif & Selesai (ditaruh di belakang, data lama aman)
- [x] `/close_ticket`: chat 1 = alamat lengkap, chat 2 = template laporan (PIC, NIK, KETERANGAN dari `.env`)

## Rincian Prompt #7 (checklist)

- [x] Pembersih teks paste: link markdown `[teks](tel:..)` jadi teks biasa, karakter tak terlihat dibuang
- [x] Nomor layanan = angka 8-16 digit sebelum `CP :` (tidak tergantung posisi kolom/format)
- [x] Tiket lama yang layanan-nya kosong terisi otomatis saat dikirim ulang

## Rincian Prompt #8 (checklist)

- [x] `/close_ticket` membaca tab Aktif dan menampilkan daftar per tanggal: `/INCxxxx | nama | ODP`
- [x] `/INCxxxx` otomatis jadi teks biru (bot_command Telegram); tap = kirim INC sebagai input
- [x] Tap bisa berulang dari daftar yang sama (handler global), input manual tetap bisa
- [x] Setelah close: tampil sisa tiket aktif, lalu 2 chat (alamat + template)

## Rincian Prompt #9 (checklist)

- [x] `/aktif` & hasil `/add_tiket`: header, tiap tiket, dan blok rute dipisah 1 baris kosong
- [x] `/close_ticket`: tiap INC (teks biru) dipisah 1 baris kosong
- [x] `/start`, alur add, ringkasan proses, dan status close ikut dirapikan
- [x] Pesan panjang dipecah per tiket (tidak terpotong); batas 4096 karakter & 100 entitas Telegram dijaga

## Rincian Prompt #10 (checklist)

- [x] `process()` (tiket `BARU`): kolom `link_maps` di Google Sheets kini menyimpan `t["url"]` asli murni dari tiket, bukan hasil generate ulang `maps_url(t["coord"])` (`https://www.google.com/maps?q=lat,lng`)
- [x] `process()` (tiket `LAMA`): mempertahankan `t["url"]` asli dan memperbarui kolom `link_maps` di sheet `Aktif` jika sebelumnya masih menyimpan format `maps?q=`
- [x] `row_to_item()`: memuat `url` langsung dari kolom `link_maps` di spreadsheet agar `/aktif` memakai link asli tiket
- [x] `build_report()`: tautan `🔗 Buka Maps` memakai `t.get("url")` asli dari tiket tanpa diubah ke koordinat, sementara alamat hasil pencarian (`reverse_geocode`) tetap tampil terpisah di baris `📍`
- [x] Bagian lain kode (`parse_input`, routing, `close_template`, UI Telegram) tidak diubah sama sekali; koordinat tetap hanya dipakai di belakang layar untuk pencarian alamat & pengurutan rute

## Catatan / Revisi

- Tiket lama (sebelum v3) belum punya nomor layanan. Akan terisi otomatis jika tiketnya dikirim ulang lewat /add_tiket.
- Tiket yang sudah tersimpan di tab `Aktif` sebelum v3.4 masih berisi link `https://www.google.com/maps?q=...` di kolom `link_maps`; akan otomatis terganti ke link asli tiket jika tiketnya dikirim ulang lewat `/add_tiket`, dan semua tiket baru otomatis menyimpan link asli murni dari tiket.
- Penyebab pasti `/aktif` kosong belum bisa dipastikan dari sisi Claude (tidak bisa melihat spreadsheet). Jika masih gagal, kirim pesan error dari bot/terminal.
