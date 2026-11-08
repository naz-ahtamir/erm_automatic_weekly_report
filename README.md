# Agen Otomatis Laporan Mingguan ERM ESG (WeChat)

Agen ini berjalan otomatis untuk men-generate laporan Excel sesuai template **`ERM 印尼安全组周报数据 (6.25-7.1).xlsx`** dan mengirimkannya lewat **WeChat Desktop** setiap **Kamis pagi pukul 08:00**.

---

## Fitur Utama
1. **Jadwal Otomatis (Kamis 08:00 Pagi)**:
   - Mengambil rentang data dari **Hari Kamis minggu lalu s/d Hari Rabu minggu ini**.
2. **Sumber Data & Filter Otomatis**:
   - **Laporan Non-Conformity / Hazard** ([Google Sheets 1](https://docs.google.com/spreadsheets/d/1vJVOkzRPLIuAWntBHyTFFR4b_95qUggZqVzUYA85R5I/edit?gid=0#gid=0)):
     - Hanya mengambil baris dengan departemen yang mengandung kata **`ESG`**.
     - Hanya mengambil temuan kategori **`BAHAYA TERSEMBUNYI` / `Hazard`** (mengabaikan temuan `PELANGGARAN 6S`).
     - Memetakan departemen otomatis ke **`ESG高压`** (jika ada kata `HPAL`/`高压`) dan **`ESG后处理`** (untuk ESG lainnya) sesuai contoh template.
   - **Laporan Safety Talk / Pelatihan** ([Google Sheets 2](https://docs.google.com/spreadsheets/d/1mglnfdiGn4-mShHS_RO4ul3q7C6bshkJCEPbAc23tL8/edit?gid=1399125371#gid=1399125371)):
     - Hanya mengambil baris dengan departemen yang mengandung kata **`ESG`**.
3. **Unduh & Tempel Foto Otomatis (`=DISPIMG`)**:
   - Mengunduh foto temuan, foto perbaikan, foto pelatihan, dan foto absensi dari **Tally (`storage.tally.so`)** serta **Google Drive (`drive.google.com`)** dan menyisipkannya langsung ke dalam sel Excel (`xl/cellimages.xml`) persis seperti template.
4. **Pengiriman Otomatis via WeChat Desktop**:
   - Menggunakan aplikasi WeChat Desktop (`/usr/bin/wechat`) untuk mencari nama kontak/grup tujuan, melampirkan file Excel `.xlsx` yang baru dibuat, dan mengirimkan ringkasan laporan.

---

## Cara Mengatur Nama Grup / Kontak WeChat Tujuan

Buka file [`config.json`](file:///mnt/sda3/naz/Software%20Engineer/erm_report/config.json) dan ubah bagian `"target_chat"` sesuai nama kontak atau grup WeChat tujuan Anda:

```json
  "wechat": {
    "enabled": true,
    "mode": "desktop",
    "target_chat": "文件传输助手",
    "send_summary_text": true
  }
```
> Secara default `"target_chat"` diisi `"文件传输助手"` (File Transfer) agar aman saat uji coba. Ganti dengan nama grup atau kontak WeChat yang Anda inginkan.

---

## Cara Menjalankan Agen

### 1. Menjalankan Agen Standby Otomatis (Tunggu Hari Kamis Pukul 08:00)
Cukup jalankan perintah berikut di terminal (dan pastikan aplikasi WeChat Desktop sudah login):
```bash
./start_agent.sh
```
atau:
```bash
python3 agent.py
```
Agen akan terus berjalan (standby) dan otomatis men-generate serta mengirim laporan setiap hari Kamis jam 08:00 pagi.

### 2. Langsung Generate & Kirim ke WeChat Sekarang Juga (Tes Langsung)
```bash
python3 agent.py --run-now
```
Atau sambil menentukan nama penerima WeChat tertentu:
```bash
python3 agent.py --run-now --target-chat "文件传输助手"
```

### 3. Hanya Generate File Excel Tanpa Kirim WeChat
```bash
python3 agent.py --generate-only
```
Atau untuk rentang tanggal tertentu (misal menguji periode `2026-06-25` s/d `2026-07-01`):
```bash
python3 agent.py --generate-only --start-date 2026-06-25 --end-date 2026-07-01
```
Hasil file Excel disimpan di dalam folder `output/`.
