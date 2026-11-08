#!/usr/bin/env python3
"""
Generator Laporan Mingguan ERM ESG (Kamis minggu lalu s/d Rabu minggu ini).
Mengambil data dari Google Sheets (Non-Conformity/Hazard & Safety Talk/Training),
memfilter khusus departemen ESG & kategori Bahaya Tersembunyi (Hazard),
mengunduh foto dari Tally & Google Drive, dan menghasilkan file Excel (.xlsx)
sesuai template 'ERM 印尼安全组周报数据 (M.D-M.D).xlsx'.
"""

import concurrent.futures
import csv
import datetime as dt
import hashlib
import io
import json
import logging
import os
import re
import subprocess
import urllib.parse
import urllib.request
import uuid
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("erm_report.generator")

EXCEL_EPOCH = dt.date(1899, 12, 30)
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def date_to_excel_serial(d: dt.date) -> int:
    """Konversi datetime.date ke nomor seri tanggal Excel."""
    return (d - EXCEL_EPOCH).days


def parse_flexible_date(val: str) -> Optional[dt.date]:
    """Parse tanggal dari berbagai format di Google Sheets (YYYY-MM-DD, M/D/YYYY, D/M/YYYY)."""
    if not val:
        return None
    s = val.strip().split(" ")[0].strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def get_thursday_to_wednesday_range(ref_date: Optional[dt.date] = None) -> Tuple[dt.date, dt.date]:
    """
    Hitung rentang laporan: Kamis minggu lalu s/d Rabu minggu ini.
    - Jika dijalankan pada hari Kamis (weekday == 3), maka:
      end_date = Rabu kemarin (ref_date - 1 hari)
      start_date = Kamis minggu lalu (ref_date - 7 hari)
    - Jika dijalankan pada hari selain Kamis (misal Jumat/Sabtu/Minggu/Senin/Selasa/Rabu),
      maka diambil siklus Kamis-Rabu terakhir yang sudah selesai (atau jika Rabu, siklus s/d Rabu tersebut).
    """
    if ref_date is None:
        ref_date = dt.date.today()

    # weekday(): Monday=0, Tuesday=1, Wednesday=2, Thursday=3, Friday=4, Saturday=5, Sunday=6
    if ref_date.weekday() == 3:  # Kamis
        end_date = ref_date - dt.timedelta(days=1)  # Rabu minggu ini (kemarin)
        start_date = ref_date - dt.timedelta(days=7)  # Kamis minggu lalu
    else:
        # Cari hari Rabu terakhir (termasuk hari ini jika hari ini Rabu)
        days_since_wed = (ref_date.weekday() - 2) % 7
        end_date = ref_date - dt.timedelta(days=days_since_wed)
        start_date = end_date - dt.timedelta(days=6)

    return start_date, end_date


def map_esg_department(raw_dept: str, raw_area: str = "") -> str:
    """
    Petakan nama departemen ke 'ESG高压' (jika ada kata HPAL/高压)
    atau 'ESG后处理' (untuk ESG lainnya) sesuai template.
    """
    combined = f"{raw_dept} {raw_area}".upper()
    if "HPAL" in combined or "高压" in combined:
        return "ESG高压"
    return "ESG后处理"


def map_hazard_Finding_and_type(row: List[str]) -> Tuple[str, str]:
    """
    Ambil deskripsi temuan dari kolom [13..27] (atau [28..43] sebagai cadangan)
    dan petakan ke jenis bahaya tersembunyi (隐患类型) sesuai sheet '隐患类型勿删'.
    """
    # Pemetaan indeks kolom Google Sheet 1 ke 隐患类型 di template
    col_to_type = {
        13: "特殊作业 Pekerjaan Khusus",  # 1. Pekerjaan Panas
        14: "特殊作业 Pekerjaan Khusus",  # 2. Ruang Terbatas
        15: "特殊作业 Pekerjaan Khusus",  # 3. Pekerjaan Pengangkatan
        16: "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya",  # 4. Keamanan Berkendara
        17: "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya",  # 5. Pekerjaan Ketinggian
        18: "电气安全 Keselamatan Listrik",  # 6. Pekerjaan Listrik
        19: "特殊作业 Pekerjaan Khusus",  # 7. Pekerjaan Penggalian
        20: "特殊作业 Pekerjaan Khusus",  # 8. Pembatasan Akses Jalan
        21: "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya",  # 9. Manajemen Lokasi
        22: "设备设施 Peralatan dan Fasilitas",  # 10. Hand Tools & Power Tools
        23: "设备设施 Peralatan dan Fasilitas",  # 11. Alat-Alat Berat & Mesin Bergerak
        24: "职业健康监测 Pemantauan Kesehatan Kerja",  # 12. Kesehatan & P3K
        25: "标识标牌 Signboard dan Label keselamatan",  # 13. Komunikasi dan Rambu K3
        26: "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya",  # 14. Traffic
        27: "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya",  # 15. Lain-Lain
    }

    descriptions = []
    mapped_type = "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya"

    for idx in range(13, min(28, len(row))):
        val = (row[idx] or "").strip()
        if val:
            descriptions.append(val)
            if mapped_type == "隐患排查与治理 Pemeriksaan dan Penanggulangan Bahaya" and idx in col_to_type:
                mapped_type = col_to_type[idx]

    if not descriptions:
        for idx in range(28, min(44, len(row))):
            val = (row[idx] or "").strip()
            if val:
                descriptions.append(val)

    desc_text = "; ".join(descriptions).upper()
    return desc_text, mapped_type


def fetch_csv_text(url: str, cache_path: Optional[Path] = None, timeout: int = 90) -> str:
    """Unduh CSV dari Google Sheets menggunakan curl atau urllib."""
    logger.info(f"Mengunduh data dari: {url}")
    try:
        cmd = [
            "curl",
            "-sL",
            "--retry",
            "3",
            "--max-time",
            str(timeout),
            "-A",
            "Mozilla/5.0 (X11; Linux x86_64)",
            url,
        ]
        res = subprocess.run(cmd, capture_output=True, timeout=timeout + 10)
        if res.returncode == 0 and res.stdout:
            text = res.stdout.decode("utf-8", errors="replace")
            if cache_path:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                cache_path.write_text(text, encoding="utf-8")
            return text
    except Exception as e:
        logger.warning(f"curl gagal ({e}), mencoba urllib...")

    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64)"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        text = resp.read().decode("utf-8", errors="replace")
        if cache_path:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(text, encoding="utf-8")
        return text


def extract_first_url(raw_cell: str) -> str:
    """Ekstrak URL pertama jika di dalam satu sel terdapat satu atau beberapa URL."""
    if not raw_cell:
        return ""
    urls = re.findall(r"https?://[^\s,;\"']+", raw_cell.strip())
    return urls[0] if urls else ""


def extract_gdrive_id(url: str) -> Optional[str]:
    """Ekstrak file ID dari URL Google Drive."""
    if "drive.google.com" not in url:
        return None
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    return None


def download_image_bytes(url: str, cache_dir: Path, timeout: int = 25) -> Optional[bytes]:
    """
    Unduh gambar dari Tally atau Google Drive dengan caching lokal.
    Untuk Google Drive, gunakan endpoint thumbnail (w1200) terlebih dahulu agar cepat & ukuran optimal,
    lalu fallback ke uc?export=download jika diperlukan.
    """
    clean_url = extract_first_url(url)
    if not clean_url:
        return None

    cache_dir.mkdir(parents=True, exist_ok=True)
    url_hash = hashlib.md5(clean_url.encode("utf-8")).hexdigest()
    cached_file = cache_dir / f"{url_hash}.jpg"
    if cached_file.exists() and cached_file.stat().st_size > 500:
        return cached_file.read_bytes()

    candidate_urls = []
    gdrive_id = extract_gdrive_id(clean_url)
    if gdrive_id:
        candidate_urls.append(f"https://drive.google.com/thumbnail?id={gdrive_id}&sz=w1200")
        candidate_urls.append(f"https://drive.google.com/uc?export=download&id={gdrive_id}")
    else:
        candidate_urls.append(clean_url)

    for target_url in candidate_urls:
        try:
            cmd = [
                "curl",
                "-sL",
                "--retry",
                "2",
                "--max-time",
                str(timeout),
                "-A",
                "Mozilla/5.0",
                target_url,
            ]
            res = subprocess.run(cmd, capture_output=True, timeout=timeout + 5)
            data = res.stdout
            # Validasi apakah byte yang diunduh adalah gambar (JPEG/PNG/WebP), bukan halaman HTML
            if data and len(data) > 500 and not data[:30].lstrip().startswith((b"<!DOCTYPE", b"<html", b"<HTML")):
                cached_file.write_bytes(data)
                return data
        except Exception as e:
            logger.debug(f"Gagal mengunduh gambar dari {target_url}: {e}")

    return None


class ERMReportGenerator:
    def __init__(self, base_dir: Path, config: Dict[str, Any]):
        self.base_dir = base_dir
        self.config = config
        self.cache_dir = self.base_dir / "cache"
        self.photos_cache_dir = self.cache_dir / "photos"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.photos_cache_dir.mkdir(parents=True, exist_ok=True)

    def fetch_and_filter_hazards(
        self, start_date: dt.date, end_date: dt.date, use_cached_csv: bool = False
    ) -> List[Dict[str, Any]]:
        """
        Ambil dan filter data Non-Conformity (Sheet 1):
        - Tanggal inspeksi dalam [start_date, end_date]
        - Departemen mengandung kata 'ESG'
        - Kategori temuan adalah Hazard / Bahaya Tersembunyi
        """
        csv_cache = self.cache_dir / "non_conformity.csv"
        url = self.config["data_sources"]["non_conformity_csv_url"]
        if use_cached_csv and csv_cache.exists():
            text = csv_cache.read_text(encoding="utf-8", errors="replace")
        else:
            text = fetch_csv_text(url, csv_cache)

        dept_kw = self.config["filters"].get("department_keyword", "ESG").lower()
        type_kws = [
            k.lower()
            for k in self.config["filters"].get(
                "non_conformity_type_keywords", ["BAHAYA TERSEMBUNYI", "HAZARD", "隐患"]
            )
        ]

        reader = csv.reader(io.StringIO(text))
        headers = next(reader, [])
        records: List[Dict[str, Any]] = []

        for row in reader:
            if len(row) < 13:
                continue
            insp_date = parse_flexible_date(row[3])
            if not insp_date or not (start_date <= insp_date <= end_date):
                continue

            raw_dept = (row[8] or "").strip()
            if dept_kw not in raw_dept.lower():
                continue

            finding_type_raw = (row[12] or "").strip()
            if not any(tk in finding_type_raw.lower() for tk in type_kws):
                continue

            raw_area = (row[6] or "").strip()
            mapped_dept = map_esg_department(raw_dept, raw_area)
            bengkel = raw_dept.upper() if raw_dept else "ESG"
            desc, hazard_category = map_hazard_Finding_and_type(row)

            # Level hazard: sesuai template, umumnya '一般隐患' (atau jika ingin tetap '一般隐患')
            level_raw = (row[11] or "").strip()
            hazard_level = "一般隐患"

            pic = (row[52] if len(row) > 52 else "").strip().upper()
            safety_officer = (row[4] or "").strip()
            rekomendasi = ""  # Di template kolom J dikosongkan
            deadline_date = parse_flexible_date(row[51] if len(row) > 51 else "") or insp_date
            status_raw = (row[53] if len(row) > 53 else "").strip().upper()
            kondisi_perbaikan = "已整改" if status_raw in ("CLSD", "CLOSED", "SELESAI", "DONE", "已整改") else "未整改"
            status_perbaikan = "未核对"
            keterangan = (row[55] if len(row) > 55 else "").strip()

            records.append(
                {
                    "date": insp_date,
                    "departemen": mapped_dept,
                    "bengkel": bengkel,
                    "hazard_photo_url": row[9] if len(row) > 9 else "",
                    "deskripsi": desc,
                    "jenis_bahaya": hazard_category,
                    "level_bahaya": hazard_level,
                    "pic": pic,
                    "safety_officer": safety_officer,
                    "rekomendasi": rekomendasi,
                    "deadline": deadline_date,
                    "after_photo_url": row[54] if len(row) > 54 else "",
                    "kondisi_perbaikan": kondisi_perbaikan,
                    "status_perbaikan": status_perbaikan,
                    "keterangan": keterangan,
                }
            )

        records.sort(key=lambda r: r["date"])
        logger.info(
            f"Ditemukan {len(records)} temuan Hazard/Bahaya Tersembunyi ESG pada periode {start_date} s/d {end_date}."
        )
        return records

    def fetch_and_filter_trainings(
        self, start_date: dt.date, end_date: dt.date, use_cached_csv: bool = False
    ) -> List[Dict[str, Any]]:
        """
        Ambil dan filter data Safety Talk / Pelatihan (Sheet 2):
        - Tanggal kegiatan dalam [start_date, end_date]
        - Departemen mengandung kata 'ESG'
        """
        csv_cache = self.cache_dir / "training.csv"
        url = self.config["data_sources"]["training_csv_url"]
        if use_cached_csv and csv_cache.exists():
            text = csv_cache.read_text(encoding="utf-8", errors="replace")
        else:
            text = fetch_csv_text(url, csv_cache)

        dept_kw = self.config["filters"].get("department_keyword", "ESG").lower()
        reader = csv.reader(io.StringIO(text))
        headers = next(reader, [])
        records: List[Dict[str, Any]] = []

        for row in reader:
            if len(row) < 11:
                continue
            act_date = parse_flexible_date(row[5])
            if not act_date or not (start_date <= act_date <= end_date):
                continue

            raw_dept_target = (row[10] if len(row) > 10 else "").strip()
            raw_dept_org = (row[1] if len(row) > 1 else "").strip()
            if dept_kw not in raw_dept_target.lower() and dept_kw not in raw_dept_org.lower():
                continue

            mapped_target = map_esg_department(raw_dept_target or raw_dept_org)
            jenis_pelatihan = (row[3] if len(row) > 3 else "").strip()
            isi_pelatihan = (row[4] if len(row) > 4 else "").strip()
            pemateri = re.sub(r"\s+", " ", (row[12] if len(row) > 12 else "").strip())

            try:
                tka_count = int(float((row[7] if len(row) > 7 else "0").strip() or 0))
            except ValueError:
                tka_count = 0

            try:
                tki_count = int(float((row[8] if len(row) > 8 else "0").strip() or 0))
            except ValueError:
                tki_count = 0

            records.append(
                {
                    "date": act_date,
                    "departemen": raw_dept_org or "HSE/应急管理部",
                    "penerima": mapped_target,
                    "jenis_pelatihan": jenis_pelatihan,
                    "topik": "",
                    "isi_pelatihan": isi_pelatihan,
                    "pemateri": pemateri,
                    "jabatan": "=",
                    "tka": tka_count,
                    "tki": tki_count,
                    "training_photo_url": row[13] if len(row) > 13 else "",
                    "absensi_photo_url": row[16] if len(row) > 16 else "",
                    "pengirim": pemateri,
                }
            )

        records.sort(key=lambda r: r["date"])
        logger.info(f"Ditemukan {len(records)} data Pelatihan/Safety Talk ESG pada periode {start_date} s/d {end_date}.")
        return records

    def generate_report(
        self,
        start_date: Optional[dt.date] = None,
        end_date: Optional[dt.date] = None,
        use_cached_csv: bool = False,
    ) -> Dict[str, Any]:
        """
        Generate file laporan Excel (.xlsx) untuk periode [start_date, end_date].
        Mengembalikan dictionary berisi informasi path file dan jumlah temuan.
        """
        if start_date is None or end_date is None:
            start_date, end_date = get_thursday_to_wednesday_range()

        hazards = self.fetch_and_filter_hazards(start_date, end_date, use_cached_csv=use_cached_csv)
        trainings = self.fetch_and_filter_trainings(start_date, end_date, use_cached_csv=use_cached_csv)

        # Unduh semua foto secara paralel jika diaktifkan
        download_photos = self.config["report"].get("download_photos", True)
        timeout = self.config["report"].get("photo_timeout_seconds", 25)
        max_workers = self.config["report"].get("max_concurrent_downloads", 6)

        url_to_bytes: Dict[str, bytes] = {}
        if download_photos:
            urls_to_fetch = set()
            for h in hazards:
                if h["hazard_photo_url"]:
                    urls_to_fetch.add(h["hazard_photo_url"])
                if h["after_photo_url"]:
                    urls_to_fetch.add(h["after_photo_url"])
            for t in trainings:
                if t["training_photo_url"]:
                    urls_to_fetch.add(t["training_photo_url"])
                if t["absensi_photo_url"]:
                    urls_to_fetch.add(t["absensi_photo_url"])

            if urls_to_fetch:
                logger.info(f"Mengunduh {len(urls_to_fetch)} foto dokumentasi...")
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
                    future_map = {
                        pool.submit(download_image_bytes, u, self.photos_cache_dir, timeout): u
                        for u in urls_to_fetch
                    }
                    for fut in concurrent.futures.as_completed(future_map):
                        u = future_map[fut]
                        try:
                            b = fut.result()
                            if b:
                                url_to_bytes[u] = b
                        except Exception as e:
                            logger.warning(f"Gagal mengunduh foto {u}: {e}")
                logger.info(f"Berhasil mengunduh {len(url_to_bytes)}/{len(urls_to_fetch)} foto.")

        # Siapkan nama file output
        filename = self.config["report"]["filename_pattern"].format(
            start_m=start_date.month,
            start_d=start_date.day,
            end_m=end_date.month,
            end_d=end_date.day,
        )
        output_dir = self.base_dir / self.config["report"].get("output_dir", "output")
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / filename

        template_path = self.base_dir / self.config["report"]["template_file"]
        if not template_path.exists():
            raise FileNotFoundError(f"Template Excel tidak ditemukan: {template_path}")

        self._build_xlsx_from_template(
            template_path=template_path,
            output_path=output_path,
            hazards=hazards,
            trainings=trainings,
            url_to_bytes=url_to_bytes,
        )

        logger.info(f"Laporan berhasil dibuat: {output_path}")
        return {
            "output_path": output_path,
            "filename": filename,
            "start_date": start_date,
            "end_date": end_date,
            "hazard_count": len(hazards),
            "training_count": len(trainings),
        }

    def _build_xlsx_from_template(
        self,
        template_path: Path,
        output_path: Path,
        hazards: List[Dict[str, Any]],
        trainings: List[Dict[str, Any]],
        url_to_bytes: Dict[str, bytes],
    ) -> None:
        """
        Bangun file .xlsx dengan memodifikasi XML dari template asli agar seluruh
        format, rumus, tombol navigasi, tabel organisasi, dan gambar =DISPIMG(...)
        tetap 100% utuh seperti di WPS/Excel.
        """
        ET.register_namespace("", MAIN_NS)
        ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        media_files: Dict[str, bytes] = {}
        cell_images_meta: List[Tuple[str, str, str]] = []  # (img_id, rId, media_target)

        def register_image(raw_url: str) -> Optional[str]:
            if not raw_url or raw_url not in url_to_bytes:
                return None
            idx = len(media_files) + 1
            media_name = f"image{idx}.jpeg"
            media_files[f"xl/media/{media_name}"] = url_to_bytes[raw_url]
            img_id = "ID_" + uuid.uuid4().hex.upper()
            r_id = f"rId{idx}"
            cell_images_meta.append((img_id, r_id, f"media/{media_name}"))
            return img_id

        with zipfile.ZipFile(template_path, "r") as zin:
            # 1. Baca sharedStrings.xml
            ss_xml = ET.fromstring(zin.read("xl/sharedStrings.xml"))
            shared_strings: List[str] = []
            str_to_idx: Dict[str, int] = {}
            for si in ss_xml.findall(f"{{{MAIN_NS}}}si"):
                text = "".join(t.text or "" for t in si.findall(f".//{{{MAIN_NS}}}t"))
                idx = len(shared_strings)
                shared_strings.append(text)
                if text not in str_to_idx:
                    str_to_idx[text] = idx

            def get_ss_idx(val: str) -> int:
                if val in str_to_idx:
                    return str_to_idx[val]
                idx = len(shared_strings)
                shared_strings.append(val)
                str_to_idx[val] = idx
                si_el = ET.SubElement(ss_xml, f"{{{MAIN_NS}}}si")
                t_el = ET.SubElement(si_el, f"{{{MAIN_NS}}}t")
                if val.startswith(" ") or val.endswith(" ") or "\n" in val:
                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t_el.text = val
                return idx

            # Helper untuk membuat elemen <c>
            def make_cell(
                col_letter: str,
                row_num: int,
                style_id: str,
                val: Any = None,
                val_type: str = "s",
                dispimg_id: Optional[str] = None,
            ) -> ET.Element:
                c_el = ET.Element(f"{{{MAIN_NS}}}c")
                c_el.set("r", f"{col_letter}{row_num}")
                c_el.set("s", style_id)
                if dispimg_id:
                    c_el.set("t", "str")
                    f_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}f")
                    f_el.text = f'_xlfn.DISPIMG("{dispimg_id}",1)'
                    v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                    v_el.text = f'=DISPIMG("{dispimg_id}",1)'
                elif val is not None and val != "":
                    if val_type == "s":
                        c_el.set("t", "s")
                        v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                        v_el.text = str(get_ss_idx(str(val)))
                    elif val_type == "n":
                        v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                        v_el.text = str(val)
                return c_el

            # 2. Update sheet2.xml (ERM隐患登记表 (印尼安全组))
            s2_xml = ET.fromstring(zin.read("xl/worksheets/sheet2.xml"))
            s2_data = s2_xml.find(f"{{{MAIN_NS}}}sheetData")
            if s2_data is not None:
                # Pertahankan baris header 1..4
                existing_rows = list(s2_data.findall(f"{{{MAIN_NS}}}row"))
                s2_data.clear()
                for r_el in existing_rows:
                    r_idx = int(r_el.attrib.get("r", "0"))
                    if r_idx <= 4:
                        s2_data.append(r_el)

                # Tambahkan baris data temuan mulai dari baris 5
                for idx, h in enumerate(hazards):
                    r_num = 5 + idx
                    row_el = ET.SubElement(
                        s2_data,
                        f"{{{MAIN_NS}}}row",
                        {
                            "r": str(r_num),
                            "ht": "80",
                            "customHeight": "1",
                            "spans": "1:15",
                        },
                    )
                    img_before_id = register_image(h["hazard_photo_url"])
                    img_after_id = register_image(h["after_photo_url"])

                    row_el.append(make_cell("A", r_num, "230", date_to_excel_serial(h["date"]), "n"))
                    row_el.append(make_cell("B", r_num, "162", h["departemen"], "s"))
                    row_el.append(make_cell("C", r_num, "231", h["bengkel"], "s"))
                    if img_before_id:
                        row_el.append(make_cell("D", r_num, "15", dispimg_id=img_before_id))
                    else:
                        row_el.append(make_cell("D", r_num, "231", h["hazard_photo_url"], "s"))
                    row_el.append(make_cell("E", r_num, "231", h["deskripsi"], "s"))
                    row_el.append(make_cell("F", r_num, "231", h["jenis_bahaya"], "s"))
                    row_el.append(make_cell("G", r_num, "231", h["level_bahaya"], "s"))
                    row_el.append(make_cell("H", r_num, "231", h["pic"], "s"))
                    row_el.append(make_cell("I", r_num, "231", h["safety_officer"], "s"))
                    row_el.append(make_cell("J", r_num, "232", h["rekomendasi"], "s"))
                    row_el.append(make_cell("K", r_num, "230", date_to_excel_serial(h["deadline"]), "n"))
                    if img_after_id:
                        row_el.append(make_cell("L", r_num, "15", dispimg_id=img_after_id))
                    else:
                        row_el.append(make_cell("L", r_num, "231", h["after_photo_url"], "s"))
                    row_el.append(make_cell("M", r_num, "231", h["kondisi_perbaikan"], "s"))
                    row_el.append(make_cell("N", r_num, "231", h["status_perbaikan"], "s"))
                    row_el.append(make_cell("O", r_num, "231", h["keterangan"], "s"))

                # Tambahkan beberapa baris kosong berformat seperti template
                empty_start = 5 + len(hazards)
                for r_num in range(empty_start, max(empty_start + 15, 30)):
                    row_el = ET.SubElement(
                        s2_data,
                        f"{{{MAIN_NS}}}row",
                        {"r": str(r_num), "ht": "80", "customHeight": "1", "spans": "1:15"},
                    )
                    for col_l, st in [
                        ("A", "230"),
                        ("B", "162"),
                        ("C", "231"),
                        ("D", "15"),
                        ("E", "231"),
                        ("F", "231"),
                        ("G", "231"),
                        ("H", "231"),
                        ("I", "231"),
                        ("J", "232"),
                        ("K", "230"),
                        ("L", "15"),
                        ("M", "231"),
                        ("N", "231"),
                        ("O", "231"),
                    ]:
                        row_el.append(make_cell(col_l, r_num, st))

            # 3. Update sheet4.xml (安全培训数据统计表)
            s4_xml = ET.fromstring(zin.read("xl/worksheets/sheet4.xml"))
            s4_data = s4_xml.find(f"{{{MAIN_NS}}}sheetData")
            if s4_data is not None:
                existing_rows4 = list(s4_data.findall(f"{{{MAIN_NS}}}row"))
                s4_data.clear()
                for r_el in existing_rows4:
                    r_idx = int(r_el.attrib.get("r", "0"))
                    if r_idx <= 2:
                        s4_data.append(r_el)

                for idx, t in enumerate(trainings):
                    r_num = 3 + idx
                    row_el = ET.SubElement(
                        s4_data,
                        f"{{{MAIN_NS}}}row",
                        {
                            "r": str(r_num),
                            "s": "116",
                            "customFormat": "1",
                            "ht": "80",
                            "customHeight": "1",
                            "spans": "1:14",
                        },
                    )
                    img_train_id = register_image(t["training_photo_url"])
                    img_abs_id = register_image(t["absensi_photo_url"])

                    row_el.append(make_cell("A", r_num, "133", idx + 1, "n"))
                    row_el.append(make_cell("B", r_num, "126", date_to_excel_serial(t["date"]), "n"))
                    row_el.append(make_cell("C", r_num, "126", t["departemen"], "s"))
                    row_el.append(make_cell("D", r_num, "162", t["penerima"], "s"))
                    row_el.append(make_cell("E", r_num, "126", t["jenis_pelatihan"], "s"))
                    row_el.append(make_cell("F", r_num, "126", t["topik"], "s"))
                    row_el.append(make_cell("G", r_num, "126", t["isi_pelatihan"], "s"))
                    row_el.append(make_cell("H", r_num, "126", t["pemateri"], "s"))
                    row_el.append(make_cell("I", r_num, "126", t["jabatan"], "s"))
                    row_el.append(make_cell("J", r_num, "163", t["tka"], "n"))
                    row_el.append(make_cell("K", r_num, "163", t["tki"], "n"))
                    if img_train_id:
                        row_el.append(make_cell("L", r_num, "126", dispimg_id=img_train_id))
                    else:
                        row_el.append(make_cell("L", r_num, "126", t["training_photo_url"], "s"))
                    if img_abs_id:
                        row_el.append(make_cell("M", r_num, "126", dispimg_id=img_abs_id))
                    else:
                        row_el.append(make_cell("M", r_num, "126", t["absensi_photo_url"], "s"))
                    row_el.append(make_cell("N", r_num, "126", t["pengirim"], "s"))

                # Tambahkan baris kosong bernomor urut (sampai 25 baris)
                empty4_start = 3 + len(trainings)
                for r_num in range(empty4_start, max(empty4_start + 15, 28)):
                    seq_no = r_num - 2
                    row_el = ET.SubElement(
                        s4_data,
                        f"{{{MAIN_NS}}}row",
                        {
                            "r": str(r_num),
                            "s": "116",
                            "customFormat": "1",
                            "ht": "80",
                            "customHeight": "1",
                            "spans": "1:14",
                        },
                    )
                    row_el.append(make_cell("A", r_num, "133", seq_no, "n"))
                    for col_l, st in [
                        ("B", "136"),
                        ("C", "162"),
                        ("D", "133"),
                        ("E", "138"),
                        ("F", "125"),
                        ("G", "125"),
                        ("H", "133"),
                        ("I", "133"),
                        ("J", "164"),
                        ("K", "164"),
                        ("L", "165"),
                        ("M", "165"),
                        ("N", "139"),
                    ]:
                        row_el.append(make_cell(col_l, r_num, st))

            # 4. Pangkas ribuan baris kosong di sheet5.xml & sheet6.xml agar ukuran file ringan
            s5_xml = ET.fromstring(zin.read("xl/worksheets/sheet5.xml"))
            s5_data = s5_xml.find(f"{{{MAIN_NS}}}sheetData")
            if s5_data is not None:
                for r_el in list(s5_data.findall(f"{{{MAIN_NS}}}row")):
                    if int(r_el.attrib.get("r", "0")) > 30:
                        s5_data.remove(r_el)

            s6_xml = ET.fromstring(zin.read("xl/worksheets/sheet6.xml"))
            s6_data = s6_xml.find(f"{{{MAIN_NS}}}sheetData")
            if s6_data is not None:
                for r_el in list(s6_data.findall(f"{{{MAIN_NS}}}row")):
                    if int(r_el.attrib.get("r", "0")) > 30:
                        s6_data.remove(r_el)

            # 5. Bangun xl/cellimages.xml dan xl/_rels/cellimages.xml.rels
            cellimages_xml_str = self._build_cellimages_xml(cell_images_meta)
            cellimages_rels_str = self._build_cellimages_rels_xml(cell_images_meta)

            # Update jumlah sharedStrings
            ss_xml.set("count", str(len(shared_strings)))
            ss_xml.set("uniqueCount", str(len(shared_strings)))

            # 6. Tulis semua ke output_path (.xlsx)
            with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    fname = item.filename
                    if fname.startswith("xl/media/"):
                        continue
                    elif fname == "xl/sharedStrings.xml":
                        zout.writestr(item, ET.tostring(ss_xml, encoding="utf-8", xml_declaration=True))
                    elif fname == "xl/worksheets/sheet2.xml":
                        zout.writestr(item, ET.tostring(s2_xml, encoding="utf-8", xml_declaration=True))
                    elif fname == "xl/worksheets/sheet4.xml":
                        zout.writestr(item, ET.tostring(s4_xml, encoding="utf-8", xml_declaration=True))
                    elif fname == "xl/worksheets/sheet5.xml":
                        zout.writestr(item, ET.tostring(s5_xml, encoding="utf-8", xml_declaration=True))
                    elif fname == "xl/worksheets/sheet6.xml":
                        zout.writestr(item, ET.tostring(s6_xml, encoding="utf-8", xml_declaration=True))
                    elif fname == "xl/cellimages.xml":
                        zout.writestr(item, cellimages_xml_str.encode("utf-8"))
                    elif fname == "xl/_rels/cellimages.xml.rels":
                        zout.writestr(item, cellimages_rels_str.encode("utf-8"))
                    else:
                        zout.writestr(item, zin.read(fname))

                for media_path, img_bytes in media_files.items():
                    zout.writestr(media_path, img_bytes)

    @staticmethod
    def _build_cellimages_xml(cell_images_meta: List[Tuple[str, str, str]]) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">',
        ]
        for idx, (img_id, r_id, _) in enumerate(cell_images_meta, start=1):
            parts.append(
                f"<etc:cellImage><xdr:pic>"
                f'<xdr:nvPicPr><xdr:cNvPr id="{idx}" name="{img_id}"/>'
                f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                f'<xdr:blipFill><a:blip r:embed="{r_id}"/>'
                f"<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
                f'<xdr:spPr><a:xfrm><a:off x="1265555" y="524510"/><a:ext cx="3723005" cy="2395855"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
                f"</xdr:pic></etc:cellImage>"
            )
        parts.append("</etc:cellImages>")
        return "".join(parts)

    @staticmethod
    def _build_cellimages_rels_xml(cell_images_meta: List[Tuple[str, str, str]]) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ]
        for _, r_id, target in cell_images_meta:
            parts.append(
                f'<Relationship Id="{r_id}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="{target}"/>'
            )
        parts.append("</Relationships>")
        return "".join(parts)
