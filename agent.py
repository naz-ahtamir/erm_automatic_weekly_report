#!/usr/bin/env python3
"""
Agen Otomatis Laporan Mingguan ERM ESG (Kamis Pagi Jam 08:00).
- Berjalan terus (daemon/standby) dan otomatis men-generate serta mengirim laporan
  setiap hari Kamis pukul 08:00 pagi lewat WeChat.
- Rentang data laporan: Hari Kamis minggu lalu s/d Hari Rabu minggu ini.
- Sumber data:
  1. Non-Conformity / Hazard (khusus departemen mengandung 'ESG' & kategori 'Bahaya Tersembunyi')
  2. Safety Talk / Pelatihan (khusus departemen mengandung 'ESG')
"""

import argparse
import datetime as dt
import json
import logging
import sys
import time
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from report_generator import ERMReportGenerator, get_thursday_to_wednesday_range, parse_flexible_date
from wechat_sender import WeChatSender

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
LOGS_DIR = BASE_DIR / "logs"
STATE_FILE = LOGS_DIR / "agent_state.json"

DAY_NAME_TO_WEEKDAY = {
    "monday": 0,
    "senin": 0,
    "tuesday": 1,
    "selasa": 1,
    "wednesday": 2,
    "rabu": 2,
    "thursday": 3,
    "kamis": 3,
    "friday": 4,
    "jumat": 4,
    "saturday": 5,
    "sabtu": 5,
    "sunday": 6,
    "minggu": 6,
}


def setup_logging() -> logging.Logger:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_DIR / "agent.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )
    return logging.getLogger("erm_report.agent")


def load_config() -> Dict[str, Any]:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"File konfigurasi tidak ditemukan: {CONFIG_PATH}")
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def load_state() -> Dict[str, Any]:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_state(state: Dict[str, Any]) -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def get_next_scheduled_run(now: dt.datetime, target_weekday: int, hour: int, minute: int) -> dt.datetime:
    """Hitung waktu eksekusi terjadwal berikutnya (misal Kamis pukul 08:00:00)."""
    days_ahead = (target_weekday - now.weekday()) % 7
    candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0) + dt.timedelta(days=days_ahead)
    if candidate <= now:
        candidate += dt.timedelta(days=7)
    return candidate


def execute_job(
    config: Dict[str, Any],
    start_date: Optional[dt.date] = None,
    end_date: Optional[dt.date] = None,
    send_wechat: bool = True,
    use_cached_csv: bool = False,
    logger: Optional[logging.Logger] = None,
) -> Dict[str, Any]:
    if logger is None:
        logger = logging.getLogger("erm_report.agent")

    if start_date is None or end_date is None:
        start_date, end_date = get_thursday_to_wednesday_range(dt.date.today())

    logger.info("=" * 65)
    logger.info(f"MEMULAI PROSES LAPORAN MINGGUAN ERM ESG")
    logger.info(f"Periode Data : Kamis ({start_date}) s/d Rabu ({end_date})")
    logger.info("=" * 65)

    generator = ERMReportGenerator(BASE_DIR, config)
    report_info = generator.generate_report(
        start_date=start_date,
        end_date=end_date,
        use_cached_csv=use_cached_csv,
    )

    logger.info(
        f"Hasil Generate: {report_info['filename']} | "
        f"Temuan Hazard ESG: {report_info['hazard_count']} | "
        f"Pelatihan ESG: {report_info['training_count']}"
    )

    if send_wechat:
        sender = WeChatSender(config)
        ok = sender.send_report(report_info)
        report_info["wechat_sent"] = ok
    else:
        logger.info("Mode --generate-only aktif: Melewati pengiriman WeChat.")
        report_info["wechat_sent"] = False

    return report_info


def run_daemon(config: Dict[str, Any], logger: logging.Logger) -> None:
    sched = config.get("schedule", {})
    day_str = str(sched.get("day_of_week", "thursday")).lower()
    target_weekday = DAY_NAME_TO_WEEKDAY.get(day_str, 3)  # Default: 3 = Kamis
    target_hour = int(sched.get("hour", 8))
    target_minute = int(sched.get("minute", 0))

    now = dt.datetime.now()
    next_run = get_next_scheduled_run(now, target_weekday, target_hour, target_minute)
    next_start, next_end = get_thursday_to_wednesday_range(next_run.date())

    logger.info("=" * 65)
    logger.info("🤖 AGEN OTOMATIS LAPORAN ERM ESG AKTIF & STANDBY")
    logger.info(f"📅 Jadwal Rutin     : Setiap Hari Kamis pukul {target_hour:02d}:{target_minute:02d}")
    logger.info(f"⏰ Eksekusi Berikut : {next_run.strftime('%Y-%m-%d %H:%M:%S')} (Kamis)")
    logger.info(f"📊 Periode Laporan  : {next_start} (Kamis lalu) s/d {next_end} (Rabu ini)")
    logger.info(f"💬 Tujuan WeChat    : {config.get('wechat', {}).get('target_chat', '文件传输助手')}")
    logger.info("   (Tekan Ctrl+C untuk menghentikan agen)")
    logger.info("=" * 65)

    while True:
        try:
            # Reload config setiap siklus agar perubahan config.json langsung terbaca tanpa restart
            config = load_config()
            sched = config.get("schedule", {})
            day_str = str(sched.get("day_of_week", "thursday")).lower()
            target_weekday = DAY_NAME_TO_WEEKDAY.get(day_str, 3)
            target_hour = int(sched.get("hour", 8))
            target_minute = int(sched.get("minute", 0))

            now = dt.datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            state = load_state()
            last_run_date = state.get("last_run_date", "")

            # Cek apakah hari ini adalah hari Kamis dan sudah memasuki jam >= 08:00 serta belum dijalankan hari ini
            target_time_today = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
            if now.weekday() == target_weekday and now >= target_time_today and last_run_date != today_str:
                logger.info(f"Waktu jadwal tercapai ({now.strftime('%Y-%m-%d %H:%M:%S')})! Menjalankan tugas mingguan...")
                start_d, end_d = get_thursday_to_wednesday_range(now.date())
                info = execute_job(config, start_date=start_d, end_date=end_d, send_wechat=True, logger=logger)
                state["last_run_date"] = today_str
                state["last_output_file"] = str(info["output_path"])
                state["last_wechat_sent"] = bool(info.get("wechat_sent", False))
                save_state(state)

                next_run = get_next_scheduled_run(dt.datetime.now(), target_weekday, target_hour, target_minute)
                logger.info(f"Tugas selesai. Jadwal berikutnya: {next_run.strftime('%Y-%m-%d %H:%M:%S')}")

            time.sleep(20)
        except KeyboardInterrupt:
            logger.info("Agen dihentikan oleh pengguna (Ctrl+C).")
            break
        except Exception as e:
            logger.error(f"Error pada loop agen: {e}", exc_info=True)
            time.sleep(60)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Agen Otomatis Laporan Mingguan ERM ESG (Kamis Pagi Jam 08:00 via WeChat)"
    )
    parser.add_argument(
        "--run-now",
        action="store_true",
        help="Langsung generate laporan dan kirim ke WeChat sekarang juga (tanpa menunggu hari Kamis jam 8)",
    )
    parser.add_argument(
        "--generate-only",
        action="store_true",
        help="Langsung generate file Excel (.xlsx) sekarang tanpa mengirimkannya ke WeChat",
    )
    parser.add_argument(
        "--start-date",
        type=str,
        default=None,
        help="Tanggal awal laporan (format YYYY-MM-DD, default: Kamis minggu lalu)",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        default=None,
        help="Tanggal akhir laporan (format YYYY-MM-DD, default: Rabu minggu ini)",
    )
    parser.add_argument(
        "--target-chat",
        type=str,
        default=None,
        help="Override nama kontak/grup WeChat tujuan (misal: 'ERM安全组' atau '文件传输助手')",
    )
    parser.add_argument(
        "--use-cache",
        action="store_true",
        help="Gunakan cache CSV lokal jika sudah terunduh sebelumnya",
    )

    args = parser.parse_args()
    logger = setup_logging()
    config = load_config()

    if args.target_chat:
        config.setdefault("wechat", {})["target_chat"] = args.target_chat

    start_d = parse_flexible_date(args.start_date) if args.start_date else None
    end_d = parse_flexible_date(args.end_date) if args.end_date else None

    if args.run_now or args.generate_only:
        execute_job(
            config=config,
            start_date=start_d,
            end_date=end_d,
            send_wechat=not args.generate_only,
            use_cached_csv=args.use_cache,
            logger=logger,
        )
    else:
        run_daemon(config, logger)


if __name__ == "__main__":
    main()
