#!/usr/bin/env python3
"""
Modul Pengirim Laporan via WeChat:
1. Mode 'desktop' (Default): Otomatisasi UI aplikasi WeChat Desktop Linux (/usr/bin/wechat)
   menggunakan wmctrl, xdotool, dan xclip untuk mencari kontak/grup tujuan,
   menempelkan file laporan Excel (.xlsx), dan mengirimkan ringkasan teks.
2. Mode 'webhook' (Opsional): Mengirimkan file .xlsx dan ringkasan ke Grup WeCom / 企业微信
   melalui Webhook URL jika dikonfigurasi.
"""

import json
import logging
import os
import subprocess
import time
from pathlib import Path
from typing import Any, Dict, Optional

logger = logging.getLogger("erm_report.wechat")


class WeChatSender:
    def __init__(self, config: Dict[str, Any]):
        self.wechat_cfg = config.get("wechat", {})
        self.enabled = self.wechat_cfg.get("enabled", True)
        self.mode = self.wechat_cfg.get("mode", "desktop").lower()
        self.target_chat = self.wechat_cfg.get("target_chat", "文件传输助手")
        self.send_summary_text = self.wechat_cfg.get("send_summary_text", True)
        self.summary_template = self.wechat_cfg.get(
            "summary_template",
            "【ERM 印尼安全组周报数据】\nPeriode: {start_date} s/d {end_date}\nTemuan Hazard ESG: {hazard_count}\nSafety Talk ESG: {training_count}",
        )
        self.webhook_url = self.wechat_cfg.get("webhook_url", "").strip()

    def format_summary(self, report_info: Dict[str, Any]) -> str:
        start_date = report_info["start_date"]
        end_date = report_info["end_date"]
        return self.summary_template.format(
            start_date=start_date.strftime("%Y-%m-%d"),
            end_date=end_date.strftime("%Y-%m-%d"),
            start_m=start_date.month,
            start_d=start_date.day,
            end_m=end_date.month,
            end_d=end_date.day,
            hazard_count=report_info.get("hazard_count", 0),
            training_count=report_info.get("training_count", 0),
            filename=report_info.get("filename", ""),
        )

    def send_report(self, report_info: Dict[str, Any]) -> bool:
        if not self.enabled:
            logger.info("Pengiriman WeChat dinonaktifkan di config.json.")
            return True

        file_path = Path(report_info["output_path"]).resolve()
        if not file_path.exists():
            logger.error(f"File laporan tidak ditemukan untuk dikirim: {file_path}")
            return False

        summary_text = self.format_summary(report_info)

        if self.mode == "webhook" and self.webhook_url:
            return self._send_via_webhook(file_path, summary_text)
        return self._send_via_linux_desktop(file_path, summary_text)

    @staticmethod
    def _find_wechat_window_id() -> Optional[str]:
        """Cari Window ID aplikasi WeChat di X11 menggunakan wmctrl atau xdotool."""
        try:
            res = subprocess.run(["wmctrl", "-l", "-x"], capture_output=True, text=True, timeout=5)
            if res.returncode == 0:
                for line in res.stdout.splitlines():
                    low = line.lower()
                    if "wechat" in low or "weixin" in low or "微信" in line:
                        parts = line.split()
                        if parts:
                            return parts[0]
        except Exception as e:
            logger.debug(f"wmctrl gagal mencari window WeChat: {e}")

        for pattern in ("WeChat", "微信", "wechat", "weixin"):
            try:
                res = subprocess.run(
                    ["xdotool", "search", "--onlyvisible", "--name", pattern],
                    capture_output=True,
                    text=True,
                    timeout=5,
                )
                if res.returncode == 0 and res.stdout.strip():
                    wids = res.stdout.strip().splitlines()
                    return wids[-1]
            except Exception:
                continue
        return None

    @staticmethod
    def _set_clipboard_text(text: str) -> None:
        """Salin teks UTF-8 ke clipboard X11 menggunakan xclip."""
        proc = subprocess.Popen(
            ["xclip", "-selection", "clipboard", "-i"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        proc.communicate(input=text.encode("utf-8"), timeout=5)

    @staticmethod
    def _set_clipboard_file(file_path: Path) -> None:
        """
        Salin file ke clipboard X11 dengan MIME type text/uri-list
        agar Ctrl+V di aplikasi WeChat Desktop Linux menempelkan file lampiran.
        """
        file_uri = file_path.resolve().as_uri() + "\r\n"
        proc = subprocess.Popen(
            ["xclip", "-selection", "clipboard", "-t", "text/uri-list", "-i"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        proc.communicate(input=file_uri.encode("utf-8"), timeout=5)

    def _send_via_linux_desktop(self, file_path: Path, summary_text: str) -> bool:
        """
        Otomatisasi pengiriman file .xlsx + ringkasan teks lewat aplikasi WeChat Desktop Linux.
        """
        logger.info(f"Menyiapkan pengiriman WeChat Desktop ke: '{self.target_chat}'")
        env = os.environ.copy()
        if "DISPLAY" not in env:
            env["DISPLAY"] = ":0.0"

        wid = self._find_wechat_window_id()
        if not wid:
            logger.info("Jendela WeChat belum terbuka, menjalankan /usr/bin/wechat...")
            try:
                subprocess.Popen(
                    ["/usr/bin/wechat"],
                    env=env,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except Exception as e:
                logger.error(f"Gagal menjalankan /usr/bin/wechat: {e}")
                return False

            for _ in range(15):
                time.sleep(1.0)
                wid = self._find_wechat_window_id()
                if wid:
                    break

        if not wid:
            logger.error(
                "Tidak dapat menemukan jendela WeChat yang aktif. Pastikan Anda sudah login ke aplikasi WeChat Desktop."
            )
            return False

        try:
            # 1. Aktifkan jendela WeChat ke paling depan
            subprocess.run(["wmctrl", "-i", "-a", wid], env=env, timeout=5)
            time.sleep(0.4)
            subprocess.run(["xdotool", "windowactivate", "--sync", wid], env=env, timeout=5)
            time.sleep(0.5)

            # 2. Tekan Ctrl+F untuk fokus ke kolom pencarian WeChat, lalu cari target_chat
            self._set_clipboard_text(self.target_chat)
            time.sleep(0.2)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+f"], env=env, timeout=5)
            time.sleep(0.5)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+a", "BackSpace"], env=env, timeout=5)
            time.sleep(0.2)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+v"], env=env, timeout=5)
            time.sleep(1.2)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "Return"], env=env, timeout=5)
            time.sleep(0.8)

            # 3. Salin file laporan .xlsx ke clipboard dan tempelkan (Ctrl+V) lalu kirim (Return)
            logger.info(f"Menempelkan file laporan ke WeChat: {file_path.name}")
            self._set_clipboard_file(file_path)
            time.sleep(0.3)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+v"], env=env, timeout=5)
            time.sleep(0.8)
            subprocess.run(["xdotool", "key", "--clearmodifiers", "Return"], env=env, timeout=5)
            time.sleep(1.0)

            # 4. Kirimkan pesan ringkasan teks jika diaktifkan
            if self.send_summary_text and summary_text:
                logger.info("Mengirimkan ringkasan laporan teks ke WeChat...")
                self._set_clipboard_text(summary_text)
                time.sleep(0.3)
                subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+v"], env=env, timeout=5)
                time.sleep(0.5)
                subprocess.run(["xdotool", "key", "--clearmodifiers", "Return"], env=env, timeout=5)
                time.sleep(0.5)

            logger.info(f"Laporan berhasil dikirim ke WeChat ('{self.target_chat}')!")
            return True

        except Exception as e:
            logger.error(f"Terjadi kesalahan saat mengotomatisasi WeChat Desktop: {e}")
            return False

    def _send_via_webhook(self, file_path: Path, summary_text: str) -> bool:
        """Kirim file dan teks melalui WeCom (Enterprise WeChat) Webhook URL."""
        key = ""
        if "key=" in self.webhook_url:
            key = self.webhook_url.split("key=")[-1].split("&")[0]
        if not key:
            logger.error("Webhook URL tidak memiliki parameter key= yang valid.")
            return False

        upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"
        try:
            res = subprocess.run(
                ["curl", "-sS", "-F", f"media=@{file_path}", upload_url],
                capture_output=True,
                text=True,
                timeout=60,
            )
            data = json.loads(res.stdout)
            media_id = data.get("media_id")
            if not media_id:
                logger.error(f"Gagal mengunggah file ke WeCom Webhook: {data}")
                return False

            send_file_payload = json.dumps({"msgtype": "file", "file": {"media_id": media_id}})
            subprocess.run(
                [
                    "curl",
                    "-sS",
                    "-H",
                    "Content-Type: application/json",
                    "-d",
                    send_file_payload,
                    self.webhook_url,
                ],
                capture_output=True,
                timeout=30,
            )

            if self.send_summary_text and summary_text:
                send_txt_payload = json.dumps({"msgtype": "text", "text": {"content": summary_text}})
                subprocess.run(
                    [
                        "curl",
                        "-sS",
                        "-H",
                        "Content-Type: application/json",
                        "-d",
                        send_txt_payload,
                        self.webhook_url,
                    ],
                    capture_output=True,
                    timeout=30,
                )
            logger.info("Laporan berhasil dikirim via WeCom Webhook!")
            return True
        except Exception as e:
            logger.error(f"Gagal mengirim via Webhook: {e}")
            return False
