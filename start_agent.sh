#!/usr/bin/env bash
set -e

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

chmod +x agent.py report_generator.py wechat_sender.py

echo "================================================================="
echo "  Menjalankan Agen Otomatis Laporan Mingguan ERM ESG"
echo "  Folder Kerja : $DIR"
echo "================================================================="

exec python3 "$DIR/agent.py" "$@"
