#!/bin/bash
# Script untuk deploy bot ke VPS menggunakan systemd

# Keluar jika ada error
set -e

echo "🚀 Memulai setup VPS untuk UltraJurnal Bot..."

# 1. Update dan install system dependencies (sesuai dengan Dockerfile)
echo "📦 Menginstal dependensi sistem..."
sudo apt-get update
sudo apt-get install -y ghostscript libmagic1 libmagic-dev libgomp1 curl build-essential

# 2. Install uv (package manager Python) jika belum ada
if ! command -v uv &> /dev/null; then
    echo "📦 Menginstal uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    export PATH="$HOME/.cargo/bin:$PATH"
fi

# 3. Setup Virtual Environment dan Install Python dependencies
echo "🐍 Menyiapkan virtual environment Python..."
uv venv .venv
uv sync --frozen --no-dev

# 4. Buat folder Hasil jika belum ada
mkdir -p Hasil

# 5. Dapatkan path absolut (working directory saat ini) dan username
APP_DIR=$(pwd)
CURRENT_USER=$(whoami)

# 6. Buat file systemd service
SERVICE_FILE="/tmp/ultrajurnal-bot.service"
echo "⚙️ Membuat file systemd service..."
cat <<EOF > $SERVICE_FILE
[Unit]
Description=UltraJurnal Discord PDF Converter Bot
After=network.target

[Service]
User=$CURRENT_USER
WorkingDirectory=$APP_DIR
ExecStart=$APP_DIR/.venv/bin/python $APP_DIR/discord_bot.py
Restart=always
RestartSec=5
EnvironmentFile=$APP_DIR/.env
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF

# 7. Pindahkan ke systemd direktori dan aktifkan
echo "🔄 Mendaftarkan service ke systemd..."
sudo mv $SERVICE_FILE /etc/systemd/system/ultrajurnal-bot.service
sudo systemctl daemon-reload
sudo systemctl enable ultrajurnal-bot.service
sudo systemctl restart ultrajurnal-bot.service

echo "✅ Deployment selesai! Bot sekarang berjalan sebagai systemd service."
echo "📜 Untuk melihat log secara real-time, gunakan perintah:"
echo "   sudo journalctl -u ultrajurnal-bot.service -f"
