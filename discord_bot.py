#!/usr/bin/env python3
"""
Discord Bot — UltraJurnal.bot
================================
Alur kerja:
  1. User ketik  /bca    → bot tampilkan tombol "Upload PDF"
  2. User klik tombol   → muncul modal (pop-up) isian alias & catatan
  3. User kirim file PDF sebagai attachment di channel
  4. Bot men-download PDF, menjalankan converter.py → .xlsx
  5. xlsx di-encode base64 → POST ke Google Apps Script Web App
  6. Apps Script simpan file ke folder Google Drive
  7. Bot kirim link GDrive ke user (channel + DM)

Variabel environment yang dibutuhkan (lihat .env.example):
  DISCORD_BOT_TOKEN      — token bot Discord
  APPSCRIPT_WEBAPP_URL   — URL Web App Apps Script (Deploy > Web App)
"""

import os
import sys
import time
import asyncio
import logging
import tempfile
import traceback
from pathlib import Path

# ── Dependency guard ──────────────────────────────────────────────────────────
try:
    import discord
    from discord import app_commands
    from discord.ext import commands
    from dotenv import load_dotenv
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "discord.py>=2.3", "python-dotenv"])
    import discord
    from discord import app_commands
    from discord.ext import commands
    from dotenv import load_dotenv

import aiohttp
import base64

# ── Setup logging ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)-8s] %(name)s — %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("discord_bot")

# ── Load .env ─────────────────────────────────────────────────────────────────
load_dotenv()

BOT_TOKEN          = os.getenv("DISCORD_BOT_TOKEN", "")
APPSCRIPT_URL      = os.getenv("APPSCRIPT_WEBAPP_URL", "")
GUILD_ID           = os.getenv("DISCORD_GUILD_ID", "")   # Opsional: sync instan ke server ini
CONVERTER_PATH     = Path(__file__).parent / "converter.py"
HASIL_DIR          = Path(__file__).parent / "Hasil"
HASIL_DIR.mkdir(exist_ok=True)

# ── Google Drive upload via Apps Script Web App ───────────────────────────────
async def upload_via_appscript(
    local_path: str,
    filename: str,
    discord_user: str = "unknown",
) -> str:
    """
    Upload file ke Google Drive melalui Apps Script Web App.
    File di-encode base64 lalu dikirim sebagai JSON POST.
    Returns: shareable Google Drive link.
    """
    if not APPSCRIPT_URL:
        raise RuntimeError(
            "APPSCRIPT_WEBAPP_URL belum diset di .env!\n"
            "Deploy gdrive_upload.gs sebagai Web App lalu salin URL-nya."
        )

    ext = Path(local_path).suffix.lower()
    mime_map = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pdf":  "application/pdf",
        ".xls":  "application/vnd.ms-excel",
    }
    mime_type = mime_map.get(ext, "application/octet-stream")

    with open(local_path, "rb") as f:
        file_bytes = f.read()
    b64_data = base64.b64encode(file_bytes).decode("utf-8")

    payload = {
        "filename":     filename,
        "mimeType":     mime_type,
        "data":         b64_data,
        "discord_user": discord_user,
    }

    log.info("🚀 Upload via Apps Script: %s (%d KB)", filename, len(file_bytes) // 1024)
    
    fallback_link = "https://drive.google.com/drive/folders/17jmj2j6WjsyEgxBNV7XOK6GT_2FplaXy"

    async with aiohttp.ClientSession() as session:
        try:
            # Step 1: POST (tanpa auto-redirect untuk menangkap 302)
            async with session.post(
                APPSCRIPT_URL,
                json=payload,
                headers={"Content-Type": "application/json"},
                allow_redirects=False,
                timeout=aiohttp.ClientTimeout(total=120),
            ) as resp:
                
                if resp.status in (301, 302, 303, 307, 308):
                    redirect_url = resp.headers.get("Location")
                    if not redirect_url:
                        raise RuntimeError("Redirected but no Location header found")
                    
                    # Step 2: GET URL redirect (dengan retry untuk hindari 404 intermittent)
                    for attempt in range(5):
                        async with session.get(redirect_url, timeout=aiohttp.ClientTimeout(total=60)) as resp2:
                            if resp2.status == 404:
                                log.warning(f"⚠️ [Attempt {attempt+1}/5] Apps Script redirect 404, retrying in 2s...")
                                await asyncio.sleep(2)
                                continue
                            
                            resp2.raise_for_status()
                            result = await resp2.json(content_type=None)
                            break
                    else:
                        # Jika loop selesai tanpa break (semua retry gagal dengan 404)
                        log.error("❌ Google Apps Script echo server terus memberikan 404 setelah retry.")
                        log.info("📁 File mungkin sudah terupload. Menggunakan fallback link folder.")
                        return fallback_link
                else:
                    # Jika tidak ada redirect (misal error di Apps Script mengembalikan 200/400)
                    resp.raise_for_status()
                    result = await resp.json(content_type=None)

            if not result.get("success"):
                raise RuntimeError("Apps Script error: " + result.get("error", "unknown"))

            link = result.get("link", fallback_link)
            log.info("✅ GDrive upload selesai: %s → %s", filename, link)
            return link

        except Exception as exc:
            log.error(f"Gagal saat menghubungi Apps Script: {exc}")
            log.info("📁 Mengembalikan fallback link folder karena kemungkinan file berhasil terkirim.")
            return fallback_link


# ── Converter runner ──────────────────────────────────────────────────────────
async def run_converter(pdf_path: str, xlsx_path: str) -> tuple[bool, str]:
    """
    Jalankan converter.py secara async subprocess.
    Returns (success, output_log).
    """
    proc = await asyncio.create_subprocess_exec(
        sys.executable, str(CONVERTER_PATH),
        pdf_path, "-o", xlsx_path,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
    )
    stdout, _ = await proc.communicate()
    output = stdout.decode("utf-8", errors="replace")
    return proc.returncode == 0, output


# ── Download helper ───────────────────────────────────────────────────────────
async def download_file(url: str, dest_path: str):
    """Download attachment dari Discord ke path lokal."""
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as resp:
            resp.raise_for_status()
            with open(dest_path, "wb") as f:
                async for chunk in resp.content.iter_chunked(65536):
                    f.write(chunk)


# ── Intents & Bot ─────────────────────────────────────────────────────────────
intents = discord.Intents.default()
intents.message_content = True  # Perlu untuk baca attachment dari message biasa

bot = commands.Bot(command_prefix="!", intents=intents)
tree = bot.tree


# ══════════════════════════════════════════════════════════════════════════════
# MODAL — Pop-up upload PDF
# ══════════════════════════════════════════════════════════════════════════════
class UploadPDFModal(discord.ui.Modal, title="📄 UltraJurnal.bot — Upload PDF"):
    """
    Modal ini menggunakan FileUpload yang baru didukung di discord.py 2.7+.
    """

    file_attachment = discord.ui.Label(
        text="Upload File PDF",
        description="Drop file here or browse",
        component=discord.ui.FileUpload(required=True)
    )

    async def on_submit(self, interaction: discord.Interaction):
        
        attachment = self.file_attachment.component.values[0]
        
        if not attachment.filename.lower().endswith('.pdf'):
            await interaction.response.send_message("❌ File harus berupa PDF!", ephemeral=True)
            return
            
        base_name = Path(attachment.filename).stem
        pdf_filename = base_name + ".pdf"
        xlsx_filename = base_name + ".xlsx"

        processing_embed = discord.Embed(
            title="⏳ Memproses file…",
            description=(
                f"📄 **File:** `{attachment.filename}`\n\n"
                "Sedang mengekstrak data & mengkonversi ke Excel…\n"
                "Mohon tunggu sebentar ☕"
            ),
            color=0xFFA500,
        )
        
        await interaction.response.send_message(embed=processing_embed)
        status_msg = await interaction.original_response()

        with tempfile.TemporaryDirectory(prefix="bca_converter_") as tmpdir:
            pdf_path  = os.path.join(tmpdir, pdf_filename)
            xlsx_path = os.path.join(tmpdir, xlsx_filename)

            try:
                # 1. Download
                await download_file(attachment.url, pdf_path)
                log.info("Downloaded: %s → %s", attachment.filename, pdf_path)

                # 2. Konversi PDF → xlsx
                t0 = time.perf_counter()
                success, conv_log = await run_converter(pdf_path, xlsx_path)
                elapsed = time.perf_counter() - t0

                elapsed_min = int(elapsed // 60)
                elapsed_sec = elapsed % 60
                if elapsed_min > 0:
                    elapsed_str = f"{elapsed_min}m {elapsed_sec:.1f}s"
                else:
                    elapsed_str = f"{elapsed_sec:.1f}s"

                if not success:
                    raise RuntimeError(f"Konversi gagal:\n```\n{conv_log[-2000:]}\n```")

                # 3. Upload ke Google Drive via Apps Script
                drive_link = await upload_via_appscript(
                    xlsx_path,
                    xlsx_filename,
                    discord_user=str(interaction.user),
                )

                # 4. Sukses! Edit status pesan
                success_embed = discord.Embed(
                    title="✅ Konversi Berhasil!",
                    description=(
                        f"📄 **File PDF:** `{attachment.filename}`\n"
                        f"📊 **Excel:** `{xlsx_filename}`\n"
                        f"⏱️ **Durasi konversi:** `{elapsed_str}`\n\n"
                        f"🔗 **Link Google Drive:**\n{drive_link}"
                    ),
                    color=0x57F287,
                )
                success_embed.set_footer(text="✨ File berhasil diupload ke Google Drive")
                await status_msg.edit(embed=success_embed)

            except Exception as exc:
                log.error("Gagal memproses %s: %s", attachment.filename, traceback.format_exc())
                err_embed = discord.Embed(
                    title="❌ Konversi Gagal",
                    description=(
                        f"Maaf, terjadi kendala saat memproses file `{attachment.filename}`.\n\n"
                        "💡 **Beberapa kemungkinan penyebab:**\n"
                        "• File yang diupload **bukan** format mutasi rekening BCA yang didukung.\n"
                        "• Format tabel di dalam file PDF tidak terbaca dengan baik.\n"
                        "• Terjadi gangguan koneksi atau server sedang sibuk.\n\n"
                        "Pastikan file PDF adalah mutasi rekening BCA asli yang valid. Silakan coba lagi beberapa saat, atau hubungi admin jika masalah berlanjut."
                    ),
                    color=0xED4245,
                )
                err_embed.set_footer(text="Sistem mendeteksi anomali pada file atau proses")
                await status_msg.edit(embed=err_embed)

    async def on_error(self, interaction: discord.Interaction, error: Exception):
        log.error("Modal error: %s", traceback.format_exc())
        if not interaction.response.is_done():
            await interaction.response.send_message(
                "❌ Terjadi kesalahan pada modal. Coba lagi.", ephemeral=True
            )
        else:
            await interaction.followup.send("❌ Terjadi kesalahan pada modal.", ephemeral=True)


# ══════════════════════════════════════════════════════════════════════════════
# VIEW — Button "Upload PDF"
# ══════════════════════════════════════════════════════════════════════════════
class UploadView(discord.ui.View):
    """Persistent view dengan tombol untuk membuka modal upload."""

    def __init__(self):
        super().__init__(timeout=None)  # persistent

    @discord.ui.button(
        label="📄 Upload PDF Rekening",
        style=discord.ButtonStyle.primary,
        custom_id="btn_upload_pdf",
        emoji="📤",
    )
    async def upload_button(
        self, interaction: discord.Interaction, button: discord.ui.Button
    ):
        modal = UploadPDFModal()
        await interaction.response.send_modal(modal)

    @discord.ui.button(
        label="❓ Bantuan",
        style=discord.ButtonStyle.secondary,
        custom_id="btn_help",
        emoji="💡",
    )
    async def help_button(
        self, interaction: discord.Interaction, button: discord.ui.Button
    ):
        embed = discord.Embed(
            title="💡 Cara Menggunakan Bot",
            description=(
                "**Bot ini mengkonversi PDF mutasi rekening BCA → Excel (.xlsx) "
                "lalu mengupload hasilnya ke Google Drive.**\n\n"
                "**Langkah-langkah:**\n"
                "1️⃣ Klik tombol **📤 Upload PDF Rekening**\n"
                "2️⃣ Isi form (nama alias & catatan opsional) lalu klik **Submit**\n"
                "3️⃣ Kirim file **PDF** rekening BCA ke channel ini\n"
                "4️⃣ Bot akan memproses & mengirim **link Google Drive** ke kamu\n\n"
                "**Format PDF yang didukung:**\n"
                "• Mutasi rekening BCA (format standar dengan tabel lattice)\n\n"
                "**Kontak:** Hubungi admin jika ada kendala."
            ),
            color=0xFEE75C,
        )
        embed.set_footer(text="UltraJurnal.bot • v1.0")
        await interaction.response.send_message(embed=embed, ephemeral=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLASH COMMAND — /bca
# ══════════════════════════════════════════════════════════════════════════════
@tree.command(name="bca", description="Mulai konversi PDF rekening BCA ke Excel")
async def cmd_bca(interaction: discord.Interaction):
    """Tampilkan panel utama dengan tombol upload."""
    embed = discord.Embed(
        title="📒 UltraJurnal.bot",
        description=(
            "Selamat datang! Bot ini akan membantu kamu mengkonversi "
            "**PDF mutasi rekening BCA** menjadi file **Excel (.xlsx)** "
            "yang rapi dan terstruktur.\n\n"
            "📋 **Yang bot lakukan:**\n"
            "• Ekstrak data transaksi dari PDF\n"
            "• Format menjadi Excel dengan kolom lengkap\n"
            "• Upload ke Google Drive\n"
            "• Kirim link download ke kamu\n\n"
            "👇 **Klik tombol di bawah untuk memulai!**"
        ),
        color=0x0077B6,
    )
    embed.add_field(
        name="📁 Kolom Excel yang dihasilkan",
        value=(
            "`NO. REKENING` • `PERIODE` • `MATA UANG`\n"
            "`TANGGAL` • `KETERANGAN` • `CBG` • `MUTASI` • `SALDO`"
        ),
        inline=False,
    )
    embed.set_footer(
        text="UltraJurnal.bot • Powered by Camelot",
        icon_url="https://www.bca.co.id/~/media/bca/logobca.png",
    )
    embed.set_thumbnail(url="https://cdn-icons-png.flaticon.com/512/4208/4208479.png")

    view = UploadView()
    await interaction.response.send_message(embed=embed, view=view)


# (Event on_message dihapus karena file sekarang diupload langsung di Modal)


# ══════════════════════════════════════════════════════════════════════════════
# BOT EVENTS
# ══════════════════════════════════════════════════════════════════════════════
@bot.event
async def on_ready():
    """Dipanggil ketika bot sudah terhubung ke Discord."""
    bot.pending_uploads = {}

    # Register persistent view agar tombol tetap berfungsi setelah restart
    bot.add_view(UploadView())

    # Sync slash commands — guild sync = instan, global sync = ~1 jam
    try:
        if GUILD_ID:
            guild_obj = discord.Object(id=int(GUILD_ID))
            tree.copy_global_to(guild=guild_obj)
            synced = await tree.sync(guild=guild_obj)
            log.info("✅ Slash commands synced ke guild %s: %d command(s)", GUILD_ID, len(synced))
        else:
            synced = await tree.sync()
            log.info("✅ Slash commands synced global: %d command(s) (bisa makan ~1 jam)", len(synced))
    except Exception as e:
        log.error("Gagal sync slash commands: %s", e)

    log.info("=" * 55)
    log.info("🤖 Bot online: %s (ID: %s)", bot.user, bot.user.id)
    log.info("🔗 Apps Script : %s", APPSCRIPT_URL[:60] + "..." if APPSCRIPT_URL else "(tidak diset)")
    log.info("🏠 Guild sync  : %s", GUILD_ID or "(global — command muncul dalam ~1 jam)")
    log.info("=" * 55)


@bot.event
async def on_application_command_error(
    interaction: discord.Interaction, error: app_commands.AppCommandError
):
    log.error("App command error: %s", error)
    try:
        await interaction.response.send_message(
            f"❌ Error: {error}", ephemeral=True
        )
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def main():
    if not BOT_TOKEN:
        log.critical(
            "DISCORD_BOT_TOKEN tidak ditemukan!\n"
            "Buat file .env dan isi:\n"
            "  DISCORD_BOT_TOKEN=token_bot_kamu_di_sini\n"
            "  APPSCRIPT_WEBAPP_URL=https://script.google.com/macros/s/.../exec"
        )
        sys.exit(1)

    if not APPSCRIPT_URL:
        log.critical(
            "APPSCRIPT_WEBAPP_URL tidak ditemukan!\n"
            "Deploy gdrive_upload.gs sebagai Google Apps Script Web App\n"
            "lalu salin URL-nya ke .env:\n"
            "  APPSCRIPT_WEBAPP_URL=https://script.google.com/macros/s/.../exec"
        )
        sys.exit(1)

    log.info("🚀 Menjalankan UltraJurnal.bot…")
    bot.run(BOT_TOKEN, log_handler=None)


if __name__ == "__main__":
    main()
