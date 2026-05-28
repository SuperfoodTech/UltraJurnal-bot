#!/usr/bin/env bash
# =============================================================================
#  BCA PDF to XLSX Converter
#  Double-click untuk menjalankan
# =============================================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_DIR="$SCRIPT_DIR/Src"
OUT_DIR="$SCRIPT_DIR/Hasil"

# ── Jika di-double-click dari file manager (bukan dari terminal),
#    buka terminal emulator baru dan jalankan ulang script di dalamnya
if [ -z "$TERM" ] || [ "$TERM" = "dumb" ]; then
    for TE in gnome-terminal xterm konsole xfce4-terminal mate-terminal lxterminal tilix; do
        if command -v "$TE" &>/dev/null; then
            case "$TE" in
                gnome-terminal) exec gnome-terminal -- bash "$0" ;;
                xfce4-terminal) exec xfce4-terminal -e "bash \"$0\"" ;;
                konsole)        exec konsole -e bash "$0" ;;
                *)              exec "$TE" -e bash "$0" ;;
            esac
        fi
    done
fi

# ── Warna & style ─────────────────────────────────────────────────────────────
R='\033[0;31m'    # red
G='\033[0;32m'    # green
Y='\033[1;33m'    # yellow
B='\033[0;34m'    # blue
C='\033[0;36m'    # cyan
W='\033[1;37m'    # white bold
D='\033[2m'       # dim
NC='\033[0m'      # reset
BOLD='\033[1m'
LINE="${D}────────────────────────────────────────────────────────────${NC}"

clear

# ── Banner ────────────────────────────────────────────────────────────────────
echo -e "${C}"
echo "  ██████╗  ██████╗ █████╗      ██████╗ ██████╗ ███╗   ██╗██╗   ██╗"
echo "  ██╔══██╗██╔════╝██╔══██╗    ██╔════╝██╔═══██╗████╗  ██║██║   ██║"
echo "  ██████╔╝██║     ███████║    ██║     ██║   ██║██╔██╗ ██║██║   ██║"
echo "  ██╔══██╗██║     ██╔══██║    ██║     ██║   ██║██║╚██╗██║╚██╗ ██╔╝"
echo "  ██████╔╝╚██████╗██║  ██║    ╚██████╗╚██████╔╝██║ ╚████║ ╚████╔╝ "
echo "  ╚═════╝  ╚═════╝╚═╝  ╚═╝     ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝  ╚═══╝ "
echo -e "${NC}"
echo -e "${W}              PDF → XLSX Mutasi Rekening Converter${NC}"
echo -e "${D}                    powered by Camelot + uv${NC}"
echo ""
echo -e "$LINE"

# ── Cek uv ───────────────────────────────────────────────────────────────────
if ! command -v uv &>/dev/null; then
    echo -e "\n${R}✖  'uv' tidak ditemukan.${NC}"
    echo -e "   Install dengan perintah:"
    echo -e "   ${Y}curl -LsSf https://astral.sh/uv/install.sh | sh${NC}\n"
    read -rp "   Tekan ENTER untuk keluar …" ; exit 1
fi

# ── Pastikan folder Src & Hasil ada ──────────────────────────────────────────
mkdir -p "$SRC_DIR" "$OUT_DIR"

# ── Cari file PDF di folder Src ───────────────────────────────────────────────
mapfile -t PDF_FILES < <(find "$SRC_DIR" -maxdepth 1 -iname "*.pdf" | sort)

echo ""
if [ ${#PDF_FILES[@]} -eq 0 ]; then
    echo -e "  ${Y}⚠  Tidak ada file PDF di folder Src/.${NC}"
    echo -e "  ${D}   Letakkan file PDF di:${NC}"
    echo -e "  ${W}   $SRC_DIR${NC}\n"
    read -rp "  Tekan ENTER untuk keluar …"
    exit 0
fi

# ── Menu pilih file ───────────────────────────────────────────────────────────
echo -e "  ${W}📂 File PDF tersedia di folder Src/:${NC}\n"
for i in "${!PDF_FILES[@]}"; do
    fname=$(basename "${PDF_FILES[$i]}")
    size=$(du -h "${PDF_FILES[$i]}" | cut -f1)
    printf "  ${C}  [%d]${NC}  %-40s ${D}%s${NC}\n" "$((i+1))" "$fname" "$size"
done

echo ""
echo -e "  ${C}  [A]${NC}  Konversi semua file sekaligus"
echo -e "  ${C}  [Q]${NC}  Keluar"
echo ""
echo -e "$LINE"
echo ""
printf "  ${W}Pilih nomor file (1-%d), A, atau Q: ${NC}" "${#PDF_FILES[@]}"
read -r PILIHAN

echo ""

convert_file() {
    local pdf="$1"
    local fname
    fname=$(basename "$pdf")
    local outfile="$OUT_DIR/${fname%.pdf}.xlsx"

    START_TIME=$(date +"%H:%M:%S")
    SECONDS=0
    echo -e "  ${Y}⟳  [${START_TIME}] Memproses:${NC} ${W}$fname${NC}"
    echo -e "  ${D}   Output  → $outfile${NC}"
    echo -e "  ${D}   Logs:${NC}"

    # Jalankan via uv secara langsung agar log real-time muncul di layar
    uv run --project "$SCRIPT_DIR" "$SCRIPT_DIR/converter.py" "$pdf" -o "$outfile"
    STATUS=$?

    ELAPSED=$SECONDS
    END_TIME=$(date +"%H:%M:%S")
    ELAPSED_MIN=$((ELAPSED / 60))
    ELAPSED_SEC=$((ELAPSED % 60))
    if [ $ELAPSED_MIN -gt 0 ]; then
        ELAPSED_STR="${ELAPSED_MIN}m ${ELAPSED_SEC}s"
    else
        ELAPSED_STR="${ELAPSED_SEC}s"
    fi

    if [ $STATUS -eq 0 ]; then
        echo -e "\n  ${G}✔  [${END_TIME}] Berhasil!${NC}  ${D}→ $(basename "$outfile")${NC}"
        echo -e "  ${D}   ⏱  Durasi konversi: ${W}${ELAPSED_STR}${NC}"
    else
        echo -e "\n  ${R}✖  [${END_TIME}] Gagal mengkonversi $fname${NC}"
        echo -e "  ${D}   ⏱  Durasi berjalan: ${W}${ELAPSED_STR}${NC}"
    fi
    echo ""
}

PILIHAN_UPPER="${PILIHAN^^}"

if [ "$PILIHAN_UPPER" = "Q" ]; then
    echo -e "  ${D}Keluar.${NC}\n"
    exit 0

elif [ "$PILIHAN_UPPER" = "A" ]; then
    echo -e "$LINE"
    echo -e "  ${W}Mengkonversi semua ${#PDF_FILES[@]} file …${NC}\n"
    COUNT_OK=0
    for pdf in "${PDF_FILES[@]}"; do
        convert_file "$pdf"
        ((COUNT_OK++))
    done
    echo -e "$LINE"
    echo -e "\n  ${G}✔  Selesai! ${COUNT_OK} file dikonversi.${NC}"
    echo -e "  ${D}   Hasil tersimpan di: ${W}$OUT_DIR${NC}\n"

elif [[ "$PILIHAN" =~ ^[0-9]+$ ]] && [ "$PILIHAN" -ge 1 ] && [ "$PILIHAN" -le "${#PDF_FILES[@]}" ]; then
    SELECTED="${PDF_FILES[$((PILIHAN-1))]}"
    echo -e "$LINE"
    convert_file "$SELECTED"
    echo -e "$LINE"
    echo -e "\n  ${G}✔  Selesai!${NC}"
    echo -e "  ${D}   Hasil tersimpan di: ${W}$OUT_DIR${NC}\n"

else
    echo -e "  ${R}✖  Pilihan tidak valid: '${PILIHAN}'${NC}\n"
fi

# ── Tanya apakah ingin konversi lagi ─────────────────────────────────────────
echo -e "$LINE"
echo ""
printf "  Konversi file lain? (y/N): "
read -r LAGI
if [[ "${LAGI,,}" == "y" ]]; then
    exec bash "$0"
fi

echo ""
echo -e "  ${D}Tekan ENTER untuk menutup …${NC}"
read -r
