#!/usr/bin/env bash
# convert-to-pdf.sh — Convert all .docx files in output/ to PDF using LibreOffice
# Usage: ./convert-to-pdf.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT_DIR="$SCRIPT_DIR/output"

echo ""
echo "╔══════════════════════════════════════╗"
echo "║       PDF Converter                  ║"
echo "╚══════════════════════════════════════╝"
echo ""

# ── Locate LibreOffice ────────────────────────────────────────────────────────
SOFFICE=""
candidates=(
    "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    "/usr/local/bin/soffice"
    "/usr/bin/soffice"
    "/opt/libreoffice/program/soffice"
)

for candidate in "${candidates[@]}"; do
    if [[ -f "$candidate" ]]; then
        SOFFICE="$candidate"
        break
    fi
done

if [[ -z "$SOFFICE" ]]; then
    SOFFICE="$(command -v soffice 2>/dev/null || true)"
fi

if [[ -z "$SOFFICE" ]]; then
    echo "Error: LibreOffice not found." >&2
    echo "Install it with:  brew install --cask libreoffice" >&2
    exit 1
fi

echo "Using LibreOffice: $SOFFICE"
echo ""

# ── Find .docx files in output/ ───────────────────────────────────────────────
if [[ ! -d "$OUTPUT_DIR" ]]; then
    echo "No output/ folder found. Run ./generate.sh first."
    exit 1
fi

mapfile -t docx_files < <(find "$OUTPUT_DIR" -maxdepth 1 -name "*.docx" ! -name ".tmp_*" | sort)

if [[ ${#docx_files[@]} -eq 0 ]]; then
    echo "No .docx files found in output/. Run ./generate.sh first."
    exit 1
fi

echo "Found ${#docx_files[@]} file(s) to convert:"
for f in "${docx_files[@]}"; do
    echo "  $(basename "$f")"
done
echo ""

# ── Ask: convert all or pick specific ones? ───────────────────────────────────
echo "Convert:"
echo "  1) All files (default)"
echo "  2) Choose specific files"
echo ""
read -rp "Choice [1/2]: " choice
choice="${choice:-1}"

selected_files=()

if [[ "$choice" == "2" ]]; then
    echo ""
    for i in "${!docx_files[@]}"; do
        echo "  $((i+1))) $(basename "${docx_files[$i]}")"
    done
    echo ""
    read -rp "Enter numbers separated by spaces (e.g. 1 3): " picks
    for pick in $picks; do
        idx=$((pick - 1))
        if [[ $idx -lt 0 || $idx -ge ${#docx_files[@]} ]]; then
            echo "Error: '$pick' is out of range." >&2
            exit 1
        fi
        selected_files+=("${docx_files[$idx]}")
    done
else
    selected_files=("${docx_files[@]}")
fi

# ── Convert ───────────────────────────────────────────────────────────────────
echo ""
success=0
failed=0

for docx in "${selected_files[@]}"; do
    base="$(basename "$docx")"
    pdf="$OUTPUT_DIR/${base%.docx}.pdf"
    echo -n "Converting $base ... "

    if "$SOFFICE" --headless --convert-to pdf --outdir "$OUTPUT_DIR" "$docx" > /dev/null 2>&1; then
        echo "done → $(basename "$pdf")"
        ((success++))
    else
        echo "FAILED"
        ((failed++))
    fi
done

echo ""
echo "Converted: $success file(s)"
[[ $failed -gt 0 ]] && echo "Failed:    $failed file(s)"
echo "PDFs saved to: $OUTPUT_DIR"
