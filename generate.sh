#!/usr/bin/env bash
# generate.sh — Interactive invoice generator
# Usage: ./generate.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "╔══════════════════════════════════════╗"
echo "║        Invoice Generator             ║"
echo "╚══════════════════════════════════════╝"
echo ""

# ── Ask: current month or specific months? ───────────────────────────────────
echo "Generate invoice for:"
echo "  1) Current month (default)"
echo "  2) Specific month(s)"
echo ""
read -rp "Choice [1/2]: " choice
choice="${choice:-1}"

months_arg=""

if [[ "$choice" == "2" ]]; then
    echo ""
    echo "Enter one or more months in YYYYMM format, separated by spaces."
    echo "Example: 202603 202604 202605"
    echo ""
    read -rp "Month(s): " raw_input

    # Validate each token is exactly 6 digits
    for token in $raw_input; do
        if [[ ! "$token" =~ ^[0-9]{6}$ ]]; then
            echo "Error: '$token' is not a valid YYYYMM value. Exiting." >&2
            exit 1
        fi
    done

    months_arg="$raw_input"
fi

# ── Confirm before running ────────────────────────────────────────────────────
echo ""
if [[ -z "$months_arg" ]]; then
    current_month=$(date +%Y%m)
    echo "Will generate invoice for: $current_month (current month)"
else
    echo "Will generate invoices for: $months_arg"
fi

echo ""
read -rp "Proceed? [Y/n]: " confirm
confirm="${confirm:-Y}"
if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
    echo "Aborted."
    exit 0
fi

# ── Run the .NET app ──────────────────────────────────────────────────────────
echo ""
if [[ -z "$months_arg" ]]; then
    dotnet run
else
    dotnet run -- $months_arg
fi

echo ""
echo "Done. Check the output/ folder for generated files."
echo "Run ./convert-to-pdf.sh to convert .docx files to PDF."
