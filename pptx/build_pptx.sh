#!/bin/bash
# Build editable PPTX from Marp markdown
#
# Usage:
#   ./pptx/build_pptx.sh example.md                    # → example_editable.pptx
#   ./pptx/build_pptx.sh example.md -o output.pptx     # → output.pptx

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

INPUT="${1:?Usage: build_pptx.sh <input.md> [-o output.pptx]}"
shift

# Parse optional -o flag
OUTPUT=""
while [[ $# -gt 0 ]]; do
    case "$1" in
        -o) OUTPUT="$2"; shift 2 ;;
        *)  shift ;;
    esac
done

if [[ -z "$OUTPUT" ]]; then
    OUTPUT="${INPUT%.md}_editable.pptx"
fi

REFERENCE="$SCRIPT_DIR/reference.pptx"
INTERMEDIATE="/tmp/marp_pandoc_$$.md"

# Generate reference template if missing
if [[ ! -f "$REFERENCE" ]]; then
    echo "Generating reference template..."
    python3 "$SCRIPT_DIR/make_reference.py"
fi

# Step 1: Marp → Pandoc markdown
echo "Converting Marp → Pandoc markdown..."
python3 "$SCRIPT_DIR/marp2pandoc.py" "$INPUT" -o "$INTERMEDIATE"

# Step 2: Convert SVG assets to PNG for Pandoc
ASSET_DIR="$(dirname "$INPUT")/assets"
if [[ -d "$ASSET_DIR" ]]; then
    echo "Converting SVG assets to PNG..."
    for svg in "$ASSET_DIR"/*.svg; do
        [[ -f "$svg" ]] || continue
        png="${svg%.svg}.png"
        if [[ ! -f "$png" ]] || [[ "$svg" -nt "$png" ]]; then
            python3 -c "
import cairosvg
cairosvg.svg2png(url='$svg', write_to='$png', output_width=1400, dpi=300)
" 2>/dev/null && echo "  SVG→PNG: $(basename "$svg")"
        fi
    done
    # Rewrite .svg references to .png in intermediate markdown
    sed -i '' 's/\.svg)/.png)/g' "$INTERMEDIATE" 2>/dev/null || \
    sed -i 's/\.svg)/.png)/g' "$INTERMEDIATE" 2>/dev/null || true
fi

# Step 3: Pandoc → PPTX
echo "Converting Pandoc markdown → PPTX..."
pandoc "$INTERMEDIATE" \
    -t pptx \
    --reference-doc="$REFERENCE" \
    --slide-level=1 \
    --columns=50 \
    -o "$OUTPUT"

# Cleanup
rm -f "$INTERMEDIATE"

echo "Done: $OUTPUT"
echo ""
echo "Generated editable PPTX with structured text boxes."
echo "Open in PowerPoint / Keynote / Google Slides to edit."
