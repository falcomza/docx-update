#!/bin/bash

# Extract table content from a DOCX file to verify tables exist
# Usage: ./tools/extract_tables.sh <path-to-docx>

DOCX_FILE="${1:-outputs/table_orientation_demo.docx}"

if [ ! -f "$DOCX_FILE" ]; then
    echo "❌ Error: File not found: $DOCX_FILE"
    exit 1
fi

echo "Extracting tables from: $DOCX_FILE"
echo "=========================================="
echo

# Extract and format table content
unzip -p "$DOCX_FILE" word/document.xml 2>/dev/null | \
    grep -o '<w:tbl>.*</w:tbl>' | \
    sed 's/<w:tc>/\n  Cell: /g' | \
    grep -o '<w:t>[^<]*</w:t>' | \
    sed 's/<w:t>//g; s/<\/w:t>//g' | \
    head -50

echo
echo "✓ Table extraction complete!"
