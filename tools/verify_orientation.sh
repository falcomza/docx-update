#!/bin/bash

# Verify orientation changes in a DOCX file
# Usage: ./tools/verify_orientation.sh <path-to-docx>

DOCX_FILE="${1:-outputs/table_orientation_demo.docx}"

if [ ! -f "$DOCX_FILE" ]; then
    echo "❌ Error: File not found: $DOCX_FILE"
    exit 1
fi

echo "Analyzing: $DOCX_FILE"
echo "=========================================="
echo

# Check file type
echo "File Type:"
file "$DOCX_FILE"
echo

# Count landscape orientations
LANDSCAPE_COUNT=$(unzip -p "$DOCX_FILE" word/document.xml 2>/dev/null | grep -o 'w:orient="landscape"' | wc -l)
echo "Landscape Sections: $LANDSCAPE_COUNT"

# Count section breaks
SECTION_COUNT=$(unzip -p "$DOCX_FILE" word/document.xml 2>/dev/null | grep -o '<w:sectPr>' | wc -l)
echo "Total Sections: $SECTION_COUNT"

# Count tables
TABLE_COUNT=$(unzip -p "$DOCX_FILE" word/document.xml 2>/dev/null | grep -o '<w:tbl>' | wc -l)
echo "Tables Found: $TABLE_COUNT"

echo
echo "Page Size Information:"
echo "----------------------"

# Extract page sizes
unzip -p "$DOCX_FILE" word/document.xml 2>/dev/null | grep -o 'w:pgSz[^/]*/' | head -5

echo
echo "✓ Verification complete!"
echo
echo "Expected values for example_table_orientation.go:"
echo "  - Landscape Sections: 2"
echo "  - Total Sections: 5"
echo "  - Tables Found: 2"
