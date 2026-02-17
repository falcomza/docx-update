package docxupdater

import (
	"strconv"
	"strings"
)

// xmlEscape escapes XML special characters.
// Used by both chart XML and Excel XML generation.
func xmlEscape(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

// formatFloat formats a float64 for XML output.
// Removes trailing zeros and unnecessary decimal points.
func formatFloat(f float64) string {
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// columnLetters converts a column number (1-based) to Excel column letters.
// Examples: 1->A, 2->B, 26->Z, 27->AA, 28->AB
func columnLetters(n int) string {
	if n <= 0 {
		return "A"
	}
	var out []byte
	for n > 0 {
		n--
		out = append([]byte{byte('A' + (n % 26))}, out...)
		n /= 26
	}
	return string(out)
}

// cellRef generates an Excel cell reference from column and row numbers.
// Both col and row are 1-based. Example: cellRef(1, 1) -> "A1"
func cellRef(col, row int) string {
	return columnLetters(col) + strconv.Itoa(row)
}
