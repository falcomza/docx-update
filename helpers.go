package godocx

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
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

// normalizeHexColor normalizes a hex color code for use in Office Open XML.
// Accepts colors with or without '#' prefix. Returns empty string if invalid.
// Examples: "#FF0000" -> "FF0000", "ff0000" -> "FF0000"
func normalizeHexColor(color string) string {
	c := strings.TrimSpace(color)
	if c == "" {
		return ""
	}
	c = strings.TrimPrefix(c, "#")
	if len(c) != 6 {
		return ""
	}
	for _, ch := range c {
		if !(ch >= '0' && ch <= '9' || ch >= 'a' && ch <= 'f' || ch >= 'A' && ch <= 'F') {
			return ""
		}
	}
	return strings.ToUpper(c)
}

// getNextDocPrId finds the next available docPr ID in the document.
func (u *Updater) getNextDocPrId() (int, error) {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document: %w", err)
	}

	matches := docPrIDPattern.FindAllStringSubmatch(string(raw), -1)

	maxId := 0
	for _, match := range matches {
		if len(match) > 1 {
			id, err := strconv.Atoi(match[1])
			if err != nil {
				continue
			}
			if id > maxId {
				maxId = id
			}
		}
	}

	return maxId + 1, nil
}

// getNextDocumentRelId finds the next available relationship ID in document.xml.rels.
func (u *Updater) getNextDocumentRelId() (string, error) {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	return getNextRelIDFromFile(relsPath)
}

// getNextRelIDFromFile finds the next available relationship ID in a .rels file.
func getNextRelIDFromFile(relsPath string) (string, error) {
	raw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read rels file %s: %w", relsPath, err)
	}

	var rels relationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return "", fmt.Errorf("parse rels file %s: %w", relsPath, err)
	}

	maxId := 0
	for _, rel := range rels.Relationships {
		if matches := relIDPattern.FindStringSubmatch(rel.ID); matches != nil {
			id, err := strconv.Atoi(matches[1])
			if err != nil {
				continue
			}
			if id > maxId {
				maxId = id
			}
		}
	}

	return fmt.Sprintf("rId%d", maxId+1), nil
}
