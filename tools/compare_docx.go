//go:build ignore

package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

// Compare two DOCX files to find structural differences
// Usage: go run compare_docx.go <working.docx> <broken.docx>

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: go run compare_docx.go <working.docx> <broken.docx>")
		fmt.Println("  Compares two DOCX files to identify differences")
		os.Exit(1)
	}

	file1 := os.Args[1]
	file2 := os.Args[2]

	fmt.Printf("Comparing:\n")
	fmt.Printf("  Working: %s\n", file1)
	fmt.Printf("  Broken:  %s\n\n", file2)

	// Load both files
	docx1, err := loadDocx(file1)
	if err != nil {
		fmt.Printf("❌ Failed to load %s: %v\n", file1, err)
		os.Exit(1)
	}

	docx2, err := loadDocx(file2)
	if err != nil {
		fmt.Printf("❌ Failed to load %s: %v\n", file2, err)
		os.Exit(1)
	}

	// Compare file lists
	fmt.Println("=== File Structure ===")
	filesOnly1 := make([]string, 0)
	filesOnly2 := make([]string, 0)
	commonFiles := make([]string, 0)

	for name := range docx1 {
		if _, ok := docx2[name]; ok {
			commonFiles = append(commonFiles, name)
		} else {
			filesOnly1 = append(filesOnly1, name)
		}
	}

	for name := range docx2 {
		if _, ok := docx1[name]; !ok {
			filesOnly2 = append(filesOnly2, name)
		}
	}

	if len(filesOnly1) > 0 {
		fmt.Printf("\nFiles only in %s:\n", file1)
		for _, name := range filesOnly1 {
			fmt.Printf("  - %s\n", name)
		}
	}

	if len(filesOnly2) > 0 {
		fmt.Printf("\nFiles only in %s:\n", file2)
		for _, name := range filesOnly2 {
			fmt.Printf("  - %s\n", name)
		}
	}

	fmt.Printf("\nCommon files: %d\n", len(commonFiles))

	// Compare content of common XML files
	fmt.Println("\n=== Content Differences ===")
	differences := 0

	for _, name := range commonFiles {
		isXML := strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels")
		if !isXML {
			continue
		}

		content1 := docx1[name]
		content2 := docx2[name]

		// Skip if identical
		if string(content1) == string(content2) {
			continue
		}

		differences++
		fmt.Printf("\n--- %s ---\n", name)

		// Try to parse both as XML
		var xml1, xml2 interface{}
		err1 := xml.Unmarshal(content1, &xml1)
		err2 := xml.Unmarshal(content2, &xml2)

		if err1 != nil {
			fmt.Printf("  ⚠️  File 1 has INVALID XML: %v\n", err1)
		}
		if err2 != nil {
			fmt.Printf("  ❌ File 2 has INVALID XML: %v\n", err2)
		}

		// Size difference
		fmt.Printf("  Size: %d vs %d bytes (%+d)\n",
			len(content1), len(content2), len(content2)-len(content1))

		// Show first difference
		str1 := string(content1)
		str2 := string(content2)

		// Find first difference position
		minLen := len(str1)
		if len(str2) < minLen {
			minLen = len(str2)
		}

		firstDiff := -1
		for i := 0; i < minLen; i++ {
			if str1[i] != str2[i] {
				firstDiff = i
				break
			}
		}

		if firstDiff >= 0 {
			contextStart := firstDiff - 50
			if contextStart < 0 {
				contextStart = 0
			}
			contextEnd := firstDiff + 50

			fmt.Printf("  First difference at position %d:\n", firstDiff)

			if contextEnd <= len(str1) {
				fmt.Printf("    File 1: ...%s...\n", str1[contextStart:contextEnd])
			}
			if contextEnd <= len(str2) {
				fmt.Printf("    File 2: ...%s...\n", str2[contextStart:contextEnd])
			}
		}

		// Check for specific issues
		checkDifferences(name, content1, content2)
	}

	if differences == 0 {
		fmt.Println("✓ No content differences found in XML files")
	} else {
		fmt.Printf("\n=== Summary ===\n")
		fmt.Printf("Found %d file(s) with differences\n", differences)
	}
}

func checkDifferences(filename string, content1, content2 []byte) {
	str1 := string(content1)
	str2 := string(content2)

	// Check relationship IDs
	if strings.Contains(filename, ".rels") {
		ids1 := extractRelIds(str1)
		ids2 := extractRelIds(str2)

		if len(ids1) != len(ids2) {
			fmt.Printf("  ⚠️  Different number of relationships: %d vs %d\n", len(ids1), len(ids2))
		}

		// Check for duplicate IDs in file 2
		idCount := make(map[string]int)
		for _, id := range ids2 {
			idCount[id]++
		}
		for id, count := range idCount {
			if count > 1 {
				fmt.Printf("  ❌ Duplicate relationship ID '%s' appears %d times in file 2\n", id, count)
			}
		}
	}

	// Check namespace declarations
	ns1 := strings.Count(str1, "xmlns:")
	ns2 := strings.Count(str2, "xmlns:")
	if ns1 != ns2 {
		fmt.Printf("  ⚠️  Different number of namespace declarations: %d vs %d\n", ns1, ns2)
	}

	// Check for control characters in file 2
	for i, ch := range content2 {
		if ch < 0x20 && ch != 0x09 && ch != 0x0A && ch != 0x0D {
			fmt.Printf("  ❌ Invalid control character 0x%02X at byte %d in file 2\n", ch, i)
			break
		}
	}
}

func extractRelIds(content string) []string {
	ids := make([]string, 0)
	// Simple extraction: look for Id="rIdXXX"
	parts := strings.Split(content, `Id="`)
	for i := 1; i < len(parts); i++ {
		endIdx := strings.Index(parts[i], `"`)
		if endIdx > 0 {
			ids = append(ids, parts[i][:endIdx])
		}
	}
	return ids
}

func loadDocx(path string) (map[string][]byte, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	result := make(map[string][]byte)
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}

		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}

		result[f.Name] = content
	}

	return result, nil
}
