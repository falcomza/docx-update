//go:build ignore

package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// Debug tool to investigate Word corruption issues
// Usage: go run debug_tool.go <path-to-docx-file>

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_tool.go <path-to-docx-file>")
		os.Exit(1)
	}

	docxPath := os.Args[1]
	fmt.Printf("Analyzing: %s\n\n", docxPath)

	// Open the DOCX file (it's a ZIP archive)
	r, err := zip.OpenReader(docxPath)
	if err != nil {
		fmt.Printf("❌ Failed to open DOCX: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	fmt.Println("=== DOCX Structure ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}
	fmt.Println()

	errors := 0

	// Check critical files
	errors += checkXMLFile(r, "[Content_Types].xml", "Content Types")
	errors += checkXMLFile(r, "word/document.xml", "Main Document")
	errors += checkXMLFile(r, "word/_rels/document.xml.rels", "Document Relationships")
	errors += checkXMLFile(r, "_rels/.rels", "Package Relationships")

	// Check all chart files
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "word/charts/chart") && strings.HasSuffix(f.Name, ".xml") {
			errors += checkXMLFile(r, f.Name, fmt.Sprintf("Chart: %s", filepath.Base(f.Name)))
		}
		if strings.HasPrefix(f.Name, "word/charts/_rels/") && strings.HasSuffix(f.Name, ".rels") {
			errors += checkXMLFile(r, f.Name, fmt.Sprintf("Chart Rel: %s", filepath.Base(f.Name)))
		}
	}

	// Check for common problematic files
	checkXMLFile(r, "word/header1.xml", "Header")
	checkXMLFile(r, "word/footer1.xml", "Footer")

	fmt.Println("\n=== Summary ===")
	if errors == 0 {
		fmt.Println("✓ No XML parsing errors found")
		fmt.Println("\nNext steps:")
		fmt.Println("1. Extract the DOCX manually and inspect XML for:")
		fmt.Println("   - Invalid characters (control chars like 0x00-0x1F)")
		fmt.Println("   - Missing namespace declarations")
		fmt.Println("   - Incorrect relationship IDs")
		fmt.Println("2. Compare with a known-good DOCX from your library")
		fmt.Println("3. Use Office Open XML SDK 2.5 Productivity Tool")
		fmt.Println("4. Check specific areas mentioned in error details above")
	} else {
		fmt.Printf("❌ Found %d XML error(s) - see details above\n", errors)
	}
}

func checkXMLFile(r *zip.ReadCloser, path string, description string) int {
	fmt.Printf("Checking %s... ", description)

	f := findFile(r, path)
	if f == nil {
		fmt.Println("⚠️  Not found (may be optional)")
		return 0
	}

	rc, err := f.Open()
	if err != nil {
		fmt.Printf("❌ Can't open: %v\n", err)
		return 1
	}
	defer rc.Close()

	content, err := io.ReadAll(rc)
	if err != nil {
		fmt.Printf("❌ Can't read: %v\n", err)
		return 1
	}

	// Try to parse as XML
	var result interface{}
	if err := xml.Unmarshal(content, &result); err != nil {
		fmt.Printf("❌ INVALID XML: %v\n", err)

		// Show problematic area
		contentStr := string(content)
		if len(contentStr) > 500 {
			fmt.Printf("   First 500 chars: %s...\n", contentStr[:500])
		} else {
			fmt.Printf("   Content: %s\n", contentStr)
		}

		// Check for common issues
		checkCommonIssues(content, path)
		return 1
	}

	// Additional validation
	warnings := validateXMLContent(content, path)
	if len(warnings) > 0 {
		fmt.Printf("⚠️  Warnings:\n")
		for _, w := range warnings {
			fmt.Printf("   - %s\n", w)
		}
		return 0 // Warnings, not errors
	}

	fmt.Println("✓ Valid")
	return 0
}

func checkCommonIssues(content []byte, path string) {
	contentStr := string(content)

	// Check for control characters
	for i, ch := range content {
		if ch < 0x20 && ch != 0x09 && ch != 0x0A && ch != 0x0D {
			fmt.Printf("   ⚠️  Invalid control character 0x%02X at byte %d\n", ch, i)
			break
		}
	}

	// Check for unescaped XML characters in text content
	if strings.Contains(contentStr, "<w:t>") {
		// This is simplified - just checking patterns
		if strings.Contains(contentStr, "<w:t>&") && !strings.Contains(contentStr, "&amp;") &&
			!strings.Contains(contentStr, "&lt;") && !strings.Contains(contentStr, "&gt;") {
			fmt.Println("   ⚠️  Possible unescaped & character in text")
		}
	}

	// Check for unclosed tags (basic check)
	openTags := strings.Count(contentStr, "<")
	closeTags := strings.Count(contentStr, ">")
	if openTags != closeTags {
		fmt.Printf("   ⚠️  Mismatched angle brackets: %d < vs %d >\n", openTags, closeTags)
	}
}

func validateXMLContent(content []byte, path string) []string {
	var warnings []string
	contentStr := string(content)

	// Check for namespace declarations in relationship files
	if strings.Contains(path, ".rels") {
		if !strings.Contains(contentStr, "http://schemas.openxmlformats.org/package/2006/relationships") {
			warnings = append(warnings, "Missing standard relationships namespace")
		}
	}

	// Check for required namespaces in document.xml
	if strings.Contains(path, "document.xml") && !strings.Contains(path, ".rels") {
		requiredNS := []string{
			"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		}
		for _, ns := range requiredNS {
			if !strings.Contains(contentStr, ns) {
				warnings = append(warnings, fmt.Sprintf("Missing namespace: %s", ns))
			}
		}
	}

	// Check for empty relationship IDs
	if strings.Contains(path, ".rels") {
		if strings.Contains(contentStr, `Id=""`) || strings.Contains(contentStr, `Id=''`) {
			warnings = append(warnings, "Empty relationship ID found")
		}
	}

	return warnings
}

func findFile(r *zip.ReadCloser, path string) *zip.File {
	for _, f := range r.File {
		if f.Name == path {
			return f
		}
	}
	return nil
}
