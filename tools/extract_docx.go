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

// Extract and pretty-print XML files from a DOCX for inspection
// Usage: go run extract_docx.go <path-to-docx> [output-directory]

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run extract_docx.go <path-to-docx> [output-directory]")
		fmt.Println("  Extracts and pretty-prints all XML files for manual inspection")
		os.Exit(1)
	}

	docxPath := os.Args[1]
	outputDir := "docx_extracted"
	if len(os.Args) > 2 {
		outputDir = os.Args[2]
	}

	fmt.Printf("Extracting: %s\n", docxPath)
	fmt.Printf("Output directory: %s\n\n", outputDir)

	// Open DOCX
	r, err := zip.OpenReader(docxPath)
	if err != nil {
		fmt.Printf("❌ Failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Create output directory
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		fmt.Printf("❌ Failed to create output directory: %v\n", err)
		os.Exit(1)
	}

	extractedCount := 0
	prettyPrintedCount := 0

	// Extract all files
	for _, f := range r.File {
		destPath := filepath.Join(outputDir, f.Name)

		// Create directory structure
		if err := os.MkdirAll(filepath.Dir(destPath), 0755); err != nil {
			fmt.Printf("⚠️  Failed to create directory for %s: %v\n", f.Name, err)
			continue
		}

		// Open source file
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("⚠️  Failed to open %s: %v\n", f.Name, err)
			continue
		}

		// Read content
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			fmt.Printf("⚠️  Failed to read %s: %v\n", f.Name, err)
			continue
		}

		// If it's XML or rels, try to pretty-print it
		isXML := strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels")

		if isXML {
			// Try to parse and pretty-print
			var doc interface{}
			if err := xml.Unmarshal(content, &doc); err != nil {
				// Not valid XML, write as-is but note the error
				fmt.Printf("⚠️  %s - Invalid XML (saved as-is): %v\n", f.Name, err)
			} else {
				// Pretty-print it
				prettyContent, err := xml.MarshalIndent(doc, "", "  ")
				if err == nil {
					content = []byte(xml.Header + string(prettyContent))
					prettyPrintedCount++
					fmt.Printf("✓ %s (pretty-printed)\n", f.Name)
				} else {
					fmt.Printf("✓ %s (XML valid but couldn't pretty-print)\n", f.Name)
				}
			}
		} else {
			fmt.Printf("✓ %s (binary)\n", f.Name)
		}

		// Write to file
		if err := os.WriteFile(destPath, content, 0644); err != nil {
			fmt.Printf("⚠️  Failed to write %s: %v\n", destPath, err)
			continue
		}

		extractedCount++
	}

	fmt.Printf("\n=== Summary ===\n")
	fmt.Printf("✓ Extracted %d files to: %s\n", extractedCount, outputDir)
	fmt.Printf("✓ Pretty-printed %d XML files\n", prettyPrintedCount)
	fmt.Println("\nYou can now inspect the files manually.")
	fmt.Println("Look for:")
	fmt.Println("  - Invalid characters in text elements")
	fmt.Println("  - Missing namespace declarations")
	fmt.Println("  - Broken relationship references")
	fmt.Println("  - Malformed chart or table XML")

	// List key files to check
	fmt.Println("\nKey files to inspect:")
	keyFiles := []string{
		"[Content_Types].xml",
		"word/document.xml",
		"word/_rels/document.xml.rels",
	}

	for _, kf := range keyFiles {
		fullPath := filepath.Join(outputDir, kf)
		if fileExists(fullPath) {
			fmt.Printf("  - %s\n", fullPath)
		}
	}

	// List all charts
	chartDir := filepath.Join(outputDir, "word", "charts")
	if entries, err := os.ReadDir(chartDir); err == nil && len(entries) > 0 {
		fmt.Println("\nChart files found:")
		for _, entry := range entries {
			if !entry.IsDir() {
				fmt.Printf("  - %s\n", filepath.Join(chartDir, entry.Name()))
			}
		}
	}
}

func fileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
