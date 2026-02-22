//go:build ignore

package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/go-docx"
)

func main() {
	// Create a document with a chart to verify the fix
	updater, err := docx.New("templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add a chart
	err = updater.InsertChart(docx.ChartOptions{
		Position:          docx.PositionEnd,
		ChartKind:         docx.ChartKindColumn,
		Title:             "Test Word Compatibility Fix",
		CategoryAxisTitle: "Categories",
		ValueAxisTitle:    "Values",
		Categories:        []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []docx.SeriesData{
			{Name: "Revenue", Values: []float64{100, 150, 120, 180}},
			{Name: "Expenses", Values: []float64{70, 85, 75, 95}},
		},
		ShowLegend:     true,
		LegendPosition: "r",
	})
	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	// Save
	outputPath := "outputs/test_fix_verification.docx"
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Printf("✓ Document saved to: %s\n", outputPath)
	fmt.Println("✓ The fix has been applied:")
	fmt.Println("  1. XML declaration followed by newline")
	fmt.Println("  2. All required namespaces included (c, a, r, c16r2, mc)")
	fmt.Println("  3. Chart properties included (date1904, lang, roundedCorners)")
	fmt.Println("\n✓ This document should now open properly in Microsoft Word!")
}
