//go:build ignore

package main

import (
	"log"

	docxupdater "github.com/falcomza/docx-update"
)

func main() {
	// Create a simple demo showing conditional cell coloring
	updater, err := docxupdater.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	err = updater.AddHeading(1, "Conditional Cell Color Demo", docxupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	// Simple table with status-based coloring
	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Server", Alignment: docxupdater.CellAlignLeft},
			{Title: "Status", Alignment: docxupdater.CellAlignCenter},
			{Title: "Response Time", Alignment: docxupdater.CellAlignRight},
		},
		Rows: [][]string{
			{"Web Server", "Critical", "3000ms"},
			{"Database", "Normal", "45ms"},
			{"API Gateway", "Warning", "500ms"},
			{"Cache", "Normal", "12ms"},
			{"Load Balancer", "Critical", "Timeout"},
			{"Auth Service", "Non-critical", "250ms"},
		},
		HeaderBold:       true,
		HeaderBackground: "2F5496",
		HeaderStyle: docxupdater.CellStyle{
			FontColor: "FFFFFF",
		},
		// THIS IS THE NEW FEATURE - Conditional cell coloring!
		ConditionalStyles: map[string]docxupdater.CellStyle{
			"Critical": {
				Background: "FF0000", // Red background
				FontColor:  "FFFFFF", // White text
				Bold:       true,
			},
			"Warning": {
				Background: "FFA500", // Orange background
				FontColor:  "000000",
				Bold:       true,
			},
			"Non-critical": {
				Background: "FFD966", // Light orange/amber
				FontColor:  "000000",
			},
			"Normal": {
				Background: "00B050", // Green background
				FontColor:  "FFFFFF",
			},
		},
		BorderStyle:    docxupdater.BorderSingle,
		TableAlignment: docxupdater.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert table: %v", err)
	}

	// Save
	outputPath := "./outputs/demo_conditional_colors.docx"
	err = updater.Save(outputPath)
	if err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	log.Printf("✓ Demo created: %s", outputPath)
	log.Println("\nOpen the document to see:")
	log.Println("  - 'Critical' cells with RED background")
	log.Println("  - 'Warning' cells with ORANGE background")
	log.Println("  - 'Non-critical' cells with AMBER background")
	log.Println("  - 'Normal' cells with GREEN background")
}
