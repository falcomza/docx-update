//go:build ignore

package main

import (
	"fmt"
	"log"
	"time"

	godocx "github.com/falcomza/go-docx"
)

func main() {
	// Open the template document
	updater, err := godocx.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add a title for the report
	err = updater.AddHeading(1, "Monthly Sales Report - "+time.Now().Format("January 2006"), godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}

	// Add subtitle
	err = updater.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Generated on: " + time.Now().Format("January 2, 2006"),
		Style:    godocx.StyleSubtitle,
		Position: godocx.PositionEnd,
		Italic:   true,
	})
	if err != nil {
		log.Fatalf("Failed to add subtitle: %v", err)
	}

	// Add section heading
	err = updater.AddHeading(2, "Sales by Region", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add section heading: %v", err)
	}

	// Create a professional sales table
	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Region", Alignment: godocx.CellAlignLeft},
			{Title: "Q1 Sales", Alignment: godocx.CellAlignRight},
			{Title: "Q2 Sales", Alignment: godocx.CellAlignRight},
			{Title: "Q3 Sales", Alignment: godocx.CellAlignRight},
			{Title: "Q4 Sales", Alignment: godocx.CellAlignRight},
			{Title: "Total", Alignment: godocx.CellAlignRight},
		},
		Rows: [][]string{
			{"North America", "$125,000", "$132,000", "$145,000", "$158,000", "$560,000"},
			{"Europe", "$98,000", "$105,000", "$112,000", "$120,000", "$435,000"},
			{"Asia Pacific", "$87,000", "$95,000", "$108,000", "$115,000", "$405,000"},
			{"Latin America", "$45,000", "$48,000", "$52,000", "$55,000", "$200,000"},
			{"Middle East", "$32,000", "$35,000", "$38,000", "$41,000", "$146,000"},
		},
		HeaderBold:        true,
		HeaderBackground:  "2E75B5",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "E7E6E6",
		BorderStyle:       godocx.BorderSingle,
		BorderSize:        6,
		BorderColor:       "2E75B5",
		TableAlignment:    godocx.AlignCenter,
		RepeatHeader:      true,
		RowStyle: godocx.CellStyle{
			FontSize: 20, // 10pt
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert sales table: %v", err)
	}

	// Add another section
	err = updater.AddHeading(2, "Top Performers", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add performers heading: %v", err)
	}

	// Create employee performance table
	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Rank", Alignment: godocx.CellAlignCenter},
			{Title: "Employee Name", Alignment: godocx.CellAlignLeft},
			{Title: "Department", Alignment: godocx.CellAlignLeft},
			{Title: "Sales", Alignment: godocx.CellAlignRight},
			{Title: "Target", Alignment: godocx.CellAlignRight},
			{Title: "Achievement", Alignment: godocx.CellAlignCenter},
		},
		ColumnWidths: []int{600, 2000, 1500, 1200, 1200, 1000}, // Custom widths
		Rows: [][]string{
			{"1", "Sarah Johnson", "North America", "$45,000", "$35,000", "129%"},
			{"2", "Michael Chen", "Asia Pacific", "$42,000", "$33,000", "127%"},
			{"3", "Emma Williams", "Europe", "$38,000", "$30,000", "127%"},
			{"4", "David Martinez", "Latin America", "$35,000", "$28,000", "125%"},
			{"5", "Lisa Anderson", "North America", "$33,000", "$27,000", "122%"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "DEEBF7",
		BorderStyle:       godocx.BorderSingle,
		BorderSize:        4,
		TableAlignment:    godocx.AlignCenter,
		RowStyle: godocx.CellStyle{
			FontSize: 20,
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert performers table: %v", err)
	}

	// Add product inventory section
	err = updater.AddHeading(2, "Product Inventory Status", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add inventory heading: %v", err)
	}

	// Create inventory table with custom styling
	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Product Code", Alignment: godocx.CellAlignLeft},
			{Title: "Product Name", Alignment: godocx.CellAlignLeft},
			{Title: "Category", Alignment: godocx.CellAlignLeft},
			{Title: "In Stock", Alignment: godocx.CellAlignRight},
			{Title: "Status", Alignment: godocx.CellAlignCenter},
		},
		Rows: [][]string{
			{"PRD-001", "Wireless Mouse", "Electronics", "245", "✓ Available"},
			{"PRD-002", "USB Keyboard", "Electronics", "12", "⚠ Low Stock"},
			{"PRD-003", "Monitor 24\"", "Electronics", "0", "✗ Out of Stock"},
			{"PRD-004", "Office Chair", "Furniture", "78", "✓ Available"},
			{"PRD-005", "Standing Desk", "Furniture", "34", "✓ Available"},
			{"PRD-006", "Desk Lamp", "Accessories", "156", "✓ Available"},
			{"PRD-007", "Notebook A4", "Stationery", "2", "⚠ Low Stock"},
			{"PRD-008", "Pen Set", "Stationery", "345", "✓ Available"},
		},
		HeaderBold:        true,
		HeaderBackground:  "70AD47",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "E2EFD9",
		BorderStyle:       godocx.BorderSingle,
		BorderSize:        6,
		BorderColor:       "70AD47",
		TableAlignment:    godocx.AlignCenter,
		RepeatHeader:      true,
	})
	if err != nil {
		log.Fatalf("Failed to insert inventory table: %v", err)
	}

	// Add footer note
	err = updater.AddText("──────────────────────────────────────────────────────────", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add separator: %v", err)
	}

	err = updater.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Note: All figures are in USD. Report generated automatically by the sales tracking system.",
		Style:    godocx.StyleNormal,
		Position: godocx.PositionEnd,
		Italic:   true,
	})
	if err != nil {
		log.Fatalf("Failed to add footer note: %v", err)
	}

	// Save the document
	outputPath := "./outputs/table_example_output.docx"
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("✅ SUCCESS!")
	fmt.Printf("📄 Output saved to: %s\n", outputPath)
	fmt.Println("\nCreated tables:")
	fmt.Println("  • Sales by Region (with header repeat)")
	fmt.Println("  • Top Performers (custom column widths)")
	fmt.Println("  • Product Inventory Status (with status indicators)")
}
