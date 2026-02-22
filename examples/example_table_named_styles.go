//go:build ignore

package main

import (
	"fmt"
	"log"

	godocx "github.com/falcomza/go-docx"
)

func main() {
	// Open the template document
	updater, err := godocx.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add title
	err = updater.AddHeading(1, "Named Word Styles in Tables", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}

	err = updater.AddText("This document demonstrates using Word's built-in named styles in tables instead of direct formatting.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add intro: %v", err)
	}

	// Example 1: Using built-in Word styles
	err = updater.AddHeading(2, "1. Using Built-in Word Styles", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Feature"},
			{Title: "Description"},
		},
		Rows: [][]string{
			{"Named Styles", "References Word styles defined in the document"},
			{"Consistency", "Maintains corporate style guidelines"},
			{"Template-based", "Styles can be customized in the template"},
		},
		HeaderStyleName:   "Heading2", // Word's Heading 2 style
		RowStyleName:      "BodyText", // Word's Body Text style
		HeaderBackground:  "4472C4",
		AlternateRowColor: "E7E6E6",
		BorderStyle:       godocx.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert built-in styles table: %v", err)
	}

	err = updater.AddText("↑ Header uses 'Heading2' style, rows use 'BodyText' style.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 2: Mixing named styles with direct formatting
	err = updater.AddHeading(2, "2. Mixing Named Styles + Direct Formatting", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Quarter", Alignment: godocx.CellAlignCenter},
			{Title: "Revenue", Alignment: godocx.CellAlignRight},
			{Title: "Growth", Alignment: godocx.CellAlignRight},
		},
		Rows: [][]string{
			{"Q1 2026", "$250,000", "+12%"},
			{"Q2 2026", "$280,000", "+15%"},
			{"Q3 2026", "$310,000", "+18%"},
			{"Q4 2026", "$340,000", "+21%"},
		},
		HeaderStyleName:  "Heading3", // Named style
		HeaderBold:       true,       // Plus direct bold
		HeaderBackground: "2E75B5",   // Plus direct background
		HeaderAlignment:  godocx.CellAlignCenter,
		RowStyleName:     "Normal", // Named style
		RowStyle: godocx.CellStyle{ // Plus direct formatting
			FontSize: 20, // 10pt
		},
		AlternateRowColor: "DEEBF7",
		BorderStyle:       godocx.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert mixed styles table: %v", err)
	}

	err = updater.AddText("↑ Combines 'Heading3' and 'Normal' styles with custom colors and formatting.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 3: Using Normal style (most common)
	err = updater.AddHeading(2, "3. Using Normal Style (Default)", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Name"},
			{Title: "Department"},
			{Title: "Status"},
		},
		Rows: [][]string{
			{"Alice Johnson", "Engineering", "Active"},
			{"Bob Smith", "Marketing", "Active"},
			{"Carol White", "Sales", "Active"},
		},
		HeaderStyleName:   "Heading1", // Heading style for header
		RowStyleName:      "Normal",   // Most commonly used
		HeaderBold:        true,
		HeaderBackground:  "70AD47",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "E2EFD9",
		BorderStyle:       godocx.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert normal style table: %v", err)
	}

	err = updater.AddText("↑ 'Normal' is Word's default paragraph style and most commonly used.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 4: Direct formatting only (no named styles)
	err = updater.AddHeading(2, "4. Direct Formatting Only (No Named Styles)", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Item"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"Direct Bold", "No style reference"},
			{"Custom Size", "Explicit formatting"},
		},
		// No HeaderStyleName or RowStyleName - uses direct formatting only
		HeaderBold:       true,
		HeaderBackground: "C65911",
		HeaderAlignment:  godocx.CellAlignCenter,
		RowStyle: godocx.CellStyle{
			FontSize:  22, // 11pt
			FontColor: "1F4E78",
		},
		AlternateRowColor: "FCE4D6",
		BorderStyle:       godocx.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert direct formatting table: %v", err)
	}

	err = updater.AddText("↑ No named styles - all formatting is applied directly.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 5: Custom style names (if defined in template)
	err = updater.AddHeading(2, "5. Custom Styles (If Defined in Template)", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Style Type"},
			{Title: "Example"},
		},
		Rows: [][]string{
			{"Custom Header", "CompanyHeader (if defined)"},
			{"Custom Body", "CompanyBody (if defined)"},
			{"Table Style", "Can also use table-specific styles"},
		},
		HeaderStyleName:   "CompanyHeader", // Custom style (if exists)
		RowStyleName:      "CompanyBody",   // Custom style (if exists)
		HeaderBackground:  "7030A0",
		AlternateRowColor: "E9D8F4",
		BorderStyle:       godocx.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert custom styles table: %v", err)
	}

	err = updater.AddText("↑ If your template defines custom styles, you can reference them by name.", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Summary table
	err = updater.AddHeading(2, "Common Word Styles", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add summary heading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Style Name", Alignment: godocx.CellAlignLeft},
			{Title: "Description", Alignment: godocx.CellAlignLeft},
			{Title: "Common Use", Alignment: godocx.CellAlignLeft},
		},
		Rows: [][]string{
			{"Normal", "Default paragraph style", "Data rows, general content"},
			{"Heading1", "Top-level heading", "Table headers, section titles"},
			{"Heading2", "Second-level heading", "Table headers, subsections"},
			{"Heading3", "Third-level heading", "Table headers, minor sections"},
			{"BodyText", "Body text paragraph", "Data rows, content text"},
			{"Title", "Document title style", "Special headers"},
			{"Subtitle", "Document subtitle", "Secondary headers"},
			{"IntenseQuote", "Emphasized quote", "Highlighted content"},
		},
		HeaderStyleName:   "Heading2",
		RowStyleName:      "Normal",
		HeaderBold:        true,
		HeaderBackground:  "44546A",
		HeaderAlignment:   godocx.CellAlignCenter,
		AlternateRowColor: "D6DCE4",
		BorderStyle:       godocx.BorderSingle,
		RowStyle: godocx.CellStyle{
			FontSize: 18, // 9pt
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert summary table: %v", err)
	}

	// Add benefits section
	err = updater.AddHeading(2, "Benefits of Named Styles", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add benefits heading: %v", err)
	}

	benefits := []godocx.ParagraphOptions{
		{
			Text:     "✓ Consistency: All tables using the same style name will update together",
			Style:    godocx.StyleNormal,
			Position: godocx.PositionEnd,
		},
		{
			Text:     "✓ Template-based: Styles can be customized in the template document",
			Style:    godocx.StyleNormal,
			Position: godocx.PositionEnd,
		},
		{
			Text:     "✓ Corporate branding: Use company-defined styles for consistent branding",
			Style:    godocx.StyleNormal,
			Position: godocx.PositionEnd,
		},
		{
			Text:     "✓ Flexible: Mix named styles with direct formatting as needed",
			Style:    godocx.StyleNormal,
			Position: godocx.PositionEnd,
		},
		{
			Text:     "✓ Easy updates: Change the style definition once, affects all instances",
			Style:    godocx.StyleNormal,
			Position: godocx.PositionEnd,
		},
	}

	if err := updater.InsertParagraphs(benefits); err != nil {
		log.Fatalf("Failed to insert benefits: %v", err)
	}

	// Save the document
	outputPath := "./outputs/table_named_styles_examples.docx"
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("✅ SUCCESS!")
	fmt.Printf("📄 Output saved to: %s\n", outputPath)
	fmt.Println("\nCreated examples:")
	fmt.Println("  • Built-in Word styles (Heading2, BodyText)")
	fmt.Println("  • Mixed named styles + direct formatting")
	fmt.Println("  • Normal style (most common)")
	fmt.Println("  • Direct formatting only (no styles)")
	fmt.Println("  • Custom styles (template-defined)")
	fmt.Println("  • Common Word styles reference table")
	fmt.Println("  • Benefits of using named styles")
}
