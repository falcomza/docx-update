package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/docx-update"
)

func main() {
	// Use empty_template.docx for clean output, or docx_template.docx for existing template
	templatePath := "./templates/empty_template.docx"
	outputPath := "./outputs/table_orientation_demo.docx"

	fmt.Println("Opening template:", templatePath)
	fmt.Println("NOTE: Tables will be inserted after any existing content in the template")
	updater, err := docx.New(templatePath)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Step 1: Add initial content in portrait orientation
	fmt.Println("Adding initial content in portrait...")
	err = updater.AddText("Document Introduction", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}

	err = updater.AddText("This document demonstrates dynamic page orientation changes. The following section contains a wide table that requires landscape orientation for proper display.", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add introduction: %v", err)
	}

	// Step 2: Add heading (on portrait page)
	err = updater.AddText("Wide Data Table (Landscape View)", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add landscape section heading: %v", err)
	}

	// Step 3: End portrait section (intro + heading stay in portrait)
	fmt.Println("Ending portrait section...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutLetterPortrait(),
	})
	if err != nil {
		log.Fatalf("Failed to end portrait section: %v", err)
	}

	// Step 4: Insert table (new section, will end as landscape)
	fmt.Println("Inserting employee table...")
	err = updater.InsertTable(docx.TableOptions{
		Position: docx.PositionEnd,
		Columns: []docx.ColumnDefinition{
			{Title: "Employee ID", Alignment: docx.CellAlignLeft},
			{Title: "Full Name", Alignment: docx.CellAlignLeft},
			{Title: "Department", Alignment: docx.CellAlignLeft},
			{Title: "Position", Alignment: docx.CellAlignLeft},
			{Title: "Location", Alignment: docx.CellAlignLeft},
			{Title: "Hire Date", Alignment: docx.CellAlignCenter},
			{Title: "Salary", Alignment: docx.CellAlignRight},
			{Title: "Performance", Alignment: docx.CellAlignCenter},
		},
		Rows: [][]string{
			{"EMP001", "John Smith", "Engineering", "Senior Developer", "New York", "2020-01-15", "$95,000", "Excellent"},
			{"EMP002", "Jane Doe", "Marketing", "Marketing Manager", "Los Angeles", "2019-06-20", "$87,500", "Very Good"},
			{"EMP003", "Bob Johnson", "Sales", "Sales Director", "Chicago", "2018-03-10", "$102,000", "Excellent"},
			{"EMP004", "Alice Williams", "Engineering", "Tech Lead", "San Francisco", "2021-09-01", "$110,000", "Outstanding"},
			{"EMP005", "Charlie Brown", "Finance", "Financial Analyst", "Boston", "2022-02-14", "$72,000", "Good"},
			{"EMP006", "Diana Prince", "HR", "HR Manager", "Seattle", "2017-11-30", "$83,000", "Very Good"},
			{"EMP007", "Evan Davis", "Engineering", "Junior Developer", "Austin", "2023-05-22", "$65,000", "Good"},
			{"EMP008", "Fiona Green", "Operations", "Operations Lead", "Denver", "2020-08-17", "$91,000", "Excellent"},
		},
		HeaderBold:        true,
		HeaderBackground:  "2E75B5",
		HeaderAlignment:   docx.CellAlignCenter,
		AlternateRowColor: "E7E6E6",
		BorderStyle:       docx.BorderSingle,
		BorderSize:        6,
		BorderColor:       "2E75B5",
		TableAlignment:    docx.AlignCenter,
		RepeatHeader:      true,
	})
	if err != nil {
		log.Fatalf("Failed to insert table: %v", err)
	}

	// Step 5: End landscape section (table is in landscape)
	fmt.Println("Ending landscape section...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutLetterLandscape(),
	})
	if err != nil {
		log.Fatalf("Failed to end landscape section: %v", err)
	}

	// Step 6: Add analysis text (new portrait section)
	fmt.Println("Adding conclusion in portrait...")
	err = updater.AddText("Analysis and Summary", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add conclusion heading: %v", err)
	}

	err = updater.AddText("The employee data table above shows our current workforce across multiple departments and locations. All employees demonstrate strong performance metrics with salary ranges appropriate for their roles and experience levels.", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add conclusion text: %v", err)
	}

	err = updater.AddText("", docx.PositionEnd) // Empty line
	err = updater.AddText("Appendix: International Format Example", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add appendix heading: %v", err)
	}

	err = updater.AddText("Quarterly Sales Report (A4 Landscape)", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add quarterly heading: %v", err)
	}

	// Step 7: End portrait section (analysis + appendix headings are in portrait)
	fmt.Println("Ending portrait section...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutLetterPortrait(),
	})
	if err != nil {
		log.Fatalf("Failed to end portrait section: %v", err)
	}

	// Step 8: Insert quarterly table (new section, will end as A4 landscape)
	fmt.Println("Inserting quarterly sales table...")
	err = updater.InsertTable(docx.TableOptions{
		Position: docx.PositionEnd,
		Columns: []docx.ColumnDefinition{
			{Title: "Region", Alignment: docx.CellAlignLeft},
			{Title: "Q1 Revenue", Alignment: docx.CellAlignRight},
			{Title: "Q2 Revenue", Alignment: docx.CellAlignRight},
			{Title: "Q3 Revenue", Alignment: docx.CellAlignRight},
			{Title: "Q4 Revenue", Alignment: docx.CellAlignRight},
			{Title: "Total", Alignment: docx.CellAlignRight},
			{Title: "Growth %", Alignment: docx.CellAlignCenter},
		},
		Rows: [][]string{
			{"North America", "$2,450,000", "$2,680,000", "$2,920,000", "$3,100,000", "$11,150,000", "+8.5%"},
			{"Europe", "$1,890,000", "$2,010,000", "$2,150,000", "$2,340,000", "$8,390,000", "+7.2%"},
			{"Asia Pacific", "$3,120,000", "$3,450,000", "$3,780,000", "$4,020,000", "$14,370,000", "+9.1%"},
			{"Latin America", "$1,230,000", "$1,320,000", "$1,410,000", "$1,550,000", "$5,510,000", "+6.8%"},
			{"Middle East", "$890,000", "$950,000", "$1,020,000", "$1,100,000", "$3,960,000", "+7.5%"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   docx.CellAlignCenter,
		AlternateRowColor: "DEEBF7",
		BorderStyle:       docx.BorderDouble,
		BorderSize:        6,
		BorderColor:       "4472C4",
		TableAlignment:    docx.AlignCenter,
		RepeatHeader:      true,
	})
	if err != nil {
		log.Fatalf("Failed to insert quarterly table: %v", err)
	}

	// Step 9: End A4 landscape section (quarterly table is in A4 landscape)
	fmt.Println("Ending A4 landscape section...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutA4Landscape(),
	})
	if err != nil {
		log.Fatalf("Failed to end A4 landscape section: %v", err)
	}

	// Step 10: Add final note (new portrait section)
	err = updater.AddText("End of Report", docx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add end note: %v", err)
	}

	// Step 11: End final section as portrait (template has no sectPr, so we add it)
	fmt.Println("Ending final section...")
	err = updater.InsertSectionBreak(docx.BreakOptions{
		Position:    docx.PositionEnd,
		SectionType: docx.SectionBreakNextPage,
		PageLayout:  docx.PageLayoutLetterPortrait(),
	})
	if err != nil {
		log.Fatalf("Failed to end final section: %v", err)
	}

	// Step 12: Save the document
	fmt.Println("Saving document to:", outputPath)
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("\n✓ Document created successfully!")
	fmt.Println("\nDocument Structure (5 sections):")
	fmt.Println("  1. PORTRAIT - Introduction + Employee table heading")
	fmt.Println("  2. LANDSCAPE (Letter) - Employee table (8 columns)")
	fmt.Println("  3. PORTRAIT - Analysis + Appendix + Quarterly table heading")
	fmt.Println("  4. LANDSCAPE (A4) - Quarterly sales table (7 columns)")
	fmt.Println("  5. PORTRAIT - End of Report")
	fmt.Println("\nPattern:")
	fmt.Println("  Portrait text → Landscape table → Portrait text → Landscape table → Portrait conclusion")
	fmt.Println("\nOutput:", outputPath)
}
