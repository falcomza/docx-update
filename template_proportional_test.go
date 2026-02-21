package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update"
)

// TestProportionalTableWithTemplate demonstrates inserting a proportional-width table
// into the actual template document
func TestProportionalTableWithTemplate(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	if _, err := os.Stat(templatePath); err != nil {
		t.Skipf("Template not found: %s", templatePath)
	}

	outputPath := filepath.Join("outputs", "template_proportional_table.docx")

	// Open the template
	updater, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Insert a table with proportional column widths
	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Product ID"},
			{Title: "Product Description"},
			{Title: "Unit Price"},
		},
		Rows: [][]string{
			{"SKU001", "High-performance professional camera with advanced autofocus", "$1,299.99"},
			{"SKU002", "Item", "$29.99"},
			{"SKU003", "Mid-range laptop suitable for business applications", "$649.99"},
			{"SKU004", "X", "$9.99"},
		},
		ProportionalColumnWidths: true,
		TableWidth:               5000, // 100%
		TableWidthType:           docxupdater.TableWidthPercentage,
		HeaderBold:               true,
		HeaderBackground:         "4472C4",
		AlternateRowColor:         "F2F2F2",
	})
	if err != nil {
		t.Fatalf("InsertTable with proportional widths failed: %v", err)
	}

	// Save the output
	if err := updater.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the table was created
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:tbl>") {
		t.Error("Table element not found in document.xml")
	}
	if !strings.Contains(docXML, "Product ID") {
		t.Error("Table header not found in document.xml")
	}
	if !strings.Contains(docXML, "SKU001") {
		t.Error("Table data not found in document.xml")
	}

	// Extract and log grid column widths for verification
	widths := extractGridColumnWidths(docXML)
	if len(widths) != 3 {
		t.Errorf("Expected 3 columns, got %d", len(widths))
	}

	t.Logf("Proportional table column widths (twips): %v", widths)

	// Verify that columns are not all equal (one should be wider due to longer content)
	if len(widths) == 3 {
		// Description column should be much wider than ID or Price
		if widths[1] <= widths[0] || widths[1] <= widths[2] {
			t.Errorf("Expected Description column (index 1) to be widest: %v", widths)
		}
	}

	t.Logf("Successfully inserted proportional table into template: %s", outputPath)
}

// TestTableWithTemplateFixedWidth inserts a fixed-width proportional table
func TestTableWithTemplateFixedWidth(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	if _, err := os.Stat(templatePath); err != nil {
		t.Skipf("Template not found: %s", templatePath)
	}

	outputPath := filepath.Join("outputs", "template_fixed_proportional_table.docx")

	updater, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Insert table with fixed width and proportional columns
	err = updater.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "Category"},
			{Title: "Extended Description Here"},
			{Title: "Status"},
		},
		Rows: [][]string{
			{"A", "This is a much longer description field", "Active"},
			{"B", "Short", "Inactive"},
		},
		ProportionalColumnWidths: true,
		TableWidth:               8640, // 6 inches
		TableWidthType:           docxupdater.TableWidthFixed,
		HeaderBold:               true,
		HeaderAlignment:          docxupdater.CellAlignCenter,
		TableAlignment:           docxupdater.AlignCenter,
	})
	if err != nil {
		t.Fatalf("InsertTable failed: %v", err)
	}

	if err := updater.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:tbl>") {
		t.Error("Table element not found")
	}

	widths := extractGridColumnWidths(docXML)
	t.Logf("Fixed-width proportional table (6\") column widths (twips): %v", widths)

	if len(widths) == 3 {
		// Middle column should dominate due to longer content
		if widths[1] < 4000 { // Should be most of the 6" table
			t.Errorf("Middle column should be much wider: %v", widths)
		}
	}

	t.Logf("Successfully inserted fixed-width proportional table into template: %s", outputPath)
}

// TestComparisonEqualsVsProportional compares equal vs proportional sizing on template
func TestComparisonEqualVsProportionalOnTemplate(t *testing.T) {
	templatePath := "templates/docx_template.docx"
	if _, err := os.Stat(templatePath); err != nil {
		t.Skipf("Template not found: %s", templatePath)
	}

	// Test 1: Equal width (default)
	outputEqualPath := filepath.Join("outputs", "template_equal_width_table.docx")
	updater1, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template for equal test: %v", err)
	}
	defer updater1.Cleanup()

	err = updater1.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "A"},
			{Title: "B"},
			{Title: "C"},
		},
		Rows: [][]string{
			{"Short", "Very long content that should get equal space anyway", "Med"},
		},
		HeaderBold: true,
	})
	if err != nil {
		t.Fatalf("Equal width table failed: %v", err)
	}

	updater1.Save(outputEqualPath)
	docXMLEqual := readZipEntry(t, outputEqualPath, "word/document.xml")
	widthsEqual := extractGridColumnWidths(docXMLEqual)

	// Test 2: Proportional width
	outputPropPath := filepath.Join("outputs", "template_proportional_width_table.docx")
	updater2, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template for proportional test: %v", err)
	}
	defer updater2.Cleanup()

	err = updater2.InsertTable(docxupdater.TableOptions{
		Position: docxupdater.PositionEnd,
		Columns: []docxupdater.ColumnDefinition{
			{Title: "A"},
			{Title: "B"},
			{Title: "C"},
		},
		Rows: [][]string{
			{"Short", "Very long content that should get equal space anyway", "Med"},
		},
		ProportionalColumnWidths: true,
		HeaderBold:               true,
	})
	if err != nil {
		t.Fatalf("Proportional width table failed: %v", err)
	}

	updater2.Save(outputPropPath)
	docXMLProp := readZipEntry(t, outputPropPath, "word/document.xml")
	widthsProp := extractGridColumnWidths(docXMLProp)

	t.Logf("Equal width distribution:       %v", widthsEqual)
	t.Logf("Proportional width distribution: %v", widthsProp)

	// Verify they're different
	if len(widthsEqual) == 3 && len(widthsProp) == 3 {
		allEqual := widthsEqual[0] == widthsEqual[1] && widthsEqual[1] == widthsEqual[2]
		notAllEqual := !(widthsProp[0] == widthsProp[1] && widthsProp[1] == widthsProp[2])

		if !allEqual {
			t.Error("Equal width distribution should have all equal widths")
		}
		if !notAllEqual {
			t.Error("Proportional width distribution should NOT have all equal widths")
		}
		if widthsProp[1] <= widthsProp[0] || widthsProp[1] <= widthsProp[2] {
			t.Error("Column B (longest content) should be widest in proportional distribution")
		}
	}
}
