package docxupdater_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	docxupdater "github.com/falcomza/docx-update/src"
)

// TestSetCoreProperties tests setting core document properties
func TestSetCoreProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_core_properties.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set core properties
	props := docxupdater.CoreProperties{
		Title:          "Test Document Title",
		Subject:        "Document Testing",
		Creator:        "John Doe",
		Keywords:       "test, docx, properties",
		Description:    "This is a test document for properties",
		Category:       "Testing",
		LastModifiedBy: "Jane Smith",
		Revision:       "2",
		Created:        time.Date(2026, 1, 1, 10, 0, 0, 0, time.UTC),
		Modified:       time.Date(2026, 2, 16, 14, 30, 0, 0, time.UTC),
	}

	err = u.SetCoreProperties(props)
	if err != nil {
		t.Fatalf("Failed to set core properties: %v", err)
	}

	// Save document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}

	t.Log("Core properties set successfully")
}

// TestGetCoreProperties tests retrieving core properties
func TestGetCoreProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set properties first
	originalProps := docxupdater.CoreProperties{
		Title:       "Get Test Title",
		Creator:     "Test Author",
		Keywords:    "read, test",
		Description: "Test description",
	}

	err = u.SetCoreProperties(originalProps)
	if err != nil {
		t.Fatalf("Failed to set core properties: %v", err)
	}

	// Get properties
	props, err := u.GetCoreProperties()
	if err != nil {
		t.Fatalf("Failed to get core properties: %v", err)
	}

	// Verify values
	if props.Title != originalProps.Title {
		t.Errorf("Title mismatch: expected %q, got %q", originalProps.Title, props.Title)
	}
	if props.Creator != originalProps.Creator {
		t.Errorf("Creator mismatch: expected %q, got %q", originalProps.Creator, props.Creator)
	}
	if props.Keywords != originalProps.Keywords {
		t.Errorf("Keywords mismatch: expected %q, got %q", originalProps.Keywords, props.Keywords)
	}
	if props.Description != originalProps.Description {
		t.Errorf("Description mismatch: expected %q, got %q", originalProps.Description, props.Description)
	}

	t.Logf("Retrieved properties: Title=%q, Creator=%q", props.Title, props.Creator)
}

// TestSetAppProperties tests setting application properties
func TestSetAppProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_app_properties.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set app properties
	appProps := docxupdater.AppProperties{
		Company:     "TechVenture Inc",
		Manager:     "Alice Johnson",
		Application: "Microsoft Word",
		AppVersion:  "16.0000",
	}

	err = u.SetAppProperties(appProps)
	if err != nil {
		t.Fatalf("Failed to set app properties: %v", err)
	}

	// Save document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}

	t.Log("App properties set successfully")
}

// TestSetCustomProperties tests setting custom properties
func TestSetCustomProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_custom_properties.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set custom properties with various types
	customProps := []docxupdater.CustomProperty{
		{Name: "ProjectName", Value: "Alpha Project", Type: "lpwstr"},
		{Name: "Version", Value: 2, Type: "i4"},
		{Name: "Budget", Value: 150000.50, Type: "r8"},
		{Name: "IsApproved", Value: true, Type: "bool"},
		{Name: "Department", Value: "Engineering"},
		{Name: "Priority", Value: 1},
		{Name: "Confidence", Value: 0.95},
	}

	err = u.SetCustomProperties(customProps)
	if err != nil {
		t.Fatalf("Failed to set custom properties: %v", err)
	}

	// Save document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}

	// Verify custom.xml was created
	customXMLPath := filepath.Join(u.TempDir(), "docProps", "custom.xml")
	if _, err := os.Stat(customXMLPath); os.IsNotExist(err) {
		t.Error("custom.xml was not created")
	}

	t.Logf("Set %d custom properties successfully", len(customProps))
}

// TestCustomPropertiesInferTypes tests type inference for custom properties
func TestCustomPropertiesInferTypes(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_custom_inferred_types.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Custom properties without explicit type (type will be inferred)
	customProps := []docxupdater.CustomProperty{
		{Name: "StringProp", Value: "Hello World"},
		{Name: "IntProp", Value: 42},
		{Name: "FloatProp", Value: 3.14159},
		{Name: "BoolProp", Value: false},
		{Name: "DateProp", Value: time.Date(2026, 2, 16, 12, 0, 0, 0, time.UTC)},
	}

	err = u.SetCustomProperties(customProps)
	if err != nil {
		t.Fatalf("Failed to set custom properties: %v", err)
	}

	// Save and verify
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	t.Log("Custom properties with inferred types set successfully")
}

// TestUpdateExistingProperties tests updating existing properties
func TestUpdateExistingProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set initial properties
	props1 := docxupdater.CoreProperties{
		Title:   "Original Title",
		Creator: "Original Author",
	}
	err = u.SetCoreProperties(props1)
	if err != nil {
		t.Fatalf("Failed to set initial properties: %v", err)
	}

	// Update properties
	props2 := docxupdater.CoreProperties{
		Title:       "Updated Title",
		Creator:     "Updated Author",
		Description: "Added description",
	}
	err = u.SetCoreProperties(props2)
	if err != nil {
		t.Fatalf("Failed to update properties: %v", err)
	}

	// Verify updated values
	retrieved, err := u.GetCoreProperties()
	if err != nil {
		t.Fatalf("Failed to get properties: %v", err)
	}

	if retrieved.Title != props2.Title {
		t.Errorf("Title not updated: expected %q, got %q", props2.Title, retrieved.Title)
	}
	if retrieved.Creator != props2.Creator {
		t.Errorf("Creator not updated: expected %q, got %q", props2.Creator, retrieved.Creator)
	}
	if retrieved.Description != props2.Description {
		t.Errorf("Description not added: expected %q, got %q", props2.Description, retrieved.Description)
	}

	t.Log("Properties updated successfully")
}

// TestEmptyProperties tests handling of empty property values
func TestEmptyProperties(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set properties with some empty values
	props := docxupdater.CoreProperties{
		Title:   "Non-empty Title",
		Creator: "", // Empty
		Subject: "Non-empty Subject",
		// Description left unset
	}

	err = u.SetCoreProperties(props)
	if err != nil {
		t.Fatalf("Failed to set properties: %v", err)
	}

	// Should succeed without error
	t.Log("Empty properties handled successfully")
}

// TestPropertiesXMLEscaping tests that special characters are properly escaped
func TestPropertiesXMLEscaping(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_properties_escaping.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Properties with special XML characters
	props := docxupdater.CoreProperties{
		Title:       "Test & Demo <Document>",
		Description: "Testing \"quotes\" and 'apostrophes'",
		Keywords:    "xml, <tags>, & symbols",
	}

	err = u.SetCoreProperties(props)
	if err != nil {
		t.Fatalf("Failed to set properties with special chars: %v", err)
	}

	// Save and verify it doesn't break XML
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify the document can be opened again
	u2, err := docxupdater.New(outputPath)
	if err != nil {
		t.Fatalf("Failed to reopen document with escaped properties: %v", err)
	}
	defer u2.Cleanup()

	t.Log("XML escaping in properties working correctly")
}

// TestCompletePropertiesWorkflow tests setting all property types together
func TestCompletePropertiesWorkflow(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_complete_properties.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	// Ensure output directory exists
	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil {
		t.Fatalf("Failed to create output directory: %v", err)
	}

	t.Log("Step 1: Opening document...")
	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	t.Log("Step 2: Setting core properties...")
	coreProps := docxupdater.CoreProperties{
		Title:          "Annual Financial Report 2026",
		Subject:        "Q4 Financial Results",
		Creator:        "Finance Department",
		Keywords:       "finance, annual, report, 2026, Q4",
		Description:    "Comprehensive financial report for fiscal year 2026",
		Category:       "Financial Reports",
		LastModifiedBy: "CFO Office",
		Revision:       "3",
		Created:        time.Date(2026, 1, 15, 9, 0, 0, 0, time.UTC),
		Modified:       time.Now(),
	}
	err = u.SetCoreProperties(coreProps)
	if err != nil {
		t.Fatalf("Failed to set core properties: %v", err)
	}

	t.Log("Step 3: Setting app properties...")
	appProps := docxupdater.AppProperties{
		Company:     "Global Finance Corp",
		Manager:     "Sarah Williams",
		Application: "Microsoft Word",
		AppVersion:  "16.0000",
	}
	err = u.SetAppProperties(appProps)
	if err != nil {
		t.Fatalf("Failed to set app properties: %v", err)
	}

	t.Log("Step 4: Setting custom properties...")
	customProps := []docxupdater.CustomProperty{
		{Name: "FiscalYear", Value: 2026},
		{Name: "Quarter", Value: "Q4"},
		{Name: "Revenue", Value: 12500000.75},
		{Name: "IsAudited", Value: true},
		{Name: "ProjectCode", Value: "FIN-2026-Q4"},
		{Name: "DepartmentID", Value: 101},
		{Name: "ConfidentialityLevel", Value: "High"},
		{Name: "ApprovalDate", Value: time.Date(2026, 2, 15, 0, 0, 0, 0, time.UTC)},
	}
	err = u.SetCustomProperties(customProps)
	if err != nil {
		t.Fatalf("Failed to set custom properties: %v", err)
	}

	t.Log("Step 5: Saving document...")
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify output
	info, err := os.Stat(outputPath)
	if err != nil {
		t.Fatalf("Failed to stat output file: %v", err)
	}

	t.Logf("✓ Document created successfully: %s", outputPath)
	t.Logf("  File size: %d bytes", info.Size())
	t.Log("  All property types set!")
}

// TestPropertiesWithExistingContent tests that properties work alongside other modifications
func TestPropertiesWithExistingContent(t *testing.T) {
	templatePath := "../templates/docx_template.docx"
	outputPath := "../outputs/test_properties_with_content.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Add some content
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This document has both content and properties.",
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("Failed to insert paragraph: %v", err)
	}

	// Set properties
	props := docxupdater.CoreProperties{
		Title:   "Document with Content and Properties",
		Creator: "Integration Test",
	}
	err = u.SetCoreProperties(props)
	if err != nil {
		t.Fatalf("Failed to set properties: %v", err)
	}

	// Save
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	t.Log("Document with content and properties created successfully")
}

// TestPropertiesVerifyFiles verifies that property files are created correctly
func TestPropertiesVerifyFiles(t *testing.T) {
	templatePath := "../templates/docx_template.docx"

	if _, err := os.Stat(templatePath); os.IsNotExist(err) {
		t.Skipf("Template file not found: %s", templatePath)
	}

	u, err := docxupdater.New(templatePath)
	if err != nil {
		t.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// Set all property types
	u.SetCoreProperties(docxupdater.CoreProperties{Title: "Test"})
	u.SetAppProperties(docxupdater.AppProperties{Company: "TestCo"})
	u.SetCustomProperties([]docxupdater.CustomProperty{{Name: "Test", Value: "Value"}})

	// Check that files exist
	coreXML := filepath.Join(u.TempDir(), "docProps", "core.xml")
	if _, err := os.Stat(coreXML); os.IsNotExist(err) {
		t.Error("core.xml not found")
	} else {
		content, _ := os.ReadFile(coreXML)
		if !strings.Contains(string(content), "Test") {
			t.Error("core.xml doesn't contain expected title")
		}
		t.Log("✓ core.xml verified")
	}

	appXML := filepath.Join(u.TempDir(), "docProps", "app.xml")
	if _, err := os.Stat(appXML); os.IsNotExist(err) {
		t.Error("app.xml not found")
	} else {
		content, _ := os.ReadFile(appXML)
		if !strings.Contains(string(content), "TestCo") {
			t.Error("app.xml doesn't contain expected company")
		}
		t.Log("✓ app.xml verified")
	}

	customXML := filepath.Join(u.TempDir(), "docProps", "custom.xml")
	if _, err := os.Stat(customXML); os.IsNotExist(err) {
		t.Error("custom.xml not found")
	} else {
		content, _ := os.ReadFile(customXML)
		if !strings.Contains(string(content), "Test") {
			t.Error("custom.xml doesn't contain expected property")
		}
		t.Log("✓ custom.xml verified")
	}
}
