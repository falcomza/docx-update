package godocx_test

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

// Test custom error types
func TestDocxError(t *testing.T) {
	err := godocx.NewChartNotFoundError(5)
	if err == nil {
		t.Fatal("Expected error, got nil")
	}

	docxErr, ok := err.(*godocx.DocxError)
	if !ok {
		t.Fatal("Expected DocxError type")
	}

	if docxErr.Code != godocx.ErrCodeChartNotFound {
		t.Errorf("Expected code %s, got %s", godocx.ErrCodeChartNotFound, docxErr.Code)
	}

	if docxErr.Context["index"] != 5 {
		t.Errorf("Expected context index 5, got %v", docxErr.Context["index"])
	}
}

// Test text replacement
func TestReplaceText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	// Create minimal test document
	if err := createMinimalDoc(inputPath, "Hello World. This is a test."); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	// Open document
	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	// Replace text
	opts := godocx.DefaultReplaceOptions()
	opts.MatchCase = false

	count, err := u.ReplaceText("World", "Universe", opts)
	if err != nil {
		t.Fatalf("Failed to replace text: %v", err)
	}

	if count != 1 {
		t.Errorf("Expected 1 replacement, got %d", count)
	}

	// Save document
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify replacement
	u2, err := godocx.New(outputPath)
	if err != nil {
		t.Fatalf("Failed to reopen document: %v", err)
	}
	defer u2.Cleanup()

	text, err := u2.GetText()
	if err != nil {
		t.Fatalf("Failed to get text: %v", err)
	}

	if !strings.Contains(text, "Universe") {
		t.Errorf("Expected text to contain 'Universe', got: %s", text)
	}

	if strings.Contains(text, "World") {
		t.Errorf("Expected text not to contain 'World', got: %s", text)
	}
}

// Test regex replacement
func TestReplaceTextRegex(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	// Create test document with numbers
	if err := createMinimalDoc(inputPath, "Call 123-456-7890 or 987-654-3210"); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	// Replace phone numbers with regex
	pattern := regexp.MustCompile(`\d{3}-\d{3}-\d{4}`)
	opts := godocx.DefaultReplaceOptions()

	count, err := u.ReplaceTextRegex(pattern, "[REDACTED]", opts)
	if err != nil {
		t.Fatalf("Failed to replace text: %v", err)
	}

	if count != 2 {
		t.Errorf("Expected 2 replacements, got %d", count)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
}

// Test read operations
func TestGetText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	testText := "This is a test paragraph."
	if err := createMinimalDoc(inputPath, testText); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	text, err := u.GetText()
	if err != nil {
		t.Fatalf("Failed to get text: %v", err)
	}

	if !strings.Contains(text, testText) {
		t.Errorf("Expected text to contain '%s', got: %s", testText, text)
	}
}

// Test find text
func TestFindText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	testText := "TODO: Implement feature. TODO: Write tests."
	if err := createMinimalDoc(inputPath, testText); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	opts := godocx.DefaultFindOptions()
	matches, err := u.FindText("TODO:", opts)
	if err != nil {
		t.Fatalf("Failed to find text: %v", err)
	}

	if len(matches) != 2 {
		t.Errorf("Expected 2 matches, got %d", len(matches))
	}

	for i, match := range matches {
		if match.Text != "TODO:" {
			t.Errorf("Match %d: expected 'TODO:', got '%s'", i, match.Text)
		}
	}
}

// Test hyperlink insertion
func TestInsertHyperlink(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := createMinimalDoc(inputPath, "Test document"); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	opts := godocx.DefaultHyperlinkOptions()
	opts.Position = godocx.PositionEnd

	err = u.InsertHyperlink("Visit Example", "https://example.com", opts)
	if err != nil {
		t.Fatalf("Failed to insert hyperlink: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify hyperlink was added
	if !fileExists(outputPath) {
		t.Error("Output file was not created")
	}
}

// Test invalid URL
func TestInvalidURL(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := createMinimalDoc(inputPath, "Test document"); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	opts := godocx.DefaultHyperlinkOptions()

	err = u.InsertHyperlink("Invalid Link", "not-a-url", opts)
	if err == nil {
		t.Fatal("Expected error for invalid URL, got nil")
	}

	docxErr, ok := err.(*godocx.DocxError)
	if !ok {
		t.Fatal("Expected DocxError type")
	}

	if docxErr.Code != godocx.ErrCodeInvalidURL {
		t.Errorf("Expected code %s, got %s", godocx.ErrCodeInvalidURL, docxErr.Code)
	}
}

// Test header creation
func TestSetHeader(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := createMinimalDoc(inputPath, "Test document"); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	content := godocx.HeaderFooterContent{
		LeftText:   "Company Name",
		CenterText: "Document Title",
		RightText:  "Page",
		PageNumber: true,
	}

	opts := godocx.DefaultHeaderOptions()

	err = u.SetHeader(content, opts)
	if err != nil {
		t.Fatalf("Failed to set header: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify header file was created
	headerPath := filepath.Join(u.TempDir(), "word", "header3.xml")
	if !fileExists(headerPath) {
		t.Error("Header file was not created")
	}
}

// Test footer creation
func TestSetFooter(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := createMinimalDoc(inputPath, "Test document"); err != nil {
		t.Fatalf("Failed to create test doc: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer u.Cleanup()

	content := godocx.HeaderFooterContent{
		CenterText:       "Page ",
		PageNumber:       true,
		PageNumberFormat: "X of Y",
	}

	opts := godocx.DefaultFooterOptions()

	err = u.SetFooter(content, opts)
	if err != nil {
		t.Fatalf("Failed to set footer: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify footer file was created
	footerPath := filepath.Join(u.TempDir(), "word", "footer3.xml")
	if !fileExists(footerPath) {
		t.Error("Footer file was not created")
	}
}

// Helper to check if file exists
func fileExists(path string) bool {
	_, err := os.Stat(path)
	return !os.IsNotExist(err)
}

// Helper to create a minimal DOCX with test text
func createMinimalDoc(path, text string) error {
	// Use the existing test helper from other tests
	u, err := godocx.New("templates/docx_template.docx")
	if err != nil {
		return err
	}
	defer u.Cleanup()

	// Insert test text
	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     text,
		Position: godocx.PositionEnd,
	})
	if err != nil {
		return err
	}

	return u.Save(path)
}
