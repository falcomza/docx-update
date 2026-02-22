package godocx_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

func TestInsertPageBreakAtEnd(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a page break at the end
	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertPageBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify the page break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Error("Page break not found in document.xml")
	}
}

func TestInsertPageBreakAtBeginning(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a page break at the beginning
	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertPageBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the page break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Error("Page break not found in document.xml")
	}
}

func TestInsertPageBreakAfterText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add anchor text
	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "First section content",
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Insert page break after the text
	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionAfterText,
		Anchor:   "First section",
	})
	if err != nil {
		t.Fatalf("InsertPageBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the page break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Error("Page break not found in document.xml")
	}
}

func TestInsertPageBreakBeforeText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add anchor text
	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Second section content",
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	// Insert page break before the text
	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionBeforeText,
		Anchor:   "Second section",
	})
	if err != nil {
		t.Fatalf("InsertPageBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the page break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Error("Page break not found in document.xml")
	}
}

func TestInsertSectionBreakNextPage(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a section break (next page)
	err = u.InsertSectionBreak(godocx.BreakOptions{
		Position:    godocx.PositionEnd,
		SectionType: godocx.SectionBreakNextPage,
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the section break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:sectPr>") {
		t.Error("Section break not found in document.xml")
	}
	if !strings.Contains(docXML, `w:val="nextPage"`) {
		t.Error("Section break type not set correctly")
	}
}

func TestInsertSectionBreakContinuous(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a continuous section break
	err = u.InsertSectionBreak(godocx.BreakOptions{
		Position:    godocx.PositionEnd,
		SectionType: godocx.SectionBreakContinuous,
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the section break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:sectPr>") {
		t.Error("Section break not found in document.xml")
	}
	if !strings.Contains(docXML, `w:val="continuous"`) {
		t.Error("Continuous section break type not set correctly")
	}
}

func TestInsertSectionBreakEvenPage(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert an even page section break
	err = u.InsertSectionBreak(godocx.BreakOptions{
		Position:    godocx.PositionEnd,
		SectionType: godocx.SectionBreakEvenPage,
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the section break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:val="evenPage"`) {
		t.Error("Even page section break type not set correctly")
	}
}

func TestInsertSectionBreakOddPage(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert an odd page section break
	err = u.InsertSectionBreak(godocx.BreakOptions{
		Position:    godocx.PositionEnd,
		SectionType: godocx.SectionBreakOddPage,
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify the section break was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:val="oddPage"`) {
		t.Error("Odd page section break type not set correctly")
	}
}

func TestInsertMultiplePageBreaks(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add some content and page breaks
	u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Page 1 content",
		Position: godocx.PositionEnd,
	})

	u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionEnd,
	})

	u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Page 2 content",
		Position: godocx.PositionEnd,
	})

	u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionEnd,
	})

	u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Page 3 content",
		Position: godocx.PositionEnd,
	})

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify multiple page breaks were added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	count := strings.Count(docXML, `<w:br w:type="page"/>`)
	if count < 2 {
		t.Errorf("Expected at least 2 page breaks, found %d", count)
	}
}

func TestInsertSectionBreakDefaultType(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert section break without specifying type (should default to nextPage)
	err = u.InsertSectionBreak(godocx.BreakOptions{
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertSectionBreak failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify default section break type is nextPage
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:val="nextPage"`) {
		t.Error("Default section break type not set to nextPage")
	}
}

func TestInsertPageBreakInvalidAnchor(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := godocx.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Try to insert page break with non-existent anchor
	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionAfterText,
		Anchor:   "nonexistent text",
	})
	if err == nil {
		t.Error("Expected error for non-existent anchor, got nil")
	}
}
