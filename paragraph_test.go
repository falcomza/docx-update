package docxupdater_test

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	docxupdater "github.com/falcomza/docx-update"
)

func TestInsertParagraphAtEnd(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Insert a paragraph at the end
	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "This is a test paragraph",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Read document.xml and verify the paragraph was added
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "This is a test paragraph") {
		t.Error("Paragraph text not found in document.xml")
	}
}

func TestInsertParagraphAtBeginning(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Beginning paragraph",
		Style:    docxupdater.StyleHeading1,
		Position: docxupdater.PositionBeginning,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Beginning paragraph") {
		t.Error("Paragraph text not found in document.xml")
	}
	if !strings.Contains(docXML, "Heading1") {
		t.Error("Heading1 style not found in document.xml")
	}
}

func TestInsertParagraphWithFormatting(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:      "Bold and italic text",
		Style:     docxupdater.StyleNormal,
		Position:  docxupdater.PositionEnd,
		Bold:      true,
		Italic:    true,
		Underline: true,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:b/>") {
		t.Error("Bold formatting not found")
	}
	if !strings.Contains(docXML, "<w:i/>") {
		t.Error("Italic formatting not found")
	}
	if !strings.Contains(docXML, "<w:u") {
		t.Error("Underline formatting not found")
	}
}

func TestInsertParagraphWithAlignment(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:      "Centered paragraph",
		Style:     docxupdater.StyleNormal,
		Alignment: docxupdater.ParagraphAlignCenter,
		Position:  docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:      "Justified paragraph",
		Style:     docxupdater.StyleNormal,
		Alignment: docxupdater.ParagraphAlignJustify,
		Position:  docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:jc w:val="center"/>`) {
		t.Error("Center alignment not found")
	}
	if !strings.Contains(docXML, `<w:jc w:val="both"/>`) {
		t.Error("Justify alignment not found")
	}
}

func TestInsertParagraphAtEndBeforeSectPr(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithSectPr(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Inserted before sectPr",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	insertedPos := strings.Index(docXML, "Inserted before sectPr")
	sectPrPos := strings.Index(docXML, "<w:sectPr")
	if insertedPos == -1 {
		t.Fatal("inserted paragraph text not found")
	}
	if sectPrPos == -1 {
		t.Fatal("sectPr not found")
	}
	if insertedPos > sectPrPos {
		t.Error("paragraph was inserted after sectPr; expected before sectPr")
	}
}

func TestInsertParagraphWithLineBreaksAndTabs(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Line 1\nLine 2\tTabbed",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "<w:br/>") {
		t.Error("line break (w:br) not found")
	}
	if !strings.Contains(docXML, "<w:tab/>") {
		t.Error("tab (w:tab) not found")
	}
	if !strings.Contains(docXML, "Line 1") || !strings.Contains(docXML, "Line 2") || !strings.Contains(docXML, "Tabbed") {
		t.Error("expected text segments not found")
	}
}

func TestInsertParagraphAfterTextWithSplitRuns(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithSplitRunAnchor(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Inserted After",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionAfterText,
		Anchor:   "Hello World",
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	anchorEnd := strings.Index(docXML, "</w:t></w:r></w:p>")
	insertedPos := strings.Index(docXML, "Inserted After")
	if insertedPos == -1 {
		t.Fatal("inserted paragraph not found")
	}
	if anchorEnd == -1 {
		t.Fatal("anchor paragraph end not found")
	}
	if insertedPos < anchorEnd {
		t.Error("paragraph inserted before split-run anchor paragraph end")
	}
}

func TestInsertParagraphBeforeTextWithSplitRuns(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithSplitRunAnchor(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Inserted Before",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionBeforeText,
		Anchor:   "Hello World",
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	insertedPos := strings.Index(docXML, "Inserted Before")
	anchorStart := strings.Index(docXML, "<w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t> World</w:t></w:r></w:p>")
	if insertedPos == -1 {
		t.Fatal("inserted paragraph not found")
	}
	if anchorStart == -1 {
		t.Fatal("anchor paragraph start not found")
	}
	if insertedPos > anchorStart {
		t.Error("paragraph inserted after split-run anchor paragraph start")
	}
}

func TestInsertParagraphAfterTextWithNormalizedWhitespaceAnchor(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithSplitRunAnchor(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Inserted Normalized",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionAfterText,
		Anchor:   "Hello\n\tWorld",
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Inserted Normalized") {
		t.Fatal("inserted paragraph not found")
	}
}

func TestInsertParagraphAfterTextWithTabBreakAnchor(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocxWithTabBreakAnchor(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "Inserted After TabBreak",
		Style:    docxupdater.StyleNormal,
		Position: docxupdater.PositionAfterText,
		Anchor:   "Alpha Beta Gamma",
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Inserted After TabBreak") {
		t.Fatal("inserted paragraph not found")
	}
}

func TestAddHeading(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	if err := u.AddHeading(1, "Main Title", docxupdater.PositionEnd); err != nil {
		t.Fatalf("AddHeading failed: %v", err)
	}

	if err := u.AddHeading(2, "Subtitle", docxupdater.PositionEnd); err != nil {
		t.Fatalf("AddHeading failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, "Main Title") {
		t.Error("Heading 1 text not found")
	}
	if !strings.Contains(docXML, "Subtitle") {
		t.Error("Heading 2 text not found")
	}
}

func TestInsertMultipleParagraphs(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	paragraphs := []docxupdater.ParagraphOptions{
		{
			Text:     "First paragraph",
			Style:    docxupdater.StyleHeading1,
			Position: docxupdater.PositionEnd,
		},
		{
			Text:     "Second paragraph with details",
			Style:    docxupdater.StyleNormal,
			Position: docxupdater.PositionEnd,
		},
		{
			Text:     "Third paragraph conclusion",
			Style:    docxupdater.StyleNormal,
			Position: docxupdater.PositionEnd,
			Bold:     true,
		},
	}

	if err := u.InsertParagraphs(paragraphs); err != nil {
		t.Fatalf("InsertParagraphs failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")
	for _, para := range paragraphs {
		if !strings.Contains(docXML, para.Text) {
			t.Errorf("Paragraph text %q not found in document", para.Text)
		}
	}
}

func TestInsertParagraphEmptyText(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := docxupdater.New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(docxupdater.ParagraphOptions{
		Text:     "",
		Position: docxupdater.PositionEnd,
	})
	if err == nil {
		t.Error("Expected error for empty text, got nil")
	}
}

func buildFixtureDocxWithSectPr(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Existing paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureDocxWithSplitRunAnchor(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t> World</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func buildFixtureDocxWithTabBreakAnchor(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alpha</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Beta</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Gamma</w:t></w:r></w:p></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}
