package docxupdater

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestInsertBulletList(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add bullet list
	err = u.AddBulletItem("First bullet point", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	err = u.AddBulletItem("Second bullet point", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	err = u.AddBulletItem("Nested bullet", 1, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletItem failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify numbering.xml was created
	numberingXML := readZipEntry(t, outputPath, "word/numbering.xml")
	if numberingXML == "" {
		t.Fatal("numbering.xml not found")
	}

	// Verify it contains bullet list definition
	if !strings.Contains(numberingXML, `w:numId="1"`) {
		t.Error("Bullet list numbering ID not found")
	}
	if !strings.Contains(numberingXML, `w:numFmt w:val="bullet"`) {
		t.Error("Bullet format not found")
	}

	// Verify document.xml has numPr elements
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:numPr>`) {
		t.Error("numPr element not found in document")
	}
	if !strings.Contains(docXML, `<w:numId w:val="1"/>`) {
		t.Error("numId for bullet list not found")
	}
	if !strings.Contains(docXML, `<w:ilvl w:val="0"/>`) {
		t.Error("ilvl level 0 not found")
	}
	if !strings.Contains(docXML, `<w:ilvl w:val="1"/>`) {
		t.Error("ilvl level 1 not found")
	}

	// Verify text content
	if !strings.Contains(docXML, "First bullet point") {
		t.Error("First bullet text not found")
	}
	if !strings.Contains(docXML, "Second bullet point") {
		t.Error("Second bullet text not found")
	}
	if !strings.Contains(docXML, "Nested bullet") {
		t.Error("Nested bullet text not found")
	}
}

func TestInsertNumberedList(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add numbered list
	err = u.AddNumberedItem("First item", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedItem failed: %v", err)
	}

	err = u.AddNumberedItem("Second item", 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedItem failed: %v", err)
	}

	err = u.AddNumberedItem("Nested item", 1, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedItem failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify numbering.xml was created
	numberingXML := readZipEntry(t, outputPath, "word/numbering.xml")
	if numberingXML == "" {
		t.Fatal("numbering.xml not found")
	}

	// Verify it contains numbered list definition
	if !strings.Contains(numberingXML, `w:numId="2"`) {
		t.Error("Numbered list numbering ID not found")
	}
	if !strings.Contains(numberingXML, `w:numFmt w:val="decimal"`) {
		t.Error("Decimal number format not found")
	}

	// Verify document.xml has numPr elements
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `<w:numId w:val="2"/>`) {
		t.Error("numId for numbered list not found")
	}

	// Verify text content
	if !strings.Contains(docXML, "First item") {
		t.Error("First item text not found")
	}
	if !strings.Contains(docXML, "Second item") {
		t.Error("Second item text not found")
	}
	if !strings.Contains(docXML, "Nested item") {
		t.Error("Nested item text not found")
	}
}

func TestAddBulletList(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add bullet list in batch
	items := []string{
		"Introduction to the topic",
		"Key benefits and advantages",
		"Implementation details",
		"Best practices",
	}

	err = u.AddBulletList(items, 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddBulletList failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify all items are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	for _, item := range items {
		if !strings.Contains(docXML, item) {
			t.Errorf("List item not found: %s", item)
		}
	}

	// Count numPr occurrences (should be 4)
	count := strings.Count(docXML, `<w:numPr>`)
	if count != 4 {
		t.Errorf("Expected 4 numPr elements, got %d", count)
	}
}

func TestAddNumberedList(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add numbered list in batch
	items := []string{
		"Follow step one carefully",
		"Complete step two thoroughly",
		"Verify step three results",
	}

	err = u.AddNumberedList(items, 0, PositionEnd)
	if err != nil {
		t.Fatalf("AddNumberedList failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify all items are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	for _, item := range items {
		if !strings.Contains(docXML, item) {
			t.Errorf("List item not found: %s", item)
		}
	}

	// Verify numbered list numbering ID
	if !strings.Contains(docXML, `<w:numId w:val="2"/>`) {
		t.Error("Numbered list numId not found")
	}
}

func TestMultiLevelList(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Create multi-level list
	u.AddNumberedItem("Main topic 1", 0, PositionEnd)
	u.AddNumberedItem("Subtopic 1.a", 1, PositionEnd)
	u.AddNumberedItem("Subtopic 1.b", 1, PositionEnd)
	u.AddNumberedItem("Sub-subtopic 1.b.i", 2, PositionEnd)
	u.AddNumberedItem("Main topic 2", 0, PositionEnd)

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify different levels are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	if !strings.Contains(docXML, `<w:ilvl w:val="0"/>`) {
		t.Error("Level 0 not found")
	}
	if !strings.Contains(docXML, `<w:ilvl w:val="1"/>`) {
		t.Error("Level 1 not found")
	}
	if !strings.Contains(docXML, `<w:ilvl w:val="2"/>`) {
		t.Error("Level 2 not found")
	}

	// Verify text content
	if !strings.Contains(docXML, "Main topic 1") {
		t.Error("Main topic 1 not found")
	}
	if !strings.Contains(docXML, "Subtopic 1.a") {
		t.Error("Subtopic 1.a not found")
	}
	if !strings.Contains(docXML, "Sub-subtopic 1.b.i") {
		t.Error("Sub-subtopic not found")
	}
}

func TestMixedBulletAndNumberedLists(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Add heading
	u.AddHeading(1, "Project Requirements", PositionEnd)

	// Add numbered list
	u.AddNumberedItem("Functional requirements", 0, PositionEnd)
	u.AddNumberedItem("Non-functional requirements", 0, PositionEnd)

	// Add some text
	u.AddText("Key deliverables include:", PositionEnd)

	// Add bullet list
	u.AddBulletItem("Documentation", 0, PositionEnd)
	u.AddBulletItem("Source code", 0, PositionEnd)
	u.AddBulletItem("Test cases", 0, PositionEnd)

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify both list types are present
	docXML := readZipEntry(t, outputPath, "word/document.xml")

	if !strings.Contains(docXML, `<w:numId w:val="1"/>`) {
		t.Error("Bullet list numId not found")
	}
	if !strings.Contains(docXML, `<w:numId w:val="2"/>`) {
		t.Error("Numbered list numId not found")
	}

	// Verify heading
	if !strings.Contains(docXML, "Project Requirements") {
		t.Error("Heading not found")
	}

	// Verify all list items
	expectedItems := []string{
		"Functional requirements",
		"Non-functional requirements",
		"Documentation",
		"Source code",
		"Test cases",
	}

	for _, item := range expectedItems {
		if !strings.Contains(docXML, item) {
			t.Errorf("List item not found: %s", item)
		}
	}
}

func TestStyleBasedListsStillWork(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Use old style-based approach (should still work)
	err = u.InsertParagraph(ParagraphOptions{
		Text:     "Style-based bullet",
		Style:    StyleListBullet,
		Position: PositionEnd,
	})
	if err != nil {
		t.Fatalf("Style-based list failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify style is applied
	docXML := readZipEntry(t, outputPath, "word/document.xml")
	if !strings.Contains(docXML, `w:pStyle w:val="ListBullet"`) {
		t.Error("Style-based list not found")
	}

	// Should NOT have numPr (since we're using style, not ListType)
	if strings.Contains(docXML, `<w:numPr>`) {
		t.Error("numPr should not be present for style-based lists")
	}
}

func TestListLevelBounds(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "input.docx")
	outputPath := filepath.Join(tempDir, "output.docx")

	if err := os.WriteFile(inputPath, buildFixtureDocx(t), 0o644); err != nil {
		t.Fatalf("write input fixture: %v", err)
	}

	u, err := New(inputPath)
	if err != nil {
		t.Fatalf("New failed: %v", err)
	}
	defer u.Cleanup()

	// Test negative level (should be clamped to 0)
	u.AddBulletItem("Negative level test", -5, PositionEnd)

	// Test level > 8 (should be clamped to 8)
	u.AddBulletItem("High level test", 15, PositionEnd)

	// Test valid boundary levels
	u.AddBulletItem("Level 0", 0, PositionEnd)
	u.AddBulletItem("Level 8", 8, PositionEnd)

	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	docXML := readZipEntry(t, outputPath, "word/document.xml")

	// Verify levels are within bounds
	if strings.Contains(docXML, `<w:ilvl w:val="-5"/>`) {
		t.Error("Negative level should be clamped")
	}
	if strings.Contains(docXML, `<w:ilvl w:val="15"/>`) {
		t.Error("Level > 8 should be clamped")
	}

	// Verify valid levels are present
	if !strings.Contains(docXML, `<w:ilvl w:val="0"/>`) {
		t.Error("Level 0 not found")
	}
	if !strings.Contains(docXML, `<w:ilvl w:val="8"/>`) {
		t.Error("Level 8 not found")
	}
}

// Helper functions for tests
func buildFixtureDocx(t *testing.T) []byte {
	t.Helper()

	docx := &bytes.Buffer{}
	docxZip := zip.NewWriter(docx)

	addZipEntry(t, docxZip, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`)
	addZipEntry(t, docxZip, "word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`)
	addZipEntry(t, docxZip, "word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`)

	if err := docxZip.Close(); err != nil {
		t.Fatalf("close docx zip: %v", err)
	}

	return docx.Bytes()
}

func addZipEntry(t *testing.T, w *zip.Writer, path, content string) {
	t.Helper()
	entry, err := w.Create(path)
	if err != nil {
		t.Fatalf("create zip entry %s: %v", path, err)
	}
	if _, err := entry.Write([]byte(content)); err != nil {
		t.Fatalf("write zip entry %s: %v", path, err)
	}
}

func readZipEntry(t *testing.T, zipPath, entryPath string) string {
	t.Helper()
	return string(readZipEntryBytes(t, zipPath, entryPath))
}

func readZipEntryBytes(t *testing.T, zipPath, entryPath string) []byte {
	t.Helper()

	r, err := zip.OpenReader(zipPath)
	if err != nil {
		t.Fatalf("open zip %s: %v", zipPath, err)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == entryPath {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open entry %s: %v", entryPath, err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("read entry %s: %v", entryPath, err)
			}
			return b
		}
	}

	t.Fatalf("entry not found: %s", entryPath)
	return nil
}
