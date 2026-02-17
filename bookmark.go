package docxupdater

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// BookmarkOptions defines options for bookmark creation
type BookmarkOptions struct {
	// Position where to insert the bookmark
	Position InsertPosition

	// Anchor text for position-based insertion (for PositionAfterText/PositionBeforeText)
	Anchor string

	// Style to apply to the bookmarked text paragraph
	Style ParagraphStyle

	// Whether the bookmark should be invisible (default: true)
	// Word bookmarks are typically invisible markers
	Hidden bool
}

// DefaultBookmarkOptions returns bookmark options with sensible defaults
func DefaultBookmarkOptions() BookmarkOptions {
	return BookmarkOptions{
		Position: PositionEnd,
		Style:    StyleNormal,
		Hidden:   true,
	}
}

// CreateBookmark creates a bookmark at a specific location with optional text
// If text is provided, the bookmark wraps the text; otherwise it's an empty bookmark marker
func (u *Updater) CreateBookmark(name string, opts BookmarkOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if err := validateBookmarkName(name); err != nil {
		return err
	}

	// Get next available bookmark ID
	bookmarkID, err := u.getNextBookmarkID()
	if err != nil {
		return fmt.Errorf("get next bookmark ID: %w", err)
	}

	// Generate bookmark XML (empty marker)
	bookmarkXML := generateEmptyBookmarkXML(name, bookmarkID, opts)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return NewXMLParseError("document.xml", err)
	}

	// Insert bookmark at specified position
	updated, err := insertBookmarkAtPosition(raw, bookmarkXML, opts)
	if err != nil {
		return fmt.Errorf("insert bookmark: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return NewXMLWriteError("document.xml", err)
	}

	return nil
}

// CreateBookmarkWithText creates a bookmark that wraps specific text content
func (u *Updater) CreateBookmarkWithText(name, text string, opts BookmarkOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if text == "" {
		return NewValidationError("text", "bookmark text cannot be empty")
	}
	if err := validateBookmarkName(name); err != nil {
		return err
	}

	// Apply style defaults
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	// Get next available bookmark ID
	bookmarkID, err := u.getNextBookmarkID()
	if err != nil {
		return fmt.Errorf("get next bookmark ID: %w", err)
	}

	// Generate bookmark XML wrapping text
	bookmarkXML := generateBookmarkWithTextXML(name, text, bookmarkID, opts)

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return NewXMLParseError("document.xml", err)
	}

	// Insert bookmark at specified position
	updated, err := insertBookmarkAtPosition(raw, bookmarkXML, opts)
	if err != nil {
		return fmt.Errorf("insert bookmark with text: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return NewXMLWriteError("document.xml", err)
	}

	return nil
}

// WrapTextInBookmark finds existing text in the document and wraps it with a bookmark
func (u *Updater) WrapTextInBookmark(name, anchorText string) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if anchorText == "" {
		return NewValidationError("anchorText", "anchor text cannot be empty")
	}
	if err := validateBookmarkName(name); err != nil {
		return err
	}

	// Get next available bookmark ID
	bookmarkID, err := u.getNextBookmarkID()
	if err != nil {
		return fmt.Errorf("get next bookmark ID: %w", err)
	}

	// Read document.xml
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return NewXMLParseError("document.xml", err)
	}

	// Wrap the text in bookmark tags
	updated, err := wrapExistingTextInBookmark(raw, name, anchorText, bookmarkID)
	if err != nil {
		return fmt.Errorf("wrap text in bookmark: %w", err)
	}

	// Write updated document
	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return NewXMLWriteError("document.xml", err)
	}

	return nil
}

// getNextBookmarkID finds the next available bookmark ID in the document
func (u *Updater) getNextBookmarkID() (int, error) {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return 0, fmt.Errorf("read document: %w", err)
	}

	// Find all bookmark IDs using the pattern from constants
	matches := bookmarkIDPattern.FindAllStringSubmatch(string(raw), -1)

	maxID := 0
	for _, match := range matches {
		if len(match) > 1 {
			var id int
			fmt.Sscanf(match[1], "%d", &id)
			if id > maxID {
				maxID = id
			}
		}
	}

	return maxID + 1, nil
}

// validateBookmarkName validates a bookmark name according to Word specifications
// Rules:
// - Must start with a letter
// - Can contain letters, digits, and underscores
// - No spaces or special characters (except underscore)
// - Maximum 40 characters
// - Cannot be a reserved word
func validateBookmarkName(name string) error {
	if name == "" {
		return NewValidationError("name", "bookmark name cannot be empty")
	}

	// Check length
	if len(name) > 40 {
		return NewValidationError("name", "bookmark name must be 40 characters or less")
	}

	// Check first character is a letter
	if !isLetter(rune(name[0])) {
		return NewValidationError("name", "bookmark name must start with a letter")
	}

	// Check for valid characters (letters, digits, underscores)
	validNamePattern := regexp.MustCompile(`^[a-zA-Z][a-zA-Z0-9_]*$`)
	if !validNamePattern.MatchString(name) {
		return NewValidationError("name", "bookmark name can only contain letters, digits, and underscores")
	}

	// Check for reserved names (Word reserved bookmark names)
	reservedNames := []string{
		"_Toc", "_Hlt", "_Ref", "_GoBack",
	}
	for _, reserved := range reservedNames {
		if strings.HasPrefix(name, reserved) {
			return NewValidationError("name", fmt.Sprintf("bookmark name cannot start with reserved prefix '%s'", reserved))
		}
	}

	return nil
}

// isLetter checks if a rune is a letter
func isLetter(r rune) bool {
	return (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z')
}

// generateEmptyBookmarkXML creates XML for an empty bookmark marker (just a position marker)
func generateEmptyBookmarkXML(name string, id int, opts BookmarkOptions) []byte {
	var buf strings.Builder

	buf.WriteString("<w:p>")

	// Add paragraph properties if style is specified
	if opts.Style != "" {
		buf.WriteString("<w:pPr>")
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, opts.Style))
		buf.WriteString("</w:pPr>")
	}

	// Add bookmark start and end (empty bookmark - just marks a position)
	buf.WriteString(fmt.Sprintf(`<w:bookmarkStart w:id="%d" w:name="%s"/>`, id, escapeXMLAttribute(name)))
	buf.WriteString(fmt.Sprintf(`<w:bookmarkEnd w:id="%d"/>`, id))

	buf.WriteString("</w:p>")

	return []byte(buf.String())
}

// generateBookmarkWithTextXML creates XML for a bookmark that wraps text content
func generateBookmarkWithTextXML(name, text string, id int, opts BookmarkOptions) []byte {
	var buf strings.Builder

	buf.WriteString("<w:p>")

	// Add paragraph properties
	if opts.Style != "" {
		buf.WriteString("<w:pPr>")
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, opts.Style))
		buf.WriteString("</w:pPr>")
	}

	// Bookmark start
	buf.WriteString(fmt.Sprintf(`<w:bookmarkStart w:id="%d" w:name="%s"/>`, id, escapeXMLAttribute(name)))

	// Text run
	buf.WriteString("<w:r>")
	buf.WriteString("<w:t")
	// Preserve spaces if text starts or ends with space
	if strings.HasPrefix(text, " ") || strings.HasSuffix(text, " ") {
		buf.WriteString(` xml:space="preserve"`)
	}
	buf.WriteString(">")
	buf.WriteString(xmlEscape(text))
	buf.WriteString("</w:t>")
	buf.WriteString("</w:r>")

	// Bookmark end
	buf.WriteString(fmt.Sprintf(`<w:bookmarkEnd w:id="%d"/>`, id))

	buf.WriteString("</w:p>")

	return []byte(buf.String())
}

// wrapExistingTextInBookmark wraps existing text in the document with bookmark tags
func wrapExistingTextInBookmark(docXML []byte, name, anchorText string, id int) ([]byte, error) {
	docStr := string(docXML)

	// Find the anchor text
	textIdx := strings.Index(docStr, anchorText)
	if textIdx == -1 {
		return nil, fmt.Errorf("anchor text not found in document: %s", anchorText)
	}

	// Work backwards to find the opening <w:r> tag before the text
	runStartIdx := strings.LastIndex(docStr[:textIdx], "<w:r")
	if runStartIdx == -1 {
		return nil, fmt.Errorf("could not find run tag before anchor text")
	}

	// Find the start of the opening <w:r> tag (could be <w:r> or <w:r ...)
	runStartEnd := strings.Index(docStr[runStartIdx:], ">")
	if runStartEnd == -1 {
		return nil, fmt.Errorf("malformed run tag")
	}
	runStartEnd += runStartIdx + 1

	// Find the closing </w:r> tag after the text
	runEndIdx := strings.Index(docStr[textIdx:], "</w:r>")
	if runEndIdx == -1 {
		return nil, fmt.Errorf("could not find closing run tag after anchor text")
	}
	runEndIdx += textIdx + len("</w:r>")

	// Create bookmark tags
	bookmarkStart := fmt.Sprintf(`<w:bookmarkStart w:id="%d" w:name="%s"/>`, id, escapeXMLAttribute(name))
	bookmarkEnd := fmt.Sprintf(`<w:bookmarkEnd w:id="%d"/>`, id)

	// Build the result with bookmarks wrapping the run
	var result strings.Builder
	result.WriteString(docStr[:runStartIdx])
	result.WriteString(bookmarkStart)
	result.WriteString(docStr[runStartIdx:runEndIdx])
	result.WriteString(bookmarkEnd)
	result.WriteString(docStr[runEndIdx:])

	return []byte(result.String()), nil
}

// insertBookmarkAtPosition inserts bookmark at the specified position
func insertBookmarkAtPosition(docXML, bookmarkXML []byte, opts BookmarkOptions) ([]byte, error) {
	switch opts.Position {
	case PositionBeginning:
		return insertAtBodyStart(docXML, bookmarkXML)
	case PositionEnd:
		return insertAtBodyEnd(docXML, bookmarkXML)
	case PositionAfterText:
		if opts.Anchor == "" {
			return nil, NewValidationError("anchor", "anchor text required for PositionAfterText")
		}
		return insertAfterText(docXML, bookmarkXML, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return nil, NewValidationError("anchor", "anchor text required for PositionBeforeText")
		}
		return insertBeforeText(docXML, bookmarkXML, opts.Anchor)
	default:
		return insertAtBodyEnd(docXML, bookmarkXML)
	}
}
