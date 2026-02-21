# Paragraph Insertion Feature

## Overview
Added comprehensive paragraph insertion functionality to the docx-chart-updater library, allowing insertion of styled paragraphs with custom formatting.

## New Files Created
- `src/paragraph.go` - Core paragraph insertion implementation
- `tests/paragraph_test.go` - Comprehensive test suite
- `examples/example_paragraph.go` - Example usage

## API

### Main Function
```go
func (u *Updater) InsertParagraph(opts ParagraphOptions) error
```

### ParagraphOptions Structure
```go
type ParagraphOptions struct {
    Text      string         // Required: The text content
    Style     ParagraphStyle // Paragraph style (default: Normal)
    Position  InsertPosition // Where to insert
    Anchor    string         // Text anchor for relative positioning
    Bold      bool           // Bold formatting
    Italic    bool           // Italic formatting
    Underline bool           // Underline formatting
}
```

### Available Styles
- `StyleNormal` - Normal paragraph
- `StyleHeading1`, `StyleHeading2`, `StyleHeading3` - Heading levels
- `StyleTitle`, `StyleSubtitle` - Title styles
- `StyleQuote`, `StyleIntense` - Quote styles
- `StyleListNumber`, `StyleListBullet` - List styles

### Insert Positions
- `PositionBeginning` - Start of document
- `PositionEnd` - End of document  
- `PositionAfterText` - After specified anchor text
- `PositionBeforeText` - Before specified anchor text

### Convenience Functions
```go
func (u *Updater) AddHeading(level int, text string, position InsertPosition) error
func (u *Updater) AddText(text string, position InsertPosition) error
func (u *Updater) InsertParagraphs(paragraphs []ParagraphOptions) error
```

## Usage Examples

### Basic Usage
```go
u, _ := updater.New("template.docx")
defer u.Cleanup()

// Add a heading
u.AddHeading(1, "Report Title", updater.PositionBeginning)

// Add normal text
u.AddText("This is the introduction.", updater.PositionEnd)

u.Save("output.docx")
```

### Custom Formatting
```go
u.InsertParagraph(updater.ParagraphOptions{
    Text:      "Important Note:",
    Style:     updater.StyleNormal,
    Position:  updater.PositionEnd,
    Bold:      true,
    Underline: true,
})
```

### Batch Insertion
```go
paragraphs := []updater.ParagraphOptions{
    {
        Text:     "Section 1",
        Style:    updater.StyleHeading2,
        Position: updater.PositionEnd,
    },
    {
        Text:     "Content here.",
        Style:    updater.StyleNormal,
        Position: updater.PositionEnd,
    },
}
u.InsertParagraphs(paragraphs)
```

## Test Coverage
- ✅ Insert at beginning
- ✅ Insert at end
- ✅ Custom formatting (bold, italic, underline)
- ✅ Heading styles
- ✅ Multiple paragraphs
- ✅ Error handling (empty text)

All tests passing (6/6).

## Implementation Details
- Uses proper WordprocessingML XML structure
- Preserves whitespace with `xml:space="preserve"` when needed
- Supports all standard paragraph styles
- Efficient byte manipulation for insertion
- XML-escaped content to prevent injection issues

## Future Enhancements
- Text anchoring (PositionAfterText/PositionBeforeText) implementation
- Support for custom paragraph properties (spacing, alignment, indentation)
- Rich text with mixed formatting within a paragraph
- Table insertion
- List management (numbered/bulleted)
