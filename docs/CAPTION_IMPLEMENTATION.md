# Caption Implementation Summary

## Overview

Successfully implemented comprehensive caption functionality for figures (charts) and tables in the docx-updater library. Captions use Microsoft Word's standard SEQ (sequence) fields for automatic numbering and follow Word's caption conventions.

## Implementation Details

### Files Created/Modified

#### New Files
1. **src/caption.go** - Complete caption implementation with:
   - `CaptionOptions` struct for configuring captions
   - `generateCaptionXML()` - Generates Word caption paragraph XML
   - `generateSEQFieldXML()` - Creates SEQ fields for auto-numbering
   - `ValidateCaptionOptions()` - Options validation
   - `DefaultCaptionOptions()` - Sensible defaults for figures/tables
   - `FormatCaptionText()` - Preview caption text

2. **tests/caption_test.go** - Comprehensive test suite:
   - Caption insertion with charts and tables
   - Auto-numbering with SEQ fields
   - Manual numbering
   - Caption positioning (before/after)
   - Alignment options
   - Validation logic
   - Multiple captions in one document

3. **examples/example_captions.go** - Demonstration examples:
   - Charts with captions (before and after)
   - Tables with captions
   - Default options usage
   - Manual numbering
   - Custom alignment
   - Multiple captioned objects

4. **CAPTION_FEATURE.md** - Complete documentation:
   - Usage guide
   - API reference
   - Best practices
   - Technical details
   - Troubleshooting

#### Modified Files
1. **src/chart.go** - Added:
   - `Caption *CaptionOptions` field to `ChartOptions`
   - Caption handling in `insertChartDrawing()`

2. **src/table.go** - Added:
   - `Caption *CaptionOptions` field to `TableOptions`
   - Caption handling in `insertTableAtPosition()`

3. **README.md** - Updated with:
   - Caption feature in features list
   - caption.go in project structure
   - Caption usage section with examples
   - Link to detailed documentation

## How Captions Work in MS Word

### Caption Structure
Captions in Word are paragraphs with:
- **Label**: "Figure" or "Table"
- **Number**: Auto-generated using SEQ fields
- **Description**: User-provided text
- **Style**: "Caption" style by default

Format: `Figure 1: Description text`

### SEQ Fields
SEQ (sequence) fields provide automatic numbering:
```xml
<w:r>
  <w:fldChar w:fldCharType="begin"/>
</w:r>
<w:r>
  <w:instrText> SEQ Figure \* ARABIC </w:instrText>
</w:r>
<w:r>
  <w:fldChar w:fldCharType="separate"/>
</w:r>
<w:r>
  <w:t>1</w:t>  <!-- Calculated by Word -->
</w:r>
<w:r>
  <w:fldChar w:fldCharType="end"/>
</w:r>
```

Word features:
- Maintains separate sequences for "Figure" and "Table"
- Auto-updates when captions are added/removed/reordered
- Supports cross-referencing
- Updates with F9 key

### Caption Positioning
- **Figures (Charts)**: Typically positioned **after** (below) the figure
- **Tables**: Typically positioned **before** (above) the table
- Both are conventions in professional documents

## API Design

### CaptionOptions Structure
```go
type CaptionOptions struct {
    Type        CaptionType      // Figure or Table
    Position    CaptionPosition  // Before or After
    Description string           // Caption text (max 500 chars)
    Style       string           // Word style (default: "Caption")
    AutoNumber  bool             // Use SEQ fields (default: true)
    Alignment   CellAlignment    // Left, Center, Right
    ManualNumber int             // Manual number (when AutoNumber=false)
}
```

### Usage Patterns

**Simple Caption:**
```go
Caption: &godocx.CaptionOptions{
    Type:        godocx.CaptionFigure,
    Description: "Sales performance chart",
    AutoNumber:  true,
    Position:    godocx.CaptionAfter,
}
```

**Using Defaults:**
```go
caption := godocx.DefaultCaptionOptions(godocx.CaptionTable)
caption.Description = "Budget breakdown"
```

**Custom Styling:**
```go
Caption: &godocx.CaptionOptions{
    Type:        godocx.CaptionTable,
    Description: "Q4 Results",
    AutoNumber:  true,
    Position:    godocx.CaptionBefore,
    Alignment:   godocx.CellAlignCenter,
    Style:       "Heading 2",  // Custom style
}
```

## Technical Implementation

### Caption XML Generation
1. Create paragraph (`<w:p>`) with Caption style
2. Add label text run ("Figure " or "Table ")
3. Insert SEQ field for numbering (3 runs: begin, instruction, end)
4. Add separator (": ")
5. Add description text run

### Integration with Charts/Tables
1. Generate caption XML paragraph
2. Generate chart/table XML
3. Combine based on position:
   - `CaptionBefore`: [Caption][Object]
   - `CaptionAfter`: [Object][Caption]
4. Insert combined content at document position

### Validation
- Caption type must be "Figure" or "Table"
- Position must be "before" or "after"
- Description max 500 characters
- Defaults applied when not specified

## Test Coverage

All tests passing (100% success rate):
- Basic caption insertion
- Auto-numbering (SEQ fields)
- Manual numbering
- Caption positioning
- Alignment options
- Validation logic
- Default options
- Multiple captions per document
- Integration with chart/table insertion

## Compatibility

### Word Versions
- Word 2007+: Full support
- Word 2003: Limited (no SEQ fields)
- Word Online: Full support
- LibreOffice Writer: Partial support

### Document Formats
- .docx (OpenXML): Full support
- .doc (binary): Not supported

## Best Practices

1. **Use AutoNumber**: Lets Word manage numbering automatically
2. **Follow Positioning Conventions**:
   - Figures: Caption after (below)
   - Tables: Caption before (above)
3. **Keep Descriptions Concise**: Under 500 characters
4. **Use Default Style**: "Caption" style for consistency
5. **Update Fields in Word**: Press Ctrl+A then F9 to refresh numbers

## Future Enhancements (Optional)

Potential future additions:
- Custom caption labels (e.g., "Diagram", "Chart")
- Chapter-based numbering (e.g., "Figure 3.1")
- Cross-reference support
- Caption list generation
- Multi-language support
- Custom separators (e.g., "–" instead of ":")

## Examples

### Chart with Caption
```go
u.InsertChart(godocx.ChartOptions{
    Title: "Sales Trend",
    // ... chart data ...
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionFigure,
        Description: "Monthly sales for 2024",
        AutoNumber:  true,
        Position:    godocx.CaptionAfter,
    },
})
```

### Table with Caption
```go
u.InsertTable(godocx.TableOptions{
    Columns: /* ... */,
    Rows:    /* ... */,
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionTable,
        Description: "Product comparison",
        AutoNumber:  true,
        Position:    godocx.CaptionBefore,
    },
})
```

## Summary

The caption implementation is:
- ✅ **Complete**: All core features implemented
- ✅ **Tested**: Comprehensive test coverage
- ✅ **Documented**: Full API documentation and examples
- ✅ **Standards-compliant**: Uses Word's native SEQ fields
- ✅ **Flexible**: Supports various positioning and styling options
- ✅ **Easy to use**: Simple API with sensible defaults

The feature integrates seamlessly with existing chart and table functionality, maintaining the library's consistent API design.
