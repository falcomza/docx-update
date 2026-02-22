# Caption Feature Documentation

## Overview

The caption functionality allows you to automatically add captions to charts and tables in Word documents. Captions include automatic numbering using Word's built-in SEQ fields, custom descriptions, and positioning options.

## How Captions Work in MS Word

In Microsoft Word, captions are special paragraphs that:
- Use the "Caption" style by default
- Include the caption label (e.g., "Figure" or "Table")
- Support automatic numbering through SEQ (sequence) fields
- Can be positioned before or after the object
- Are commonly used for cross-referencing

### Caption Format

A typical caption follows this format:
```
Figure 1: Description of the figure
Table 2: Description of the table
```

### SEQ Fields

Word uses SEQ (sequence) fields for automatic numbering. When you insert captions with `AutoNumber: true`, the library generates SEQ fields like:
```xml
{ SEQ Figure \* ARABIC }
```

Word automatically:
- Calculates the correct number when the document opens
- Updates all numbers when you add, remove, or reorder captioned items
- Supports separate numbering sequences for different caption types

## Usage

### Basic Caption with Chart

```go
err := u.InsertChart(godocx.ChartOptions{
    Position: godocx.PositionEnd,
    Title:    "Quarterly Sales",
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []godocx.SeriesData{
        {Name: "Revenue", Values: []float64{250, 280, 310, 290}},
    },
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionFigure,
        Description: "Quarterly sales performance",
        AutoNumber:  true,
        Position:    godocx.CaptionAfter,
    },
})
```

### Basic Caption with Table

```go
err := u.InsertTable(godocx.TableOptions{
    Position: godocx.PositionEnd,
    Columns: []godocx.ColumnDefinition{
        {Title: "Product"},
        {Title: "Sales"},
    },
    Rows: [][]string{
        {"Product A", "$50,000"},
        {"Product B", "$45,000"},
    },
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionTable,
        Description: "Product sales summary",
        AutoNumber:  true,
        Position:    godocx.CaptionBefore,  // Tables typically have captions above
    },
})
```

## CaptionOptions

### Fields

| Field | Type | Description | Default |
|-------|------|-------------|---------|
| `Type` | `CaptionType` | Caption type: `CaptionFigure` or `CaptionTable` | Required |
| `Position` | `CaptionPosition` | Position: `CaptionBefore` or `CaptionAfter` | `CaptionAfter` for figures, `CaptionBefore` for tables |
| `Description` | `string` | Text after the number (max 500 chars) | "" (optional) |
| `Style` | `string` | Word style name | "Caption" |
| `AutoNumber` | `bool` | Use SEQ fields for auto-numbering | `true` |
| `Alignment` | `CellAlignment` | Text alignment | `CellAlignLeft` |
| `ManualNumber` | `int` | Manual number (when AutoNumber is false) | 0 |

### Caption Types

```go
// For charts, diagrams, images
godocx.CaptionFigure

// For tables
godocx.CaptionTable
```

### Caption Positions

```go
// Caption appears before (above) the object
godocx.CaptionBefore

// Caption appears after (below) the object
godocx.CaptionAfter
```

### Alignment Options

```go
godocx.CellAlignLeft    // Left-aligned caption
godocx.CellAlignCenter  // Centered caption
godocx.CellAlignRight   // Right-aligned caption
```

## Advanced Examples

### Using Default Caption Options

The `DefaultCaptionOptions` function returns sensible defaults:

```go
// For figures: AutoNumber=true, Position=CaptionAfter
captionOpts := godocx.DefaultCaptionOptions(godocx.CaptionFigure)
captionOpts.Description = "Sales trend analysis"

err := u.InsertChart(godocx.ChartOptions{
    Position:   godocx.PositionEnd,
    Title:      "Sales Trend",
    Categories: []string{"Jan", "Feb", "Mar"},
    Series: []godocx.SeriesData{
        {Name: "Sales", Values: []float64{100, 120, 140}},
    },
    Caption: &captionOpts,
})
```

### Centered Caption

```go
err := u.InsertTable(godocx.TableOptions{
    Position: godocx.PositionEnd,
    Columns: []godocx.ColumnDefinition{
        {Title: "Department"},
        {Title: "Budget"},
    },
    Rows: [][]string{
        {"Marketing", "$100,000"},
        {"R&D", "$200,000"},
    },
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionTable,
        Description: "Department budget allocation",
        AutoNumber:  true,
        Position:    godocx.CaptionBefore,
        Alignment:   godocx.CellAlignCenter,
    },
})
```

### Manual Numbering

For cases where you want to control the numbers manually:

```go
err := u.InsertChart(godocx.ChartOptions{
    Position:   godocx.PositionEnd,
    Title:      "Special Chart",
    Categories: []string{"A", "B"},
    Series: []godocx.SeriesData{
        {Name: "Data", Values: []float64{10, 20}},
    },
    Caption: &godocx.CaptionOptions{
        Type:         godocx.CaptionFigure,
        Description:  "Appendix chart",
        AutoNumber:   false,
        ManualNumber: 99,  // Custom number
        Position:     godocx.CaptionAfter,
    },
})
```

### Custom Style

```go
err := u.InsertTable(godocx.TableOptions{
    Position: godocx.PositionEnd,
    Columns:  /* ... */,
    Rows:     /* ... */,
    Caption: &godocx.CaptionOptions{
        Type:        godocx.CaptionTable,
        Description: "Custom styled caption",
        Style:       "Heading 2",  // Use a different Word style
        AutoNumber:  true,
    },
})
```

## Best Practices

### Caption Positioning

1. **Figures (Charts, Images)**: Typically use `CaptionAfter` - captions appear below the figure
2. **Tables**: Typically use `CaptionBefore` - captions appear above the table
3. These are conventions in academic and professional documents, but you can customize as needed

### Auto-Numbering vs Manual

- **Use AutoNumber (recommended)**: Word handles numbering automatically, updates when you reorder, and supports cross-references
- **Use Manual**: Only when you need specific numbers or are working outside normal sequential numbering

### Description Guidelines

- Keep descriptions concise but informative (max 500 characters enforced)
- Use sentence case: "Quarterly sales by region" not "QUARTERLY SALES BY REGION"
- Don't duplicate information already in the chart title or table headers

### Caption Styles

The default "Caption" style is recommended because:
- It's a standard Word style that exists in all documents
- Users can customize it globally through their Word styles
- It maintains consistency across documents

## Validation

The `ValidateCaptionOptions` function checks:
- Caption type is valid (`CaptionFigure` or `CaptionTable`)
- Position is valid (`before` or `after`)
- Description is not too long (max 500 characters)

Validation is automatically called when inserting charts or tables with captions.

## Utility Functions

### FormatCaptionText

Preview what a caption will look like:

```go
opts := godocx.CaptionOptions{
    Type:        godocx.CaptionFigure,
    Description: "Sales data",
    AutoNumber:  true,
}

text := godocx.FormatCaptionText(opts)
// Returns: "Figure #: Sales data"
```

## Complete Example

See `examples/example_captions.go` for a comprehensive demonstration including:
- Charts with captions before and after
- Tables with various caption styles
- Manual and automatic numbering
- Different alignment options

Run it with:
```bash
cd examples
go run example_captions.go
```

## Technical Details

### XML Structure

Generated caption XML follows Word's standard format:

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Caption"/>
    <w:jc w:val="center"/>  <!-- If centered -->
  </w:pPr>
  <w:r>
    <w:t>Figure </w:t>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:instrText xml:space="preserve"> SEQ Figure \* ARABIC </w:instrText>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:t>1</w:t>  <!-- Placeholder, Word recalculates -->
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
  <w:r>
    <w:t xml:space="preserve">: </w:t>
  </w:r>
  <w:r>
    <w:t>Description text</w:t>
  </w:r>
</w:p>
```

### Caption Insertion Logic

1. Generate caption XML paragraph
2. Generate chart/table XML
3. Combine based on `CaptionPosition`:
   - `CaptionBefore`: [Caption] [Object]
   - `CaptionAfter`: [Object] [Caption]
4. Insert combined content at the specified document position

## Testing

Run caption tests:
```bash
go test ./tests -run Caption -v
```

Test coverage includes:
- Auto-numbering with SEQ fields
- Manual numbering
- Caption positioning (before/after)
- Various alignment options
- Validation logic
- Default options
- Multiple captions in one document

## Troubleshooting

### Caption numbers not updating

In Word, press `Ctrl+A` (Select All) then `F9` to update all fields.

### Caption appears in wrong location

Check the `Position` field - `CaptionBefore` vs `CaptionAfter`.

### Custom style not working

Ensure the style name matches exactly as defined in Word (case-sensitive). Use "Caption" if unsure.

### Numbers restarting unexpectedly

Each caption type (Figure, Table) maintains a separate sequence. This is standard Word behavior.
