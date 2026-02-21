# Table Insertion Feature

This document provides comprehensive documentation for the table insertion functionality in the docx-chart-updater library.

## Overview

The table insertion feature allows you to create professional tables in DOCX documents with full control over:
- Column structure and widths
- Header styling and formatting
- Row data and alternate row colors
- Border styles and colors
- Cell alignment and padding
- Header row repetition on new pages

## API Reference

### InsertTable Function

```go
func (u *Updater) InsertTable(opts TableOptions) error
```

Main function for inserting tables into a document.

### TableOptions Structure

```go
type TableOptions struct {
    // Position where to insert the table
    Position InsertPosition  // PositionBeginning, PositionEnd, PositionAfterText, PositionBeforeText
    Anchor   string          // Text anchor for relative positioning

    // Column definitions
    Columns                  []ColumnDefinition // Column titles and properties
    ColumnWidths             []int              // Optional: column widths in twips (1/1440 inch)
    ProportionalColumnWidths bool               // Optional: size columns proportionally based on content

    // Data rows (excluding header)
    Rows [][]string // Each inner slice is a row of cell data

    // Header styling
    HeaderStyle      CellStyle         // Style for header row
    HeaderStyleName  string            // Word paragraph style name (e.g., "Heading1", "Normal")
    RepeatHeader     bool              // Repeat header row on each page
    HeaderBackground string            // Hex color for header background (e.g., "4472C4")
    HeaderBold       bool              // Make header text bold
    HeaderAlignment  CellAlignment     // Header text alignment

    // Row styling
    RowStyle          CellStyle         // Style for data rows
    RowStyleName      string            // Word paragraph style name (e.g., "Normal", "BodyText")
    AlternateRowColor string            // Hex color for alternate rows (e.g., "F2F2F2")
    RowAlignment      CellAlignment     // Default cell alignment for data rows
    VerticalAlign     VerticalAlignment // Vertical alignment in cells

    // Row height
    HeaderRowHeight int           // Header row height in twips, 0 for auto
    HeaderHeightRule RowHeightRule // Header height rule (auto, atLeast, exact)
    RowHeight        int           // Data row height in twips, 0 for auto
    RowHeightRule    RowHeightRule // Data row height rule (auto, atLeast, exact)

    // Table properties
    TableAlignment TableAlignment // Table alignment on page
    TableWidthType TableWidthType // Width type: auto, percentage, or fixed (default: percentage)
    TableWidth     int            // Width value: 0 for auto, 5000 for 100% (pct), or twips (dxa)
    TableStyle     TableStyle     // Predefined table style

    // Border properties
    BorderStyle BorderStyle // Border style
    BorderSize  int         // Border width in eighths of a point (default: 4 = 0.5pt)
    BorderColor string      // Hex color for borders (default: "000000")

    // Cell properties
    CellPadding int  // Cell padding in twips (default: 108 = 0.075")
    AutoFit     bool // Auto-fit content
}
```

### Column Definition

```go
type ColumnDefinition struct {
    Title     string        // Column header title
    Width     int           // Optional: width in twips, 0 for auto
    Alignment CellAlignment // Optional: alignment for this column
    Bold      bool          // Make header bold
}
```

### Cell Style

```go
type CellStyle struct {
    Bold       bool
    Italic     bool
    FontSize   int    // Font size in half-points (e.g., 20 = 10pt)
    FontColor  string // Hex color (e.g., "000000")
    Background string // Hex color for cell background
}
```

## Enumerations

### TableAlignment

```go
const (
    AlignLeft   TableAlignment = "left"
    AlignCenter TableAlignment = "center"
    AlignRight  TableAlignment = "right"
)
```

### CellAlignment

```go
const (
    CellAlignLeft   CellAlignment = "start"
    CellAlignCenter CellAlignment = "center"
    CellAlignRight  CellAlignment = "end"
)
```

### VerticalAlignment

```go
const (
    VerticalAlignTop    VerticalAlignment = "top"
    VerticalAlignCenter VerticalAlignment = "center"
    VerticalAlignBottom VerticalAlignment = "bottom"
)
```

### BorderStyle

```go
const (
    BorderSingle BorderStyle = "single"
    BorderDouble BorderStyle = "double"
    BorderDashed BorderStyle = "dashed"
    BorderDotted BorderStyle = "dotted"
    BorderNone   BorderStyle = "none"
)
```

### TableWidthType

```go
const (
    TableWidthAuto       TableWidthType = "auto"  // Auto-fit to content
    TableWidthPercentage TableWidthType = "pct"   // Percentage (5000 = 100%, default)
    TableWidthFixed      TableWidthType = "dxa"   // Fixed width in twips
)
```

**Width Values:**
- **Percentage mode**: 5000 = 100% (default, spans between margins), 2500 = 50%, 3750 = 75%
- **Fixed mode**: Width in twips (1440 = 1 inch, 7200 = 5 inches)
- **Auto mode**: Width value is ignored, table fits to content

### RowHeightRule

```go
const (
    RowHeightAuto    RowHeightRule = "auto"    // Auto height based on content (default)
    RowHeightAtLeast RowHeightRule = "atLeast" // Minimum height, can grow
    RowHeightExact   RowHeightRule = "exact"   // Fixed height, no growth
)
```

**Height Behaviors:**
- **Auto (default)**: Row height automatically adjusts to fit content
- **AtLeast**: Minimum height specified, but row can grow if content requires it
- **Exact**: Fixed height, content is clipped if it exceeds the height

**Common Heights (in twips, 1 inch = 1440 twips):**
- 360 = 0.25 inch (compact rows)
- 450 = 0.3125 inch (standard data rows)
- 720 = 0.5 inch (spacious rows)
- 900 = 0.625 inch (large header rows)
- 1080 = 0.75 inch (extra spacious)
- 1440 = 1.0 inch (very tall rows)

### TableStyle

```go
const (
    TableStyleGrid         TableStyle = "TableGrid"
    TableStyleGridAccent1  TableStyle = "LightShading-Accent1"
    TableStyleGridAccent2  TableStyle = "MediumShading1-Accent1"
    TableStylePlain        TableStyle = "TableNormal"
    TableStyleColorful     TableStyle = "ColorfulGrid-Accent1"
    TableStyleProfessional TableStyle = "LightGrid-Accent1"
)
```

## Usage Examples

### Basic Table

```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Name"},
        {Title: "Age"},
        {Title: "City"},
    },
    Rows: [][]string{
        {"Alice", "30", "New York"},
        {"Bob", "25", "Los Angeles"},
        {"Charlie", "35", "Chicago"},
    },
    HeaderBold: true,
})
```

### Width Configuration Examples

**Default Width (100% - Spans Between Margins)**
```go
// Default behavior: 100% width spanning between left and right margins
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Column 1"},
        {Title: "Column 2"},
    },
    Rows: [][]string{
        {"Data 1", "Data 2"},
    },
    HeaderBold: true,
    // TableWidthType defaults to TableWidthPercentage
    // TableWidth defaults to 5000 (100%)
})
```

**Custom Percentage Width (50%)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Code"},
        {Title: "Status"},
    },
    Rows: [][]string{
        {"A001", "Active"},
    },
    TableWidthType: TableWidthPercentage,
    TableWidth:     2500, // 50% (5000 = 100%)
    TableAlignment: AlignCenter,
    HeaderBold:     true,
})
```

**Fixed Width (5 inches)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Name"},
        {Title: "Email"},
    },
    Rows: [][]string{
        {"John Doe", "john@example.com"},
    },
    TableWidthType: TableWidthFixed,
    TableWidth:     7200, // 5 inches (1440 twips per inch)
    HeaderBold:     true,
})
```

**Auto Width (Fits Content)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "#"},
        {Title: "Item"},
    },
    Rows: [][]string{
        {"1", "Short"},
        {"2", "Text"},
    },
    TableWidthType: TableWidthAuto,
    HeaderBold:     true,
})
```

### Styled Table with Colors

```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Product", Alignment: CellAlignLeft},
        {Title: "Price", Alignment: CellAlignRight},
        {Title: "Stock", Alignment: CellAlignCenter},
    },
    Rows: [][]string{
        {"Laptop", "$999", "15"},
        {"Mouse", "$29", "50"},
        {"Keyboard", "$79", "30"},
    },
    HeaderBold:        true,
    HeaderBackground:  "4472C4",        // Blue
    HeaderAlignment:   CellAlignCenter,
    AlternateRowColor: "F2F2F2",        // Light gray
    BorderStyle:       BorderSingle,
    BorderSize:        6,
    BorderColor:       "2E75B5",
    TableAlignment:    AlignCenter,
})
```

### Table with Custom Column Widths

```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Code"},
        {Title: "Description"},
        {Title: "Status"},
    },
    ColumnWidths: []int{1000, 3000, 1000}, // In twips (1440 = 1 inch)
    Rows: [][]string{
        {"A1", "First item with long description", "Active"},
        {"B2", "Second item", "Pending"},
    },
    HeaderBold:       true,
    HeaderBackground: "70AD47",
})
```

### Table with Proportional Column Widths (Content-Based)

The proportional column width feature automatically sizes columns based on content length. Columns with longer content receive wider space, proportionally scaled to the total table width.

**How It Works:**
- Measures content length in each column (header + data cells)
- Distributes available width proportionally based on content
- Maintains total table width constraints (respects percentage/fixed width)
- Useful for data-driven tables where content varies significantly

**Basic Proportional Sizing (100% Width)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "ID"},                // Short content
        {Title: "Product Description"}, // Long content
        {Title: "Price"},             // Medium content
    },
    Rows: [][]string{
        {"1", "High-quality professional camera with advanced features", "$999"},
        {"2", "Basic item", "$29"},
        {"3", "Mid-range product suitable for most uses", "$199"},
    },
    ProportionalColumnWidths: true,  // Enable proportional sizing
    TableWidth:               5000,   // 100% width (default)
    TableWidthType:           TableWidthPercentage,
    HeaderBold:               true,
    HeaderBackground:         "4472C4",
})
```

Result: "Product Description" column gets ~7x more width than "ID" column due to longer content.

**Proportional Sizing with Fixed Table Width (6 Inches)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "A"},           // Very short
        {Title: "Middle Column"}, // Long
        {Title: "B"},           // Very short
    },
    Rows: [][]string{
        {"X", "This is a much longer piece of text", "Y"},
        {"1", "Extended content here", "2"},
    },
    ProportionalColumnWidths: true,
    TableWidthType:           TableWidthFixed,
    TableWidth:               8640, // 6 inches (1440 twips/inch)
    HeaderBold:               true,
})
```

Result: Middle column spans ~90% of table width. Narrow "A" and "B" columns get minimal space.

**Proportional Sizing with Percentage Width (50%)**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Short"},
        {Title: "This is a much longer header"},
        {Title: "Med"},
    },
    Rows: [][]string{
        {"X", "Extended content text here", "Data"},
    },
    ProportionalColumnWidths: true,
    TableWidthType:           TableWidthPercentage,
    TableWidth:               2500, // 50% width
    TableAlignment:           AlignCenter,
    HeaderBold:               true,
})
```

Result: Within half-page width, middle column still gets 60-70% of available space.

**Precedence: Explicit > Proportional > Equal**

Explicit column widths take precedence over proportional sizing:
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "A"},
        {Title: "B"},
        {Title: "C"},
    },
    Rows: [][]string{
        {"Short", "Much longer content here", "Med"},
    },
    // EXPLICIT widths take precedence - proportional is ignored
    ColumnWidths: []int{1500, 3000, 1500},
    ProportionalColumnWidths: true, // Ignored due to explicit ColumnWidths
    HeaderBold: true,
})
```

### Row Height Examples

**Auto Height (Default)**
```go
// Rows automatically adjust to content height
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Description"},
        {Title: "Notes"},
    },
    Rows: [][]string{
        {"Short text", "Auto"},
        {"Longer text that wraps", "Grows as needed"},
    },
    HeaderBold: true,
    // RowHeightRule defaults to RowHeightAuto
})
```

**Exact Height (Fixed)**
```go
// Fixed height rows - content clipped if too large
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Item"},
        {Title: "Value"},
    },
    Rows: [][]string{
        {"Row 1", "Data 1"},
        {"Row 2", "Data 2"},
    },
    HeaderRowHeight:  720,  // 0.5 inch header
    HeaderHeightRule: RowHeightExact,
    RowHeight:        450,  // 0.3125 inch data rows
    RowHeightRule:    RowHeightExact,
    HeaderBold:       true,
})
```

**Minimum Height (At Least)**
```go
// Minimum height that can grow
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Content"},
    },
    Rows: [][]string{
        {"Short"},
        {"This is much longer content that expands the row"},
    },
    RowHeight:     500,  // Minimum 500 twips
    RowHeightRule: RowHeightAtLeast,
    HeaderBold:    true,
})
```

### Table with Repeating Header

```go
// Create table with many rows - header will repeat on each page
rows := make([][]string, 50)
for i := range rows {
    rows[i] = []string{
        fmt.Sprintf("Item %d", i+1),
        fmt.Sprintf("Description %d", i+1),
        fmt.Sprintf("$%d.00", (i+1)*10),
    }
}

err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Item"},
        {Title: "Description"},
        {Title: "Price"},
    },
    Rows:             rows,
    HeaderBold:       true,
    RepeatHeader:     true,              // Repeat on each page
    HeaderBackground: "2E75B5",
    TableAlignment:   AlignCenter,
})
```

### Professional Financial Table

```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Quarter", Alignment: CellAlignCenter},
        {Title: "Revenue", Alignment: CellAlignRight},
        {Title: "Expenses", Alignment: CellAlignRight},
        {Title: "Profit", Alignment: CellAlignRight},
        {Title: "Margin", Alignment: CellAlignRight},
    },
    ColumnWidths: []int{1200, 1500, 1500, 1500, 1000},
    Rows: [][]string{
        {"Q1 2026", "$500,000", "$350,000", "$150,000", "30%"},
        {"Q2 2026", "$550,000", "$375,000", "$175,000", "32%"},
        {"Q3 2026", "$600,000", "$400,000", "$200,000", "33%"},
        {"Q4 2026", "$650,000", "$425,000", "$225,000", "35%"},
    },
    HeaderBold:        true,
    HeaderBackground:  "2E75B5",
    HeaderAlignment:   CellAlignCenter,
    AlternateRowColor: "DEEBF7",
    BorderStyle:       BorderDouble,
    BorderSize:        8,
    BorderColor:       "2E75B5",
    TableAlignment:    AlignCenter,
    RowStyle: CellStyle{
        FontSize: 20, // 10pt
    },
})
```

### Named Word Styles

Tables can reference Word's built-in or custom paragraph styles instead of (or in addition to) direct formatting.

**Using Built-in Word Styles**
```go
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Feature"},
        {Title: "Description"},
    },
    Rows: [][]string{
        {"Named Styles", "References Word styles defined in document"},
        {"Consistency", "Maintains corporate style guidelines"},
    },
    HeaderStyleName:   "Heading2",    // Word's Heading 2 style
    RowStyleName:      "BodyText",    // Word's Body Text style
    HeaderBackground:  "4472C4",
    AlternateRowColor: "E7E6E6",
})
```

**Mixing Named Styles with Direct Formatting**
```go
// Combine named styles with custom colors and formatting
err := updater.InsertTable(TableOptions{
    Position: PositionEnd,
    Columns: []ColumnDefinition{
        {Title: "Quarter"},
        {Title: "Revenue"},
    },
    Rows: [][]string{
        {"Q1 2026", "$250,000"},
        {"Q2 2026", "$280,000"},
    },
    HeaderStyleName:   "Heading3",    // Named style
    HeaderBold:        true,          // Plus direct bold
    HeaderBackground:  "2E75B5",      // Plus direct background
    RowStyleName:      "Normal",      // Named style
    RowStyle: CellStyle{              // Plus direct formatting
        FontSize: 20, // 10pt
    },
    AlternateRowColor: "DEEBF7",
})
```

**Common Word Style Names**
- `Normal` - Default paragraph style (most common for data rows)
- `Heading1`, `Heading2`, `Heading3` - Heading styles (common for headers)
- `BodyText` - Body text paragraph style
- `Title`, `Subtitle` - Document title styles
- `IntenseQuote` - Emphasized quote style
- Custom styles defined in your template document

**Benefits of Named Styles**
- **Consistency**: All tables using the same style name update together
- **Template-based**: Styles can be customized in the template document
- **Corporate branding**: Use company-defined styles for consistent branding
- **Flexible**: Mix named styles with direct formatting as needed
- **Easy updates**: Change the style definition once, affects all instances

**Examples**: See `examples/example_table_named_styles.go` for comprehensive demonstrations.

## Measurement Units

### Twips
Tables use **twips** (twentieth of a point) for measurements:
- 1 twip = 1/1440 inch
- 1 inch = 1440 twips
- Letter page width (8.5") ≈ 12,240 twips
- A4 page width (8.27") ≈ 11,906 twips

Common conversions:
- 0.5 inch = 720 twips
- 1 inch = 1440 twips
- 2 inches = 2880 twips
- 3 inches = 4320 twips

### Font Sizes
Font sizes are specified in **half-points**:
- 18 = 9pt
- 20 = 10pt
- 22 = 11pt
- 24 = 12pt
- 28 = 14pt

### Border Sizes
Border widths are in **eighths of a point**:
- 4 = 0.5pt (thin)
- 6 = 0.75pt (medium)
- 8 = 1pt (thick)
- 12 = 1.5pt (very thick)

## Color Specification

All colors use **hex format without the # prefix**:
- `"000000"` - Black
- `"FFFFFF"` - White
- `"FF0000"` - Red
- `"00FF00"` - Green
- `"0000FF"` - Blue
- `"4472C4"` - Professional Blue
- `"70AD47"` - Professional Green
- `"F2F2F2"` - Light Gray
- `"E7E6E6"` - Alternate Row Gray

## Validation

The library performs validation on table options:

1. **Column Count**: At least one column is required
2. **Row Consistency**: All rows must have the same number of cells as columns
3. **Column Widths**: If specified, must match the number of columns
4. **Anchor Text**: Required when using `PositionAfterText` or `PositionBeforeText`

## Word Document Structure

Tables are created using WordprocessingML XML:

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblBorders>...</w:tblBorders>
        <w:tblCellMar>...</w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w="1666"/>
        <w:gridCol w:w="1666"/>
        <w:gridCol w:w="1668"/>
    </w:tblGrid>
    <w:tr><!-- header row -->
        <w:trPr>
            <w:tblHeader/><!-- repeat on each page -->
        </w:trPr>
        <w:tc><!-- header cell -->
            <w:tcPr>
                <w:shd w:fill="4472C4"/>
            </w:tcPr>
            <w:p>
                <w:r>
                    <w:rPr><w:b/></w:rPr>
                    <w:t>Header Text</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr><!-- data row -->
        <w:tc><!-- data cell -->
            <w:p>
                <w:r>
                    <w:t>Cell Data</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>
```

## Test Coverage

The table feature includes comprehensive test coverage:

1. **TestInsertBasicTable** - Basic table insertion with minimal options
2. **TestInsertTableWithStyling** - Full styling options (colors, borders, alignment)
3. **TestInsertTableWithRepeatHeader** - Header repetition on multiple pages
4. **TestInsertTableInvalidRows** - Validation of mismatched column counts
5. **TestInsertTableNoColumns** - Validation of empty column definitions
6. **TestInsertTableCustomWidths** - Custom column width specifications

All tests verify XML structure and content in the generated document.

## Best Practices

1. **Column Widths**: Let Word auto-calculate widths unless you have specific requirements
2. **Colors**: Use professional color schemes (blues, grays) for business documents
3. **Border Sizes**: Keep borders thin (4-6) for professional appearance
4. **Header Repetition**: Enable `RepeatHeader` for tables that may span multiple pages
5. **Alternate Rows**: Use subtle colors (F2F2F2, E7E6E6) for readability
6. **Cell Alignment**: Align numbers right, text left, headers center
7. **Cell Padding**: Default padding (108 twips) works well for most tables

## Common Use Cases

### Data Reports
- Financial statements
- Sales reports
- Inventory lists
- Performance metrics

### Project Documentation
- Task lists
- Resource allocation
- Timeline tables
- Requirement matrices

### Business Documents
- Price lists
- Product catalogs
- Comparison tables
- Contact directories

## Error Handling

```go
if err := updater.InsertTable(opts); err != nil {
    // Common errors:
    // - "at least one column is required"
    // - "row N has X cells, expected Y"
    // - "column widths count must match columns count"
    // - "anchor text required for PositionAfterText"
    log.Fatalf("Failed to insert table: %v", err)
}
```

## Performance Considerations

- **Large Tables**: Tables with many rows (100+) generate large XML structures
- **Column Widths**: Auto-calculation is fast, custom widths are slightly faster
- **Styling**: Alternate row colors add minimal overhead
- **Memory**: Each table row consumes ~100-200 bytes in memory during generation

## Future Enhancements

Potential future features:
- Merged cells (colspan/rowspan)
- Cell borders (individual cell border control)
- Conditional formatting based on cell values
- Table templates/presets
- Excel-like formulas in cells
- Nested tables
- Image insertion in cells
- Text rotation in cells
