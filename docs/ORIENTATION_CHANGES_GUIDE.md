# Page Orientation Changes with Tables - Complete Guide

This guide explains how to use section breaks with orientation changes when inserting tables in DOCX documents.

## Quick Start

```go
package main

import (
    "log"
    docx "github.com/falcomza/go-docx"
)

func main() {
    updater, _ := docx.New("template.docx")
    defer updater.Cleanup()

    // Add content in portrait
    updater.AddText("Introduction", docx.PositionEnd)

    // Switch to landscape BEFORE inserting wide table
    updater.InsertSectionBreak(docx.BreakOptions{
        Position:    docx.PositionEnd,
        SectionType: docx.SectionBreakNextPage,
        PageLayout:  docx.PageLayoutLetterLandscape(),
    })

    // Insert wide table in landscape
    updater.InsertTable(docx.TableOptions{
        Position: docx.PositionEnd,
        Columns: []docx.ColumnDefinition{
            {Title: "Col1"},
            {Title: "Col2"},
            // ... more columns
        },
        Rows: [][]string{
            {"Data1", "Data2"},
        },
    })

    // Switch back to portrait AFTER table
    updater.InsertSectionBreak(docx.BreakOptions{
        Position:    docx.PositionEnd,
        SectionType: docx.SectionBreakNextPage,
        PageLayout:  docx.PageLayoutLetterPortrait(),
    })

    // Continue with portrait content
    updater.AddText("Conclusion", docx.PositionEnd)

    updater.Save("output.docx")
}
```

## The Three-Step Pattern

### Step 1: Insert Section Break (Switch to Landscape)

```go
err := updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,  // New page required for orientation change
    PageLayout:  docx.PageLayoutLetterLandscape(),
})
```

**Key Points:**
- Use `SectionBreakNextPage` for orientation changes (not `SectionBreakContinuous`)
- The landscape section starts on a new page
- Choose appropriate page layout helper function

### Step 2: Insert Table

```go
err := updater.InsertTable(docx.TableOptions{
    Position: docx.PositionEnd,
    Columns: []docx.ColumnDefinition{
        {Title: "Employee ID", Alignment: docx.CellAlignLeft},
        {Title: "Full Name", Alignment: docx.CellAlignLeft},
        {Title: "Department", Alignment: docx.CellAlignLeft},
        {Title: "Position", Alignment: docx.CellAlignLeft},
        {Title: "Location", Alignment: docx.CellAlignCenter},
        {Title: "Salary", Alignment: docx.CellAlignRight},
        {Title: "Performance", Alignment: docx.CellAlignCenter},
    },
    Rows: [][]string{
        {"EMP001", "John Smith", "Engineering", "Developer", "NY", "$95,000", "Excellent"},
        {"EMP002", "Jane Doe", "Marketing", "Manager", "LA", "$87,500", "Very Good"},
    },
    HeaderBold:        true,
    HeaderBackground:  "2E75B5",
    HeaderAlignment:   docx.CellAlignCenter,
    AlternateRowColor: "E7E6E6",
    BorderStyle:       docx.BorderSingle,
    TableAlignment:    docx.AlignCenter,
    RepeatHeader:      true,
})
```

**Key Points:**
- Table is inserted in the current (landscape) section
- Wide tables fit better in landscape orientation
- Use `RepeatHeader: true` for multi-page tables

### Step 3: Insert Section Break (Return to Portrait)

```go
err := updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
```

**Key Points:**
- Returns to portrait orientation on a new page
- Following content will be in portrait mode
- Previous section properties are not affected

## Available Page Layouts

### Pre-configured Layouts

| Function | Paper Size | Orientation | Margins |
|----------|-----------|-------------|---------|
| `PageLayoutLetterPortrait()` | US Letter (8.5" × 11") | Portrait | 1" all around |
| `PageLayoutLetterLandscape()` | US Letter (11" × 8.5") | Landscape | 1" all around |
| `PageLayoutA4Portrait()` | A4 (210mm × 297mm) | Portrait | Default |
| `PageLayoutA4Landscape()` | A4 (297mm × 210mm) | Landscape | Default |
| `PageLayoutA3Portrait()` | A3 (297mm × 420mm) | Portrait | Default |
| `PageLayoutA3Landscape()` | A3 (420mm × 297mm) | Landscape | Default |
| `PageLayoutLegalPortrait()` | Legal (8.5" × 14") | Portrait | Default |

### Custom Layouts

```go
customLayout := &docx.PageLayoutOptions{
    PageWidth:    docx.PageWidthLetter,      // Width in twips (1/1440")
    PageHeight:   docx.PageHeightLetter,
    Orientation:  docx.OrientationLandscape,
    MarginTop:    docx.MarginNarrow,         // 0.5"
    MarginRight:  docx.MarginNarrow,
    MarginBottom: docx.MarginNarrow,
    MarginLeft:   docx.MarginNarrow,
    MarginHeader: docx.MarginHeaderFooterDefault,
    MarginFooter: docx.MarginHeaderFooterDefault,
    MarginGutter: 0,
}
```

## Section Break Types

| Type | Behavior | Use Case |
|------|----------|----------|
| `SectionBreakNextPage` | Start on new page | **Orientation changes**, new chapters |
| `SectionBreakContinuous` | Same page | Different margins/columns (same orientation) |
| `SectionBreakEvenPage` | Next even page | Double-sided printing (chapters start on right) |
| `SectionBreakOddPage` | Next odd page | Double-sided printing (chapters start on left) |

**Important:** Only `SectionBreakNextPage` should be used for orientation changes!

## Constants Reference

### Page Sizes (in twips, 1 twip = 1/1440 inch)

```go
const (
    // US Letter: 8.5" × 11"
    PageWidthLetter  = 12240
    PageHeightLetter = 15840

    // US Legal: 8.5" × 14"
    PageWidthLegal  = 12240
    PageHeightLegal = 20160

    // A4: 210mm × 297mm
    PageWidthA4  = 11906
    PageHeightA4 = 16838

    // A3: 297mm × 420mm
    PageWidthA3  = 16838
    PageHeightA3 = 23811
)
```

### Margin Sizes (in twips)

```go
const (
    MarginDefault             = 1440  // 1.0 inch
    MarginNarrow             = 720   // 0.5 inch
    MarginWide               = 2160  // 1.5 inches
    MarginHeaderFooterDefault = 720   // 0.5 inch
)
```

### Orientation Values

```go
const (
    OrientationPortrait  PageOrientation = "portrait"
    OrientationLandscape PageOrientation = "landscape"
)
```

## Complete Working Example

See `examples/example_table_orientation.go` for a complete working example that demonstrates:

1. Portrait introduction section
2. Landscape section with 8-column employee table
3. Portrait analysis section
4. A4 landscape section with 7-column quarterly sales table
5. Portrait conclusion section

Run the example:

```bash
go run examples/example_table_orientation.go
```

Verify the output:

```bash
./tools/verify_orientation.sh outputs/table_orientation_demo.docx
```

## Common Patterns

### Pattern 1: Single Wide Table

```go
// Portrait intro
updater.AddText("Introduction", docx.PositionEnd)

// Landscape for wide table
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterLandscape(),
})
updater.InsertTable(wideTableOptions)

// Back to portrait
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
updater.AddText("Conclusion", docx.PositionEnd)
```

### Pattern 2: Multiple Landscape Sections

```go
// Portrait section
updater.AddText("Chapter 1", docx.PositionEnd)

// First landscape section
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterLandscape(),
})
updater.InsertTable(table1Options)

// Back to portrait
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
updater.AddText("Chapter 2", docx.PositionEnd)

// Second landscape section
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutA4Landscape(),  // Different paper size
})
updater.InsertTable(table2Options)

// Final portrait section
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
updater.AddText("Conclusion", docx.PositionEnd)
```

### Pattern 3: Mixed Letter and A4 Formats

```go
// Start with Letter portrait
updater.SetPageLayout(*docx.PageLayoutLetterPortrait())
updater.AddText("US Letter section", docx.PositionEnd)

// Switch to A4 landscape for international report
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutA4Landscape(),
})
updater.InsertTable(internationalTableOptions)

// Back to Letter portrait
updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
```

## Troubleshooting

### Problem: Orientation not changing

**Solution:** Make sure you're using `SectionBreakNextPage`, not `SectionBreakContinuous`

```go
// ❌ Wrong - won't change orientation
updater.InsertSectionBreak(docx.BreakOptions{
    SectionType: docx.SectionBreakContinuous,
    PageLayout:  docx.PageLayoutLetterLandscape(),
})

// ✅ Correct - changes orientation on new page
updater.InsertSectionBreak(docx.BreakOptions{
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterLandscape(),
})
```

### Problem: Table appears in wrong orientation

**Cause:** Table inserted before section break

**Solution:** Insert section break BEFORE the table

```go
// ❌ Wrong order
updater.InsertTable(tableOptions)  // Still in portrait!
updater.InsertSectionBreak(landscapeBreak)

// ✅ Correct order
updater.InsertSectionBreak(landscapeBreak)
updater.InsertTable(tableOptions)  // Now in landscape
```

### Problem: All subsequent pages in landscape

**Cause:** Missing section break after table to return to portrait

**Solution:** Always add section break after landscape content

```go
// ✅ Complete workflow
updater.InsertSectionBreak(landscapeBreak)
updater.InsertTable(tableOptions)
updater.InsertSectionBreak(portraitBreak)  // ← Don't forget this!
```

## Testing Your Implementation

Use the provided verification script:

```bash
./tools/verify_orientation.sh your_output.docx
```

Expected output for correct implementation:
```
Analyzing: your_output.docx
==========================================

File Type:
your_output.docx: Microsoft OOXML

Landscape Sections: 2
Total Sections: 5
Tables Found: 2

Page Size Information:
----------------------
w:pgSz w:w="12240" w:h="15840"/              ← Portrait
w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/  ← Landscape
w:pgSz w:w="12240" w:h="15840"/              ← Portrait
w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/  ← A4 Landscape
w:pgSz w:w="12240" w:h="15840"/              ← Portrait

✓ Verification complete!
```

## API Reference

See the complete API documentation:

- **Main Documentation:** `docs/API_DOCUMENTATION.md`
- **Fiber Integration:** `docs/API_REFERENCE.md`
- **Section Breaks:** `breaks.go:40-78`
- **Page Layout Types:** `types.go:79-247`

## Related Examples

- `examples/example_table_orientation.go` - Complete working example (this guide)
- `examples/example_page_layout.go` - Page layout demonstrations
- `examples/example_table.go` - Table styling and formatting
- `examples/example_breaks.go` - Section and page break patterns
