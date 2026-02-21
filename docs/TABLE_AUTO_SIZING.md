# Table Column Auto-Sizing and Page Margin Constraints

## Overview

Tables in Word documents automatically size their columns while respecting page margins. The docx-chart-updater library now properly implements this behavior by calculating column widths to be constrained by the available page width (page width minus left and right margins).

## How Table Sizing Works

### Available Page Width Concept

The **available width** is the usable space for content between the left and right margins:

```
Available Width = Page Width - Left Margin - Right Margin
```

For a standard US Letter portrait page (8.5" × 11") with 1" margins:
```
Available Width = 12240 - 1440 - 1440 = 9360 twips
```

(Note: 1 twip = 1/1440 inch, so 1" = 1440 twips)

### Width Modes

Tables support three width modes in `TableOptions`:

#### 1. Percentage Mode (Default)

- **Type**: `TableWidthPercentage`
- **Usage**: Table width is specified as a percentage of available width
- **Values**: 5000 = 100%, 2500 = 50%, etc.
- **Column Calculation**: 
  ```
  Column Width = (Available Width × Table Percentage) / 5000 / Number of Columns
  ```

**Example**: Default table (100% width) with 3 columns on Letter portrait:
```
Column Width = (9360 × 5000) / 5000 / 3 = 3120 twips
```

```go
err := u.InsertTable(docxupdater.TableOptions{
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Name"},
        {Title: "Email"},
        {Title: "Phone"},
    },
    Rows: [][]string{
        {"Alice", "alice@example.com", "555-0101"},
    },
    HeaderBold: true,
    // Uses Default: 100% width, auto-sized columns
})
```

#### 2. Fixed Mode

- **Type**: `TableWidthFixed`
- **Usage**: Table has an absolute width in twips
- **Values**: Width in twips (e.g., 7200 for 5 inches)
- **Column Calculation**: 
  ```
  Column Width = Table Width / Number of Columns
  ```

**Example**: 5-inch wide table with 3 columns:
```
Column Width = 7200 / 3 = 2400 twips
```

```go
err := u.InsertTable(docxupdater.TableOptions{
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Col 1"},
        {Title: "Col 2"},
        {Title: "Col 3"},
    },
    Rows: [][]string{
        {"A", "B", "C"},
    },
    HeaderBold: true,
    TableWidthType: docxupdater.TableWidthFixed,
    TableWidth:     7200, // 5 inches
})
```

#### 3. Auto Mode

- **Type**: `TableWidthAuto`
- **Usage**: Word automatically sizes based on content
- **Column Calculation**: Uses a reasonable default width
- **Note**: Actual sizing determined by Word when document is opened

```go
err := u.InsertTable(docxupdater.TableOptions{
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Col 1"},
        {Title: "Col 2"},
    },
    Rows: [][]string{
        {"A", "B"},
    },
    HeaderBold: true,
    TableWidthType: docxupdater.TableWidthAuto,
})
```

### Explicit Column Widths

When you specify column widths explicitly, they are used as-is without applying general constraints:

```go
err := u.InsertTable(docxupdater.TableOptions{
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Narrow"},
        {Title: "Wide"},
        {Title: "Medium"},
    },
    ColumnWidths: []int{1500, 3000, 2000}, // Explicit widths in twips
    Rows: [][]string{
        {"A", "B", "C"},
    },
    HeaderBold: true,
})
```

## AvailableWidth Parameter

For percentage-mode tables, the library uses a default available width calculation:

```
Default = 9360 twips (Letter portrait with 1" margins)
```

However, if you have different page layouts or margins, you can **explicitly specify** the available width:

```go
// Custom page layout with narrow margins (0.5" each)
// Available width = 12240 - 720 - 720 = 10800
err := u.InsertTable(docxupdater.TableOptions{
    Columns: []docxupdater.ColumnDefinition{
        {Title: "A"},
        {Title: "B"},
        {Title: "C"},
    },
    Rows: [][]string{
        {"1", "2", "3"},
    },
    HeaderBold:     true,
    AvailableWidth: 10800, // For narrow (0.5") margins
})
```

### Common Available Widths

| Page Format | Margins | Available Width |
|---|---|---|
| Letter (8.5" × 11") | 1" (default) | 9360 twips |
| Letter (8.5" × 11") | 0.5" (narrow) | 10800 twips |
| Letter (8.5" × 11") | 1.5" (wide) | 7920 twips |
| A4 (210mm × 297mm) | Default | ~8200 twips |
| A3 (297mm × 420mm) | Default | ~12000 twips |

## Common Use Cases

### Scenario 1: Full-Width Table with Default Margins

```go
// Automatically spans full page between left and right margins
err := u.InsertTable(docxupdater.TableOptions{
    Position: docxupdater.PositionEnd,
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Product"},
        {Title: "Price"},
        {Title: "Stock"},
        {Title: "Status"},
    },
    Rows: [][]string{
        {"Widget A", "$19.99", "150", "In Stock"},
        {"Widget B", "$29.99", "0", "Out of Stock"},
    },
    HeaderBold: true,
    // Columns auto-sized to fill available width
})
```

**Result**: 4 columns × 2340 twips each = full width

### Scenario 2: Half-Width Table (Side-by-Side Layout)

```go
err := u.InsertTable(docxupdater.TableOptions{
    Position: docxupdater.PositionEnd,
    Columns: []docxupdater.ColumnDefinition{
        {Title: "Col 1"},
        {Title: "Col 2"},
    },
    Rows: [][]string{
        {"A", "B"},
    },
    HeaderBold: true,
    TableWidthType: docxupdater.TableWidthPercentage,
    TableWidth:     2500, // 50% width
    // Can place two such tables side-by-side
})
```

**Result**: 2 columns × 2340 twips each = 50% width

### Scenario 3: Fixed-Width Table

```go
err := u.InsertTable(docxupdater.TableOptions{
    Position: docxupdater.PositionEnd,
    Columns: []docxupdater.ColumnDefinition{
        {Title: "ID"},
        {Title: "Name"},
    },
    Rows: [][]string{
        {"001", "Item One"},
    },
    HeaderBold: true,
    TableWidthType: docxupdater.TableWidthFixed,
    TableWidth:     4320, // 3 inches wide
    // Table size independent of page margins
})
```

**Result**: Fixed width of 3 inches

## Page Margin Constants

The package provides constants for common margin sizes:

```go
const (
    MarginDefault             = 1440  // 1.0 inch
    MarginNarrow              = 720   // 0.5 inch
    MarginWide                = 2160  // 1.5 inches
    MarginHeaderFooterDefault = 720   // 0.5 inch
)
```

Page size constants:

```go
const (
    PageWidthLetter  = 12240  // 8.5"
    PageHeightLetter = 15840  // 11"
    PageWidthA4      = 11906  // 210mm
    PageHeightA4     = 16838  // 297mm
    // ... more sizes
)
```

## Implementation Details

### Column Width Calculation Formula

**Percentage Mode**:
```
Grid Column Width = (Available Width × Table Width Percentage) / 5000 / Num Columns
```

**Fixed Mode**:
```
Grid Column Width = Table Width / Num Columns
```

**Auto Mode**:
```
Grid Column Width = 11520 / Num Columns  // ~8 inches distributed
```

### Word XML Structure

The implementation sets proper WordprocessingML elements:

```xml
<w:tbl>
  <w:tblPr>
    <!-- Table width (percentage, fixed, or auto) -->
    <w:tblW w:w="5000" w:type="pct"/>
  </w:tblPr>
  
  <!-- Column definitions with calculated widths -->
  <w:tblGrid>
    <w:gridCol w:w="3120"/>  <!-- Column 1: 3120 twips -->
    <w:gridCol w:w="3120"/>  <!-- Column 2: 3120 twips -->
    <w:gridCol w:w="3120"/>  <!-- Column 3: 3120 twips -->
  </w:tblGrid>
  
  <!-- Header and data rows -->
  <w:tr>...</w:tr>
</w:tbl>
```

## Best Practices

1. **Use percentage mode for flexible layouts**: Let tables expand/contract with different page sizes
2. **Specify AvailableWidth when using non-standard margins**: Ensures columns size correctly
3. **Use fixed widths for consistent layouts**: When table size must remain constant
4. **Test with different page layouts**: Verify tables display correctly in portrait, landscape, and various margins
5. **Explicit widths for specific requirements**: When you need precise column widths

## Testing

The auto-sizing behavior is verified by comprehensive tests:

- `TestTableAutoSizingWithPageMargins`: Standard Letter portrait
- `TestTableAutoSizingWithNarrowMargins`: Narrow 0.5" margins
- `TestTableAutoSizingWithPercentageWidths`: 50% width tables
- `TestTableFixedWidthConstrained`: Fixed-width tables
- `TestTableExplicitColumnWidths`: User-specified widths

All tests verify that generated column widths match expected calculations based on page constraints.

## Migration Notes

If you have existing code using percentage mode, the column widths will now be calculated correctly. The visible effect in Word should be that tables appear larger (filling more of the available space) since column widths are now properly constrained by page margins.

This is the **correct** behavior according to Word's auto-sizing model.
