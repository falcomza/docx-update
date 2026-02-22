# Conditional Cell Formatting

## Overview

The table functionality now supports conditional cell formatting, allowing you to dynamically style individual cells based on their content. This is particularly useful for status indicators, priority levels, performance metrics, and other data that requires visual differentiation.

## Features

- **Content-Based Styling**: Apply different backgrounds, font colors, and text formatting based on cell text
- **Case-Insensitive Matching**: Cell content is matched case-insensitively with automatic whitespace trimming
- **Flexible Style Override**: Conditional styles intelligently merge with row-level styles
- **Multiple Conditions**: Support for multiple conditional styles in the same table

## Usage

### Basic Example

```go
err := updater.InsertTable(godocx.TableOptions{
    Position: godocx.PositionEnd,
    Columns: []godocx.ColumnDefinition{
        {Title: "Service"},
        {Title: "Status"},
    },
    Rows: [][]string{
        {"Database", "Critical"},
        {"API", "Normal"},
        {"Cache", "Warning"},
    },
    // Define conditional styles
    ConditionalStyles: map[string]godocx.CellStyle{
        "Critical": {
            Background: "FF0000", // Red
            FontColor:  "FFFFFF", // White text
            Bold:       true,
        },
        "Warning": {
            Background: "FFA500", // Orange
            FontColor:  "000000",
        },
        "Normal": {
            Background: "00B050", // Green
            FontColor:  "FFFFFF",
        },
    },
})
```

### Advanced Example with Row Styles

```go
err := updater.InsertTable(godocx.TableOptions{
    Position: godocx.PositionEnd,
    Columns: []godocx.ColumnDefinition{
        {Title: "Metric"},
        {Title: "Rating"},
        {Title: "Value"},
    },
    Rows: [][]string{
        {"CPU Usage", "Good", "45%"},
        {"Memory", "Poor", "85%"},
        {"Disk I/O", "Excellent", "15%"},
    },
    // Row-level styling (applies to all cells by default)
    RowStyle: godocx.CellStyle{
        FontSize:   20,       // 10pt
        Background: "F2F2F2", // Light gray default
    },
    AlternateRowColor: "E7E6E6",
    // Conditional styles override row styles for matching cells
    ConditionalStyles: map[string]godocx.CellStyle{
        "Excellent": {
            Background: "00B050", // Green (overrides row background)
            FontColor:  "FFFFFF",
            Bold:       true,
        },
        "Good": {
            Background: "92D050", // Light green
        },
        "Poor": {
            Background: "FF0000", // Red
            FontColor:  "FFFFFF",
            Bold:       true,
        },
    },
})
```

## Style Merging Rules

When a cell's content matches a conditional style key:

1. **Background Color**: Conditional background overrides row background if specified
2. **Font Color**: Conditional font color overrides row font color if specified
3. **Font Size**: Conditional font size overrides row font size if specified (>0)
4. **Bold/Italic**: Uses OR logic - cell is bold/italic if either row or conditional style specifies it

### Examples:

| Scenario | Row Style | Conditional Style | Result |
|----------|-----------|-------------------|--------|
| Background only | `Background: "F2F2F2"` | `Background: "FF0000"` | Red background, row's other styles |
| Font color only | `FontColor: "000000"` | `FontColor: "FFFFFF"` | White text, row's other styles |
| Bold combination | `Bold: false` | `Bold: true` | **Bold** (OR logic) |
| Complete override | `FontSize: 20` | `Background: "FF0000", FontColor: "FFFFFF", Bold: true` | All conditional styles applied |

## Matching Behavior

- **Case-Insensitive**: "Critical", "CRITICAL", and "critical" all match
- **Whitespace Trimming**: "  Normal  " matches "Normal"
- **Exact Match**: Only exact matches after normalization are styled
- **First Match Wins**: If multiple keys could match (edge case), first found is used

## Use Cases

### Status Indicators
```go
ConditionalStyles: map[string]godocx.CellStyle{
    "Critical":     {Background: "FF0000", FontColor: "FFFFFF", Bold: true},
    "Warning":      {Background: "FFA500", FontColor: "000000"},
    "Non-critical": {Background: "FFD966", FontColor: "000000"},
    "Normal":       {Background: "00B050", FontColor: "FFFFFF"},
}
```

### Priority Levels
```go
ConditionalStyles: map[string]godocx.CellStyle{
    "High":   {Background: "FF6B6B", FontColor: "000000", Bold: true},
    "Medium": {Background: "FFE066", FontColor: "000000"},
    "Low":    {Background: "B4C7E7", FontColor: "000000"},
}
```

### Performance Ratings
```go
ConditionalStyles: map[string]godocx.CellStyle{
    "Excellent": {Background: "00B050", FontColor: "FFFFFF", Bold: true},
    "Good":      {Background: "92D050", FontColor: "000000"},
    "Fair":      {Background: "FFC000", FontColor: "000000"},
    "Poor":      {Background: "FF0000", FontColor: "FFFFFF", Bold: true},
}
```

## Color Reference

Common colors in hex format (without #):

| Color | Hex Code | Use Case |
|-------|----------|----------|
| Red | `FF0000` | Critical, Error, High priority |
| Orange | `FFA500` | Warning, Medium priority |
| Yellow | `FFD966` | Caution, Attention needed |
| Green | `00B050` | Success, Normal, Good |
| Light Green | `92D050` | Acceptable, Fair |
| Blue | `4472C4` | Information, Headers |
| Light Gray | `F2F2F2` | Default, Neutral |
| White | `FFFFFF` | Text on dark backgrounds |
| Black | `000000` | Text on light backgrounds |

## Complete Example

See [examples/example_conditional_cell_colors.go](../examples/example_conditional_cell_colors.go) for a comprehensive demonstration including:
- System status monitoring with color-coded statuses
- Issue tracking with priority-based formatting
- Performance metrics with threshold-based coloring

## API Reference

### ConditionalStyles Field

```go
type TableOptions struct {
    // ... other fields ...
    
    // Conditional cell styling based on content
    // Map keys are matched case-insensitively against cell text
    // Matching cells will have their style overridden by the map value
    // Non-empty conditional values take precedence over row-level styling
    ConditionalStyles map[string]CellStyle
}
```

### CellStyle Structure

```go
type CellStyle struct {
    Bold       bool
    Italic     bool
    FontSize   int    // Font size in half-points (e.g., 20 = 10pt)
    FontColor  string // Hex color (e.g., "000000")
    Background string // Hex color for cell background
}
```

## Testing

The feature includes comprehensive tests:
- `TestInsertTableWithConditionalCellColors`: Verifies multiple conditional styles work correctly
- `TestInsertTableConditionalCaseInsensitive`: Ensures case-insensitive matching
- `TestInsertTableConditionalWithRowStyle`: Tests style merging behavior

Run tests:
```bash
go test -v -run TestInsertTableConditional
```

## Limitations

- Only exact text matches are supported (no regex or partial matching)
- Conditional styles apply to the entire cell content, not individual words
- Map iteration order in Go is random, but in practice this only matters if you have duplicate keys (which is impossible in a map)

## Tips

1. **Use descriptive keys**: Make your conditional style keys match your data exactly
2. **Consider whitespace**: The matcher trims whitespace, so "  Critical  " will match "Critical"
3. **Use contrasting colors**: Ensure text is readable on colored backgrounds
4. **Test accessibility**: Consider color-blind users when choosing color schemes
5. **Combine with row styles**: Use row styles for defaults and conditionals for exceptions
