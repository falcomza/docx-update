# Chart Copying Guide

## Overview

The `CopyChart` method allows you to dynamically duplicate existing charts in a DOCX document. This is particularly useful when:

- You need to generate reports with variable numbers of charts
- Each subsystem/section requires similar chart formatting
- The exact number of charts is determined at runtime

## Basic Usage

```go
// Copy chart 1 and insert after paragraph containing "Section 2"
newChartIndex, err := updater.CopyChart(1, "Section 2")
if err != nil {
    log.Fatal(err)
}

// Update the new chart with your data
err = updater.UpdateChart(newChartIndex, myData)
```

## How It Works

When you call `CopyChart(sourceIndex, afterText)`, the library:

1. **Finds the next available chart index** by scanning existing chart files
2. **Copies the chart XML** (`word/charts/chart1.xml` → `word/charts/chart2.xml`)
3. **Copies the chart relationships** (`word/charts/_rels/chart1.xml.rels` → `word/charts/_rels/chart2.xml.rels`)
4. **Copies the embedded Excel workbook** with a unique filename
5. **Updates the chart relationships** to point to the new workbook
6. **Inserts chart reference into document.xml** after the specified text
7. **Adds document relationship** in `word/_rels/document.xml.rels`
8. **Updates content types** in `[Content_Types].xml`

## Parameters

### `sourceChartIndex` (int)
The 1-based index of the chart to copy. For example:
- `1` = first chart in the document
- `2` = second chart, etc.

### `afterText` (string)
Text within a paragraph after which to insert the new chart. The library:
- Searches for this exact text in `document.xml`
- Finds the enclosing paragraph (`<w:p>...</w:p>`)
- Inserts the new chart immediately after that paragraph

**Important**: The text must exist in the document, or an error will be returned.

## Return Value

Returns the new chart's index (1-based integer), which you can use with `UpdateChart()` to populate it with data.

## Best Practices

### 1. Use a Template Chart

Keep one "template" chart in your DOCX with the desired formatting (colors, fonts, axis labels, etc.). Copy this chart for each section.

```go
// Chart 1 is the template
for i, section := range sections {
    var chartIndex int
    
    if i == 0 {
        chartIndex = 1 // Use template directly
    } else {
        chartIndex, _ = updater.CopyChart(1, section.Marker)
    }
    
    updater.UpdateChart(chartIndex, section.Data)
}
```

### 2. Choose Unique Marker Text

Use distinct paragraph text for insertion points to avoid ambiguity:

```
✓ Good: "## Database Performance", "## Authentication Metrics"
✗ Bad:  "Performance", "Metrics" (may appear multiple times)
```

### 3. Error Handling

Always check for errors when copying charts:

```go
newIndex, err := updater.CopyChart(1, markerText)
if err != nil {
    if strings.Contains(err.Error(), "not found") {
        // Marker text doesn't exist in document
        log.Printf("Warning: Skipping chart - marker not found: %s", markerText)
    } else {
        return fmt.Errorf("failed to copy chart: %w", err)
    }
}
```

### 4. Performance Considerations

Chart copying involves:
- File I/O operations (copying chart XML and workbooks)
- XML parsing and manipulation
- Document structure updates

For large batches, consider:
- Copying charts in sequence (not parallel) to avoid race conditions
- Monitoring memory usage if creating 50+ charts
- Testing with your specific chart complexity

## Common Patterns

### Pattern 1: Fixed Sections with Variable Data

```go
sections := []string{"Overview", "Authentication", "Database", "API"}

for i, section := range sections {
    chartIndex := i + 1
    if i > 0 {
        chartIndex, _ = updater.CopyChart(1, section)
    }
    
    data := fetchDataForSection(section)
    updater.UpdateChart(chartIndex, data)
}
```

### Pattern 2: Dynamic Subsystems

```go
// Number of subsystems determined at runtime
subsystems := discoverSubsystems()

for i, subsystem := range subsystems {
    var chartIndex int
    
    if i == 0 {
        chartIndex = 1
    } else {
        // All charts inserted after the same marker
        chartIndex, _ = updater.CopyChart(1, "Report Content")
    }
    
    updater.UpdateChart(chartIndex, subsystem.GenerateChartData())
}
```

### Pattern 3: Conditional Chart Generation

```go
chartIndex := 1

for _, subsystem := range subsystems {
    if !subsystem.HasData() {
        continue // Skip subsystems without data
    }
    
    if chartIndex > 1 {
        chartIndex, _ = updater.CopyChart(1, subsystem.Name)
    }
    
    updater.UpdateChart(chartIndex, subsystem.Data)
    chartIndex++
}
```

## Limitations

1. **Text-based positioning**: Chart insertion uses text search, which requires the marker text to exist
2. **Sequential insertion**: All charts are inserted after their respective markers, but relative order depends on marker positions in the document
3. **Template dependency**: The source chart must exist and be valid
4. **Single insertion point per call**: Each `CopyChart` call inserts one chart; use a loop for multiple copies

## Troubleshooting

### Error: "text not found in document"

**Cause**: The `afterText` parameter doesn't match any text in the document.

**Solution**: 
- Open the DOCX and verify the exact text exists
- Check for typos, case sensitivity, or extra whitespace
- Use a more unique text snippet

### Error: "chart file does not exist"

**Cause**: The source chart index is invalid.

**Solution**:
- Verify your template has a chart at the specified index
- Chart indices are 1-based (first chart = 1)

### Charts appear in unexpected locations

**Cause**: The marker text appears multiple times in the document.

**Solution**:
- Use more specific marker text
- Consider adding unique identifiers to each section

## Advanced: Custom Chart Positioning

If you need more control over chart positioning, you can:

1. **Use XML paths** instead of text search (requires modifying the library)
2. **Pre-process the document** to add unique marker paragraphs
3. **Post-process with additional tools** to reorder elements

## Example: Multi-Page Report

```go
type ReportSection struct {
    Title  string
    Marker string
    Data   ChartData
}

sections := []ReportSection{
    {Title: "Q1 Results", Marker: "Q1 PERFORMANCE", Data: q1Data},
    {Title: "Q2 Results", Marker: "Q2 PERFORMANCE", Data: q2Data},
    {Title: "Q3 Results", Marker: "Q3 PERFORMANCE", Data: q3Data},
    {Title: "Q4 Results", Marker: "Q4 PERFORMANCE", Data: q4Data},
}

u, _ := updater.New("annual_report_template.docx")
defer u.Cleanup()

for i, section := range sections {
    chartIndex := 1
    
    if i > 0 {
        chartIndex, err = u.CopyChart(1, section.Marker)
        if err != nil {
            log.Printf("Skipping %s: %v", section.Title, err)
            continue
        }
    }
    
    section.Data.ChartTitle = section.Title
    u.UpdateChart(chartIndex, section.Data)
}

u.Save("annual_report_2024.docx")
```

## See Also

- [README.md](README.md) - Main documentation
- [example_multi_subsystem.go](example_multi_subsystem.go) - Complete working example
- [src/chart_copy_test.go](src/chart_copy_test.go) - Test cases demonstrating usage
