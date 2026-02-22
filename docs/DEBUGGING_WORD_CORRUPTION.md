# Debugging Word Corruption Issues

When LibreOffice opens a .docx file correctly but Microsoft Word shows "unreadable content", it indicates malformed XML that Word's stricter parser rejects.

## Quick Start - Automated Check

Run the debugging tool on your corrupted file:

```bash
go run debug_tool.go "BKR02-Alarm_Groups_Analytics_20251001-20251231 (17).docx"
```

This will:
- ✓ Validate all XML files in the DOCX
- ✓ Check for common XML issues
- ✓ Report specific errors with locations

## Common Issues and Fixes

### 1. **Invalid XML Characters**

**Problem:** Control characters (0x00-0x1F except tab, newline, carriage return) in XML

**How to find:**
```bash
# Extract DOCX and check for invalid chars
7z x your_file.docx -o./extracted
grep -r --text $'[\x00-\x08\x0B\x0C\x0E-\x1F]' ./extracted/word/
```

**Fix in library:** Escape or remove invalid characters before writing to XML
```go
// Add to helpers.go or create sanitizer
func sanitizeXMLText(s string) string {
    var result strings.Builder
    for _, r := range s {
        // Only allow valid XML 1.0 characters
        if r == 0x09 || r == 0x0A || r == 0x0D || (r >= 0x20 && r <= 0xD7FF) ||
           (r >= 0xE000 && r <= 0xFFFD) || (r >= 0x10000 && r <= 0x10FFFF) {
            result.WriteRune(r)
        }
    }
    return result.String()
}
```

### 2. **Missing or Incorrect Namespace Declarations**

**Problem:** Elements without proper namespace declarations

**Check:**
```bash
# Extract and inspect document.xml
unzip -p your_file.docx word/document.xml | head -20
```

Must have:
```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
```

**Fix:** Ensure namespace declarations are preserved when manipulating XML

### 3. **Broken Relationship IDs**

**Problem:** References to non-existent relationship IDs

**Check:**
```bash
# Compare relationship IDs
unzip -p your_file.docx word/_rels/document.xml.rels
unzip -p your_file.docx word/document.xml | grep 'r:id='
```

All `r:id="rId123"` in document.xml must exist in document.xml.rels

**Common in your library:** Chart relationships, image relationships

### 4. **Duplicate Relationship IDs**

**Problem:** Same rId used twice in .rels file

**Fix in library:** Check your relationship generation code:
```go
// In helpers.go or similar - ensure unique IDs
func getNextRelIDFromFile(relsPath string) (string, error) {
    // Parse relationships and return next rIdN
    // using max existing relationship ID + 1.
    // Reuse this helper across chart/image/header/footer/hyperlink flows.
}
```

### 5. **Invalid Content Types**

**Problem:** Missing or incorrect entries in [Content_Types].xml

**Check:**
```bash
unzip -p your_file.docx [Content_Types].xml
```

Each part must have a content type:
- Charts: `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`
- Images: `image/png`, `image/jpeg`, etc.

### 6. **Malformed Chart XML**

**Problem:** Charts are common corruption sources

**Check specific chart:**
```bash
unzip -p your_file.docx word/charts/chart1.xml
```

**Common issues:**
- Missing `<c:chartSpace>` root
- Unclosed tags in chart data
- Invalid number formatting

### 7. **Excel Embedded Workbooks**

**Problem:** Chart data workbooks with corrupt formulas or ranges

**Check:**
```bash
unzip -p your_file.docx word/embeddings/Microsoft_Excel_Worksheet1.xlsx
unzip -p Microsoft_Excel_Worksheet1.xlsx xl/worksheets/sheet1.xml
```

## Manual Inspection Method

### Step 1: Extract DOCX
```bash
# Create directory and extract
mkdir extracted_docx
cd extracted_docx
unzip ../your_file.docx
```

### Step 2: Validate Each XML File
```bash
# Windows PowerShell
Get-ChildItem -Recurse -Filter *.xml | ForEach-Object {
    Write-Host "Checking $($_.Name)..."
    $xml = New-Object System.Xml.XmlDocument
    try {
        $xml.Load($_.FullName)
        Write-Host "  ✓ Valid" -ForegroundColor Green
    } catch {
        Write-Host "  ❌ Invalid: $($_.Exception.Message)" -ForegroundColor Red
    }
}
```

### Step 3: Compare with Working DOCX
Create a simple document with your library that Word opens correctly, then compare:
```bash
diff -r good_docx/ bad_docx/
```

## Official Microsoft Tools

### Office Open XML SDK 2.5 Productivity Tool
Download: https://github.com/OfficeDev/Open-XML-SDK

This tool can:
- Validate DOCX files
- Show validation errors with exact locations
- Compare document structures
- Generate C# code from valid DOCX

### Steps:
1. Install the productivity tool
2. Open your corrupted DOCX
3. Click "Validate" in the ribbon
4. Review errors (will show exact element paths)

## Testing in Your Library

Add validation to your test suite:

```go
// Test helper to validate DOCX can be opened by stricter parsers
func TestDocxValidity(t *testing.T) {
    // Generate your document
    updater, err := New("templates/docx_template.docx")
    require.NoError(t, err)
    defer updater.Cleanup()
    
    // ... your operations ...
    
    err = updater.Save("test_output.docx")
    require.NoError(t, err)
    
    // Validate XML structure
    r, err := zip.OpenReader("test_output.docx")
    require.NoError(t, err)
    defer r.Close()
    
    for _, f := range r.File {
        if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
            rc, err := f.Open()
            require.NoError(t, err)
            
            content, err := io.ReadAll(rc)
            rc.Close()
            require.NoError(t, err)
            
            // Validate XML
            var doc interface{}
            err = xml.Unmarshal(content, &doc)
            require.NoError(t, err, "Invalid XML in %s", f.Name)
        }
    }
}
```

## Specific Areas to Review in Your Library

Based on your codebase:

1. **chart_xml.go** - Chart XML manipulation
   - Line 144: `updateChartXMLContent` - check namespace prefix handling
   - Ensure all generated XML is well-formed

2. **chart.go / helpers.go** - Chart insertion + relationship handling
   - Verify generated chart XML output is well-formed
   - Check shared relationship ID generation helper

3. **excel_handler.go** - Embedded Excel workbooks
   - Lines 122, 130: Unmarshal operations
   - Line 282: Marshal operation - might introduce issues

4. **Replace operations** - Text manipulation
   - Check if replacements break XML structure
   - Ensure special chars are escaped

## Next Steps

1. Run `debug_tool.go` on your corrupted file
2. If no XML errors found, use Microsoft's Productivity Tool
3. Compare corrupted DOCX with a minimal working example
4. Check the specific operation that created the file
5. Add XML validation to your test suite

## Prevention

Add this to all XML writing code:
```go
// Validate before writing
func validateXML(data []byte) error {
    var doc interface{}
    return xml.Unmarshal(data, &doc)
}
```
