# Section Break Orientation Fix

## Problem

Tables were appearing in portrait sections instead of landscape sections, despite using `InsertSectionBreak` with landscape page layout options.

## Root Cause

The issue was with how section breaks work in OpenXML:

1. **Section properties apply to content BEFORE the `<w:sectPr>` element**, not after
2. A `<w:sectPr>` element **ends** the current section and defines its properties
3. Content AFTER a section break is in the NEXT section

## Example of Incorrect Understanding

If you want a landscape table, you might think:
```go
InsertSectionBreak(landscape)  // Start landscape section
InsertTable()                   // Table in landscape
```

But this creates:
- Previous section ends with landscape properties
- Table goes in NEXT section (not defined yet)

## Correct Pattern

To have content in a specific orientation:

```go
// 1. Add content
InsertTable()

// 2. Insert section break WITH THE ORIENTATION YOU WANT FOR THAT CONTENT
InsertSectionBreak(landscape)  // Ends current section as LANDSCAPE

// 3. Add more content (in next section)
AddText("Analysis")

// 4. End that section
InsertSectionBreak(portrait)  // Ends analysis section as PORTRAIT
```

## Document Structure

For this pattern:
- Section 1 (portrait): Introduction + heading
- Section 2 (landscape): Table
- Section 3 (portrait): Analysis

You need:
```go
AddText("Introduction")
AddText("Heading")
InsertSectionBreak(portrait)   // Ends Section 1 as portrait

InsertTable()
InsertSectionBreak(landscape)  // Ends Section 2 as landscape

AddText("Analysis")
InsertSectionBreak(portrait)   // Ends Section 3 as portrait
```

## Template Considerations

The empty template must NOT have a document-level `<w:sectPr>` at the beginning, otherwise it creates an extra empty section. The template should:

1. Have NO `<w:sectPr>` in the `<w:body>` initially
2. Let the code insert ALL section breaks
3. Include a final section break at the end to close the last section

## Fixed Files

1. **`templates/empty_template.docx`**: Updated to have no initial sectPr
2. **`examples/example_table_orientation.go`**: Corrected section break insertion order
3. **`paragraph.go`**: Already fixed `insertAtBodyEnd()` to handle section breaks correctly

## Verification

Run the example and verify:
```bash
go run examples/example_table_orientation.go
./tools/verify_orientation.sh ./outputs/table_orientation_demo.docx
```

Expected output:
- Landscape Sections: 2
- Total Sections: 5
- Tables Found: 2

Both tables should be in landscape sections (Sections 2 and 4).
