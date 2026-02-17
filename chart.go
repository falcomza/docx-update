package docxupdater

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// ChartKind defines the type of chart
type ChartKind string

const (
	ChartKindColumn ChartKind = "barChart"  // Column chart (vertical bars)
	ChartKindBar    ChartKind = "barChart"  // Bar chart (horizontal bars)
	ChartKindLine   ChartKind = "lineChart" // Line chart
	ChartKindPie    ChartKind = "pieChart"  // Pie chart
	ChartKindArea   ChartKind = "areaChart" // Area chart
)

// ChartOptions defines comprehensive options for chart creation
type ChartOptions struct {
	// Position where to insert the chart
	Position InsertPosition
	Anchor   string // Text anchor for relative positioning

	// Chart type (default: Column)
	ChartKind ChartKind

	// Chart titles
	Title             string // Main chart title
	CategoryAxisTitle string // X-axis title (horizontal axis)
	ValueAxisTitle    string // Y-axis title (vertical axis)

	// Data
	Categories   []string     // Category labels (X-axis)
	Series       []SeriesData // Data series with names and values
	ShowLegend   bool         // Show legend (default: true)
	LegendPosition string     // Legend position: "r" (right), "l" (left), "t" (top), "b" (bottom)

	// Chart dimensions (default: spans between margins)
	Width  int // Width in EMUs (English Metric Units), 0 for default (6099523 = ~6.5")
	Height int // Height in EMUs, 0 for default (3340467 = ~3.5")

	// Caption options (nil for no caption)
	Caption *CaptionOptions
}

// InsertChart creates a new chart and inserts it into the document
func (u *Updater) InsertChart(opts ChartOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}

	// Validate options
	if err := validateChartOptions(opts); err != nil {
		return fmt.Errorf("invalid chart options: %w", err)
	}

	// Apply defaults
	opts = applyChartDefaults(opts)

	// Find next available chart index
	chartIndex := u.findNextChartIndex()

	// Create chart XML file
	chartPath := filepath.Join(u.tempDir, "word", "charts", fmt.Sprintf("chart%d.xml", chartIndex))
	if err := u.createChartXML(chartPath, opts); err != nil {
		return fmt.Errorf("create chart xml: %w", err)
	}

	// Create embedded workbook
	workbookPath := filepath.Join(u.tempDir, "word", "embeddings", fmt.Sprintf("Microsoft_Excel_Worksheet%d.xlsx", chartIndex))
	if err := u.createEmbeddedWorkbook(workbookPath, opts); err != nil {
		return fmt.Errorf("create embedded workbook: %w", err)
	}

	// Create chart relationships file
	chartRelsPath := filepath.Join(u.tempDir, "word", "charts", "_rels", fmt.Sprintf("chart%d.xml.rels", chartIndex))
	if err := u.createChartRelationships(chartRelsPath, workbookPath); err != nil {
		return fmt.Errorf("create chart relationships: %w", err)
	}

	// Add chart relationship to document.xml.rels
	relID, err := u.addChartRelationship(chartIndex)
	if err != nil {
		return fmt.Errorf("add chart relationship: %w", err)
	}

	// Insert chart drawing into document
	if err := u.insertChartDrawing(chartIndex, relID, opts); err != nil {
		return fmt.Errorf("insert chart drawing: %w", err)
	}

	// Update content types
	if err := u.addContentTypeOverride(chartIndex); err != nil {
		return fmt.Errorf("add content type: %w", err)
	}

	return nil
}

// validateChartOptions validates chart creation options
func validateChartOptions(opts ChartOptions) error {
	if len(opts.Categories) == 0 {
		return fmt.Errorf("categories cannot be empty")
	}
	if len(opts.Series) == 0 {
		return fmt.Errorf("at least one series is required")
	}

	// Validate series
	for i, series := range opts.Series {
		if strings.TrimSpace(series.Name) == "" {
			return fmt.Errorf("series[%d] name cannot be empty", i)
		}
		if len(series.Values) != len(opts.Categories) {
			return fmt.Errorf("series[%d] values length (%d) must match categories length (%d)", i, len(series.Values), len(opts.Categories))
		}
	}

	return nil
}

// applyChartDefaults sets default values for unspecified options
func applyChartDefaults(opts ChartOptions) ChartOptions {
	if opts.ChartKind == "" {
		opts.ChartKind = ChartKindColumn
	}
	if opts.Width == 0 {
		opts.Width = 6099523 // ~6.5 inches (spans between margins on letter-size page)
	}
	if opts.Height == 0 {
		opts.Height = 3340467 // ~3.5 inches
	}
	if opts.ShowLegend && opts.LegendPosition == "" {
		opts.LegendPosition = "r" // Right by default
	}
	return opts
}

// createChartXML generates the chart XML file
func (u *Updater) createChartXML(chartPath string, opts ChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(chartPath), 0o755); err != nil {
		return fmt.Errorf("create charts directory: %w", err)
	}

	xml := generateChartXML(opts)

	if err := os.WriteFile(chartPath, xml, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// generateChartXML creates the chart XML content
func generateChartXML(opts ChartOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	
	buf.WriteString(`<c:chart>`)

	// Chart title
	if opts.Title != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx>`)
		buf.WriteString(`<c:rich>`)
		buf.WriteString(`<a:bodyPr/>`)
		buf.WriteString(`<a:lstStyle/>`)
		buf.WriteString(`<a:p>`)
		buf.WriteString(`<a:pPr><a:defRPr/></a:pPr>`)
		buf.WriteString(`<a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.Title))
		buf.WriteString(`</a:t></a:r>`)
		buf.WriteString(`</a:p>`)
		buf.WriteString(`</c:rich>`)
		buf.WriteString(`</c:tx>`)
		buf.WriteString(`<c:layout/>`)
		buf.WriteString(`<c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}

	buf.WriteString(`<c:autoTitleDeleted val="0"/>`)
	buf.WriteString(`<c:plotArea>`)
	buf.WriteString(`<c:layout/>`)

	// Generate chart type specific content
	switch opts.ChartKind {
	case ChartKindColumn:
		buf.WriteString(generateColumnChartXML(opts))
	default:
		buf.WriteString(generateColumnChartXML(opts)) // Default to column
	}

	// Category axis
	buf.WriteString(`<c:catAx>`)
	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
	buf.WriteString(`<c:delete val="0"/>`)
	buf.WriteString(`<c:axPos val="b"/>`)
	if opts.CategoryAxisTitle != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.CategoryAxisTitle))
		buf.WriteString(`</a:t></a:r></a:p></c:rich></c:tx>`)
		buf.WriteString(`<c:layout/><c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}
	buf.WriteString(`<c:numFmt formatCode="General" sourceLinked="1"/>`)
	buf.WriteString(`<c:majorTickMark val="out"/>`)
	buf.WriteString(`<c:minorTickMark val="none"/>`)
	buf.WriteString(`<c:tickLblPos val="nextTo"/>`)
	buf.WriteString(`<c:crossAx val="2071991240"/>`)
	buf.WriteString(`<c:crosses val="autoZero"/>`)
	buf.WriteString(`<c:auto val="1"/>`)
	buf.WriteString(`<c:lblAlgn val="ctr"/>`)
	buf.WriteString(`<c:lblOffset val="100"/>`)
	buf.WriteString(`</c:catAx>`)

	// Value axis
	buf.WriteString(`<c:valAx>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
	buf.WriteString(`<c:delete val="0"/>`)
	buf.WriteString(`<c:axPos val="l"/>`)
	if opts.ValueAxisTitle != "" {
		buf.WriteString(`<c:title>`)
		buf.WriteString(`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>`)
		buf.WriteString(xmlEscape(opts.ValueAxisTitle))
		buf.WriteString(`</a:t></a:r></a:p></c:rich></c:tx>`)
		buf.WriteString(`<c:layout/><c:overlay val="0"/>`)
		buf.WriteString(`</c:title>`)
	}
	buf.WriteString(`<c:numFmt formatCode="General" sourceLinked="1"/>`)
	buf.WriteString(`<c:majorTickMark val="out"/>`)
	buf.WriteString(`<c:minorTickMark val="none"/>`)
	buf.WriteString(`<c:tickLblPos val="nextTo"/>`)
	buf.WriteString(`<c:crossAx val="2071991400"/>`)
	buf.WriteString(`<c:crosses val="autoZero"/>`)
	buf.WriteString(`<c:crossBetween val="between"/>`)
	buf.WriteString(`</c:valAx>`)

	buf.WriteString(`</c:plotArea>`)

	// Legend
	if opts.ShowLegend {
		buf.WriteString(`<c:legend>`)
		buf.WriteString(fmt.Sprintf(`<c:legendPos val="%s"/>`, opts.LegendPosition))
		buf.WriteString(`<c:layout/>`)
		buf.WriteString(`<c:overlay val="0"/>`)
		buf.WriteString(`</c:legend>`)
	}

	buf.WriteString(`<c:plotVisOnly val="1"/>`)
	buf.WriteString(`<c:dispBlanksAs val="gap"/>`)
	buf.WriteString(`<c:showDLblsOverMax val="0"/>`)
	
	buf.WriteString(`</c:chart>`)

	// External data reference
	buf.WriteString(`<c:externalData r:id="rId1">`)
	buf.WriteString(`<c:autoUpdate val="0"/>`)
	buf.WriteString(`</c:externalData>`)

	buf.WriteString(`</c:chartSpace>`)

	return buf.Bytes()
}

// generateColumnChartXML generates column chart specific XML
func generateColumnChartXML(opts ChartOptions) string {
	var buf bytes.Buffer

	buf.WriteString(`<c:barChart>`)
	buf.WriteString(`<c:barDir val="col"/>`) // Column direction (col=vertical, bar=horizontal)
	buf.WriteString(`<c:grouping val="clustered"/>`)
	buf.WriteString(`<c:varyColors val="0"/>`)

	// Series
	for i, series := range opts.Series {
		buf.WriteString(fmt.Sprintf(`<c:ser>
<c:idx val="%d"/>
<c:order val="%d"/>
<c:tx>
  <c:strRef>
    <c:f>Sheet1!$%s$1</c:f>
    <c:strCache>
      <c:ptCount val="1"/>
      <c:pt idx="0"><c:v>%s</c:v></c:pt>
    </c:strCache>
  </c:strRef>
</c:tx>`, i, i, columnLetter(i+1), xmlEscape(series.Name)))

		buf.WriteString(`<c:cat>
  <c:strRef>
    <c:f>Sheet1!$A$2:$A$`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)+1))
		buf.WriteString(`</c:f>
    <c:strCache>
      <c:ptCount val="`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)))
		buf.WriteString(`"/>`)
		for j, cat := range opts.Categories {
			buf.WriteString(fmt.Sprintf(`<c:pt  idx="%d"><c:v>%s</c:v></c:pt>`, j, xmlEscape(cat)))
		}
		buf.WriteString(`</c:strCache>
  </c:strRef>
</c:cat>`)

		buf.WriteString(`<c:val>
  <c:numRef>
    <c:f>Sheet1!$`)
		buf.WriteString(columnLetter(i + 1))
		buf.WriteString(`$2:$`)
		buf.WriteString(columnLetter(i + 1))
		buf.WriteString(`$`)
		buf.WriteString(fmt.Sprintf("%d", len(opts.Categories)+1))
		buf.WriteString(`</c:f>
    <c:numCache>
      <c:formatCode>General</c:formatCode>
      <c:ptCount val="`)
		buf.WriteString(fmt.Sprintf("%d", len(series.Values)))
		buf.WriteString(`"/>`)
		for j, val := range series.Values {
			buf.WriteString(fmt.Sprintf(`<c:pt idx="%d"><c:v>%g</c:v></c:pt>`, j, val))
		}
		buf.WriteString(`</c:numCache>
  </c:numRef>
</c:val>`)

		buf.WriteString(`</c:ser>`)
	}

	buf.WriteString(`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`)
	buf.WriteString(`<c:gapWidth val="150"/>`)
	buf.WriteString(`<c:overlap val="0"/>`)
	buf.WriteString(`<c:axId val="2071991400"/>`)
	buf.WriteString(`<c:axId val="2071991240"/>`)
	buf.WriteString(`</c:barChart>`)

	return buf.String()
}

// columnLetter converts column number to Excel column letter (1=A, 2=B, etc.)
func columnLetter(col int) string {
	result := ""
	for col > 0 {
		col--
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// createEmbeddedWorkbook creates the embedded Excel workbook with chart data
func (u *Updater) createEmbeddedWorkbook(workbookPath string, opts ChartOptions) error {
	if err := os.MkdirAll(filepath.Dir(workbookPath), 0o755); err != nil {
		return fmt.Errorf("create embeddings directory: %w", err)
	}

	// Create a minimal XLSX file with the chart data
	file, err := os.Create(workbookPath)
	if err != nil {
		return fmt.Errorf("create workbook file: %w", err)
	}
	defer file.Close()

	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// Create [Content_Types].xml
	if err := addZipFile(zipWriter, "[Content_Types].xml", generateWorkbookContentTypes()); err != nil {
		return err
	}

	// Create _rels/.rels
	if err := addZipFile(zipWriter, "_rels/.rels", generateWorkbookRels()); err != nil {
		return err
	}

	// Create xl/workbook.xml
	if err := addZipFile(zipWriter, "xl/workbook.xml", generateWorkbookXML()); err != nil {
		return err
	}

	// Create xl/_rels/workbook.xml.rels
	if err := addZipFile(zipWriter, "xl/_rels/workbook.xml.rels", generateWorkbookXMLRels()); err != nil {
		return err
	}

	// Create xl/worksheets/sheet1.xml with data
	if err := addZipFile(zipWriter, "xl/worksheets/sheet1.xml", generateSheetXML(opts)); err != nil {
		return err
	}

	// Create xl/styles.xml
	if err := addZipFile(zipWriter, "xl/styles.xml", generateStylesXML()); err != nil {
		return err
	}

	return nil
}

// Helper function to add file to zip
func addZipFile(zipWriter *zip.Writer, name string, content []byte) error {
	writer, err := zipWriter.Create(name)
	if err != nil {
		return fmt.Errorf("create zip entry %s: %w", name, err)
	}
	if _, err := writer.Write(content); err != nil {
		return fmt.Errorf("write zip entry %s: %w", name, err)
	}
	return nil
}

// generateWorkbookContentTypes creates the [Content_Types].xml for the embedded workbook
func generateWorkbookContentTypes() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`)
}

// generateWorkbookRels creates the _rels/.rels for the embedded workbook
func generateWorkbookRels() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)
}

// generateWorkbookXML creates the xl/workbook.xml
func generateWorkbookXML() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`)
}

// generateWorkbookXMLRels creates the xl/_rels/workbook.xml.rels
func generateWorkbookXMLRels() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)
}

// generateSheetXML creates the xl/worksheets/sheet1.xml with chart data
func generateSheetXML(opts ChartOptions) []byte {
	var buf bytes.Buffer

	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>`)

	// Header row with series names
	buf.WriteString(`<row r="1">`)
	buf.WriteString(`<c r="A1" t="str"><v></v></c>`) // Empty cell at A1
	for i, series := range opts.Series {
		col := columnLetter(i + 2) // B, C, D, etc.
		buf.WriteString(fmt.Sprintf(`<c r="%s1" t="str"><v>%s</v></c>`, col, xmlEscape(series.Name)))
	}
	buf.WriteString(`</row>`)

	// Data rows
	for i, category := range opts.Categories {
		rowNum := i + 2
		buf.WriteString(fmt.Sprintf(`<row r="%d">`, rowNum))

		// Category in column A
		buf.WriteString(fmt.Sprintf(`<c r="A%d" t="str"><v>%s</v></c>`, rowNum, xmlEscape(category)))

		// Values for each series
		for j, series := range opts.Series {
			col := columnLetter(j + 2)
			buf.WriteString(fmt.Sprintf(`<c r="%s%d"><v>%g</v></c>`, col, rowNum, series.Values[i]))
		}

		buf.WriteString(`</row>`)
	}

	buf.WriteString(`</sheetData>
</worksheet>`)

	return buf.Bytes()
}

// generateStylesXML creates a minimal xl/styles.xml
func generateStylesXML() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="0"/>
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`)
}

// createChartRelationships creates the chart relationships file
func (u *Updater) createChartRelationships(relsPath, workbookPath string) error {
	if err := os.MkdirAll(filepath.Dir(relsPath), 0o755); err != nil {
		return fmt.Errorf("create chart _rels directory: %w", err)
	}

	// Get relative path from charts directory to workbook
	chartsDir := filepath.Join(u.tempDir, "word", "charts")
	relPath, err := filepath.Rel(chartsDir, workbookPath)
	if err != nil {
		return fmt.Errorf("calculate relative path: %w", err)
	}

	// Convert to forward slashes for XML
	relPath = filepath.ToSlash(relPath)

	xml := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="%s"/>
</Relationships>`, relPath)

	if err := os.WriteFile(relsPath, []byte(xml), 0o644); err != nil {
		return fmt.Errorf("write relationships file: %w", err)
	}

	return nil
}

// insertChartDrawing inserts the chart drawing into the document
func (u *Updater) insertChartDrawing(chartIndex int, relID string, opts ChartOptions) error {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	// Generate chart drawing XML
	drawingXML, err := u.generateChartDrawingWithSize(chartIndex, relID, opts.Width, opts.Height)
	if err != nil {
		return fmt.Errorf("generate drawing xml: %w", err)
	}

	// Handle caption if specified
	contentToInsert := drawingXML
	if opts.Caption != nil {
		// Validate caption options
		if err := ValidateCaptionOptions(opts.Caption); err != nil {
			return fmt.Errorf("invalid caption options: %w", err)
		}

		// Set caption type to Figure if not already set
		if opts.Caption.Type == "" {
			opts.Caption.Type = CaptionFigure
		}

		// Generate caption XML
		captionXML := generateCaptionXML(*opts.Caption)

		// Combine chart and caption based on position
		contentToInsert = insertCaptionWithElement(raw, captionXML, drawingXML, opts.Caption.Position)
	}

	// Insert based on position
	var updated []byte
	switch opts.Position {
	case PositionBeginning:
		updated, err = insertAtBodyStart(raw, contentToInsert)
	case PositionEnd:
		updated, err = insertAtBodyEnd(raw, contentToInsert)
	case PositionAfterText:
		if opts.Anchor == "" {
			return fmt.Errorf("anchor text required for PositionAfterText")
		}
		updated, err = insertAfterText(raw, contentToInsert, opts.Anchor)
	case PositionBeforeText:
		if opts.Anchor == "" {
			return fmt.Errorf("anchor text required for PositionBeforeText")
		}
		updated, err = insertBeforeText(raw, contentToInsert, opts.Anchor)
	default:
		return fmt.Errorf("invalid insert position")
	}

	if err != nil {
		return fmt.Errorf("insert chart: %w", err)
	}

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// generateChartDrawingWithSize creates the inline drawing XML for a chart with custom dimensions
func (u *Updater) generateChartDrawingWithSize(chartIndex int, relId string, width, height int) ([]byte, error) {
	// Get a unique docPr ID (document-wide drawing object ID)
	docPrId, err := u.getNextDocPrId()
	if err != nil {
		return nil, fmt.Errorf("get next docPr id: %w", err)
	}

	// Generate unique IDs
	anchorId := ChartAnchorIDBase + uint32(chartIndex)*ChartIDIncrement
	editId := ChartEditIDBase + uint32(chartIndex)*ChartIDIncrement

	template := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="%08X" wp14:editId="%08X"><wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="15875" b="12700"/><wp:docPr id="%d" name="Chart %d"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

	return []byte(fmt.Sprintf(template, anchorId, editId, width, height, docPrId, chartIndex, relId)), nil
}
