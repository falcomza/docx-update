package docxchartupdater

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
)

// seriesElement represents a chart series
type seriesElement struct {
	XMLName xml.Name    `xml:"ser"`
	Idx     *valAttr `xml:"idx"`
	Order   *valAttr `xml:"order"`
	Tx      *txElement  `xml:"tx"`
	SpPr    *spPrElement `xml:"spPr"`
	InvertIfNegative *valAttr `xml:"invertIfNegative"`
	Cat     *catElement `xml:"cat"`
	Val     *valElement `xml:"val"`
	XVal    *catElement `xml:"xVal"` // for scatter charts
	YVal    *valElement `xml:"yVal"` // for scatter charts
	ExtLst  *extLstElement `xml:"extLst"`
}

type txElement struct {
	XMLName xml.Name     `xml:"tx"`
	V       *stringValue `xml:"v"`
	StrRef  *strRef      `xml:"strRef"`
}

type catElement struct {
	XMLName  xml.Name
	StrRef   *strRef   `xml:"strRef"`
	StrCache *strCache `xml:"strCache"`
}

type valElement struct {
	XMLName  xml.Name
	NumRef   *numRef   `xml:"numRef"`
	NumCache *numCache `xml:"numCache"`
}

type strRef struct {
	XMLName xml.Name  `xml:"strRef"`
	F       *formula  `xml:"f"`
	Cache   *strCache `xml:"strCache"`
}

type numRef struct {
	XMLName xml.Name  `xml:"numRef"`
	F       *formula  `xml:"f"`
	Cache   *numCache `xml:"numCache"`
}

type formula struct {
	XMLName xml.Name `xml:"f"`
	Value   string   `xml:",chardata"`
}

type strCache struct {
	XMLName xml.Name     `xml:"strCache"`
	PtCount *ptCount     `xml:"ptCount"`
	Pts     []stringPt   `xml:"pt"`
}

type numCache struct {
	XMLName xml.Name    `xml:"numCache"`
	PtCount *ptCount    `xml:"ptCount"`
	Pts     []numericPt `xml:"pt"`
}

type ptCount struct {
	XMLName xml.Name `xml:"ptCount"`
	Val     int      `xml:"val,attr"`
}

type stringPt struct {
	XMLName xml.Name     `xml:"pt"`
	Idx     int          `xml:"idx,attr"`
	V       *stringValue `xml:"v"`
}

type numericPt struct {
	XMLName xml.Name     `xml:"pt"`
	Idx     int          `xml:"idx,attr"`
	V       *stringValue `xml:"v"`
}

type stringValue struct {
	XMLName xml.Name `xml:"v"`
	Value   string   `xml:",chardata"`
}

// valAttr represents elements with a val attribute
type valAttr struct {
	Val int `xml:"val,attr"`
}

// spPrElement represents shape properties (including colors)
type spPrElement struct {
	XMLName xml.Name `xml:"spPr"`
	InnerXML []byte  `xml:",innerxml"`
}

type extLstElement struct {
	XMLName xml.Name `xml:"extLst"`
	InnerXML []byte  `xml:",innerxml"`
}

func updateChartXML(chartPath string, data ChartData) error {
	raw, err := os.ReadFile(chartPath)
	if err != nil {
		return fmt.Errorf("read chart xml: %w", err)
	}

	// Update titles FIRST on the original raw XML (before series rebuild)
	if data.ChartTitle != "" || data.CategoryAxisTitle != "" || data.ValueAxisTitle != "" {
		raw = updateTitles(raw, data)
	}

	// Use a streaming approach to find and update <c:ser> elements
	seriesElements, err := extractSeriesElements(raw)
	if err != nil {
		return err
	}

	if len(seriesElements) == 0 {
		return fmt.Errorf("no <ser> elements found in chart xml")
	}

	// Update each series element
	for i := range seriesElements {
		if i < len(data.Series) {
			updateSeries(&seriesElements[i], data.Categories, data.Series[i])
		}
	}

	// Rebuild the chart XML with updated series; drop any extra <ser> beyond provided data
	updated, err := rebuildChartXML(raw, seriesElements, len(data.Series))
	if err != nil {
		return err
	}

	if err := os.WriteFile(chartPath, updated, 0o644); err != nil {
		return fmt.Errorf("write chart xml: %w", err)
	}

	return nil
}

// extractSeriesElements finds all <c:ser> or similar namespace ser elements
func extractSeriesElements(raw []byte) ([]seriesElement, error) {
	var series []seriesElement
	decoder := xml.NewDecoder(bytes.NewReader(raw))

	for {
		token, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return nil, fmt.Errorf("xml token: %w", err)
		}

		if start, ok := token.(xml.StartElement); ok {
			if start.Name.Local == "ser" {
				var ser seriesElement
				if err := decoder.DecodeElement(&ser, &start); err != nil {
					return nil, fmt.Errorf("parse series: %w", err)
				}
				series = append(series, ser)
			}
		}
	}

	return series, nil
}

// rebuildChartXML rebuilds chart XML with updated series elements
// keepCount controls how many <ser> to keep; any further original series are removed
func rebuildChartXML(original []byte, updatedSeries []seriesElement, keepCount int) ([]byte, error) {
	var result bytes.Buffer
	decoder := xml.NewDecoder(bytes.NewReader(original))
	encoder := xml.NewEncoder(&result)

	seriesIdx := 0

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		if start, ok := token.(xml.StartElement); ok {
			if start.Name.Local == "ser" {
				// Always skip the original series element from the decoder
				if err := skipElement(decoder); err != nil {
					return nil, fmt.Errorf("skip original series: %w", err)
				}
				// If we still need to keep a series, encode the updated one; else drop it
				if seriesIdx < keepCount && seriesIdx < len(updatedSeries) {
					if err := encoder.EncodeElement(updatedSeries[seriesIdx], start); err != nil {
						return nil, fmt.Errorf("encode series %d: %w", seriesIdx, err)
					}
				}
				seriesIdx++
				continue
		}
		}

		if err := encoder.EncodeToken(xml.CopyToken(token)); err != nil {
			return nil, fmt.Errorf("encode token: %w", err)
		}
	}

	encoder.Flush()
	
	// TODO: Fix removeDuplicateXmlns - currently corrupts XML
	// For now, return as-is (will have duplicate xmlns but that's less harmful)
	return result.Bytes(), nil
}

// removeDuplicateXmlns removes duplicate xmlns attributes from XML
func removeDuplicateXmlns(xmlData []byte) []byte {
	// Find all xmlns declarations and remove consecutive duplicates
	xmlnsRe := regexp.MustCompile(`xmlns(:[a-zA-Z0-9]+)?="[^"]*"`)
	
	// Process each opening tag
	tagRe := regexp.MustCompile(`<[a-zA-Z0-9:]+[^>]*>`)
	
	result := tagRe.ReplaceAllFunc(xmlData, func(tag []byte) []byte {
		// Find all xmlns in this tag
		xmlnsMatches := xmlnsRe.FindAll(tag, -1)
		
		if len(xmlnsMatches) <= 1 {
			return tag // No duplicates possible
		}
		
		// Track unique xmlns
		seen := make(map[string]bool)
		
		// Remove duplicates by replacing tag content
		for _, xmlns := range xmlnsMatches {
			xmlnsStr := string(xmlns)
			if seen[xmlnsStr] {
				// This is a duplicate, remove it
				tag = bytes.Replace(tag, xmlns, []byte(""), 1)
			} else {
				seen[xmlnsStr] = true
			}
		}
		
		// Clean up extra whitespace
		tag = regexp.MustCompile(`\s+`).ReplaceAll(tag, []byte(" "))
		tag = bytes.Replace(tag, []byte(" >"), []byte(">"), -1)
		
		return tag
	})
	
	return result
}

// skipElement skips the current element in the decoder
func skipElement(decoder *xml.Decoder) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

// xmlEscape escapes special XML characters
func xmlEscape(s string) string {
	var buf bytes.Buffer
	xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

// updateFormulaRange updates an Excel formula range to match newCount
// e.g., "Sheet1!$A$2:$A$4" with newCount=5 becomes "Sheet1!$A$2:$A$6"
func updateFormulaRange(formula string, newCount int) string {
	if formula == "" || newCount <= 0 {
		return formula
	}

	// Find the range pattern like $A$2:$A$4 or A2:A4
	idx := bytes.IndexByte([]byte(formula), ':')
	if idx == -1 {
		return formula // No range, return as-is
	}

	// Parse the formula to extract sheet, column, and start row
	// Common pattern: Sheet1!$A$2:$A$4
	parts := bytes.Split([]byte(formula), []byte(":"))
	if len(parts) != 2 {
		return formula
	}

	// Extract column and starting row from first part (e.g., "Sheet1!$A$2")
	firstPart := string(parts[0])
	
	// Find the last occurrence of a letter (column) before digits (row)
	var colEnd, rowStart int
	for i := len(firstPart) - 1; i >= 0; i-- {
		if firstPart[i] >= '0' && firstPart[i] <= '9' {
			continue
		}
		if (firstPart[i] >= 'A' && firstPart[i] <= 'Z') || (firstPart[i] >= 'a' && firstPart[i] <= 'z') || firstPart[i] == '$' {
			colEnd = i + 1
			rowStart = i + 1
			// Skip $ before column letter
			for rowStart < len(firstPart) && firstPart[rowStart] == '$' {
				rowStart++
			}
			break
		}
	}

	if rowStart == 0 || rowStart >= len(firstPart) {
		return formula // Couldn't parse
	}

	// Extract the starting row number
	var startRow int
	for i := rowStart; i < len(firstPart); i++ {
		if firstPart[i] >= '0' && firstPart[i] <= '9' {
			startRow = startRow*10 + int(firstPart[i]-'0')
		} else {
			break
		}
	}

	if startRow == 0 {
		return formula // Invalid row
	}

	// Calculate end row (startRow + newCount - 1)
	endRow := startRow + newCount - 1

	// Reconstruct formula: prefix + ":" + column + endRow
	prefix := firstPart[:colEnd]
	return fmt.Sprintf("%s:%s%d", firstPart, prefix[bytes.LastIndexByte([]byte(prefix), '!')+1:], endRow)
}

func updateSeries(ser *seriesElement, categories []string, series SeriesData) {
	// Update series name
	if ser.Tx != nil {
		if ser.Tx.V != nil {
			ser.Tx.V.Value = series.Name
		} else if ser.Tx.StrRef != nil && ser.Tx.StrRef.Cache != nil {
			// If using strRef, update cache
			if len(ser.Tx.StrRef.Cache.Pts) > 0 && ser.Tx.StrRef.Cache.Pts[0].V != nil {
				ser.Tx.StrRef.Cache.Pts[0].V.Value = series.Name
			}
		}
	}

	// Update categories (or xVal for scatter)
	catElem := ser.Cat
	if catElem == nil {
		catElem = ser.XVal
	}
	if catElem != nil {
		if catElem.StrCache != nil {
			updateStrCache(catElem.StrCache, categories)
		}
		if catElem.StrRef != nil {
			// Update formula to match new range
			if catElem.StrRef.F != nil {
				catElem.StrRef.F.Value = updateFormulaRange(catElem.StrRef.F.Value, len(categories))
			}
			if catElem.StrRef.Cache != nil {
				updateStrCache(catElem.StrRef.Cache, categories)
			}
		}
	}

	// Update values (or yVal for scatter)
	valElem := ser.Val
	if valElem == nil {
		valElem = ser.YVal
	}
	if valElem != nil {
		if valElem.NumCache != nil {
			updateNumCache(valElem.NumCache, series.Values)
		}
		if valElem.NumRef != nil {
			// Update formula to match new range
			if valElem.NumRef.F != nil {
				valElem.NumRef.F.Value = updateFormulaRange(valElem.NumRef.F.Value, len(series.Values))
			}
			if valElem.NumRef.Cache != nil {
				updateNumCache(valElem.NumRef.Cache, series.Values)
			}
		}
	}
}

func updateStrCache(cache *strCache, values []string) {
	if cache.PtCount == nil {
		cache.PtCount = &ptCount{}
	}
	cache.PtCount.Val = len(values)
	cache.Pts = make([]stringPt, len(values))
	for i, v := range values {
		cache.Pts[i] = stringPt{
			Idx: i,
			V:   &stringValue{Value: v},
		}
	}
}

func updateNumCache(cache *numCache, values []float64) {
	if cache.PtCount == nil {
		cache.PtCount = &ptCount{}
	}
	cache.PtCount.Val = len(values)
	cache.Pts = make([]numericPt, len(values))
	for i, v := range values {
		cache.Pts[i] = numericPt{
			Idx: i,
			V:   &stringValue{Value: strconv.FormatFloat(v, 'f', -1, 64)},
		}
	}
}

// updateTitles updates chart and axis titles
func updateTitles(xmlData []byte, data ChartData) []byte {
	result := xmlData

	// Update main chart title (first <c:title> after <c:chart>)
	if data.ChartTitle != "" {
		result = updateTitleText(result, "<c:chart>", data.ChartTitle, 1)
	}

	// Update category axis title (in <c:catAx>)
	if data.CategoryAxisTitle != "" {
		result = updateTitleText(result, "<c:catAx>", data.CategoryAxisTitle, 1)
	}

	// Update value axis title (in <c:valAx>)
	if data.ValueAxisTitle != "" {
		result = updateTitleText(result, "<c:valAx>", data.ValueAxisTitle, 1)
	}

	return result
}

// updateTitleText updates the text content within a title element
// It finds the parent element, then locates the title's <a:t> or <t> tag and replaces its content
func updateTitleText(xmlData []byte, parentMarker string, newText string, occurrence int) []byte {
	// Find the parent element - try both with and without namespace
	parentIdx := findNthOccurrence(xmlData, []byte(parentMarker), occurrence)
	if parentIdx == -1 {
		// Try without c: prefix
		altMarker := bytes.Replace([]byte(parentMarker), []byte("c:"), []byte{}, 1)
		parentIdx = findNthOccurrence(xmlData, altMarker, occurrence)
		if parentIdx == -1 {
			return xmlData
		}
	}

	// Look for title after the parent marker (with or without namespace)
	var titleStart, titleEndLen int
	titleStart = bytes.Index(xmlData[parentIdx:], []byte("<c:title>"))
	if titleStart == -1 {
		titleStart = bytes.Index(xmlData[parentIdx:], []byte("<title "))
		if titleStart == -1 {
			titleStart = bytes.Index(xmlData[parentIdx:], []byte("<title>"))
			if titleStart == -1 {
				return xmlData
			}
		}
	}
	titleStart += parentIdx

	// Find the closing tag
	if bytes.Contains(xmlData[titleStart:titleStart+20], []byte("c:title")) {
		titleEndLen = len("</c:title>")
	} else {
		titleEndLen = len("</title>")
	}
	titleEnd := bytes.Index(xmlData[titleStart:], xmlData[titleStart:titleStart+titleEndLen])
	if titleEnd == -1 {
		return xmlData
	}
	// Find proper closing tag
	var closeTag []byte
	if bytes.Contains(xmlData[titleStart:titleStart+15], []byte("c:title")) {
		closeTag = []byte("</c:title>")
	} else {
		closeTag = []byte("</title>")
	}
	titleEnd = bytes.Index(xmlData[titleStart:], closeTag)
	if titleEnd == -1 {
		return xmlData
	}
	titleEnd += titleStart + len(closeTag)

	// Extract the title section
	titleSection := xmlData[titleStart:titleEnd]

	// Find <a:t>...</a:t> or <t>...</t> within the title section
	tStart := bytes.Index(titleSection, []byte("<a:t>"))
	var tTag string
	if tStart == -1 {
		tStart = bytes.Index(titleSection, []byte("<t>"))
		if tStart == -1 {
			// No text element found - need to insert one
			return insertTitleText(xmlData, titleStart, titleEnd, newText)
		}
		tTag = "<t>"
	} else {
		tTag = "<a:t>"
	}
	closeT := bytes.Replace([]byte(tTag), []byte("<"), []byte("</"), 1)
	tEnd := bytes.Index(titleSection[tStart:], closeT)
	if tEnd == -1 {
		return xmlData
	}
	tEnd += tStart

	// Replace the text content
	newTitleSection := append([]byte{}, titleSection[:tStart+len(tTag)]...)
	newTitleSection = append(newTitleSection, []byte(xmlEscape(newText))...)
	newTitleSection = append(newTitleSection, titleSection[tEnd:]...)

	// Reconstruct the full XML
	result := append([]byte{}, xmlData[:titleStart]...)
	result = append(result, newTitleSection...)
	result = append(result, xmlData[titleEnd:]...)

	return result
}

// findNthOccurrence finds the nth occurrence of a pattern in data
func findNthOccurrence(data []byte, pattern []byte, n int) int {
	count := 0
	offset := 0
	for {
		idx := bytes.Index(data[offset:], pattern)
		if idx == -1 {
			return -1
		}
		count++
		if count == n {
			return offset + idx
		}
		offset += idx + len(pattern)
	}
}

// insertTitleText creates a title text element when one doesn't exist
func insertTitleText(xmlData []byte, titleStart, titleEnd int, newText string) []byte {
	titleSection := xmlData[titleStart:titleEnd]

	// Look for <c:txPr> or <txPr> which contains formatting
	txPrIdx := bytes.Index(titleSection, []byte("<c:txPr>"))
	if txPrIdx == -1 {
		txPrIdx = bytes.Index(titleSection, []byte("<txPr>"))
	}

	if txPrIdx == -1 {
		// No txPr found, can't safely insert
		return xmlData
	}

	// Insert a <c:tx> element with <c:rich> containing the text before <c:txPr>
	// This follows the pattern: <c:title><c:tx><c:rich>...<a:p><a:r><a:t>TEXT</a:t></a:r></a:p>...</c:rich></c:tx><c:txPr>...
	textElement := fmt.Sprintf(
		`<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>`,
		xmlEscape(newText),
	)

	// Insert the text element before <c:txPr>
	newTitleSection := append([]byte{}, titleSection[:txPrIdx]...)
	newTitleSection = append(newTitleSection, []byte(textElement)...)
	newTitleSection = append(newTitleSection, titleSection[txPrIdx:]...)

	// Reconstruct the full XML
	result := append([]byte{}, xmlData[:titleStart]...)
	result = append(result, newTitleSection...)
	result = append(result, xmlData[titleEnd:]...)

	return result
}

