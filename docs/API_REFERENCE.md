# DOCX Update Library - API Reference

**Module**: `github.com/falcomza/go-docx`
**Go Version**: 1.25.7+
**Fiber Version**: v3.0.0+

---

## Table of Contents

1. [Package Overview](#package-overview)
2. [Installation](#installation)
3. [Exported Types](#exported-types)
4. [Exported Functions](#exported-functions)
5. [Exported Methods](#exported-methods)
6. [Fiber v3 Integration](#fiber-v3-integration)
7. [Complete Handler Examples](#complete-handler-examples)

---

## Package Overview

```go
import godocx "github.com/falcomza/go-docx"
```

The `godocx` package provides functionality for programmatically manipulating Microsoft Word (DOCX) documents. It operates directly on the OpenXML format, enabling chart updates, table insertion, image embedding, and text operations.

**Key characteristics:**
- Safe for concurrent use (each `Updater` instance uses isolated temp directories)
- 1-based indexing for chart operations
- Strict validation with descriptive error types
- Automatic cleanup via `defer updater.Cleanup()`

---

## Installation

```bash
go get github.com/falcomza/go-docx@latest
```

### Import in Fiber Application

```go
import (
    godocx "github.com/falcomza/go-docx"
    "github.com/gofiber/fiber/v3"
)
```

---

## Exported Types

### Updater

Main struct for document manipulation. Create using `New()` function.

```go
type Updater struct {
    // Contains unexported fields
}
```

**Lifecycle:**
```go
updater, err := godocx.New("input.docx")
if err != nil {
    return err
}
defer updater.Cleanup()

// Perform operations...

err = updater.Save("output.docx")
```

### ChartData

Container for chart update data.

```go
type ChartData struct {
    Categories       []string     // X-axis labels (required, non-empty)
    Series          []SeriesData  // Data series (required, at least one)
    ChartTitle        string       // Optional: main chart title
    CategoryAxisTitle string       // Optional: x-axis title
    ValueAxisTitle    string       // Optional: y-axis title
}
```

**Validation:**
- `len(Categories) > 0`
- `len(Series) > 0`
- All series names must be non-empty after trimming whitespace
- All `len(Series[i].Values) == len(Categories)`

**Example:**
```go
data := godocx.ChartData{
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []godocx.SeriesData{
        {
            Name:   "Revenue",
            Values: []float64{1000, 1500, 1200, 1800},
        },
        {
            Name:   "Expenses",
            Values: []float64{800, 900, 850, 1000},
        },
    },
    ChartTitle:        "Quarterly Results",
    CategoryAxisTitle: "Fiscal Quarter",
    ValueAxisTitle:    "USD",
}
```

### SeriesData

Single series definition.

```go
type SeriesData struct {
    Name   string    // Series label (required)
    Values []float64 // Data points (must match Categories length)
}
```

### ParagraphOptions

Options for paragraph insertion.

```go
type ParagraphOptions struct {
    Text     string         // Required: paragraph content
    Style    ParagraphStyle // Default: StyleNormal
    Position InsertPosition // Default: PositionEnd
    Anchor   string         // Required for PositionAfterText/BeforeText
    Bold     bool
    Italic   bool
    Underline bool
}
```

### ParagraphStyle

Predefined Word paragraph styles.

```go
type ParagraphStyle string

const (
    StyleNormal     ParagraphStyle = "Normal"
    StyleHeading1   ParagraphStyle = "Heading1"
    StyleHeading2   ParagraphStyle = "Heading2"
    StyleHeading3   ParagraphStyle = "Heading3"
    StyleTitle      ParagraphStyle = "Title"
    StyleSubtitle   ParagraphStyle = "Subtitle"
    StyleQuote      ParagraphStyle = "Quote"
    StyleIntense    ParagraphStyle = "IntenseQuote"
    StyleListNumber ParagraphStyle = "ListNumber"
    StyleListBullet ParagraphStyle = "ListBullet"
)
```

### InsertPosition

Location for content insertion.

```go
type InsertPosition int

const (
    PositionBeginning InsertPosition = iota // Start of document
    PositionEnd                              // End of document
    PositionAfterText                        // After anchor text
    PositionBeforeText                       // Before anchor text
)
```

### TableOptions

Comprehensive table creation options.

```go
type TableOptions struct {
    // Positioning
    Position  InsertPosition
    Anchor    string

    // Structure
    Columns      []ColumnDefinition
    ColumnWidths []int

    // Data
    Rows [][]string

    // Header styling
    HeaderStyle       CellStyle
    HeaderStyleName   string
    RepeatHeader      bool
    HeaderBackground  string // hex color
    HeaderBold        bool
    HeaderAlignment   CellAlignment

    // Row styling
    RowStyle          CellStyle
    RowStyleName      string
    AlternateRowColor string // hex color
    RowAlignment      CellAlignment
    VerticalAlign     VerticalAlignment

    // Dimensions
    HeaderRowHeight int
    HeaderHeightRule RowHeightRule
    RowHeight       int
    RowHeightRule   RowHeightRule

    // Table properties
    TableAlignment TableAlignment
    TableWidthType TableWidthType
    TableWidth     int
    TableStyle     TableStyle
    BorderStyle    BorderStyle
    BorderSize     int
    BorderColor    string

    // Spacing
    CellPadding int
    AutoFit     bool

    // Caption
    Caption *CaptionOptions
}
```

### ImageOptions

Image insertion with proportional scaling.

```go
type ImageOptions struct {
    Path     string          // Required: image file path
    Width    int             // Optional: pixels
    Height   int             // Optional: pixels
    AltText  string
    Position InsertPosition
    Anchor   string
    Caption  *CaptionOptions
}
```

**Supported formats:** PNG, JPEG, GIF, BMP, TIFF

**Scaling behavior:**
- Both `Width` and `Height` set → Use exact dimensions
- Only `Width` set → Calculate height proportionally
- Only `Height` set → Calculate width proportionally
- Neither set → Use original image dimensions

### FindOptions

Text search configuration.

```go
type FindOptions struct {
    MatchCase    bool
    WholeWord    bool
    UseRegex     bool
    MaxResults   int // 0 = unlimited
    InParagraphs bool
    InTables     bool
    InHeaders    bool
    InFooters    bool
}
```

### ReplaceOptions

Text replacement configuration.

```go
type ReplaceOptions struct {
    MatchCase       bool
    WholeWord       bool
    InParagraphs    bool
    InTables        bool
    InHeaders       bool
    InFooters       bool
    MaxReplacements int // 0 = unlimited
}
```

### TextMatch

Result from text search.

```go
type TextMatch struct {
    Text      string // Matched content
    Paragraph int    // 0-based paragraph index
    Position  int    // Character position
    Before    string // Context before (up to 50 chars)
    After     string // Context after (up to 50 chars)
}
```

### HyperlinkOptions

Hyperlink insertion options.

```go
type HyperlinkOptions struct {
    Position  InsertPosition
    Anchor    string
    Tooltip   string
    Style     ParagraphStyle
    Color     string // hex, default: "0563C1"
    Underline bool
    ScreenTip string
}
```

### HeaderFooterContent

Header/footer content structure.

```go
type HeaderFooterContent struct {
    LeftText          string
    CenterText        string
    RightText         string
    PageNumber        bool
    PageNumberFormat  string // e.g., "Page X of Y"
    Date              bool
    DateFormat        string // e.g., "MMMM d, yyyy"
}
```

### HeaderOptions / FooterOptions

Header/footer configuration.

```go
type HeaderOptions struct {
    Type             HeaderType
    DifferentFirst   bool
    DifferentOddEven bool
}

type FooterOptions struct {
    Type             FooterType
    DifferentFirst   bool
    DifferentOddEven bool
}

const (
    HeaderFirst    HeaderType = "first"
    HeaderEven     HeaderType = "even"
    HeaderDefault  HeaderType = "default"

    FooterFirst    FooterType = "first"
    FooterEven     FooterType = "even"
    FooterDefault  FooterType = "default"
)
```

### BreakOptions

Page/section break configuration.

```go
type BreakOptions struct {
    Position     InsertPosition
    Anchor       string
    SectionType  SectionBreakType
}

const (
    SectionBreakNextPage     SectionBreakType = "nextPage"
    SectionBreakContinuous   SectionBreakType = "continuous"
    SectionBreakEvenPage     SectionBreakType = "evenPage"
    SectionBreakOddPage      SectionBreakType = "oddPage"
)
```

### CoreProperties

Core document metadata properties.

```go
type CoreProperties struct {
    Title          string    // Document title
    Subject        string    // Document subject
    Creator        string    // Author/Creator name
    Keywords       string    // Keywords (comma-separated)
    Description    string    // Description/Comments
    Category       string    // Document category
    Created        time.Time // Creation date
    Modified       time.Time // Modification date
    LastModifiedBy string    // Last modifier name
    Revision       string    // Revision number/version
}
```

### AppProperties

Application-specific document properties.

```go
type AppProperties struct {
    Company     string // Company name
    Manager     string // Manager name
    Application string // Application name (typically Microsoft Word)
    AppVersion  string // Application version (e.g., "16.0000")
}
```

### CustomProperty

Custom document property with typed value.

```go
type CustomProperty struct {
    Name  string      // Property name
    Value interface{} // Property value (string, int, float64, bool, or time.Time)
    Type  string      // Type identifier (optional, auto-inferred)
}
```

**Supported types:**
- `"lpwstr"` - String value
- `"i4"` - Integer value
- `"r8"` - Float64 value
- `"bool"` - Boolean value
- `"filetime"` - Time value

### BookmarkOptions

Options for bookmark creation.

```go
type BookmarkOptions struct {
    Position InsertPosition // Where to create bookmark
    Anchor   string         // Anchor text for relative positioning
    Style    ParagraphStyle // Style for bookmarked text
    Hidden   bool           // Invisible marker (default: true)
}
```

### ChartOptions

Comprehensive options for chart creation.

```go
type ChartOptions struct {
    // Positioning
    Position InsertPosition
    Anchor   string // Text anchor for relative positioning

    // Chart type
    ChartKind ChartKind // Column, Bar, Line, Pie, Area

    // Titles
    Title             string // Main chart title
    CategoryAxisTitle string // X-axis title (horizontal axis)
    ValueAxisTitle    string // Y-axis title (vertical axis)

    // Data
    Categories []string     // Category labels (X-axis)
    Series     []SeriesData // Data series with names and values

    // Legend
    ShowLegend     bool   // Show legend (default: true)
    LegendPosition string // Legend position: "r" (right), "l" (left), "t" (top), "b" (bottom)

    // Dimensions (default: spans between margins)
    Width  int // Width in EMUs (English Metric Units), 0 for default (6099523 = ~6.5")
    Height int // Height in EMUs, 0 for default (3340467 = ~3.5")

    // Caption
    Caption *CaptionOptions
}
```

### ChartKind

Chart type enumeration.

```go
type ChartKind string

const (
    ChartKindColumn ChartKind = "barChart"  // Column chart (vertical bars)
    ChartKindBar    ChartKind = "barChart"  // Bar chart (horizontal bars)
    ChartKindLine   ChartKind = "lineChart" // Line chart
    ChartKindPie    ChartKind = "pieChart"  // Pie chart
    ChartKindArea   ChartKind = "areaChart" // Area chart
)
```

### CaptionOptions

Caption for tables/figures.

```go
type CaptionOptions struct {
    Type        CaptionType // Figure or Table
    Position    CaptionPosition // before or after
    Description string
    Style       string
    AutoNumber  bool
    Alignment   CellAlignment
    ManualNumber int
}

const (
    CaptionFigure CaptionType = "Figure"
    CaptionTable  CaptionType = "Table"
)
```

### CellStyle

Table cell styling.

```go
type CellStyle struct {
    Bold       bool
    Italic     bool
    FontSize   int    // half-points
    FontColor  string // hex
    Background string // hex
}
```

### ColumnDefinition

Table column definition.

```go
type ColumnDefinition struct {
    Title     string
    Width     int
    Alignment CellAlignment
    Bold      bool
}
```

### Table Style Constants

```go
const (
    TableStyleGrid         TableStyle = "TableGrid"
    TableStyleGridAccent1  TableStyle = "LightShading-Accent1"
    TableStylePlain        TableStyle = "TableNormal"
    TableStyleColorful     TableStyle = "ColorfulGrid-Accent1"
    TableStyleProfessional TableStyle = "LightGrid-Accent1"
)
```

### Alignment Constants

```go
const (
    CellAlignLeft   CellAlignment = "start"
    CellAlignCenter CellAlignment = "center"
    CellAlignRight  CellAlignment = "end"
)

const (
    VerticalAlignTop    VerticalAlignment = "top"
    VerticalAlignCenter VerticalAlignment = "center"
    VerticalAlignBottom VerticalAlignment = "bottom"
)

const (
    AlignLeft   TableAlignment = "left"
    AlignCenter TableAlignment = "center"
    AlignRight  TableAlignment = "right"
)
```

### DocxError

Structured error type.

```go
type DocxError struct {
    Code    ErrorCode
    Message string
    Err     error
    Context map[string]interface{}
}

func (e *DocxError) Error() string
func (e *DocxError) Unwrap() error
func (e *DocxError) WithContext(key string, value interface{}) *DocxError
```

### ErrorCode Constants

```go
const (
    ErrCodeInvalidFile      ErrorCode = "INVALID_FILE"
    ErrCodeFileNotFound     ErrorCode = "FILE_NOT_FOUND"
    ErrCodeChartNotFound    ErrorCode = "CHART_NOT_FOUND"
    ErrCodeInvalidChartData ErrorCode = "INVALID_CHART_DATA"
    ErrCodeImageNotFound    ErrorCode = "IMAGE_NOT_FOUND"
    ErrCodeTextNotFound     ErrorCode = "TEXT_NOT_FOUND"
    ErrCodeInvalidRegex     ErrorCode = "INVALID_REGEX"
    ErrCodeXMLParse         ErrorCode = "XML_PARSE"
    ErrCodeRelationship     ErrorCode = "RELATIONSHIP"
    ErrCodeValidation       ErrorCode = "VALIDATION"
    ErrCodeInvalidURL       ErrorCode = "INVALID_URL"
    ErrCodeHeaderFooter     ErrorCode = "HEADER_FOOTER"
)
```

---

## Exported Functions

### New

```go
func New(docxPath string) (*Updater, error)
```

Creates a new Updater by extracting the DOCX file to a temporary directory.

**Parameters:**
- `docxPath`: Path to the source DOCX file (must exist and be valid)

**Returns:**
- `*Updater`: Updater instance for document manipulation
- `error`: `os.ErrNotExist` if file not found, error if extraction fails or DOCX structure invalid

**Example:**
```go
updater, err := godocx.New("template.docx")
if err != nil {
    return fmt.Errorf("failed to load template: %w", err)
}
defer updater.Cleanup()
```

### DefaultFindOptions

```go
func DefaultFindOptions() FindOptions
```

Returns find options with sensible defaults.

### DefaultReplaceOptions

```go
func DefaultReplaceOptions() ReplaceOptions
```

Returns replace options with sensible defaults.

### DefaultHyperlinkOptions

```go
func DefaultHyperlinkOptions() HyperlinkOptions
```

Returns hyperlink options with sensible defaults.

### DefaultCaptionOptions

```go
func DefaultCaptionOptions(captionType CaptionType) CaptionOptions
```

Returns caption options with sensible defaults for the specified type.

### Error Constructors

```go
func NewChartNotFoundError(index int) error
func NewInvalidChartDataError(reason string) error
func NewImageNotFoundError(path string) error
func NewTextNotFoundError(text string) error
func NewInvalidRegexError(pattern string, err error) error
func NewXMLParseError(file string, err error) error
func NewValidationError(field, reason string) error
func NewInvalidURLError(url string) error
func NewHyperlinkError(reason string, err error) error
func NewHeaderFooterError(reason string, err error) error
```

---

## Exported Methods

### Chart Operations

#### UpdateChart

```go
func (u *Updater) UpdateChart(chartIndex int, data ChartData) error
```

Updates a specific chart's data and embedded workbook.

**Parameters:**
- `chartIndex`: 1-based chart index (must be ≥ 1)
- `data`: Chart data with categories, series, and optional titles

**Returns:**
- `error`: Chart not found, invalid data, workbook resolution failed

**Example:**
```go
err := updater.UpdateChart(1, godocx.ChartData{
    Categories: []string{"A", "B", "C"},
    Series: []godocx.SeriesData{
        {Name: "Series 1", Values: []float64{1, 2, 3}},
    },
})
```

#### InsertChart

```go
func (u *Updater) InsertChart(opts ChartOptions) error
```

Creates a new chart with embedded Excel workbook and inserts it into the document.

**Parameters:**
- `opts`: Chart options including type, data, positioning, and styling

**Returns:**
- `error`: Invalid options, creation failed, or insertion failed

**Validation:**
- Categories must not be empty
- At least one series required
- All series values must match categories length
- Chart type must be valid

**Example:**
```go
err := updater.InsertChart(godocx.ChartOptions{
    Position:  godocx.PositionEnd,
    ChartKind: godocx.ChartKindColumn,
    Title:     "Sales by Quarter",
    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
    Series: []godocx.SeriesData{
        {Name: "2023", Values: []float64{100, 150, 120, 180}},
        {Name: "2024", Values: []float64{120, 170, 140, 200}},
    },
    ShowLegend:     true,
    LegendPosition: "r",
})
```

### Document Operations

#### Save

```go
func (u *Updater) Save(outputPath string) error
```

Writes the modified document to the specified path.

**Parameters:**
- `outputPath`: Destination file path (creates parent directories if needed)

**Returns:**
- `error`: Write failed, directory creation failed

**Example:**
```go
err := updater.Save("./output/report.docx")
```

#### Cleanup

```go
func (u *Updater) Cleanup() error
```

Removes the temporary workspace.

**Example:**
```go
updater, err := godocx.New("input.docx")
if err != nil {
    return err
}
defer updater.Cleanup() // Always defer
```

#### TempDir

```go
func (u *Updater) TempDir() string
```

Returns the temporary directory path for debugging.

### Paragraph Operations

#### InsertParagraph

```go
func (u *Updater) InsertParagraph(opts ParagraphOptions) error
```

Inserts a single paragraph.

**Example:**
```go
updater.InsertParagraph(godocx.ParagraphOptions{
    Text:     "Section Title",
    Style:    godocx.StyleHeading2,
    Position: godocx.PositionBeginning,
})
```

#### InsertParagraphs

```go
func (u *Updater) InsertParagraphs(paragraphs []ParagraphOptions) error
```

Batch inserts multiple paragraphs.

#### AddHeading

```go
func (u *Updater) AddHeading(level int, text string, position InsertPosition) error
```

Convenience method for headings (level 1-3).

#### AddText

```go
func (u *Updater) AddText(text string, position InsertPosition) error
```

Convenience method for normal text.

### Table Operations

#### InsertTable

```go
func (u *Updater) InsertTable(opts TableOptions) error
```

Inserts a styled table with optional caption.

**Example:**
```go
updater.InsertTable(godocx.TableOptions{
    Columns: []godocx.ColumnDefinition{
        {Title: "Name", Width: 2000},
        {Title: "Value", Width: 1000},
    },
    Rows: [][]string{
        {"Item 1", "100"},
        {"Item 2", "200"},
    },
    HeaderStyle: godocx.CellStyle{
        Bold:      true,
        FontColor: "FFFFFF",
        Background: "4472C4",
    },
    TableStyle: godocx.TableStyleProfessional,
})
```

### Image Operations

#### InsertImage

```go
func (u *Updater) InsertImage(opts ImageOptions) error
```

Inserts an image with proportional scaling.

**Example:**
```go
updater.InsertImage(godocx.ImageOptions{
    Path:     "./assets/logo.png",
    Width:    300,
    Position: godocx.PositionBeginning,
})
```

### Text Search

#### FindText

```go
func (u *Updater) FindText(pattern string, opts FindOptions) ([]TextMatch, error)
```

Searches for text in the document.

**Example:**
```go
matches, err := updater.FindText("TODO", godocx.FindOptions{
    MatchCase:    true,
    InParagraphs: true,
})
```

### Text Replacement

#### ReplaceText

```go
func (u *Updater) ReplaceText(old, new string, opts ReplaceOptions) (int, error)
```

Replaces text occurrences, returns count replaced.

#### ReplaceTextRegex

```go
func (u *Updater) ReplaceTextRegex(pattern *regexp.Regexp, replacement string, opts ReplaceOptions) (int, error)
```

Regex-based replacement.

### Content Reading

#### GetText

```go
func (u *Updater) GetText() (string, error)
```

Extracts all visible text from document body.

#### GetParagraphText

```go
func (u *Updater) GetParagraphText() ([]string, error)
```

Returns text from each paragraph.

#### GetTableText

```go
func (u *Updater) GetTableText() ([][][]string, error)
```

Returns table data as `[table][row][cell]`.

### Break Operations

#### InsertPageBreak

```go
func (u *Updater) InsertPageBreak(opts BreakOptions) error
```

Inserts a page break.

#### InsertSectionBreak

```go
func (u *Updater) InsertSectionBreak(opts BreakOptions) error
```

Inserts a section break.

### Hyperlink Operations

#### InsertHyperlink

```go
func (u *Updater) InsertHyperlink(text, urlStr string, opts HyperlinkOptions) error
```

Inserts an external hyperlink.

#### InsertInternalLink

```go
func (u *Updater) InsertInternalLink(text, bookmarkName string, opts HyperlinkOptions) error
```

Inserts an internal document link.

### Header/Footer Operations

#### SetHeader

```go
func (u *Updater) SetHeader(content HeaderFooterContent, opts HeaderOptions) error
```

Sets or creates a document header.

#### SetFooter

```go
func (u *Updater) SetFooter(content HeaderFooterContent, opts FooterOptions) error
```

Sets or creates a document footer.

### Document Properties Operations

#### SetCoreProperties

```go
func (u *Updater) SetCoreProperties(props CoreProperties) error
```

Sets the core document properties (metadata).

**Parameters:**
- `props`: Core properties including title, author, subject, keywords, etc.

**Returns:**
- `error`: File write failed, XML generation failed

**Example:**
```go
err := updater.SetCoreProperties(godocx.CoreProperties{
    Title:       "Annual Report 2024",
    Subject:     "Financial Performance",
    Creator:     "Finance Department",
    Keywords:    "finance, annual, report, 2024",
    Description: "Comprehensive financial report for fiscal year 2024",
    Category:    "Financial Reports",
})
```

#### SetAppProperties

```go
func (u *Updater) SetAppProperties(props AppProperties) error
```

Sets application-specific document properties.

**Parameters:**
- `props`: Application properties including company, manager, application name

**Returns:**
- `error`: File write failed, XML generation failed

**Example:**
```go
err := updater.SetAppProperties(godocx.AppProperties{
    Company:     "Acme Corporation",
    Manager:     "John Smith",
    Application: "Microsoft Word",
    AppVersion:  "16.0000",
})
```

#### SetCustomProperties

```go
func (u *Updater) SetCustomProperties(properties []CustomProperty) error
```

Sets custom document properties with typed values.

**Parameters:**
- `properties`: Slice of custom properties with names and typed values

**Returns:**
- `error`: Type inference failed, XML generation failed, file write failed

**Supported value types:**
- `string` → lpwstr
- `int` → i4
- `float64` → r8
- `bool` → bool
- `time.Time` → filetime

**Example:**
```go
err := updater.SetCustomProperties([]godocx.CustomProperty{
    {Name: "ProjectCode", Value: "PRJ-2024-001"},
    {Name: "Budget", Value: 150000.50},
    {Name: "Approved", Value: true},
    {Name: "ReviewDate", Value: time.Date(2024, 12, 31, 0, 0, 0, 0, time.UTC)},
})
```

#### GetCoreProperties

```go
func (u *Updater) GetCoreProperties() (*CoreProperties, error)
```

Retrieves the current core document properties.

**Returns:**
- `*CoreProperties`: Current document metadata
- `error`: File not found, XML parse failed

**Example:**
```go
props, err := updater.GetCoreProperties()
if err != nil {
    return err
}
fmt.Printf("Title: %s\n", props.Title)
fmt.Printf("Author: %s\n", props.Creator)
fmt.Printf("Created: %s\n", props.Created.Format("2006-01-02"))
```

### Bookmark Operations

#### CreateBookmark

```go
func (u *Updater) CreateBookmark(name string, opts BookmarkOptions) error
```

Creates an empty bookmark marker at the specified position.

**Parameters:**
- `name`: Bookmark name (must be valid Word bookmark name)
- `opts`: Bookmark options including position and styling

**Returns:**
- `error`: Invalid name, position not found, or creation failed

**Bookmark name rules:**
- Must start with a letter
- Can contain letters, digits, and underscores
- No spaces or special characters
- Maximum 40 characters

**Example:**
```go
// Create hidden bookmark at document end
err := updater.CreateBookmark("section_start", godocx.BookmarkOptions{
    Position: godocx.PositionEnd,
    Hidden:   true,
})

// Create bookmark after specific text
err := updater.CreateBookmark("summary", godocx.BookmarkOptions{
    Position: godocx.PositionAfterText,
    Anchor:   "Executive Summary",
})
```

#### CreateBookmarkWithText

```go
func (u *Updater) CreateBookmarkWithText(name, text string, opts BookmarkOptions) error
```

Creates a bookmark that wraps specific text content.

**Parameters:**
- `name`: Bookmark name (must be valid Word bookmark name)
- `text`: Text content to bookmark
- `opts`: Bookmark options including position and styling

**Returns:**
- `error`: Invalid name, empty text, position not found, or creation failed

**Example:**
```go
err := updater.CreateBookmarkWithText(
    "important_section",
    "Critical Information",
    godocx.BookmarkOptions{
        Position: godocx.PositionAfterText,
        Anchor:   "Introduction",
        Style:    godocx.StyleHeading2,
    },
)
```

**Use cases:**
- Creating navigation targets for `InsertInternalLink()`
- Document structure markers
- Cross-reference anchors
- Table of contents generation

---

## Fiber v3 Integration

### Basic Setup

```go
package main

import (
    "github.com/gofiber/fiber/v3"
    godocx "github.com/falcomza/go-docx"
)

func main() {
    app := fiber.New(fiber.Config{
        BodyLimit: 100 * 1024 * 1024, // 100MB for DOCX files
    })

    // Routes
    app.Post("/api/documents/generate", GenerateDocumentHandler)
    app.Post("/api/documents/:chartIndex/update", UpdateChartHandler)
    app.Get("/api/documents/preview", PreviewDocumentHandler)

    app.Listen(":3000")
}
```

### Request/Response Types

```go
// Request DTOs
type UpdateChartRequest struct {
    Categories []string                `json:"categories"`
    Series     []SeriesDataRequest     `json:"series"`
    ChartTitle string                  `json:"chartTitle,omitempty"`
    CategoryAxisTitle string           `json:"categoryAxisTitle,omitempty"`
    ValueAxisTitle   string           `json:"valueAxisTitle,omitempty"`
}

type SeriesDataRequest struct {
    Name   string    `json:"name"`
    Values []float64 `json:"values"`
}

type TableInsertRequest struct {
    Columns []ColumnRequest           `json:"columns"`
    Rows    [][]string                `json:"rows"`
    Style   string                    `json:"style,omitempty"`
    Caption string                    `json:"caption,omitempty"`
}

type ColumnRequest struct {
    Title     string `json:"title"`
    Width     int    `json:"width,omitempty"`
    Alignment string `json:"alignment,omitempty"`
}

// Response DTOs
type DocumentResponse struct {
    Success      bool   `json:"success"`
    Message      string `json:"message,omitempty"`
    DownloadURL  string `json:"downloadUrl,omitempty"`
    ProcessTime  int64  `json:"processTimeMs"`
}

type ErrorResponse struct {
    Success  bool   `json:"success"`
    Code     string `json:"code,omitempty"`
    Message  string `json:"message"`
    Context  map[string]interface{} `json:"context,omitempty"`
}
```

---

## Complete Handler Examples

### Chart Update Handler

```go
func UpdateChartHandler(c *fiber.Ctx) error {
    chartIndex, err := c.ParamsInt("chartIndex")
    if err != nil || chartIndex < 1 {
        return c.Status(400).JSON(ErrorResponse{
            Success: false,
            Code:    "INVALID_CHART_INDEX",
            Message: "Chart index must be a positive integer",
        })
    }

    // Parse request body
    var req UpdateChartRequest
    if err := c.BodyParser(&req); err != nil {
        return c.Status(400).JSON(ErrorResponse{
            Success: false,
            Code:    "INVALID_REQUEST",
            Message: err.Error(),
        })
    }

    // Get uploaded template
    file, err := c.FormFile("template")
    if err != nil {
        return c.Status(400).JSON(ErrorResponse{
            Success: false,
            Code:    "NO_TEMPLATE",
            Message: "Template file required",
        })
    }

    // Generate unique output filename
    outputFilename := fmt.Sprintf("report_%d_%s.docx",
        time.Now().Unix(),
        uuid.New().String()[:8])

    // Process document
    success := processDocument(c.Context(), file, outputFilename, chartIndex, req)

    if success {
        return c.JSON(DocumentResponse{
            Success:     true,
            DownloadURL: fmt.Sprintf("/api/documents/download/%s", outputFilename),
        })
    }

    return c.Status(500).JSON(ErrorResponse{
        Success: false,
        Code:    "PROCESSING_FAILED",
        Message: "Failed to process document",
    })
}

func processDocument(
    ctx context.Context,
    file *multipart.FileHeader,
    outputFilename string,
    chartIndex int,
    req UpdateChartRequest,
) bool {
    // Save uploaded file
    tempPath := filepath.Join(os.TempDir(), file.Filename)
    if err := c.SaveFile(file, tempPath); err != nil {
        return false
    }
    defer os.Remove(tempPath)

    // Create updater
    updater, err := godocx.New(tempPath)
    if err != nil {
        return false
    }
    defer updater.Cleanup()

    // Convert request to ChartData
    data := convertToChartData(req)

    // Update chart
    if err := updater.UpdateChart(chartIndex, data); err != nil {
        log.Printf("UpdateChart error: %v", err)
        return false
    }

    // Save output
    outputPath := filepath.Join("./output", outputFilename)
    if err := updater.Save(outputPath); err != nil {
        return false
    }

    return true
}
```

### Batch Operations Handler

```go
type BatchOperation struct {
    Type     string                 `json:"type"` // "chart", "paragraph", "table"
    Payload  map[string]interface{} `json:"payload"`
}

func BatchProcessHandler(c *fiber.Ctx) error {
    var batch struct {
        Operations []BatchOperation `json:"operations"`
    }

    if err := c.BodyParser(&batch); err != nil {
        return err
    }

    file, _ := c.FormFile("document")
    tempPath := saveUploadedFile(file)

    updater, err := godocx.New(tempPath)
    if err != nil {
        return err
    }
    defer updater.Cleanup()

    // Execute operations in sequence
    for _, op := range batch.Operations {
        if err := executeOperation(updater, op); err != nil {
            return c.Status(500).JSON(ErrorResponse{
                Success: false,
                Code:    "OPERATION_FAILED",
                Message: fmt.Sprintf("Operation %s failed: %v", op.Type, err),
            })
        }
    }

    outputPath := generateOutputPath()
    updater.Save(outputPath)

    return c.JSON(DocumentResponse{
        Success:     true,
        DownloadURL: outputPath,
    })
}

func executeOperation(u *godocx.Updater, op BatchOperation) error {
    switch op.Type {
    case "chart":
        return executeChartUpdate(u, op.Payload)
    case "paragraph":
        return executeParagraphInsert(u, op.Payload)
    case "table":
        return executeTableInsert(u, op.Payload)
    default:
        return fmt.Errorf("unknown operation type: %s", op.Type)
    }
}
```

### Streaming Response Handler

```go
func StreamDocumentHandler(c *fiber.Ctx) error {
    templatePath := c.Query("template")
    if templatePath == "" {
        return c.Status(400).JSON(ErrorResponse{
            Message: "template query parameter required",
        })
    }

    // Process document...
    outputPath := processAndSave(templatePath)

    // Set headers for download
    c.Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    c.Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filepath.Base(outputPath)))

    return c.SendFile(outputPath)
}
```

### Background Job Handler

```go
type Job struct {
    ID       string
    Status   string
    Progress int
    Output   string
    Error    string
}

var jobs = sync.Map{}

func SubmitJobHandler(c *fiber.Ctx) error {
    var req UpdateChartRequest
    c.BodyParser(&req)

    file, _ := c.FormFile("template")
    jobID := uuid.New().String()

    job := &Job{ID: jobID, Status: "pending"}
    jobs.Store(jobID, job)

    go processJobAsync(jobID, file, req)

    return c.Status(202).JSON(fiber.Map{
        "jobId":  jobID,
        "status": "pending",
    })
}

func processJobAsync(jobID string, file *multipart.FileHeader, req UpdateChartRequest) {
    job, _ := jobs.Load(jobID)
    j := job.(*Job)

    defer func() {
        if r := recover(); r != nil {
            j.Status = "failed"
            j.Error = fmt.Sprintf("panic: %v", r)
        }
    }()

    tempPath := saveUploadedFileToTemp(file)

    updater, err := godocx.New(tempPath)
    if err != nil {
        j.Status = "failed"
        j.Error = err.Error()
        return
    }
    defer updater.Cleanup()

    j.Status = "processing"
    j.Progress = 25

    // Process...
    data := convertToChartData(req)
    updater.UpdateChart(1, data)
    j.Progress = 75

    outputPath := fmt.Sprintf("./jobs/%s.docx", jobID)
    updater.Save(outputPath)

    j.Status = "completed"
    j.Progress = 100
    j.Output = fmt.Sprintf("/api/documents/jobs/%s", jobID)
}

func JobStatusHandler(c *fiber.Ctx) error {
    jobID := c.Params("id")
    job, exists := jobs.Load(jobID)
    if !exists {
        return c.Status(404).JSON(ErrorResponse{Message: "Job not found"})
    }

    return c.JSON(job)
}
```

### Download Handler

```go
func DownloadDocumentHandler(c *fiber.Ctx) error {
    filename := c.Params("filename")
    if filename == "" {
        return c.Status(400).JSON(ErrorResponse{Message: "filename required"})
    }

    // Validate filename
    if !isValidDocumentFilename(filename) {
        return c.Status(400).JSON(ErrorResponse{Message: "invalid filename"})
    }

    filePath := filepath.Join("./output", filename)

    if _, err := os.Stat(filePath); os.IsNotExist(err) {
        return c.Status(404).JSON(ErrorResponse{Message: "file not found"})
    }

    c.Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    c.Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))

    return c.SendFile(filePath)
}
```

---

## Error Handling Pattern

```go
func handleError(c *fiber.Ctx, err error) error {
    // Check for DocxError
    var docxErr *godocx.DocxError
    if errors.As(err, &docxErr) {
        statusCode := getStatusCodeForError(docxErr.Code)
        return c.Status(statusCode).JSON(ErrorResponse{
            Success: false,
            Code:    string(docxErr.Code),
            Message: docxErr.Message,
            Context: docxErr.Context,
        })
    }

    // Generic error
    return c.Status(500).JSON(ErrorResponse{
        Success: false,
        Message: err.Error(),
    })
}

func getStatusCodeForError(code godocx.ErrorCode) int {
    switch code {
    case godocx.ErrCodeChartNotFound,
         godocx.ErrCodeTextNotFound,
         godocx.ErrCodeFileNotFound:
        return 404
    case godocx.ErrCodeValidation,
         godocx.ErrCodeInvalidChartData,
         godocx.ErrCodeInvalidRegex:
        return 400
    default:
        return 500
    }
}
```

---

## Middleware

### Document Validation Middleware

```go
func DocumentValidator() fiber.Handler {
    return func(c *fiber.Ctx) error {
        file, err := c.FormFile("document")
        if err != nil {
            return c.Status(400).JSON(ErrorResponse{
                Message: "document file required",
            })
        }

        // Validate file extension
        ext := strings.ToLower(filepath.Ext(file.Filename))
        if ext != ".docx" {
            return c.Status(400).JSON(ErrorResponse{
                Message: "only .docx files are supported",
            })
        }

        // Validate file size (max 50MB)
        if file.Size > 50*1024*1024 {
            return c.Status(413).JSON(ErrorResponse{
                Message: "file too large (max 50MB)",
            })
        }

        return c.Next()
    }
}
```

### Cleanup Middleware

```go
func CleanupMiddleware() fiber.Handler {
    return func(c *fiber.Ctx) error {
        // Create temp directory for this request
        tempDir := filepath.Join(os.TempDir(), fmt.Sprintf("docx-%d", time.Now().UnixNano()))
        os.MkdirAll(tempDir, 0755)
        c.Locals("tempDir", tempDir)

        err := c.Next()

        // Cleanup after request
        os.RemoveAll(tempDir)

        return err
    }
}
```

---

## Testing

### Table-Driven Test Example

```go
func TestUpdateChartHandler(t *testing.T) {
    app := fiber.New()
    app.Post("/api/charts/:index/update", UpdateChartHandler)

    tests := []struct {
        name           string
        chartIndex     string
        requestBody    UpdateChartRequest
        expectedStatus int
    }{
        {
            name:       "valid chart update",
            chartIndex: "1",
            requestBody: UpdateChartRequest{
                Categories: []string{"A", "B"},
                Series: []SeriesDataRequest{
                    {Name: "Test", Values: []float64{1, 2}},
                },
            },
            expectedStatus: 200,
        },
        {
            name:           "invalid chart index",
            chartIndex:     "0",
            expectedStatus: 400,
        },
    }

    for _, tt := range tests {
        t.Run(tt.name, func(t *testing.T) {
            body, _ := json.Marshal(tt.requestBody)

            req := httptest.NewRequest("POST", "/api/charts/"+tt.chartIndex+"/update", bytes.NewReader(body))
            req.Header.Set("Content-Type", "application/json")

            resp, err := app.Test(req)

            assert.NoError(t, err)
            assert.Equal(t, tt.expectedStatus, resp.StatusCode)
        })
    }
}
```
