package docxupdater

// ChartData defines chart categories and series values.
type ChartData struct {
	Categories []string
	Series     []SeriesData
	// Optional titles
	ChartTitle      string // Main chart title
	CategoryAxisTitle string // X-axis title
	ValueAxisTitle    string // Y-axis title
}

// SeriesData defines one chart series.
type SeriesData struct {
	Name   string
	Values []float64
}

// ImageOptions defines options for image insertion
type ImageOptions struct {
	// Path to the image file (required)
	Path string
	
	// Width in pixels (optional - if only width is set, height is calculated proportionally)
	Width int
	
	// Height in pixels (optional - if only height is set, width is calculated proportionally)
	Height int
	
	// Alternative text for accessibility (optional)
	AltText string
	
	// Position where to insert the image
	Position InsertPosition
	
	// Anchor text for position-based insertion (for PositionAfterText/PositionBeforeText)
	Anchor string
}

// ImageDimensions stores image width and height in pixels
type ImageDimensions struct {
	Width  int
	Height int
}
