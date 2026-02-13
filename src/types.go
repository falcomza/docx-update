package docxchartupdater

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
