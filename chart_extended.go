package godocx

// ChartStyle represents predefined chart styles (1-48 in Office)
type ChartStyle int

const (
	ChartStyleNone       ChartStyle = 0  // No specific style
	ChartStyle1          ChartStyle = 1  // Style 1
	ChartStyle2          ChartStyle = 2  // Style 2 (default in many templates)
	ChartStyle3          ChartStyle = 3  // Style 3
	ChartStyleColorful   ChartStyle = 10 // Colorful style
	ChartStyleMonochrome ChartStyle = 42 // Monochrome style
)

// DataLabelPosition defines where data labels appear
type DataLabelPosition string

const (
	DataLabelCenter     DataLabelPosition = "ctr"     // Center of data point
	DataLabelInsideEnd  DataLabelPosition = "inEnd"   // Inside end
	DataLabelInsideBase DataLabelPosition = "inBase"  // Inside base
	DataLabelOutsideEnd DataLabelPosition = "outEnd"  // Outside end
	DataLabelBestFit    DataLabelPosition = "bestFit" // Best fit (auto)
)

// AxisPosition defines axis position
type AxisPosition string

const (
	AxisPositionBottom AxisPosition = "b" // Bottom
	AxisPositionLeft   AxisPosition = "l" // Left
	AxisPositionRight  AxisPosition = "r" // Right
	AxisPositionTop    AxisPosition = "t" // Top
)

// TickMark defines tick mark type
type TickMark string

const (
	TickMarkCross TickMark = "cross" // Cross
	TickMarkIn    TickMark = "in"    // Inside
	TickMarkNone  TickMark = "none"  // None
	TickMarkOut   TickMark = "out"   // Outside (default)
)

// TickLabelPosition defines tick label position
type TickLabelPosition string

const (
	TickLabelHigh   TickLabelPosition = "high"   // High
	TickLabelLow    TickLabelPosition = "low"    // Low
	TickLabelNextTo TickLabelPosition = "nextTo" // Next to axis (default)
	TickLabelNone   TickLabelPosition = "none"   // No labels
)

// BarGrouping defines how bars are grouped
type BarGrouping string

const (
	BarGroupingClustered      BarGrouping = "clustered"      // Clustered (default)
	BarGroupingStacked        BarGrouping = "stacked"        // Stacked
	BarGroupingPercentStacked BarGrouping = "percentStacked" // 100% stacked
	BarGroupingStandard       BarGrouping = "standard"       // Standard
)

// BarDirection defines bar orientation
type BarDirection string

const (
	BarDirectionColumn BarDirection = "col" // Vertical bars (column chart)
	BarDirectionBar    BarDirection = "bar" // Horizontal bars (bar chart)
)

// DataLabelOptions defines options for data labels on chart elements
type DataLabelOptions struct {
	ShowValue        bool              // Show values on data points (default: false)
	ShowCategoryName bool              // Show category names (default: false)
	ShowSeriesName   bool              // Show series names (default: false)
	ShowPercent      bool              // Show percentage (for pie charts) (default: false)
	ShowLegendKey    bool              // Show legend key (default: false)
	Position         DataLabelPosition // Label position (default: bestFit)
	ShowLeaderLines  bool              // Show leader lines for pie charts (default: true)
}

// AxisOptions defines comprehensive axis customization
type AxisOptions struct {
	// Basic properties
	Title        string // Axis title
	TitleOverlay bool   // Title overlays chart area (default: false)

	// Scale properties
	Min       *float64 // Minimum value (nil for auto)
	Max       *float64 // Maximum value (nil for auto)
	MajorUnit *float64 // Major unit interval (nil for auto)
	MinorUnit *float64 // Minor unit interval (nil for auto)

	// Display properties
	Visible  bool         // Show axis (default: true)
	Position AxisPosition // Axis position

	// Tick marks
	MajorTickMark TickMark // Major tick mark style (default: out)
	MinorTickMark TickMark // Minor tick mark style (default: none)

	// Tick labels
	TickLabelPos TickLabelPosition // Tick label position (default: nextTo)

	// Number format
	NumberFormat string // Number format code (e.g., "0.00", "#,##0")

	// Gridlines
	MajorGridlines bool // Show major gridlines (default: true for value axis)
	MinorGridlines bool // Show minor gridlines (default: false)

	// Crossing
	CrossesAt *float64 // Where axis crosses (nil for auto)
}

// LegendOptions defines legend customization
type LegendOptions struct {
	Show     bool   // Show legend (default: true)
	Position string // Position: "r" (right), "l" (left), "t" (top), "b" (bottom), "tr" (top right)
	Overlay  bool   // Legend overlays chart (default: false)
}

// SeriesOptions defines per-series customization
type SeriesOptions struct {
	Name             string            // Series name
	Values           []float64         // Data values
	Color            string            // Hex color (e.g., "FF0000")
	InvertIfNegative bool              // Use different color for negative values (default: false)
	Smooth           bool              // Smooth lines (for line charts) (default: false)
	ShowMarkers      bool              // Show markers (for line charts) (default: false)
	DataLabels       *DataLabelOptions // Data labels for this series (nil for default)
}

// ChartProperties defines chart-level properties
type ChartProperties struct {
	// Appearance
	Style          ChartStyle // Chart style (0-48, 0=none, default: 2)
	RoundedCorners bool       // Use rounded corners (default: false)

	// Behavior
	Date1904 bool   // Use 1904 date system (Mac compatibility) (default: false)
	Language string // Language code (e.g., "en-US", "en-GB") (default: "en-US")

	// Display options
	PlotVisibleOnly       bool   // Plot only visible cells (default: true)
	DisplayBlanksAs       string // How to display blank cells: "gap", "zero", "span" (default: "gap")
	ShowDataLabelsOverMax bool   // Show data labels even if over max (default: false)
}

// BarChartOptions defines options specific to bar/column charts
type BarChartOptions struct {
	Direction  BarDirection // Bar direction (default: col for column)
	Grouping   BarGrouping  // Grouping type (default: clustered)
	GapWidth   int          // Gap between bar groups (0-500, default: 150)
	Overlap    int          // Overlap of bars (-100 to 100, default: 0)
	VaryColors bool         // Vary colors by point (default: false)
}

// ExtendedChartOptions defines comprehensive chart creation options with all customization
type ExtendedChartOptions struct {
	// Position and basic info
	Position InsertPosition
	Anchor   string

	// Chart type
	ChartKind ChartKind

	// Titles
	Title        string
	TitleOverlay bool

	// Data
	Categories []string
	Series     []SeriesOptions // Extended series with per-series options

	// Axes
	CategoryAxis *AxisOptions
	ValueAxis    *AxisOptions

	// Legend
	Legend *LegendOptions

	// Data labels (default for all series)
	DataLabels *DataLabelOptions

	// Chart-level properties
	Properties *ChartProperties

	// Chart type specific options
	BarChartOptions *BarChartOptions

	// Dimensions
	Width  int
	Height int

	// Caption
	Caption *CaptionOptions
}
