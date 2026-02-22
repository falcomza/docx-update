//go:build ignore

package main

import (
	"fmt"
	"log"

	godocx "github.com/falcomza/go-docx"
)

func main() {
	// Example 1: Minimal - Using defaults
	createMinimalChart()

	// Example 2: Custom axes with formatting
	createCustomAxisChart()

	// Example 3: Data labels and styling
	createDataLabelsChart()

	// Example 4: Full customization
	createFullyCustomizedChart()

	// Example 5: Financial chart with number formatting
	createFinancialChart()

	// Example 6: Scientific chart with gridlines
	createScientificChart()
}

// Example 1: Minimal chart with defaults
func createMinimalChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_minimal.docx")

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Categories: []string{"A", "B", "C"},
		Series: []godocx.SeriesOptions{
			{
				Name:   "Sales",
				Values: []float64{10, 20, 15},
			},
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created minimal chart with defaults")
}

// Example 2: Custom axes with title and formatting
func createCustomAxisChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_custom_axes.docx")

	minValue := 0.0
	maxValue := 100.0

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Title:      "Monthly Performance",
		ChartKind:  godocx.ChartKindColumn,
		Categories: []string{"Jan", "Feb", "Mar", "Apr"},
		Series: []godocx.SeriesOptions{
			{
				Name:   "Target",
				Values: []float64{80, 85, 90, 95},
				Color:  "4472C4", // Blue
			},
			{
				Name:   "Actual",
				Values: []float64{75, 90, 88, 98},
				Color:  "ED7D31", // Orange
			},
		},
		CategoryAxis: &godocx.AxisOptions{
			Title:        "Month",
			TitleOverlay: false,
		},
		ValueAxis: &godocx.AxisOptions{
			Title:          "Performance Score",
			Min:            &minValue,
			Max:            &maxValue,
			NumberFormat:   "0",
			MajorGridlines: true,
			MinorGridlines: false,
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created chart with custom axes")
}

// Example 3: Data labels and styling
func createDataLabelsChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_data_labels.docx")

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Title:      "Product Sales",
		ChartKind:  godocx.ChartKindBar,
		Categories: []string{"Product A", "Product B", "Product C", "Product D"},
		Series: []godocx.SeriesOptions{
			{
				Name:   "Q4 2025",
				Values: []float64{125, 200, 150, 175},
			},
		},
		DataLabels: &godocx.DataLabelOptions{
			ShowValue: true,
			Position:  godocx.DataLabelOutsideEnd,
		},
		BarChartOptions: &godocx.BarChartOptions{
			Direction: godocx.BarDirectionBar, // Horizontal bars
			Grouping:  godocx.BarGroupingClustered,
		},
		Properties: &godocx.ChartProperties{
			Style: godocx.ChartStyleColorful,
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created chart with data labels")
}

// Example 4: Full customization
func createFullyCustomizedChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_full_custom.docx")

	minVal := 0.0
	maxVal := 50.0
	majorUnit := 10.0

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Title:        "Comprehensive Analysis",
		TitleOverlay: false,
		ChartKind:    godocx.ChartKindLine,
		Categories:   []string{"Week 1", "Week 2", "Week 3", "Week 4"},
		Series: []godocx.SeriesOptions{
			{
				Name:        "Series A",
				Values:      []float64{10, 25, 18, 35},
				Color:       "4472C4",
				Smooth:      true,
				ShowMarkers: true,
			},
			{
				Name:        "Series B",
				Values:      []float64{15, 20, 30, 28},
				Color:       "ED7D31",
				Smooth:      true,
				ShowMarkers: true,
			},
		},
		CategoryAxis: &godocx.AxisOptions{
			Title:          "Time Period",
			Position:       godocx.AxisPositionBottom,
			MajorTickMark:  godocx.TickMarkOut,
			MinorTickMark:  godocx.TickMarkNone,
			MajorGridlines: false,
		},
		ValueAxis: &godocx.AxisOptions{
			Title:          "Value",
			Position:       godocx.AxisPositionLeft,
			Min:            &minVal,
			Max:            &maxVal,
			MajorUnit:      &majorUnit,
			NumberFormat:   "0.0",
			MajorGridlines: true,
			MinorGridlines: true,
		},
		Legend: &godocx.LegendOptions{
			Show:     true,
			Position: "r",
			Overlay:  false,
		},
		Properties: &godocx.ChartProperties{
			Style:           godocx.ChartStyle2,
			RoundedCorners:  false,
			DisplayBlanksAs: "gap",
		},
		Width:  int(6.0 * 914400), // 6 inches
		Height: int(4.0 * 914400), // 4 inches
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created fully customized chart")
}

// Example 5: Financial chart with currency formatting
func createFinancialChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_financial.docx")

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Title:      "Quarterly Revenue",
		ChartKind:  godocx.ChartKindColumn,
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []godocx.SeriesOptions{
			{
				Name:   "2025",
				Values: []float64{100000, 120000, 115000, 140000},
				Color:  "70AD47", // Green
			},
			{
				Name:   "2026 (Projected)",
				Values: []float64{110000, 135000, 125000, 160000},
				Color:  "5B9BD5", // Light Blue
			},
		},
		ValueAxis: &godocx.AxisOptions{
			Title:        "Revenue ($)",
			NumberFormat: "$#,##0",
		},
		DataLabels: &godocx.DataLabelOptions{
			ShowValue: true,
			Position:  godocx.DataLabelOutsideEnd,
		},
		BarChartOptions: &godocx.BarChartOptions{
			Grouping: godocx.BarGroupingClustered,
			GapWidth: 150,
		},
		Properties: &godocx.ChartProperties{
			Style: godocx.ChartStyle2,
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created financial chart with currency formatting")
}

// Example 6: Scientific chart with precise gridlines
func createScientificChart() {
	updater, err := godocx.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_scientific.docx")

	minVal := 0.0
	maxVal := 100.0
	majorUnit := 20.0
	minorUnit := 5.0

	err = updater.InsertChartExtended(godocx.ExtendedChartOptions{
		Title:      "Temperature Over Time",
		ChartKind:  godocx.ChartKindLine,
		Categories: []string{"0h", "6h", "12h", "18h", "24h"},
		Series: []godocx.SeriesOptions{
			{
				Name:        "Sensor A",
				Values:      []float64{20.5, 22.3, 25.7, 23.1, 21.8},
				Color:       "FF0000", // Red
				Smooth:      true,
				ShowMarkers: true,
			},
			{
				Name:        "Sensor B",
				Values:      []float64{19.8, 21.5, 24.2, 22.8, 20.9},
				Color:       "0000FF", // Blue
				Smooth:      true,
				ShowMarkers: true,
			},
		},
		CategoryAxis: &godocx.AxisOptions{
			Title: "Time (hours)",
		},
		ValueAxis: &godocx.AxisOptions{
			Title:          "Temperature (°C)",
			Min:            &minVal,
			Max:            &maxVal,
			MajorUnit:      &majorUnit,
			MinorUnit:      &minorUnit,
			NumberFormat:   "0.0",
			MajorGridlines: true,
			MinorGridlines: true,
			MajorTickMark:  godocx.TickMarkOut,
			MinorTickMark:  godocx.TickMarkOut,
		},
		Legend: &godocx.LegendOptions{
			Show:     true,
			Position: "r",
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created scientific chart with gridlines")
}
