package main

import (
	"fmt"
	"log"

	docxupdater "github.com/falcomza/docx-update"
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
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_minimal.docx")

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Categories: []string{"A", "B", "C"},
		Series: []docxupdater.SeriesOptions{
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
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_custom_axes.docx")

	minValue := 0.0
	maxValue := 100.0

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Title:      "Monthly Performance",
		ChartKind:  docxupdater.ChartKindColumn,
		Categories: []string{"Jan", "Feb", "Mar", "Apr"},
		Series: []docxupdater.SeriesOptions{
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
		CategoryAxis: &docxupdater.AxisOptions{
			Title:        "Month",
			TitleOverlay: false,
		},
		ValueAxis: &docxupdater.AxisOptions{
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
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_data_labels.docx")

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Title:      "Product Sales",
		ChartKind:  docxupdater.ChartKindBar,
		Categories: []string{"Product A", "Product B", "Product C", "Product D"},
		Series: []docxupdater.SeriesOptions{
			{
				Name:   "Q4 2025",
				Values: []float64{125, 200, 150, 175},
			},
		},
		DataLabels: &docxupdater.DataLabelOptions{
			ShowValue: true,
			Position:  docxupdater.DataLabelOutsideEnd,
		},
		BarChartOptions: &docxupdater.BarChartOptions{
			Direction: docxupdater.BarDirectionBar, // Horizontal bars
			Grouping:  docxupdater.BarGroupingClustered,
		},
		Properties: &docxupdater.ChartProperties{
			Style: docxupdater.ChartStyleColorful,
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created chart with data labels")
}

// Example 4: Full customization
func createFullyCustomizedChart() {
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_full_custom.docx")

	minVal := 0.0
	maxVal := 50.0
	majorUnit := 10.0

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Title:        "Comprehensive Analysis",
		TitleOverlay: false,
		ChartKind:    docxupdater.ChartKindLine,
		Categories:   []string{"Week 1", "Week 2", "Week 3", "Week 4"},
		Series: []docxupdater.SeriesOptions{
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
		CategoryAxis: &docxupdater.AxisOptions{
			Title:          "Time Period",
			Position:       docxupdater.AxisPositionBottom,
			MajorTickMark:  docxupdater.TickMarkOut,
			MinorTickMark:  docxupdater.TickMarkNone,
			MajorGridlines: false,
		},
		ValueAxis: &docxupdater.AxisOptions{
			Title:          "Value",
			Position:       docxupdater.AxisPositionLeft,
			Min:            &minVal,
			Max:            &maxVal,
			MajorUnit:      &majorUnit,
			NumberFormat:   "0.0",
			MajorGridlines: true,
			MinorGridlines: true,
		},
		Legend: &docxupdater.LegendOptions{
			Show:     true,
			Position: "r",
			Overlay:  false,
		},
		Properties: &docxupdater.ChartProperties{
			Style:           docxupdater.ChartStyle2,
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
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_financial.docx")

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Title:      "Quarterly Revenue",
		ChartKind:  docxupdater.ChartKindColumn,
		Categories: []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []docxupdater.SeriesOptions{
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
		ValueAxis: &docxupdater.AxisOptions{
			Title:        "Revenue ($)",
			NumberFormat: "$#,##0",
		},
		DataLabels: &docxupdater.DataLabelOptions{
			ShowValue: true,
			Position:  docxupdater.DataLabelOutsideEnd,
		},
		BarChartOptions: &docxupdater.BarChartOptions{
			Grouping: docxupdater.BarGroupingClustered,
			GapWidth: 150,
		},
		Properties: &docxupdater.ChartProperties{
			Style: docxupdater.ChartStyle2,
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created financial chart with currency formatting")
}

// Example 6: Scientific chart with precise gridlines
func createScientificChart() {
	updater, err := docxupdater.New("templates/template.docx")
	if err != nil {
		log.Fatalf("Failed to create updater: %v", err)
	}
	defer updater.Save("outputs/chart_scientific.docx")

	minVal := 0.0
	maxVal := 100.0
	majorUnit := 20.0
	minorUnit := 5.0

	err = updater.InsertChartExtended(docxupdater.ExtendedChartOptions{
		Title:      "Temperature Over Time",
		ChartKind:  docxupdater.ChartKindLine,
		Categories: []string{"0h", "6h", "12h", "18h", "24h"},
		Series: []docxupdater.SeriesOptions{
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
		CategoryAxis: &docxupdater.AxisOptions{
			Title: "Time (hours)",
		},
		ValueAxis: &docxupdater.AxisOptions{
			Title:          "Temperature (°C)",
			Min:            &minVal,
			Max:            &maxVal,
			MajorUnit:      &majorUnit,
			MinorUnit:      &minorUnit,
			NumberFormat:   "0.0",
			MajorGridlines: true,
			MinorGridlines: true,
			MajorTickMark:  docxupdater.TickMarkOut,
			MinorTickMark:  docxupdater.TickMarkOut,
		},
		Legend: &docxupdater.LegendOptions{
			Show:     true,
			Position: "r",
		},
	})

	if err != nil {
		log.Fatalf("Failed to insert chart: %v", err)
	}

	fmt.Println("✓ Created scientific chart with gridlines")
}
