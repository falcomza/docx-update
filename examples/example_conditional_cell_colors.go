//go:build ignore

package main

import (
	"log"

	godocx "github.com/falcomza/go-docx"
)

func main() {
	// Open or create a document
	updater, err := godocx.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add a title
	err = updater.AddHeading(1, "System Status Report", godocx.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	// Example 1: Simple status-based conditional formatting
	err = updater.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Service Health Monitoring",
		Style:    godocx.StyleHeading2,
		Position: godocx.PositionEnd,
	})
	if err != nil {
		log.Fatalf("Failed to add subheading: %v", err)
	}

	// Create a table with conditional cell coloring based on status
	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Service Name", Alignment: godocx.CellAlignLeft},
			{Title: "Status", Alignment: godocx.CellAlignCenter},
			{Title: "Response Time", Alignment: godocx.CellAlignRight},
			{Title: "Last Check", Alignment: godocx.CellAlignCenter},
		},
		Rows: [][]string{
			{"Database Server", "Critical", "2500ms", "2026-02-20 10:30"},
			{"Web Application", "Normal", "120ms", "2026-02-20 10:29"},
			{"API Gateway", "Warning", "450ms", "2026-02-20 10:31"},
			{"Authentication", "Normal", "85ms", "2026-02-20 10:28"},
			{"Cache Service", "Critical", "Timeout", "2026-02-20 10:32"},
			{"Email Service", "Non-critical", "300ms", "2026-02-20 10:27"},
		},
		HeaderBold:       true,
		HeaderBackground: "2F5496",
		HeaderAlignment:  godocx.CellAlignCenter,
		HeaderStyle: godocx.CellStyle{
			FontColor: "FFFFFF",
		},
		RowStyle: godocx.CellStyle{
			FontSize: 20, // 10pt
		},
		// Conditional styling: color cells based on status text
		ConditionalStyles: map[string]godocx.CellStyle{
			"Critical": {
				Background: "FF0000", // Red
				FontColor:  "FFFFFF", // White text
				Bold:       true,
			},
			"Warning": {
				Background: "FFA500", // Orange
				FontColor:  "000000", // Black text
				Bold:       true,
			},
			"Non-critical": {
				Background: "FFD966", // Light orange/amber
				FontColor:  "000000",
			},
			"Normal": {
				Background: "00B050", // Green
				FontColor:  "FFFFFF",
			},
		},
		BorderStyle:    godocx.BorderSingle,
		BorderSize:     6,
		BorderColor:    "2F5496",
		TableAlignment: godocx.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert status table: %v", err)
	}

	// Example 2: Priority-based conditional formatting
	err = updater.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Issue Tracker",
		Style:    godocx.StyleHeading2,
		Position: godocx.PositionEnd,
	})
	if err != nil {
		log.Fatalf("Failed to add subheading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Issue ID", Alignment: godocx.CellAlignCenter},
			{Title: "Description", Alignment: godocx.CellAlignLeft},
			{Title: "Priority", Alignment: godocx.CellAlignCenter},
			{Title: "Assigned To", Alignment: godocx.CellAlignLeft},
		},
		Rows: [][]string{
			{"ISS-001", "Database connection pool exhausted", "High", "Database Team"},
			{"ISS-002", "UI button alignment issue", "Low", "Frontend Team"},
			{"ISS-003", "Memory leak in background worker", "High", "Backend Team"},
			{"ISS-004", "Documentation typo on API page", "Low", "Documentation"},
			{"ISS-005", "Production server disk space at 95%", "Critical", "DevOps Team"},
			{"ISS-006", "Email notifications delayed", "Medium", "Integration Team"},
		},
		HeaderBold:        true,
		HeaderBackground:  "203864",
		AlternateRowColor: "F2F2F2",
		HeaderStyle: godocx.CellStyle{
			FontColor: "FFFFFF",
		},
		RowStyle: godocx.CellStyle{
			FontSize: 20,
		},
		// Conditional styling for priority levels
		ConditionalStyles: map[string]godocx.CellStyle{
			"Critical": {
				Background: "C00000", // Dark red
				FontColor:  "FFFFFF",
				Bold:       true,
			},
			"High": {
				Background: "FF6B6B", // Light red
				FontColor:  "000000",
				Bold:       true,
			},
			"Medium": {
				Background: "FFE066", // Yellow
				FontColor:  "000000",
			},
			"Low": {
				Background: "B4C7E7", // Light blue
				FontColor:  "000000",
			},
		},
		BorderStyle:    godocx.BorderSingle,
		TableAlignment: godocx.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert issue table: %v", err)
	}

	// Example 3: Performance metrics with threshold-based coloring
	err = updater.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Performance Metrics",
		Style:    godocx.StyleHeading2,
		Position: godocx.PositionEnd,
	})
	if err != nil {
		log.Fatalf("Failed to add subheading: %v", err)
	}

	err = updater.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Metric", Alignment: godocx.CellAlignLeft},
			{Title: "Current Value", Alignment: godocx.CellAlignRight},
			{Title: "Rating", Alignment: godocx.CellAlignCenter},
			{Title: "Target", Alignment: godocx.CellAlignRight},
		},
		Rows: [][]string{
			{"CPU Usage", "45%", "Good", "< 60%"},
			{"Memory Usage", "85%", "Fair", "< 70%"},
			{"Disk I/O", "92%", "Poor", "< 80%"},
			{"Network Latency", "15ms", "Excellent", "< 50ms"},
			{"Error Rate", "5.2%", "Poor", "< 1%"},
			{"Uptime", "99.95%", "Excellent", "> 99.9%"},
		},
		HeaderBold:       true,
		HeaderBackground: "4472C4",
		HeaderStyle: godocx.CellStyle{
			FontColor: "FFFFFF",
		},
		RowStyle: godocx.CellStyle{
			FontSize: 20,
		},
		// Conditional styling for performance ratings
		ConditionalStyles: map[string]godocx.CellStyle{
			"Excellent": {
				Background: "00B050", // Green
				FontColor:  "FFFFFF",
				Bold:       true,
			},
			"Good": {
				Background: "92D050", // Light green
				FontColor:  "000000",
			},
			"Fair": {
				Background: "FFC000", // Amber
				FontColor:  "000000",
			},
			"Poor": {
				Background: "FF0000", // Red
				FontColor:  "FFFFFF",
				Bold:       true,
			},
		},
		BorderStyle:    godocx.BorderSingle,
		BorderSize:     4,
		TableAlignment: godocx.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert metrics table: %v", err)
	}

	// Save the document
	outputPath := "./outputs/example_conditional_cell_colors.docx"
	err = updater.Save(outputPath)
	if err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	log.Printf("Document with conditional cell coloring created successfully: %s", outputPath)
}
