package main

import (
	"fmt"
	"log"
	"time"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	fmt.Println("Document Properties Example")
	fmt.Println("=============================")

	// Open template
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// 1. Set Core Properties
	fmt.Println("\n1. Setting Core Properties...")
	coreProps := updater.CoreProperties{
		Title:          "Annual Report 2026",
		Subject:        "Company Performance and Financial Results",
		Creator:        "Corporate Communications",
		Keywords:       "annual report, 2026, financial, performance, revenue",
		Description:    "Comprehensive annual report covering all aspects of company performance in fiscal year 2026",
		Category:       "Annual Reports",
		LastModifiedBy: "Review Committee",
		Revision:       "3",
		Created:        time.Date(2026, 1, 15, 9, 0, 0, 0, time.UTC),
		Modified:       time.Now(),
	}

	if err := u.SetCoreProperties(coreProps); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Title:", coreProps.Title)
	fmt.Println("   ✓ Author:", coreProps.Creator)
	fmt.Println("   ✓ Keywords:", coreProps.Keywords)

	// 2. Set Application Properties
	fmt.Println("\n2. Setting Application Properties...")
	appProps := updater.AppProperties{
		Company:     "Global Innovations Corp",
		Manager:     "Alice Johnson",
		Application: "Microsoft Word",
		AppVersion:  "16.0000",
	}

	if err := u.SetAppProperties(appProps); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Company:", appProps.Company)
	fmt.Println("   ✓ Manager:", appProps.Manager)

	// 3. Set Custom Properties (for workflow, automation, tracking)
	fmt.Println("\n3. Setting Custom Properties...")
	customProps := []updater.CustomProperty{
		// String properties
		{Name: "Department", Value: "Corporate Communications"},
		{Name: "DocumentType", Value: "Annual Report"},
		{Name: "Status", Value: "Final"},
		{Name: "Classification", Value: "Public"},

		// Numeric properties
		{Name: "FiscalYear", Value: 2026},
		{Name: "Revenue", Value: 125000000.75},
		{Name: "EmployeeCount", Value: 5420},
		{Name: "GrowthRate", Value: 12.5},

		// Boolean properties
		{Name: "IsAudited", Value: true},
		{Name: "IsBoardApproved", Value: true},
		{Name: "RequiresSignature", Value: false},

		// Date properties
		{Name: "FiscalYearEnd", Value: time.Date(2026, 12, 31, 0, 0, 0, 0, time.UTC)},
		{Name: "PublicationDate", Value: time.Date(2027, 1, 15, 0, 0, 0, 0, time.UTC)},

		// Tracking properties
		{Name: "ProjectCode", Value: "RPT-2026-ANNUAL"},
		{Name: "VersionNumber", Value: 3},
		{Name: "ReviewCycle", Value: "Q4-2026"},
	}

	if err := u.SetCustomProperties(customProps); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("   ✓ Set %d custom properties\n", len(customProps))
	for _, prop := range customProps[:5] { // Show first 5
		fmt.Printf("     - %s: %v\n", prop.Name, prop.Value)
	}
	fmt.Println("     - ...")

	// 4. Read back core properties to verify
	fmt.Println("\n4. Reading Core Properties...")
	retrievedProps, err := u.GetCoreProperties()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("   Retrieved properties:")
	fmt.Println("   - Title:", retrievedProps.Title)
	fmt.Println("   - Creator:", retrievedProps.Creator)
	fmt.Println("   - Category:", retrievedProps.Category)
	fmt.Println("   - Revision:", retrievedProps.Revision)

	// 5. Add some content to the document
	fmt.Println("\n5. Adding content to document...")
	err = u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This document has comprehensive metadata including core, application, and custom properties.",
		Position: updater.PositionEnd,
		Bold:     true,
	})
	if err != nil {
		log.Fatal(err)
	}

	// Add a table with company information
	err = u.InsertTable(updater.TableOptions{
		Columns: []updater.ColumnDefinition{
			{Title: "Metric"},
			{Title: "2026"},
			{Title: "2025"},
			{Title: "Change"},
		},
		Rows: [][]string{
			{"Revenue ($M)", "$125.0", "$111.2", "+12.4%"},
			{"Employees", "5,420", "5,100", "+6.3%"},
			{"Net Income ($M)", "$18.5", "$15.2", "+21.7%"},
		},
		Position:         updater.PositionEnd,
		TableStyle:       updater.TableStyleGridAccent1,
		HeaderBold:       true,
		HeaderBackground: "4472C4",
	})
	if err != nil {
		log.Fatal(err)
	}

	// 6. Save the document
	outputPath := "outputs/example_with_properties.docx"
	fmt.Println("\n6. Saving document...")
	if err := u.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("\n✓ Document saved successfully:", outputPath)
	fmt.Println("\nTo view properties in Word:")
	fmt.Println("  1. Open the document")
	fmt.Println("  2. Go to File > Info")
	fmt.Println("  3. Click 'Properties' > 'Advanced Properties'")
	fmt.Println("  4. View Summary, Statistics, and Custom tabs")
}
