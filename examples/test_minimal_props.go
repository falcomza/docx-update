package main

import (
	"fmt"
	"log"
	"time"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	fmt.Println("Testing Minimal Properties")
	fmt.Println("===========================\n")

	// Test 1: Only core properties (no dates)
	fmt.Println("Test 1: Core properties without dates...")
	u1, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}

	err = u1.SetCoreProperties(updater.CoreProperties{
		Title:   "Test Document",
		Creator: "Test User",
	})
	if err != nil {
		log.Fatal(err)
	}

	if err := u1.Save("../outputs/test_core_only.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Created: test_core_only.docx\n")

	// Test 2: Core properties with dates
	fmt.Println("Test 2: Core properties WITH dates...")
	u2, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}

	err = u2.SetCoreProperties(updater.CoreProperties{
		Title:    "Test with Dates",
		Creator:  "Test User",
		Created:  time.Date(2026, 1, 15, 9, 0, 0, 0, time.UTC),
		Modified: time.Date(2026, 2, 16, 10, 0, 0, 0, time.UTC),
	})
	if err != nil {
		log.Fatal(err)
	}

	if err := u2.Save("../outputs/test_core_with_dates.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Created: test_core_with_dates.docx\n")

	// Test 3: Core + App properties
	fmt.Println("Test 3: Core + App properties...")
	u3, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}

	err = u3.SetCoreProperties(updater.CoreProperties{
		Title:   "Test with App",
		Creator: "Test User",
	})
	if err != nil {
		log.Fatal(err)
	}

	err = u3.SetAppProperties(updater.AppProperties{
		Company: "Test Company",
		Manager: "Test Manager",
	})
	if err != nil {
		log.Fatal(err)
	}

	if err := u3.Save("../outputs/test_core_app.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Created: test_core_app.docx\n")

	// Test 4: Only custom properties (string only, no dates)
	fmt.Println("Test 4: Custom properties (strings only)...")
	u4, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}

	customProps := []updater.CustomProperty{
		{Name: "Department", Value: "Engineering", Type: "lpwstr"},
		{Name: "Status", Value: "Active", Type: "lpwstr"},
	}

	err = u4.SetCustomProperties(customProps)
	if err != nil {
		log.Fatal(err)
	}

	if err := u4.Save("../outputs/test_custom_strings.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Created: test_custom_strings.docx\n")

	fmt.Println("All test documents created. Please try opening each one in Word.")
}
