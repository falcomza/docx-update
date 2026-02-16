package main

import (
	"fmt"
	"log"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Create a minimal test with just basic properties (no dates)
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Test 1: Only core properties (no dates)
	fmt.Println("Test 1: Basic core properties...")
	props := updater.CoreProperties{
		Title:   "Test Title",
		Creator: "Test Author",
	}
	if err := u.SetCoreProperties(props); err != nil {
		log.Fatal(err)
	}

	if err := u.Save("outputs/test_simple_props.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Saved test_simple_props.docx")

	// Test 2: Core properties with dates
	fmt.Println("\nTest 2: Core properties with dates...")
	u2, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u2.Cleanup()

	// Skip dates for now - just set basic strings
	props2 := updater.CoreProperties{
		Title:   "Test With More Fields",
		Creator: "Test Author",
		Subject: "Test Subject",
	}
	if err := u2.SetCoreProperties(props2); err != nil {
		log.Fatal(err)
	}

	if err := u2.Save("outputs/test_props_no_dates.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Saved test_props_no_dates.docx")

	// Test 3: Just custom properties without dates
	fmt.Println("\nTest 3: Custom properties without dates...")
	u3, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u3.Cleanup()

	customProps := []updater.CustomProperty{
		{Name: "StringProp", Value: "Test String"},
		{Name: "IntProp", Value: 42},
		{Name: "FloatProp", Value: 3.14},
		{Name: "BoolProp", Value: true},
	}
	if err := u3.SetCustomProperties(customProps); err != nil {
		log.Fatal(err)
	}

	if err := u3.Save("outputs/test_custom_simple.docx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("✓ Saved test_custom_simple.docx")

	fmt.Println("\nAll tests complete. Try opening each file to find which triggers the error.")
}
