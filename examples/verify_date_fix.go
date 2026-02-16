package main

import (
	"fmt"
	"log"
	"time"

	updater "github.com/falcomza/docx-update/src"
)

func main() {
	// Create from scratch
	u, err := updater.New("tests/testdata/simple_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	err = u.SetCoreProperties(updater.CoreProperties{
		Title:    "Test Document",
		Creator:  "Test User",
		Created:  time.Date(2026, 1, 15, 9, 0, 0, 0, time.UTC),
		Modified: time.Date(2026, 2, 16, 10, 0, 0, 0, time.UTC),
	})
	if err != nil {
		log.Fatal(err)
	}

	if err := u.Save("outputs/verify_dates.docx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("✓ Created verify_dates.docx")
}
