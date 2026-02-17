//go:build ignore

package main

import (
	"log"

	updater "github.com/falcomza/docx-update"
)

func main() {
	// Open a DOCX file
	u, err := updater.New("templates/docx_template.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()

	// Example 1: Insert image at the beginning with only width specified
	// Height will be calculated proportionally
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/company_logo.png",
		Width:    400, // Only width specified
		AltText:  "Company Logo",
		Position: updater.PositionBeginning,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 2: Insert image at the end with only height specified
	// Width will be calculated proportionally
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/chart_illustration.jpg",
		Height:   300, // Only height specified
		AltText:  "Chart Illustration",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 3: Insert image with both width and height specified
	// Image will use these exact dimensions (may distort if not proportional)
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/product_photo.png",
		Width:    500,
		Height:   400,
		AltText:  "Product Photo",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 4: Insert image with no dimensions specified
	// Image will use its actual dimensions from the file
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/screenshot.png",
		AltText:  "Application Screenshot",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Example 5: Insert image after specific text
	// First, add some text to anchor to
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "See the diagram below for details:",
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	// Now insert the image after that text
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/diagram.png",
		Width:    600,
		AltText:  "Process Diagram",
		Position: updater.PositionAfterText,
		Anchor:   "See the diagram below",
	}); err != nil {
		log.Fatal(err)
	}

	// Example 6: Insert image before specific text
	if err := u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Figure 1: System Architecture",
		Style:    updater.StyleHeading3,
		Position: updater.PositionEnd,
	}); err != nil {
		log.Fatal(err)
	}

	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/architecture.png",
		Height:   450,
		AltText:  "System Architecture Diagram",
		Position: updater.PositionBeforeText,
		Anchor:   "Figure 1",
	}); err != nil {
		log.Fatal(err)
	}

	// Example 7: Insert image with auto-numbered caption (Figure 1)
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/chart.png",
		Width:    500,
		AltText:  "Sales Chart",
		Position: updater.PositionEnd,
		Caption: &updater.CaptionOptions{
			Type:        updater.CaptionFigure,
			Description: "Q1 Sales Performance",
			AutoNumber:  true,
			Position:    updater.CaptionAfter, // Caption below image
		},
	}); err != nil {
		log.Fatal(err)
	}

	// Example 8: Insert image with caption before (Figure 2)
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/process.png",
		Height:   350,
		AltText:  "Process Flow",
		Position: updater.PositionEnd,
		Caption: &updater.CaptionOptions{
			Type:        updater.CaptionFigure,
			Description: "End-to-End Process Flow",
			AutoNumber:  true,
			Position:    updater.CaptionBefore, // Caption above image
		},
	}); err != nil {
		log.Fatal(err)
	}

	// Example 9: Insert image with centered caption (Figure 3)
	if err := u.InsertImage(updater.ImageOptions{
		Path:     "images/architecture.png",
		Width:    600,
		AltText:  "System Architecture",
		Position: updater.PositionEnd,
		Caption: &updater.CaptionOptions{
			Type:        updater.CaptionFigure,
			Description: "Complete System Architecture",
			AutoNumber:  true,
			Position:    updater.CaptionAfter,
			Alignment:   updater.CellAlignCenter, // Center the caption
		},
	}); err != nil {
		log.Fatal(err)
	}

	// Save the document
	if err := u.Save("output/document_with_images.docx"); err != nil {
		log.Fatal(err)
	}

	log.Println("Document with images created successfully!")
}
