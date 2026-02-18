package main

import (
	"fmt"
	"log"

	docx "github.com/falcomza/docx-update"
)

// This example demonstrates comprehensive list functionality including:
// - Simple bullet lists
// - Numbered lists
// - Multi-level nested lists
// - Mixed bullet and numbered lists
// - Batch list operations

func main() {
	fmt.Println("=== DOCX List Examples ===\n")

	// Example 1: Simple Bullet List
	fmt.Println("📄 Example 1: Simple Bullet List")
	if err := createSimpleBulletList(); err != nil {
		log.Fatalf("Example 1 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_simple_bullet_list.docx\n")

	// Example 2: Simple Numbered List
	fmt.Println("📄 Example 2: Simple Numbered List")
	if err := createSimpleNumberedList(); err != nil {
		log.Fatalf("Example 2 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_simple_numbered_list.docx\n")

	// Example 3: Multi-Level Lists
	fmt.Println("📄 Example 3: Multi-Level Nested Lists")
	if err := createMultiLevelList(); err != nil {
		log.Fatalf("Example 3 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_multilevel_list.docx\n")

	// Example 4: Mixed Content Document
	fmt.Println("📄 Example 4: Mixed Content with Lists")
	if err := createMixedContentDocument(); err != nil {
		log.Fatalf("Example 4 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_mixed_content_lists.docx\n")

	// Example 5: Batch List Operations
	fmt.Println("📄 Example 5: Batch List Operations")
	if err := createBatchListOperations(); err != nil {
		log.Fatalf("Example 5 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_batch_lists.docx\n")

	// Example 6: Style-Based Lists (Legacy Approach)
	fmt.Println("📄 Example 6: Style-Based Lists (Legacy)")
	if err := createStyleBasedLists(); err != nil {
		log.Fatalf("Example 6 failed: %v", err)
	}
	fmt.Println("✓ Created: outputs/example_style_based_lists.docx\n")

	fmt.Println("All list examples completed successfully!")
}

// Example 1: Simple bullet list
func createSimpleBulletList() error {
	// Create new document from template
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title
	if err := u.AddHeading(1, "Simple Bullet List", docx.PositionEnd); err != nil {
		return err
	}

	// Add some text
	if err := u.AddText("Key features of our product:", docx.PositionEnd); err != nil {
		return err
	}

	// Add bullet list items
	if err := u.AddBulletItem("Fast performance", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddBulletItem("Easy to use", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddBulletItem("Secure by default", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddBulletItem("Cross-platform support", 0, docx.PositionEnd); err != nil {
		return err
	}

	return u.Save("outputs/example_simple_bullet_list.docx")
}

// Example 2: Simple numbered list
func createSimpleNumberedList() error {
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title
	if err := u.AddHeading(1, "Installation Steps", docx.PositionEnd); err != nil {
		return err
	}

	// Add numbered list items
	if err := u.AddNumberedItem("Download the installer from our website", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddNumberedItem("Run the installer with administrator privileges", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddNumberedItem("Follow the on-screen instructions", 0, docx.PositionEnd); err != nil {
		return err
	}
	if err := u.AddNumberedItem("Restart your computer to complete installation", 0, docx.PositionEnd); err != nil {
		return err
	}

	return u.Save("outputs/example_simple_numbered_list.docx")
}

// Example 3: Multi-level nested lists
func createMultiLevelList() error {
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title
	if err := u.AddHeading(1, "Project Structure", docx.PositionEnd); err != nil {
		return err
	}

	// Create nested structure with numbered lists
	u.AddNumberedItem("Backend Development", 0, docx.PositionEnd)
	u.AddNumberedItem("API Design", 1, docx.PositionEnd)
	u.AddNumberedItem("REST endpoints", 2, docx.PositionEnd)
	u.AddNumberedItem("GraphQL schema", 2, docx.PositionEnd)
	u.AddNumberedItem("Database Schema", 1, docx.PositionEnd)
	u.AddNumberedItem("User tables", 2, docx.PositionEnd)
	u.AddNumberedItem("Product catalog", 2, docx.PositionEnd)

	u.AddNumberedItem("Frontend Development", 0, docx.PositionEnd)
	u.AddNumberedItem("User Interface", 1, docx.PositionEnd)
	u.AddNumberedItem("Dashboard", 2, docx.PositionEnd)
	u.AddNumberedItem("Settings page", 2, docx.PositionEnd)
	u.AddNumberedItem("Responsive Design", 1, docx.PositionEnd)

	u.AddNumberedItem("Testing & Deployment", 0, docx.PositionEnd)
	u.AddNumberedItem("Unit tests", 1, docx.PositionEnd)
	u.AddNumberedItem("Integration tests", 1, docx.PositionEnd)
	u.AddNumberedItem("Deployment pipeline", 1, docx.PositionEnd)

	// Add a section with bullet lists
	if err := u.AddText("\nTechnologies Used:", docx.PositionEnd); err != nil {
		return err
	}

	u.AddBulletItem("Programming Languages", 0, docx.PositionEnd)
	u.AddBulletItem("Go (Backend)", 1, docx.PositionEnd)
	u.AddBulletItem("TypeScript (Frontend)", 1, docx.PositionEnd)
	u.AddBulletItem("Python (Data processing)", 1, docx.PositionEnd)

	u.AddBulletItem("Frameworks", 0, docx.PositionEnd)
	u.AddBulletItem("Gin (Go web framework)", 1, docx.PositionEnd)
	u.AddBulletItem("React (UI library)", 1, docx.PositionEnd)

	return u.Save("outputs/example_multilevel_list.docx")
}

// Example 4: Mixed content document
func createMixedContentDocument() error {
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title and introduction
	u.AddHeading(1, "Product Documentation", docx.PositionEnd)
	u.AddText("This document provides comprehensive information about our product offering.", docx.PositionEnd)

	// Section 1: Features
	u.AddHeading(2, "Key Features", docx.PositionEnd)
	u.AddText("Our product includes the following features:", docx.PositionEnd)

	u.AddBulletItem("Real-time data synchronization", 0, docx.PositionEnd)
	u.AddBulletItem("Advanced security features", 0, docx.PositionEnd)
	u.AddBulletItem("Customizable workflows", 0, docx.PositionEnd)
	u.AddBulletItem("Comprehensive API", 0, docx.PositionEnd)

	// Section 2: Getting Started
	u.AddHeading(2, "Getting Started", docx.PositionEnd)
	u.AddText("Follow these steps to get started:", docx.PositionEnd)

	u.AddNumberedItem("Create an account", 0, docx.PositionEnd)
	u.AddNumberedItem("Configure your workspace", 0, docx.PositionEnd)
	u.AddNumberedItem("Invite team members", 0, docx.PositionEnd)
	u.AddNumberedItem("Start your first project", 0, docx.PositionEnd)

	// Section 3: Requirements
	u.AddHeading(2, "System Requirements", docx.PositionEnd)
	u.AddText("Minimum requirements:", docx.PositionEnd)

	u.AddBulletItem("Operating System: Windows 10, macOS 10.15, or Linux", 0, docx.PositionEnd)
	u.AddBulletItem("Memory: 8 GB RAM", 0, docx.PositionEnd)
	u.AddBulletItem("Storage: 500 MB available space", 0, docx.PositionEnd)
	u.AddBulletItem("Network: Broadband internet connection", 0, docx.PositionEnd)

	return u.Save("outputs/example_mixed_content_lists.docx")
}

// Example 5: Batch list operations
func createBatchListOperations() error {
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title
	u.AddHeading(1, "Batch List Operations", docx.PositionEnd)

	// Use batch operations for bullet lists
	u.AddText("Team members:", docx.PositionEnd)
	teammates := []string{
		"Alice Johnson - Project Manager",
		"Bob Smith - Lead Developer",
		"Carol White - UX Designer",
		"David Brown - QA Engineer",
	}
	u.AddBulletList(teammates, 0, docx.PositionEnd)

	// Use batch operations for numbered lists
	u.AddText("\nProject milestones:", docx.PositionEnd)
	milestones := []string{
		"Q1: Requirements gathering and planning",
		"Q2: Design and prototyping",
		"Q3: Development and testing",
		"Q4: Launch and post-launch support",
	}
	u.AddNumberedList(milestones, 0, docx.PositionEnd)

	// Another batch example with technical tasks
	u.AddText("\nImmediate action items:", docx.PositionEnd)
	tasks := []string{
		"Review and merge pending pull requests",
		"Update documentation for new API endpoints",
		"Schedule team meeting for sprint planning",
		"Investigate performance issues reported by users",
		"Update dependencies to latest versions",
	}
	u.AddBulletList(tasks, 0, docx.PositionEnd)

	return u.Save("outputs/example_batch_lists.docx")
}

// Example 6: Style-based lists (legacy approach)
func createStyleBasedLists() error {
	u, err := docx.New("templates/docx_template.docx")
	if err != nil {
		return fmt.Errorf("open template: %w", err)
	}
	defer u.Cleanup()

	// Add title
	u.AddHeading(1, "Style-Based Lists (Legacy)", docx.PositionEnd)
	u.AddText("This example shows the legacy style-based approach for lists.", docx.PositionEnd)

	// Using style-based lists requires the template to have ListBullet and ListNumber styles defined
	// This is the old approach, maintained for backward compatibility

	u.AddText("\nUsing ListBullet style:", docx.PositionEnd)
	u.InsertParagraph(docx.ParagraphOptions{
		Text:     "First bullet item (style-based)",
		Style:    docx.StyleListBullet,
		Position: docx.PositionEnd,
	})
	u.InsertParagraph(docx.ParagraphOptions{
		Text:     "Second bullet item (style-based)",
		Style:    docx.StyleListBullet,
		Position: docx.PositionEnd,
	})

	u.AddText("\nUsing ListNumber style:", docx.PositionEnd)
	u.InsertParagraph(docx.ParagraphOptions{
		Text:     "First numbered item (style-based)",
		Style:    docx.StyleListNumber,
		Position: docx.PositionEnd,
	})
	u.InsertParagraph(docx.ParagraphOptions{
		Text:     "Second numbered item (style-based)",
		Style:    docx.StyleListNumber,
		Position: docx.PositionEnd,
	})

	u.AddText("\nNote: The new approach using AddBulletItem() and AddNumberedItem() is recommended.", docx.PositionEnd)

	return u.Save("outputs/example_style_based_lists.docx")
}
