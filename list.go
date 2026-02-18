package docxupdater

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// List numbering IDs
const (
	BulletListNumID   = 1 // Numbering ID for bullet lists
	NumberedListNumID = 2 // Numbering ID for numbered lists
)

// ensureNumberingXML ensures numbering.xml exists with bullet and numbered list support
func (u *Updater) ensureNumberingXML() error {
	numberingPath := filepath.Join(u.tempDir, "word", "numbering.xml")

	// Check if numbering.xml already exists
	if _, err := os.Stat(numberingPath); err == nil {
		// File exists, check if it has our list definitions
		data, err := os.ReadFile(numberingPath)
		if err != nil {
			return fmt.Errorf("read numbering.xml: %w", err)
		}

		// If it already has our numbering definitions, we're done
		if strings.Contains(string(data), fmt.Sprintf(`w:numId="%d"`, BulletListNumID)) &&
			strings.Contains(string(data), fmt.Sprintf(`w:numId="%d"`, NumberedListNumID)) {
			return nil
		}
	}

	// Create new numbering.xml with bullet and numbered list definitions
	numberingXML := generateNumberingXML()
	if err := os.WriteFile(numberingPath, []byte(numberingXML), 0o644); err != nil {
		return fmt.Errorf("write numbering.xml: %w", err)
	}

	// Update content types if needed
	if err := u.ensureNumberingContentType(); err != nil {
		return fmt.Errorf("update content types: %w", err)
	}

	// Update document.xml.rels if needed
	if err := u.ensureNumberingRelationship(); err != nil {
		return fmt.Errorf("update relationships: %w", err)
	}

	return nil
}

// generateNumberingXML creates a complete numbering.xml with bullet and numbered list definitions
func generateNumberingXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
             mc:Ignorable="w14">
  
  <!-- Abstract Numbering Definition for Bullets -->
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    
    <!-- Level 0 -->
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="●"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    
    <!-- Level 1 -->
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="○"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    
    <!-- Level 2 -->
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="■"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    
    <!-- Level 3-8 (additional levels with increasing indentation) -->
    <w:lvl w:ilvl="3">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="●"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="4">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="○"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="5">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="■"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="4320" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="6">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="●"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="7">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="○"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="8">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="■"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="6480" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  
  <!-- Abstract Numbering Definition for Numbered Lists -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    
    <!-- Level 0: 1, 2, 3... -->
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 1: a, b, c... -->
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 2: i, ii, iii... -->
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 3: 1), 2), 3)... -->
    <w:lvl w:ilvl="3">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%4)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 4: (a), (b), (c)... -->
    <w:lvl w:ilvl="4">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="(%5)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 5: (i), (ii), (iii)... -->
    <w:lvl w:ilvl="5">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="(%6)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="4320" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <!-- Level 6-8: Same as level 0-2 with more indentation -->
    <w:lvl w:ilvl="6">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%7."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="7">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%8."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    
    <w:lvl w:ilvl="8">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%9."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="6480" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  
  <!-- Concrete Numbering Instance for Bullets -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  
  <!-- Concrete Numbering Instance for Numbered Lists -->
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
  
</w:numbering>`
}

// ensureNumberingContentType adds numbering.xml to [Content_Types].xml if not present
func (u *Updater) ensureNumberingContentType() error {
	contentTypesPath := filepath.Join(u.tempDir, "[Content_Types].xml")
	data, err := os.ReadFile(contentTypesPath)
	if err != nil {
		return fmt.Errorf("read [Content_Types].xml: %w", err)
	}

	content := string(data)

	// Check if numbering override already exists
	if strings.Contains(content, `PartName="/word/numbering.xml"`) {
		return nil // Already present
	}

	// Add numbering override before </Types>
	numberingOverride := `  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
	content = strings.Replace(content, "</Types>", numberingOverride+"\n</Types>", 1)

	return os.WriteFile(contentTypesPath, []byte(content), 0o644)
}

// ensureNumberingRelationship adds numbering.xml relationship to document.xml.rels if not present
func (u *Updater) ensureNumberingRelationship() error {
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	data, err := os.ReadFile(relsPath)
	if err != nil {
		return fmt.Errorf("read document.xml.rels: %w", err)
	}

	content := string(data)

	// Check if numbering relationship already exists
	if strings.Contains(content, `Target="numbering.xml"`) {
		return nil // Already present
	}

	// Find the next available relationship ID
	nextID := u.getNextRelationshipID(content)

	// Add numbering relationship before </Relationships>
	numberingRel := fmt.Sprintf(`  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`, nextID)
	content = strings.Replace(content, "</Relationships>", numberingRel+"\n</Relationships>", 1)

	return os.WriteFile(relsPath, []byte(content), 0o644)
}
