package godocx_test

import (
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	godocx "github.com/falcomza/go-docx"
)

// Verifies that when updating a chart with fewer series than it was created with,
// extra <c:ser> elements are removed from chartN.xml
func TestUpdateDropsExtraSeries(t *testing.T) {
	tpl := "templates/docx_template.docx"
	if _, err := os.Stat(tpl); err != nil {
		t.Skip("template not present: " + tpl)
	}

	u, err := godocx.New(tpl)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer u.Cleanup()

	// Insert a new chart with two series
	err = u.InsertChart(godocx.ChartOptions{
		Position:   godocx.PositionEnd,
		ChartKind:  godocx.ChartKindColumn,
		Categories: []string{"A", "B", "C"},
		Series: []godocx.SeriesData{
			{Name: "First", Values: []float64{10, 20, 30}},
			{Name: "Second", Values: []float64{40, 50, 60}},
		},
		ShowLegend: true,
	})
	if err != nil {
		t.Fatalf("InsertChart: %v", err)
	}

	// Find the chart index that was just created
	chartsDir := filepath.Join(u.TempDir(), "word", "charts")
	entries, err := os.ReadDir(chartsDir)
	if err != nil {
		t.Fatalf("read charts dir: %v", err)
	}
	newIdx := 0
	for _, entry := range entries {
		if strings.HasPrefix(entry.Name(), "chart") && strings.HasSuffix(entry.Name(), ".xml") {
			numStr := strings.TrimPrefix(entry.Name(), "chart")
			numStr = strings.TrimSuffix(numStr, ".xml")
			idx, err := strconv.Atoi(numStr)
			if err != nil {
				continue
			}
			if idx > newIdx {
				newIdx = idx
			}
		}
	}
	if newIdx == 0 {
		t.Fatal("could not find inserted chart")
	}

	// Update with a single series (fewer than the two we created with)
	data := godocx.ChartData{
		Categories: []string{"A", "B", "C"},
		Series: []godocx.SeriesData{{
			Name:   "Only",
			Values: []float64{1, 2, 3},
		}},
	}
	if err := u.UpdateChart(newIdx, data); err != nil {
		t.Fatalf("UpdateChart: %v", err)
	}

	// Inspect chartN.xml to ensure there is only one <c:ser>
	chartPath := filepath.Join(u.TempDir(), "word", "charts",
		"chart"+strconv.Itoa(newIdx)+".xml")
	b, err := os.ReadFile(chartPath)
	if err != nil {
		t.Fatalf("read chart xml: %v", err)
	}
	count := strings.Count(string(b), "<c:ser")
	if count != 1 {
		t.Fatalf("expected exactly 1 <c:ser>, got %d", count)
	}
}
