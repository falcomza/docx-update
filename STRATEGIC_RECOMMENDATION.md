# Strategic Recommendation: Chart Manipulation Extension

## 🎯 Recommended Approach

**NO - Don't fork go-template-docx immediately.**

Instead, follow this **3-phase strategy** for maximum efficiency and minimal risk:

---

## Phase 1: Standalone Development (1-2 weeks) ⭐ START HERE

### Why Standalone First?

✅ **Faster Development** - No need to understand go-template-docx internals  
✅ **Easier Testing** - Isolated, focused testing  
✅ **Lower Risk** - Doesn't break existing functionality  
✅ **Flexible Integration** - Can integrate later in multiple ways  
✅ **Reusable** - Works independently or with any DOCX library  

### Action Steps

1. **Create New Module**
   ```bash
   mkdir docx-chart-updater
   cd docx-chart-updater
   go mod init github.com/falcomza/docx-chart-updater
   ```

2. **Use Provided POC as Base**
   - Copy `chart_manipulator_poc.go`
   - Refactor into proper package structure
   - Add comprehensive tests

3. **Implement Core Features**
   ```
   docx-chart-updater/
   ├── chart_updater.go      # Main API
   ├── excel_handler.go      # Excel manipulation
   ├── chart_xml.go          # Chart XML updates
   ├── utils.go              # ZIP/unzip utilities
   ├── types.go              # Data structures
   └── chart_updater_test.go # Tests
   ```

4. **Test Thoroughly**
   - Unit tests for each component
   - Integration tests with real DOCX files
   - Edge cases (empty data, single value, 100+ devices)

### Immediate Benefits

- ✅ Working solution in 1-2 weeks
- ✅ Can use in production immediately
- ✅ Learn the domain before committing to fork
- ✅ Validate approach before major investment

---

## Phase 2: Production Integration (1 week)

### Integration Strategy

Use **both libraries together** (wrapper approach):

```go
package main

import (
    gotemplatedocx "github.com/JJJJJJack/go-template-docx"
    chartupdater "github.com/falco/docx-chart-updater"
)

func generateReport(subsystem string, devices []DeviceAlarm) error {
    // 1. Fill template with go-template-docx
    template, _ := gotemplatedocx.NewDocxTemplateFromFilename("template.docx")
    
    templateData := SubsystemData{
        Subsystem: subsystem,
        TETRA: TetraData{Devices: devices},
    }
    template.Apply(templateData)
    
    // Save to temp file
    tempPath := "/tmp/temp_report.docx"
    template.Save(tempPath)
    
    // 2. Update charts with your library
    cm, _ := chartupdater.NewChartManipulator(tempPath)
    defer cm.Cleanup()
    
    chartData := chartupdater.ChartData{
        Categories: extractDeviceNames(devices),
        Series: []chartupdater.SeriesData{
            {Name: "Critical", Values: extractCriticalCounts(devices)},
            {Name: "Non-critical", Values: extractNonCriticalCounts(devices)},
        },
    }
    
    cm.UpdateChart(1, chartData)
    
    // 3. Save final output
    outputPath := fmt.Sprintf("reports/%s_report_%s.docx", 
        subsystem, time.Now().Format("2006-01-02"))
    cm.Save(outputPath)
    
    return nil
}
```

### Why This Works Well

- ✅ **Best of Both Worlds** - Text/tables from go-template-docx + charts from your library
- ✅ **Clean Separation** - Each library does what it's best at
- ✅ **Maintainable** - Updates to either library are independent
- ✅ **No Temp Files** - Can be optimized to work in-memory later

---

## Phase 3: Decide on Long-term Strategy (After production validation)

### Option A: Keep Standalone (Recommended for most cases)

**When to choose:**
- You only need chart updates
- Want maximum flexibility
- Prefer simple, focused libraries
- Need to support multiple DOCX libraries

**Pros:**
- ✅ Simpler maintenance
- ✅ Easier to update
- ✅ Can work with any DOCX library
- ✅ Single responsibility

### Option B: Fork and Integrate

**When to choose:**
- You use go-template-docx heavily
- Want single unified API
- Will contribute back to community
- Need deep integration

**Pros:**
- ✅ Single library to manage
- ✅ Unified API
- ✅ Can contribute to community

**Cons:**
- ⚠️ Maintain fork long-term
- ⚠️ Sync with upstream changes
- ⚠️ More complex codebase

### Option C: Contribute to go-template-docx

**When to choose:**
- After Option A is proven
- Want to help community
- Have stable, well-tested code

**Approach:**
```bash
# 1. Fork on GitHub
# 2. Create feature branch
git checkout -b feature/chart-manipulation

# 3. Add your code as new module
go-template-docx/
├── template.go           # Existing
├── chart/                # NEW
│   ├── updater.go
│   ├── excel.go
│   └── xml.go
└── examples/
    └── chart_update.go   # NEW

# 4. Submit PR
# 5. Maintain until merged
```

---

## 🎯 My Strong Recommendation: START WITH STANDALONE

### Detailed Implementation Plan

#### Week 1: Core Implementation

**Days 1-2: Project Setup**
```bash
mkdir docx-chart-updater
cd docx-chart-updater
go mod init github.com/falco/docx-chart-updater

# Create structure
mkdir -p {pkg/chart,pkg/excel,pkg/utils,examples,tests}

# Copy POC as starting point
cp chart_manipulator_poc.go pkg/chart/updater.go
```

**Days 3-5: Refactor & Enhance**

1. **Split into modules:**
   ```go
   // pkg/chart/updater.go - Main API
   type Updater struct {
       docxPath string
       tempDir  string
   }
   
   func New(docxPath string) (*Updater, error)
   func (u *Updater) UpdateChart(chartIndex int, data ChartData) error
   func (u *Updater) Save(outputPath string) error
   ```

2. **Excel handling:**
   ```go
   // pkg/excel/handler.go
   type Handler struct {
       worksheetPath string
       stringsPath   string
   }
   
   func (h *Handler) UpdateWorksheet(data ChartData) error
   func (h *Handler) UpdateSharedStrings(categories []string) error
   ```

3. **Chart XML:**
   ```go
   // pkg/chart/xml.go
   func UpdateBarChart(chartPath string, data ChartData) error
   func UpdateLineChart(chartPath string, data ChartData) error
   ```

#### Week 2: Testing & Polish

**Days 6-8: Testing**
```go
// tests/updater_test.go
func TestBasicUpdate(t *testing.T) { }
func TestMultipleSeries(t *testing.T) { }
func TestEmptyData(t *testing.T) { }
func TestLargeDataset(t *testing.T) { }
func TestInvalidTemplate(t *testing.T) { }
```

**Days 9-10: Documentation & Examples**
```go
// examples/basic_usage.go
// examples/from_database.go
// examples/multiple_charts.go
// README.md with full documentation
```

#### Week 3+: Production Integration

**Integrate with iNMS-NG:**

```go
// internal/reporting/alarm_report.go
package reporting

import (
    chartupdater "github.com/falco/docx-chart-updater"
    gotemplatedocx "github.com/JJJJJJack/go-template-docx"
)

type AlarmReportGenerator struct {
    templatePath string
    db          *sql.DB
}

func (g *AlarmReportGenerator) GenerateReport(subsystem string) error {
    // 1. Get data from PostgreSQL
    devices := g.getDeviceAlarms(subsystem)
    
    // 2. Fill template text/tables
    template, _ := gotemplatedocx.NewDocxTemplateFromFilename(g.templatePath)
    template.Apply(SubsystemData{
        Subsystem: subsystem,
        TETRA: TetraData{Devices: devices},
    })
    template.Save("/tmp/temp.docx")
    
    // 3. Update charts
    cm, _ := chartupdater.New("/tmp/temp.docx")
    defer cm.Cleanup()
    
    cm.UpdateChart(1, prepareChartData(devices))
    cm.Save(fmt.Sprintf("reports/%s_%s.docx", 
        subsystem, time.Now().Format("2006-01-02")))
    
    return nil
}
```

---

## 📊 Comparison Matrix

| Approach | Time to Production | Flexibility | Maintenance | Community Benefit |
|----------|-------------------|-------------|-------------|-------------------|
| **Standalone** | ✅ 1-2 weeks | ✅✅✅ High | ✅✅ Easy | ⚠️ Limited |
| **Fork & Extend** | ⚠️ 3-4 weeks | ⚠️ Medium | ⚠️ Complex | ✅ Medium |
| **Contribute PR** | ❌ 4-8 weeks | ⚠️ Medium | ✅✅ Shared | ✅✅✅ High |

---

## 🚀 Action Items for Next Week

### Day 1: Setup
- [ ] Create `docx-chart-updater` repository
- [ ] Initialize Go module
- [ ] Copy POC code as base
- [ ] Set up Git repository

### Day 2-3: Core Development
- [ ] Refactor POC into proper package structure
- [ ] Implement clean API
- [ ] Add error handling
- [ ] Create basic tests

### Day 4-5: Testing
- [ ] Test with your actual template
- [ ] Test with different data sizes
- [ ] Test edge cases
- [ ] Fix bugs

### Week 2: Integration
- [ ] Create integration example with go-template-docx
- [ ] Test in iNMS-NG environment
- [ ] Add PostgreSQL integration
- [ ] Deploy to test environment

### Week 3: Production
- [ ] Generate first production report
- [ ] Monitor for issues
- [ ] Gather feedback
- [ ] Plan enhancements

---

## 💡 Why This Strategy Wins

### 1. **Fast Time to Value**
You'll have working solution in 1-2 weeks vs 4-8 weeks for fork approach.

### 2. **Learn by Doing**
Understanding the problem domain before committing to maintain a fork.

### 3. **Flexibility**
Can later:
- Integrate into go-template-docx
- Keep standalone
- Switch to different DOCX library
- Contribute to community

### 4. **Lower Risk**
If approach doesn't work, you haven't invested in fork maintenance.

### 5. **Production Focus**
Focused on YOUR use case (iNMS-NG alarm reports) first.

---

## 🎓 Learning Path

### Week 1: Understand Domain
- Office Open XML structure
- Excel XLSX format
- Chart XML schemas
- ZIP file manipulation

### Week 2: Build Core
- Extract/modify/repack
- XML parsing/generation
- Error handling
- Testing strategies

### Week 3: Production Hardening
- Performance optimization
- Edge case handling
- Integration patterns
- Monitoring/logging

### Later: Community Contribution
- Clean up code
- Add more chart types
- Write comprehensive docs
- Submit PR to go-template-docx

---

## 🎯 Final Recommendation

**START WITH STANDALONE PACKAGE**

```bash
# Your next commands:
mkdir ~/projects/docx-chart-updater
cd ~/projects/docx-chart-updater
go mod init github.com/falco/docx-chart-updater

# Copy the POC
cp /path/to/chart_manipulator_poc.go ./updater.go

# Start refactoring
code .
```

**Timeline:**
- Week 1: Working standalone library
- Week 2: Integrated with iNMS-NG
- Week 3: Production deployment
- Month 2+: Decide on long-term strategy

**After 1 month of production use, you'll have:**
- ✅ Proven, battle-tested code
- ✅ Clear understanding of requirements
- ✅ Data on edge cases and issues
- ✅ Confidence to contribute to community

**THEN** decide whether to:
1. Keep standalone (probably best)
2. Fork go-template-docx
3. Contribute PR to upstream

---

## ❓ Questions to Consider

Before forking, ask yourself:

1. **Do I need deep integration?**
   - If no → Standalone
   - If yes → Consider fork

2. **Am I willing to maintain fork long-term?**
   - If no → Standalone or contribute PR
   - If yes → Fork might work

3. **Will I use go-template-docx features heavily?**
   - If no → Standalone
   - If yes → Tight integration might help

4. **Do I want to help community?**
   - If yes → Build standalone first, then contribute
   - If no → Keep standalone

---

## 🏁 Summary

**Recommended Path:**

```
Week 1-2:  Build Standalone Package
    ↓
Week 3:    Test in Production
    ↓
Month 2:   Evaluate Options
    ↓
Decision:  Keep Standalone (most likely)
   OR      Contribute to go-template-docx (if you want)
   OR      Fork (only if really needed)
```

**Why This Works:**
- ✅ Fastest to production
- ✅ Lowest risk
- ✅ Maximum learning
- ✅ Keeps options open
- ✅ Production-validated before major commitment

Ready to start? I can help you set up the initial project structure! 🚀
