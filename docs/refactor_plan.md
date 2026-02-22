# Refactor Plan (Current Codebase)

This plan replaces the previous draft and reflects the repository as it exists now.

## Current Baseline

- `chart_copy.go`, `chart_copy_test.go`, and copy-example files are already removed.
- Shared helpers previously attributed to `chart_copy.go` already exist in current files:
  - `findNextChartIndex` in `chart.go`
  - `getNextDocPrId`, `getNextDocumentRelId` in `helpers.go`
  - `copyFile` in `utils.go`
- Zip Slip protection is already present in `extractZip()` (`utils.go`).
- `go test ./...` currently fails for two categories:
  1. `examples` and `tools` packages: multiple `main redeclared` build errors.
  2. Core chart tests:
     - `TestInsertChartWithAxisTitles`
     - `TestInsertMultipleCharts`
     - `TestInsertChartMultipleSeries`
     - `TestInsertChartAtBeginning`
     - `TestUpdateDropsExtraSeries`

## Goals

1. Stabilize CI/build (`go test ./...`, `go vet ./...`).
2. Fix chart correctness regressions first.
3. Clean up duplicated internals (escaping, relationship ID logic, regex usage).
4. Remove stale `CopyChart` references from docs/examples.
5. Defer module/package rename to a dedicated follow-up PR.

## Execution Plan

## Phase 0: Stabilize Build Surface

### Step 0.1: Exclude standalone executable samples/tools from package builds

Add `//go:build ignore` to files that are standalone binaries and currently break `./...`:

- `examples/example_extended_chart.go`
- `examples/example_lists.go`
- `examples/verify_word_fix.go`
- `examples/test_simple_add.go`
- `examples/example_page_layout.go`
- `examples/test_addtext_only.go`
- `examples/test_replace_only.go`
- `examples/example_table_orientation.go`
- `tools/compare_docx.go`
- `tools/debug_tool.go`
- `tools/extract_docx.go`
- `tools/fix_empty_properties.go`
- `tools/validate_relationships.go`

Acceptance:
- `go test ./...` no longer fails with `main redeclared` errors.

### Step 0.2: Re-baseline failures after build cleanup

Run:

```bash
go test . -count=1
```

Capture only remaining library failures before making behavior changes.

---

## Phase 1: Fix Chart Regressions (P0)

### Step 1.1: Investigate chart file/index assumptions

Focus on failing tests in:
- `chart_insert_test.go`
- `chart_update_series_count_test.go`

Check whether tests are incorrectly hard-coding `word/charts/chart1.xml` when inserted chart index may differ. If needed, update tests to discover the created chart file from relationships/content rather than assuming index `1`.

### Step 1.2: Validate chart XML generation details

Trace `InsertChart` output for:
- chart title
- axis titles
- series names
- number of `<ser>` entries after `UpdateChart`

Adjust implementation and/or test fixtures based on the true intended behavior (do not weaken assertions unless behavior is intentionally changed and documented).

### Step 1.3: Add targeted regression coverage

Add/adjust tests so failures are specific and stable:
- verify inserted chart by relationship target, not filename guessing
- verify series count after update using namespace-safe matching (`<c:ser` or XML parse) rather than fragile string assumptions

Acceptance:
- `go test . -count=1` passes.

---

## Phase 2: Internal Cleanup (P1)

### Step 2.1: Consolidate XML escaping helpers

- Keep `xmlEscape` in `helpers.go`.
- Remove duplicated `escapeXML` in `replace.go`.
- Remove duplicated `escapeXMLAttribute` in `hyperlink.go`.
- Update call sites in:
  - `replace.go`
  - `hyperlink.go`
  - `headerfooter.go`
  - `bookmark.go`

Acceptance:
- No duplicate escape helpers remain.
- Tests unchanged in behavior and still pass.

### Step 2.2: Consolidate relationship ID generation

Introduce one helper in `helpers.go`, e.g.:

```go
func getNextRelIDFromFile(relsPath string) (string, error)
```

Refactor call sites in `hyperlink.go`, `headerfooter.go`, and any remaining rel-id consumers to use the shared helper.

Acceptance:
- No parallel implementations for next-rel-id lookup remain.

### Step 2.3: Remove avoidable regex recompilation in hot paths

In `replace.go`:
- Move static regexes to package-level compiled vars (likely `constants.go`).
- Compile dynamic regexes once per outer call, not inside callback loops.

Acceptance:
- No `regexp.MustCompile(...)` inside `ReplaceAllStringFunc` callbacks.

### Step 2.4: Replace `fmt.Sscanf` integer parsing with `strconv.Atoi`

Update known spots:
- `image.go`
- `hyperlink.go`
- `bookmark.go`
- any additional occurrences found by grep

Acceptance:
- No `fmt.Sscanf(..., "%d", ...)` usage for simple integer extraction remains.

---

## Phase 3: API/Docs Consistency (P1)

### Step 3.1: Resolve `CopyChart` drift

Current state: `CopyChart` references exist in docs/examples, but implementation is absent in `.go` source files.

Official removal:
- remove/replace all `CopyChart` references in docs and examples with `InsertChart` + `UpdateChart` workflow.

Acceptance:
- Public docs match exported API exactly.
- No broken usage examples.

### Step 3.2: Update documentation links/import paths only after API is stable

Update `README.md` and docs under `docs/` for API accuracy first. Keep module path rename out of this phase.

---

## Phase 4: Module/Package Rename (Completed)

Rename target:
- package: `docxupdater` -> `godocx`
- module: `github.com/falcomza/docx-update` -> `github.com/falcomza/go-docx`

Completed scope:
- `go.mod` module path updated.
- all package declarations updated (`godocx` / `godocx_test`).
- all Go import paths updated to `github.com/falcomza/go-docx`.
- docs/examples updated to the new import path and alias usage.

Post-rename validation status:
- `go test ./... -count=1` passes.
- `go vet ./...` passes.

---

## Verification Checklist

Run in order:

```bash
go test ./... -count=1
go vet ./...
```

Expected end state:
- zero compile errors across module
- zero test failures
- docs/API references aligned with actual exported functionality

## Suggested Commit Slicing

1. Build hygiene: build tags for examples/tools.
2. Chart regression fixes + related tests.
3. Internal cleanup (escape/rel-id/regex/Atoi).
4. Docs/API consistency (`CopyChart` decision).
5. Optional rename PR (module/package/imports).
