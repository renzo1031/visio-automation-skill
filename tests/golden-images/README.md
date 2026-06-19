# Golden Images

Business regression previews are managed by `scripts/test_visio_business_regression.ps1`.

## Directory Layout

- `actual/` — Latest generated PNG previews from each regression run.
- `expected/` — Approved baseline PNG previews for image diff comparison.

## How It Works

1. Each regression case renders a `.vsdx` and exports a PNG preview.
2. The script performs **structural verification**: node/edge counts, connector OneD, Glue endpoints, and route type cells.
3. The script performs **visual comparison**: pixel-level image diff between actual and expected PNGs when an expected baseline exists.
4. Failures are categorized as `structural` (missing shapes, unglued connectors, wrong route types) or `visual` (pixel diff exceeds threshold).

## First Run / Establishing Golden Baselines

Run with `-ApproveGolden` to copy all actual PNGs into `expected/`:

```powershell
pwsh scripts/test_visio_business_regression.ps1 -ApproveGolden
```

After approval, subsequent runs will compare against those baselines.

## Image Diff Threshold

The default pixel diff threshold is 5% (`-ImageDiffThreshold 0.05`) with a max per-channel delta of 30 (`-ImageDiffMaxPixelDelta 30`). Pixels that differ by more than the max delta count toward the ratio; if the ratio exceeds the threshold, the case is marked as a visual failure.

These defaults tolerate font rendering and anti-aliasing differences while catching meaningful layout regressions.

## Regression Case Coverage

| Case | Kind | Description |
| --- | --- | --- |
| approval-flowchart | flowchart | Approval flowchart with decision branch and curved loop-back |
| bpmn-approval | bpmn | BPMN diagram with Start/End Events, Task, and Gateway |
| network-architecture | network | Network architecture with client, load balancer, API, cache, database |
| dfd-order-processing | dfd | Data Flow Diagram with External Interactor, Data Process, Data Store |
| org-chart | orgchart | Organizational chart with CEO and department hierarchy |
| uml-component | uml-component | UML component diagram with service dependencies |

## Running

```powershell
# Full regression with invisible Visio
pwsh scripts/test_visio_business_regression.ps1

# Visible mode for manual inspection
pwsh scripts/test_visio_business_regression.ps1 -Visible

# Dry run (no Visio, structure check only)
pwsh scripts/test_visio_business_regression.ps1 -NoVisio

# Approve current PNGs as golden baselines
pwsh scripts/test_visio_business_regression.ps1 -ApproveGolden

# Custom pixel threshold
pwsh scripts/test_visio_business_regression.ps1 -ImageDiffThreshold 0.10 -ImageDiffMaxPixelDelta 50
```
