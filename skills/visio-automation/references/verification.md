# Verification

Use this file when you need to verify routing, glue, ShapeSheet cells, output files, or preview fidelity.

## Verification Checklist

Before claiming the diagram is complete, verify:

- The `.vsdx` file exists and has a recent timestamp.
- The document opens or is attached through ROT if already open.
- Node count and connector count are plausible.
- Connectors are `OneD` shapes.
- Default connectors should use right-angle/orthogonal routing unless the user requested straight lines.
- Image reconstruction preserves each observed connector route type: right-angle, straight, or curved.
- Straight connector requests have `ShapeRouteStyle = 2` and `ConLineRouteExt = 1`.
- Curved connector requests have `ConLineRouteExt = 2`.
- Connector endpoints remain glued; `BeginX` and `EndX` should usually contain glue formulas after connection.
- A preview export visually matches the requested layout closely enough.

Suggested verification snippet:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
$doc = Get-OpenVisioDocumentByPath 'E:\path\diagram.vsdx'
if ($null -eq $doc) {
  $visio = New-InvisibleVisioApplication
  $doc = $visio.Documents.Open('E:\path\diagram.vsdx')
}
$page = $doc.Pages.Item(1)
$connectors = @($page.Shapes) | Where-Object { $_.OneD -ne 0 }
$straight = $connectors | Where-Object {
  $_.CellExistsU('ConLineRouteExt', 0) -ne 0 -and $_.CellsU('ConLineRouteExt').ResultIU -eq 1
}
$orthogonal = $connectors | Where-Object {
  $_.CellExistsU('ConLineRouteExt', 0) -ne 0 -and $_.CellsU('ConLineRouteExt').ResultIU -eq 0
}
$curved = $connectors | Where-Object {
  $_.CellExistsU('ConLineRouteExt', 0) -ne 0 -and $_.CellsU('ConLineRouteExt').ResultIU -eq 2
}
[pscustomobject]@{
  Shapes = $page.Shapes.Count
  Connectors = $connectors.Count
  OrthogonalConnectors = $orthogonal.Count
  StraightConnectors = $straight.Count
  CurvedConnectors = $curved.Count
}
```

## Practical Notes

- Use `pwsh` for scripts containing UTF-8 Chinese text. Windows PowerShell 5 may misparse UTF-8 files without BOM.
- Run user-facing Visio work in foreground visible mode by default.
- Keep stencils hidden/read-only with `OpenEx(..., 66)` to avoid save prompts.
- Avoid generating raw `.vsdx` XML for editable diagrams.
