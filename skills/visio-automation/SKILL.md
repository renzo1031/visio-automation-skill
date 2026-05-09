---
name: visio-automation
description: Use whenever the user wants Codex to control Microsoft Visio directly, create or modify .vsdx diagrams, use Visio built-in stencils/masters/shapes/connectors, replicate an image as an editable Visio diagram, fix Visio connector routing, or avoid generated files that are hard to edit. Strongly prefer this skill for any request mentioning Visio, .vsdx, stencil, master, ShapeSheet, dynamic connector, flowchart in Visio, DFD in Visio, or "use Visio's own shapes".
---

# Visio Automation

Create and modify diagrams by controlling real Microsoft Visio through COM automation. The goal is editable Visio output: built-in master shapes, native dynamic connectors, glued endpoints, and ShapeSheet settings that survive user edits.

## When This Skill Applies

Use this skill when the user wants:

- Visible-mode control of Visio so they can watch the drawing appear.
- A `.vsdx` file made from Visio stencils/masters rather than hand-authored XML or approximate SVG.
- An image or whiteboard screenshot recreated as an editable Visio diagram.
- Connector edits such as straight vs. right-angle routing.
- Shape discovery from local Visio stencils.
- A diagram that remains easy to modify manually in Visio.

## Core Workflow

1. Confirm whether the user wants visible mode or background mode. If they already specified visible mode, use `Visio.Application` and set `$visio.Visible = $true`.
2. Inspect available stencils/masters if the diagram type is not obvious.
3. If this skill does not list the needed stencil/master, run local discovery first, then use official docs or external references only for missing API behavior.
4. Prefer built-in Visio masters by `NameU`; only draw primitive shapes when no suitable master exists.
5. Use `Page.Drop(master, x, y)` for nodes.
6. Use a Visio `Dynamic connector` master plus `GlueToPos` for connections.
7. Default to Visio's right-angle/orthogonal connector routing. Use straight routing only when the user explicitly asks for straight lines, says the connectors should not be right-angle, or provides a reference image whose connector geometry is clearly straight.
8. Set connector cells for straight lines only when requested:
   - `ShapeRouteStyle = 2`
   - `ConLineRouteExt = 1`
9. Save the `.vsdx`, then verify by reading the opened document or reopening in `Visio.InvisibleApp`.
10. Export a preview PNG when helpful and inspect it before reporting completion.

## Use The Bundled Helper Script

Load reusable helpers instead of retyping COM boilerplate:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
```

Useful functions:

- `New-VisibleVisioApplication`
- `New-InvisibleVisioApplication`
- `Open-VisioStencilReadOnly`
- `Get-VisioContentRoots`
- `Find-VisioStencilFiles`
- `Get-VisioMasters`
- `Find-VisioMasters`
- `Set-VisioTextStyle`
- `Set-VisioShapeFill`
- `Connect-VisioShapesOrthogonal`
- `Connect-VisioShapesStraight`
- `Set-VisioConnectorStraight`
- `Add-VisioLabel`
- `Get-OpenVisioDocumentByPath`

If the user has a document already open and Visio refuses to reopen it because it is locked, use `Get-OpenVisioDocumentByPath` to attach to the running document through the Windows Running Object Table.

## When The Skill Does Not Know The Needed Shape

Do not fall back immediately to drawing raw rectangles or lines. Use this recovery path:

1. Search the local master catalog in `references/master-catalog.md`.
2. If nothing matches, search installed Visio stencil files:

   ```powershell
   . "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
   Find-VisioMasters -Query 'Gateway|网关' -PreferredStencilRegex 'BPMN' | Format-Table -AutoSize
   ```

3. Use `PreferredStencilRegex` for domain terms (`BPMN`, `VSM`, `UML`, `AWS`, `Azure`, `BASFLO`) so discovery does not waste time scanning every cloud icon stencil first.
4. If discovery finds a master, open that stencil and use the discovered `NameU`.
5. If discovery does not find a shape, search Microsoft Learn for the API behavior and search installed Visio content by likely domain terms.
6. Create a tiny local proof file with the discovered master or API call.
7. If the proof works and the shape is likely reusable, append it to `references/master-catalog.md`.
8. Only use primitive `BASIC_*` shapes as a final fallback, and state that a direct built-in master was not found.

Discovery performance notes:

- Default discovery searches only stencil files (`*.vssx`, `*.vss`). Templates (`*.vstx`) can be slow or inappropriate for master enumeration, so include them only when the task explicitly asks for templates.
- If a broad query like `Gateway` returns unrelated cloud/network masters, rerun it with a domain-specific `PreferredStencilRegex`, for example `BPMN`.
- If a timed-out discovery leaves invisible `VISIO.EXE` processes, close only blank-title background Visio processes; do not close a user's visible document window.

## Stencil Selection

Read `references/master-catalog.md` when choosing shapes or when a master name is unknown.

Default choices:

- Data-flow/business-flow diagrams: `DATFLO_M.VSSX`
  - Process: `Data process`
  - External actor: `External interactor`
  - Data store: `Data store`
  - Connector: `Dynamic connector`
- Basic visual replication: `BASIC_U.VSSX` or `BASIC_M.VSSX`
  - `Rectangle`, `Circle`, `Ellipse`, `Diamond`, `Can`
- Standard flowcharts: `BASFLO_U.VSSX`, `BASFLO_M.VSSX`, or `SSFLOW_M.VSSX`
  - `Process`, `Decision`, `Start/End`, `Database`, `Dynamic connector`
- Yourdon/Coad DFD: `YOURDON_COAD_NOTATION_M.VSSX`
- Gane-Sarson DFD: `GANESA_M.VSSX`

Prefer localized `_M` stencils on a Chinese Visio install when `_U` stencils are unavailable. Still use `NameU` for master lookup when possible.

## Official Docs And Fallback Research

Read `references/official-docs.md` when you need API details, ShapeSheet cells, or an external documentation fallback plan.

Research rules:

- Use Microsoft Learn first for Visio COM/VBA and ShapeSheet behavior.
- If a method or cell is unclear, search official docs with the exact object, method, or cell name.
- If official docs do not answer it, use third-party examples only as hints and verify locally with a tiny Visio script.
- Record newly verified stencil/master names or cell values in the skill references if they are likely to be reused.
- Correct the skill itself when local testing disproves an assumption. For example, `OpenEx(..., 66)` is the preferred helper value for hidden + read-only stencils.

## Example: Create A Visible Data-Flow Diagram

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"

$visio = New-VisibleVisioApplication
$doc = $visio.Documents.Add('')
$page = $visio.ActivePage

$stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath 'DATFLO_M.VSSX'
$processMaster = $stencil.Masters.ItemU('Data process')
$externalMaster = $stencil.Masters.ItemU('External interactor')
$connectorMaster = $stencil.Masters.ItemU('Dynamic connector')

$user = $page.Drop($externalMaster, 1, 4)
$user.Text = '微信用户'
$browse = $page.Drop($processMaster, 4, 4)
$browse.Text = '浏览'

Set-VisioShapeFill -Shape $user -Fill 'RGB(70,118,242)'
Set-VisioTextStyle -Shape $user -Color 'RGB(255,255,255)' -Bold 1
Set-VisioShapeFill -Shape $browse -Fill 'RGB(239,129,219)'

Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $user -To $browse | Out-Null

$doc.SaveAs('E:\path\diagram.vsdx')
```

## Verification Checklist

Before claiming the diagram is complete, verify:

- The `.vsdx` file exists and has a recent timestamp.
- The document opens or is attached through ROT if already open.
- Node count and connector count are plausible.
- Connectors are `OneD` shapes.
- Default connectors should use right-angle/orthogonal routing unless the user requested straight lines.
- Straight connector requests have `ShapeRouteStyle = 2` and `ConLineRouteExt = 1`.
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
[pscustomobject]@{
  Shapes = $page.Shapes.Count
  Connectors = $connectors.Count
  OrthogonalConnectors = $orthogonal.Count
  StraightConnectors = $straight.Count
}
```

## Practical Notes

- Use `pwsh` for scripts containing UTF-8 Chinese text. Windows PowerShell 5 may misparse UTF-8 files without BOM.
- Keep stencils hidden/read-only with `OpenEx(..., 66)` to avoid save prompts.
- Do not close the user's visible Visio window unless they explicitly ask.
- Avoid generating raw `.vsdx` XML for editable diagrams. It often lacks the native master/connector behavior the user cares about.
- Keep the PowerShell script used for a diagram in the workspace when future edits are likely.
