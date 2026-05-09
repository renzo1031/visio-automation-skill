---
name: visio-automation
description: Use whenever the user wants Codex to control Microsoft Visio directly for any editable diagram: create or modify .vsdx files, choose built-in stencils/masters/shapes/connectors, replicate screenshots or whiteboards, discover missing masters, fix connector routing/gluing, or avoid non-native generated diagrams. Strongly prefer this skill for Visio, .vsdx, stencil, master, ShapeSheet, dynamic connector, flowchart, BPMN, UML, network, org chart, DFD, VSM, or "use Visio's own shapes".
---

# Visio Automation

Create and modify diagrams by controlling real Microsoft Visio through COM automation. The goal is editable Visio output: built-in master shapes, native dynamic connectors, glued endpoints, and ShapeSheet settings that survive user edits.

## When This Skill Applies

Use this skill when the user wants:

- Visible-mode control of Visio so they can watch the drawing appear.
- A `.vsdx` file made from Visio stencils/masters rather than hand-authored XML or approximate SVG.
- Any editable Visio diagram: flowchart, BPMN, UML, ERD, DFD, network/infrastructure, org chart, value stream, floor/engineering-style diagram, swimlane, or a custom business diagram.
- An image, whiteboard, PDF crop, slide screenshot, or hand sketch recreated as an editable Visio diagram.
- Connector edits such as straight vs. right-angle routing.
- Shape discovery from local Visio stencils.
- A diagram that remains easy to modify manually in Visio.

## Core Workflow

1. Confirm whether the user wants visible mode or background mode. If they already specified visible mode, use `Visio.Application` and set `$visio.Visible = $true`.
2. Identify the diagram family from the user's intent before choosing shapes. Do not overfit to the examples in this skill; treat them as patterns for using Visio's native objects.
3. Inspect `references/master-catalog.md` for a likely stencil/master family. If the diagram type is not listed or the local install differs, run local stencil/master discovery.
4. Choose built-in Visio masters by `NameU` whenever possible. Use primitive `BASIC_*` shapes only when no semantic master exists or the user wants a deliberately generic sketch.
5. Create reusable variables for the selected stencil, node masters, and connector master in the task script so future edits are easy.
6. Use `Page.Drop(master, x, y)` for nodes and containers.
7. Use Visio connector masters plus `GlueToPos` for connections so endpoints move with shapes.
8. Default to Visio's right-angle/orthogonal connector routing. Use straight routing only when the user explicitly asks for straight lines, says the connectors should not be right-angle, or provides a reference image whose connector geometry is clearly straight.
9. Set connector cells for straight lines only when requested:
   - `ShapeRouteStyle = 2`
   - `ConLineRouteExt = 1`
10. Save the `.vsdx`, then verify by reading the opened document or reopening in `Visio.InvisibleApp`.
11. Export a preview PNG when helpful and inspect it before reporting completion.

## General Diagram Triage

Before writing a Visio script, classify the request by semantic intent:

| User intent | Start with | Notes |
| --- | --- | --- |
| Ordinary process, approval, onboarding, test, operations flow | Basic Flowchart | Good default for generic steps, decisions, documents, data, databases, and start/end nodes. |
| Business process with pools, lanes, tasks, events, gateways | BPMN | Prefer BPMN masters instead of approximating with generic diamonds and rectangles. |
| Data movement, actors, processes, data stores | DFD or Gane-Sarson/Yourdon-Coad | Use this only when the content is actually data-flow oriented, not as a generic default. |
| Software architecture, classes, components, sequence, use cases | UML or software-related stencil discovery | If the catalog lacks the needed UML master, discover by terms such as `Class`, `Component`, `Actor`, `Lifeline`, `Use Case`. |
| Network, cloud, infrastructure, topology | Network/cloud stencil discovery | Search installed content by vendor/domain terms before drawing generic boxes. |
| Org structure or reporting lines | Org chart stencil discovery | Use native org chart masters when available; keep connectors glued. |
| Value stream, production, kanban, FIFO, shipment | VSM | Use VSM masters for manufacturing/lean terms. |
| Screenshot/whiteboard replication | Semantic family first, Basic Shapes second | Match meaning with native masters; use basic shapes only for purely visual elements. |
| Existing Visio file edits | Open or attach through ROT | Modify only the relevant shapes/connectors and avoid closing the user's visible window. |

If the request does not fit one row, combine families intentionally: for example, a cloud architecture diagram can use network/cloud masters for systems and flowchart connectors for process steps.

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

Selection guidance:

- Generic flowcharts and step-by-step procedures: `BASFLO_U.VSSX`, `BASFLO_M.VSSX`, or `SSFLOW_M.VSSX`
  - Process: `Process`
  - Decision: `Decision`
  - Start/end: `Start/End`
  - Database: `Database`
  - Connector: `Dynamic connector`
- BPMN workflows: `BPMN_M.VSSX`
  - Task: `Task`
  - Gateway: `Gateway`
  - Start/end/intermediate events: `Start Event`, `End Event`, `Intermediate Event`
  - Pools/lanes: `Pool / Lane`
  - Connector: `Sequence Flow`, `Association`, `Message Flow`, or `Dynamic connector` depending on the requested semantics
- Data-flow diagrams: `DATFLO_M.VSSX`
  - Process: `Data process`
  - External actor: `External interactor`
  - Data store: `Data store`
  - Connector: `Dynamic connector`
- Basic visual replication or generic sketching: `BASIC_U.VSSX` or `BASIC_M.VSSX`
  - `Rectangle`, `Circle`, `Ellipse`, `Diamond`, `Can`
- Yourdon/Coad DFD: `YOURDON_COAD_NOTATION_M.VSSX`
- Gane-Sarson DFD: `GANESA_M.VSSX`
- Value Stream Mapping: `VSM_M.VSSX`

For UML, network, cloud, org chart, engineering, floor plan, or other domain-specific diagrams, use `Find-VisioMasters` with domain terms and a `PreferredStencilRegex` rather than guessing a basic shape. Add verified results to `references/master-catalog.md` when they are reusable.

Prefer localized `_M` stencils on a Chinese Visio install when `_U` stencils are unavailable. Still use `NameU` for master lookup when possible.

## Official Docs And Fallback Research

Read `references/official-docs.md` when you need API details, ShapeSheet cells, or an external documentation fallback plan.

Research rules:

- Use Microsoft Learn first for Visio COM/VBA and ShapeSheet behavior.
- If a method or cell is unclear, search official docs with the exact object, method, or cell name.
- If official docs do not answer it, use third-party examples only as hints and verify locally with a tiny Visio script.
- Record newly verified stencil/master names or cell values in the skill references if they are likely to be reused.
- Correct the skill itself when local testing disproves an assumption. For example, `OpenEx(..., 66)` is the preferred helper value for hidden + read-only stencils.

## Example: Create A Visible Generic Flowchart

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"

$visio = New-VisibleVisioApplication
$doc = $visio.Documents.Add('')
$page = $visio.ActivePage

$stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath 'BASFLO_M.VSSX'
$startMaster = $stencil.Masters.ItemU('Start/End')
$processMaster = $stencil.Masters.ItemU('Process')
$connectorMaster = $stencil.Masters.ItemU('Dynamic connector')

$start = $page.Drop($startMaster, 1, 4)
$start.Text = '开始'
$step = $page.Drop($processMaster, 4, 4)
$step.Text = '处理请求'

Set-VisioShapeFill -Shape $start -Fill 'RGB(255,255,255)' -Line 'RGB(0,0,0)'
Set-VisioShapeFill -Shape $step -Fill 'RGB(255,255,255)' -Line 'RGB(0,0,0)'
Set-VisioTextStyle -Shape $start -Color 'RGB(0,0,0)' -Bold 1
Set-VisioTextStyle -Shape $step -Color 'RGB(0,0,0)'

Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $start -To $step | Out-Null

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
