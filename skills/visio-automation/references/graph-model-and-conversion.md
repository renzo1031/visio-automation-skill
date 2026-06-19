# Graph Model And Conversion

Use this file for planned graph rendering, coarse layout, natural-language flow, arrow-style text flow, Mermaid conversion, and draw.io conversion.

## Graph Model Rendering

Use the lightweight Graph Model contract when a diagram is planned structurally before drawing. This is the preferred bridge for layout, Mermaid/draw.io conversion, and image reconstruction.

Minimal usage:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
$graph = [pscustomobject]@{
  diagramType = 'flowchart'
  page = [pscustomobject]@{ width = 6.5; height = 3.0 }
  nodes = @(
    [pscustomobject]@{ id = 'start'; text = 'Start'; semanticType = 'start'; x = 1.2; y = 1.5 }
    [pscustomobject]@{ id = 'process'; text = 'Process'; semanticType = 'process'; x = 3.4; y = 1.5 }
    [pscustomobject]@{ id = 'end'; text = 'End'; semanticType = 'end'; x = 5.4; y = 1.5 }
  )
  edges = @(
    [pscustomobject]@{ from = 'start'; to = 'process'; routeType = 'orthogonal' }
    [pscustomobject]@{ from = 'process'; to = 'end'; routeType = 'orthogonal' }
  )
}
Render-VisioGraphModel -GraphModel $graph -OutputPath 'E:\path\graph.vsdx' -PreviewPath 'E:\path\graph.png' -Force
```

The renderer expects coordinates in Visio inches. Automatic positioning belongs to the layout stage; do not hide unrelated placement rules inside the renderer.

When a Graph Model omits coordinates, call `Invoke-VisioGraphLayout` first. The first implementation gives stable left-to-right or top-to-bottom positions from graph depth and node order, which is enough for smoke tests and for feeding later, richer layout engines.

`Render-VisioGraphModel` and the conversion helpers open foreground visible Visio by default. Pass `-Invisible` only for automated smoke tests, internal verification, or an explicit background/headless user request.

## Natural Language Flow Conversion

Use `Convert-NaturalLanguageFlowToVisio` when the user gives a short business-flow request instead of arrow text. The first version is a rule-based normalizer for clear approval-like and expense-like descriptions; it converts the request to Text Flow, then uses the same Graph Model, layout, routing policy, and Visio renderer.

First-version support:

- Short Chinese flow requests such as leave approval, expense approval, and similar start -> validation -> approval -> branch patterns.
- Common start actions such as `发起申请`, `提交申请`, `提交报销`, `提交报销单`, `发起请求`.
- Validation steps such as `系统校验发票`, `校验资料`, `校验单据`.
- Approval steps such as `部门经理审批`, `主管审核`, `财务复核`.
- Pass/fail branches such as `如果通过就流程结束，如果不通过就返回修改`.
- Completeness branches such as `如果资料不完整就退回补充，如果完整就财务审核`.

Current boundary: this is not a general long-paragraph semantic parser. If normalization fails or the request has many ambiguous branches, ask for arrow text or build a Graph Model directly.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
$text = '帮我画一个请假审批流程：员工发起申请，部门经理审批，如果通过就流程结束，如果不通过就返回修改。'
$flow = Convert-NaturalLanguageFlowToTextFlow -Text $text
Convert-NaturalLanguageFlowToVisio -Text $text -OutputPath 'E:\path\natural-flow.vsdx' -PreviewPath 'E:\path\natural-flow.png' -Force
```

Expense-style example:

```powershell
$text = '画一个报销流程：员工提交报销单，系统校验发票，如果资料不完整就退回补充，如果完整就财务审核，审核通过后付款并归档，审核不通过就驳回。'
$flow = Convert-NaturalLanguageFlowToTextFlow -Text $text
Convert-NaturalLanguageFlowToVisio -Text $text -OutputPath 'E:\path\expense-flow.vsdx' -PreviewPath 'E:\path\expense-flow.png' -Force
```

## Text Flow Conversion

Use `Convert-TextFlowToVisio` when the user gives a clear arrow-style business flow instead of a formal Graph Model or Mermaid block. This is a lightweight natural-language entry for ordinary flowcharts; it normalizes to Graph Model first, then uses the same layout, routing policy, and Visio renderer.

First-version support:

- Ordinary edges: `A -> B`, `A => B`, `A → B`, `A 到 B`, `A 然后 B`.
- Labeled edges: `A --通过-> B`.
- Multiple statements separated by `;` or new lines.
- Basic semantic inference for start/end/decision/document/database/process nodes.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
$text = '发起申请 -> 部门审批 -> 是否通过? --通过-> 流程结束; 是否通过? --拒绝-> 返回修改'
Convert-TextFlowToVisio -Text $text -OutputPath 'E:\path\flow.vsdx' -PreviewPath 'E:\path\flow.png' -Force
```

## Mermaid Conversion

Use the conversion helpers when the user gives Mermaid as the source format. Mermaid normalizes to Graph Model first, then renders through the same Visio renderer, so the output remains native `.vsdx` with Visio masters and dynamic connectors.

Mermaid first-version support is intentionally narrow: `graph TD/LR/TB/BT`, basic flowchart nodes, arrows, and simple edge labels.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
$mermaid = @'
graph TD
    A((开始)) --> B[审批]
    B -- yes --> C((结束))
'@
Convert-MermaidToVisio -MermaidText $mermaid -OutputPath 'E:\path\mermaid.vsdx' -PreviewPath 'E:\path\mermaid.png' -Force
```

## Draw.io Conversion

Use the conversion helpers when the user gives raw/uncompressed draw.io XML or a `.drawio` file. The conversion preserves source ids and maps basic shape styles to flowchart semantic types.

Draw.io first-version limitations:

- Supports ordinary `mxCell` vertices and edges.
- Does not parse compressed draw.io payloads.
- Does not fully support swimlanes, groups, full style catalogs, or detailed waypoint geometry yet.

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
Convert-DrawIOToVisio -DrawIOPath 'E:\path\input.drawio' -OutputPath 'E:\path\drawio.vsdx' -PreviewPath 'E:\path\drawio.png' -Force
```
