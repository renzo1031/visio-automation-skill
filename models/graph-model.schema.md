# Visio Graph Model

This document defines the lightweight graph contract used before rendering to Visio.

## Shape

```json
{
  "diagramType": "flowchart",
  "routingIntent": "balanced",
  "layout": {
    "strategy": "Flowchart",
    "direction": "LR"
  },
  "page": {
    "width": 8.5,
    "height": 6.0
  },
  "nodes": [],
  "edges": []
}
```

`layout` is optional. When nodes do not provide coordinates, `Invoke-VisioGraphLayout` can assign stable positions before rendering. Supported first-version directions are `LR`, `RL`, `TB`, and `BT`; supported strategy names are `Flowchart`, `Hierarchical`, `Tree`, `Network`, and `Radial`. The current flowchart layout uses deterministic graph-depth placement, ignores obvious reverse/loopback edges for main-rank calculation, keeps positive branches in the main reading direction, and places negative/rework branches in lower slots.

`routingIntent` is optional. When set, it controls how the renderer selects connector route types and glue points. Allowed values: `fidelity`, `balanced`, `clean`. When omitted, the intent is auto-detected from `sourceFormat` and `diagramType` — image-reconstruction resolves to `fidelity`, network/UML resolves to `clean`, everything else defaults to `balanced`. You can also pass `-RoutingIntent` to `Render-VisioGraphModel` directly.

## Nodes

Each node represents one Visio shape.

| Field | Required | Description |
| --- | --- | --- |
| `id` | yes | Stable node id used by edges. |
| `text` | no | Shape text. |
| `semanticType` | no | Domain type such as `start`, `process`, `decision`, `database`. |
| `stencil` | no | Visio stencil file, defaults to `BASFLO_M.VSSX`. |
| `master` | no | Exact Visio master `NameU`. Overrides `semanticType`. |
| `preferredMaster` | no | Alternative master hint when `master` is not set. |
| `x`, `y` | no | Shape center in Visio inches. Defaults to `1.0`, `1.0`. |
| `width`, `height` | no | Shape size in inches. |
| `style` | no | Optional fill, line, text color, text size, and bold settings. |

## Text Annotations

Optional `annotations`, `labels`, or `texts` collections can describe standalone text regions that connectors should avoid when automatic routing is enabled.

| Field | Required | Description |
| --- | --- | --- |
| `id` | no | Stable annotation id. |
| `text` | no | Displayed text or note content. |
| `x`, `y` | yes | Text region center in Visio inches. |
| `width`, `height` | yes | Text region size in Visio inches. |
| `bounds` | no | Alternative nested bounds object with `x`, `y`, `width`, and `height`. |

## Regions And Containers

Optional `containers`, `swimlanes`, `lanes`, `pools`, or `groups` collections can describe large diagram regions whose boundaries should be avoided by unrelated connectors when automatic routing is enabled. These regions are routing hints in the first version; they do not automatically render Visio container or swimlane shapes unless a renderer step explicitly creates matching nodes/shapes.

| Field | Required | Description |
| --- | --- | --- |
| `id` | no | Stable region id. |
| `text` | no | Region title, lane name, or container label. |
| `x`, `y` | yes | Region center in Visio inches. |
| `width`, `height` | yes | Region size in Visio inches. |
| `bounds` | no | Alternative nested bounds object with `x`, `y`, `width`, and `height`. |

When either endpoint of an edge is inside a region, that region is not treated as an unrelated boundary obstacle for that edge. This prevents real cross-lane or in-container connections from being over-penalized.

## Edges

Each edge renders as a native Visio connector glued to two nodes.

| Field | Required | Description |
| --- | --- | --- |
| `from` | yes | Source node id. |
| `to` | yes | Target node id. |
| `text` | no | Connector label. |
| `stencil` | no | Connector stencil, defaults to `BASFLO_M.VSSX`. |
| `connector` | no | Connector master, defaults to `Dynamic connector`. |
| `routeType` | no | `orthogonal`, `straight`, or `curved`. |
| `fromX`, `fromY` | no | Source glue position, relative to the source shape. |
| `toX`, `toY` | no | Target glue position, relative to the target shape. |
| `style` | no | Optional line color, line weight, and end arrow settings. |

When `routeType` is not set on an edge, the routing policy decides the default based on `routingIntent` and `diagramType`. Explicit `routeType` on an edge always takes priority over the routing policy defaults — the policy only fills in route types when edges leave them blank.

When `fromX/fromY/toX/toY` are not set, the renderer can choose glue points automatically. It evaluates the four common ports on each node (`top`, `bottom`, `left`, `right`) and selects the lowest-scoring pair based on route length, direction, loopback status, layout direction, routing intent, and first-version obstacle avoidance. Explicit glue coordinates always win over automatic port routing.

## Routing Policy

The routing policy controls how the renderer selects connector route types and glue points when edges do not specify them explicitly.

### Routing Intent

| Intent | Behavior |
| --- | --- |
| `fidelity` | Preserve original connector styles from the source. Used automatically for image-reconstruction. Edges without `routeType` fall back to the diagram-type default. |
| `balanced` | Default intent. Loop-back edges (edges that flow against the layout direction) use the diagram-type loopback default (usually `curved`). Other edges use the diagram-type default (usually `orthogonal`). Explicit `routeType` is preserved. |
| `clean` | Prioritize readability — orthogonal routing, predictable glue sides, fewer crossings. Loop-back edges still use diagram-type loopback defaults. Explicit `routeType` is preserved. |

### Diagram-Type Routing Defaults

| Diagram Type | Default Route | Loopback Route | Cross Edges Curved |
| --- | --- | --- | --- |
| flowchart | orthogonal | curved | no |
| bpmn / workflow | orthogonal | curved (top→top) | yes |
| network / UML | orthogonal | orthogonal | yes |
| org-chart | orthogonal | orthogonal | no |
| DFD | orthogonal | curved | yes |
| image-reconstruction | preserve original | preserve original | yes |

### Port Routing Summary

`Render-VisioGraphModel` returns routing diagnostics for each rendered edge:

| Field | Description |
| --- | --- |
| `fromSide` | Chosen source port: `top`, `bottom`, `left`, or `right`. Empty when explicit glue coordinates are used. |
| `toSide` | Chosen target port: `top`, `bottom`, `left`, or `right`. Empty when explicit glue coordinates are used. |
| `routeScore` | Internal score used to compare candidate port pairs. Lower is better. |
| `routeLength` | Manhattan distance between the chosen source and target ports in Visio inches. |
| `obstacleHits` | Number of intermediate node obstacles intersected by the chosen candidate orthogonal path. |
| `avoidancePenalty` | Score penalty added for obstacle intersections. Lower is better; `0` means the chosen candidate avoided known node obstacles. |
| `labelHits` | Number of standalone text annotation regions intersected by the chosen candidate path. |
| `labelPenalty` | Score penalty added for crossing text annotation regions. |
| `boundaryHits` | Number of unrelated container, swimlane, lane, pool, or group regions intersected by the chosen candidate path. |
| `boundaryPenalty` | Score penalty added for crossing unrelated region boundaries. |
| `crossingHits` | Number of intersections with already planned connector paths. |
| `crossingPenalty` | Score penalty added for crossing existing connector paths. |

Example: in a left-to-right approval flow, a rejection edge that goes from a lower-right revision step back to an upper-left approval step can choose `fromSide = "left"` and `toSide = "bottom"` to avoid an outer detour.

First-version obstacle avoidance treats other nodes, standalone text annotations, and unrelated region boundaries as padded high-cost rectangles and compares simple orthogonal candidate paths between candidate ports. The renderer also keeps lightweight candidate paths for earlier edges so later edges can penalize obvious connector crossings. It is not a full grid/A* router yet: exact Visio waypoint geometry is still handled by Visio's native Dynamic Connector or future routing enhancements.

### Renderer

`Render-VisioGraphModel` consumes this object model and returns a render summary. If coordinates are missing, run layout explicitly or pass a layout strategy to the renderer:

```powershell
Render-VisioGraphModel -GraphModel $graph -OutputPath .\out.vsdx -PreviewPath .\out.png -Force
Render-VisioGraphModel -GraphModel $graph -OutputPath .\out.vsdx -PreviewPath .\out.png -LayoutStrategy Flowchart -LayoutDirection LR -Force
Render-VisioGraphModel -GraphModel $graph -OutputPath .\out.vsdx -PreviewPath .\out.png -RoutingIntent clean -Force
```

The renderer preserves editable Visio behavior by using real masters, `Page.Drop`, native dynamic connectors, and `GlueToPos`. The `-RoutingIntent` parameter (or `routingIntent` field on the Graph Model) controls how connector route types and glue points are resolved when edges do not specify them.

## Text Flow Input

`Convert-TextFlowToVisioGraphModel` provides a lightweight user-facing bridge for simple business processes written as arrow text. It is meant for clear flow descriptions, not arbitrary prose.

Supported first-version forms:

```text
发起申请 -> 部门审批 -> 是否通过? --通过-> 流程结束
是否通过? --拒绝-> 返回修改
```

Rules:

- `->`, `=>`, `→`, `到`, `至`, `然后`, and `再到` create ordinary edges.
- `--标签->` creates an edge with a connector label.
- `;` or a new line separates independent statements.
- Node text ending in `?` or `？`, or starting with decision-like words such as `是否`, maps to `decision`.
- Common start/end/document/database words map to matching flowchart semantic types. Labels beginning with words such as `开始`, `发起`, `启动`, or `提交` map to `start`; other labels map to `process`.
- The output is the same Graph Model contract used by Mermaid, draw.io, image reconstruction, layout, and Visio rendering.

## Natural Language Flow Input

`Convert-NaturalLanguageFlowToVisioGraphModel` is the first ordinary-user prose bridge. It accepts short Chinese business-flow requests, normalizes them to Text Flow, then emits the same Graph Model contract.

Supported first-version forms:

```text
帮我画一个请假审批流程：员工发起申请，部门经理审批，如果通过就流程结束，如果不通过就返回修改。
画一个报销流程：员工提交报销单，系统校验发票，如果资料不完整就退回补充，如果完整就财务审核，审核通过后付款并归档，审核不通过就驳回。
```

The normalizer currently targets clear approval-like patterns:

- Start actions: `发起申请`, `提交申请`, `提交报销`, `提交报销单`, `发起请求`.
- Validation actions: `系统校验发票`, `校验资料`, `校验单据`.
- Approval actions: `审批`, `审核`, `复核`, optionally with a role such as `部门经理`.
- Branches: pass/fail phrases such as `通过` and `不通过`, plus completeness phrases such as `完整` and `不完整`.
- Clear expense-style flows can produce two decision nodes, for example `资料是否完整?` and `审核是否通过?`.
- It records `sourceFormat = "natural-language"`, `sourceText`, and `normalizedTextFlow` on the resulting graph.

Boundary: use Text Flow or a direct Graph Model for long paragraphs, multiple unrelated branches, or ambiguous node ownership.

## Image Reconstruction Contract

When screenshot or whiteboard reconstruction is used, the upstream step should provide a structured object that can be normalized into the same renderer input:

```json
{
  "diagramType": "image-reconstruction",
  "source": {
    "imagePath": "whiteboard.png",
    "width": 600,
    "height": 300
  },
  "page": {
    "width": 6.0,
    "height": 3.0
  },
  "nodes": [
    {
      "id": "capture",
      "text": "截图",
      "semanticType": "start",
      "bounds": {
        "x": 40,
        "y": 80,
        "width": 120,
        "height": 60
      }
    }
  ],
  "connectors": [
    {
      "from": "capture",
      "to": "extract",
      "routeType": "orthogonal",
      "fromPoint": { "x": 1.0, "y": 0.5 },
      "toPoint": { "x": 0.0, "y": 0.5 }
    }
  ]
}
```

Rules:

- `source.width` and `source.height` are pixel dimensions of the reference image.
- `bounds` are pixel-space rectangles in the source image.
- `fromPoint` and `toPoint` are relative glue points on each node, in the same spirit as other graph-model edge glue positions.
- `routeType` still uses `orthogonal`, `straight`, or `curved`.
- OCR and CV are out of scope for this contract; the upstream step must provide the structured data.
