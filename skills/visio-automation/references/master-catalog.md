# Visio Stencil And Master Catalog

This catalog records Visio stencils and master names that were verified on the local Windows machine. Prefer `NameU` in scripts because it is language-stable across localized Visio installs.

## Data Flow Diagrams

Use this family when the user wants a data-flow style diagram, business process diagram with external actors, data stores, and directed flows.

Stencil: `DATFLO_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 数据流程 | Data process | Process/function node |
| 外部交互方 | External interactor | External actor/system block |
| 数据存储 | Data store | Data table/store symbol |
| 对象 | Object | Generic object |
| 从中心到中心 1 | Center to center 1 | Connector variant |
| 从中心到中心 2 | Center to center 2 | Connector variant |
| 中心环绕 | Loop on center | Loop connector |
| 中心环绕 2 | Loop on center 2 | Loop connector |
| 多流程 | Multiple process | Multiple-process node |
| 车船使用税 | State | State node, label is unusual in Chinese install |
| 起始状态 | Start state | Start state |
| 停止状态 | Stop state | Stop state |
| 多状态 | Multi state | Multi-state node |
| 停止状态 2 | Stop state 2 | Alternate stop state |
| 实体关系 | Entity relationship | Entity relationship |
| 对象标注 | Object callout | Callout |
| 实体 1 | Entity 1 | Entity |
| 实体 2 | Entity 2 | Entity |
| 椭圆流程 | Oval process | Oval process |
| 椭圆流程(偏移) | Oval process (offset) | Offset oval process |
| 动态连接线 | Dynamic connector | Dynamic connector |

## Yourdon-Coad Notation

Use when `DATFLO_M.VSSX` is missing or when the user explicitly asks for Yourdon/Coad style.

Stencil: `YOURDON_COAD_NOTATION_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 具有 ID 的外部实体 | External Entity with ID | Numbered external entity |
| 数据流程 | Data Process | Process |
| 数据存储 | Data Store | Data store |
| 外部实体 | External Entity | External actor/system |

## Gane-Sarson

Use when the user asks for Gane-Sarson DFD notation or when it better matches the reference image.

Stencil: `GANESA_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 流程 | Process | Process |
| 接口 | Interface | External interface |
| 数据存储 | Data Store | Data store |
| 数据流 | Data Flow | Data flow connector |

## Basic Shapes

Use when a diagram must visually mimic an image and DFD master shapes do not match.

Stencil: `BASIC_U.VSSX` or `BASIC_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 矩形 | Rectangle | Actor blocks, boxes |
| 正方形 | Square | Square blocks |
| 圆形 | Circle | Circular process nodes |
| 椭圆形 | Ellipse | Oval nodes |
| 圆角矩形 | Rounded Rectangle | Friendly blocks |
| 菱形 | Diamond | Decisions |
| 圆柱形 | Can | Database/storage |
| 平行四边形 | Parallelogram | Input/output |

## Basic Flowchart

Use for conventional flowcharts.

Stencil: `BASFLO_U.VSSX`, `BASFLO_M.VSSX`, or `SSFLOW_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 流程 | Process | Process step |
| 判定 / 决策 | Decision | Decision |
| 子流程 | Subprocess | Subprocess |
| 开始/结束 | Start/End | Terminator |
| 文档 | Document | Document |
| 数据 | Data | Data |
| 数据库 | Database | Database |
| 外部数据 | External Data | External data |
| 动态连接线 | Dynamic connector | Dynamic connector |

## Audit Flowchart

Use when the user asks for audit/process-control diagrams.

Stencil: `AUDIT_U.VSSX` or `AUDIT_M.VSSX`

Common `NameU` values:

- `Tagged process`
- `Decision`
- `Tagged document`
- `I/O`
- `Manual operation`
- `Terminator`
- `Manual file`
- `Display`
- `Database`
- `Disk storage`
- `Data transmission`
- `Manual input`
- `Event`
- `Title block`
- `Note block`

## BPMN

Verified during the recovery-path test. Use when the user asks for BPMN, gateways, events, message flows, pools, or lanes.

Stencil: `BPMN_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 任务 | Task | BPMN activity/task |
| 网关 | Gateway | BPMN gateway |
| 中间事件 | Intermediate Event | Intermediate event |
| 结束事件 | End Event | End event |
| 开始事件 | Start Event | Start event |
| 折叠的子流程 | Collapsed Sub-Process | Collapsed subprocess |
| 展开的子流程 | Expanded Sub-Process | Expanded subprocess |
| 文本批注 | Text Annotation | Annotation |
| 序列流 | Sequence Flow | Sequence flow connector |
| 关联 | Association | Association connector |
| 消息流 | Message Flow | Message flow connector |
| 消息 | Message | Message |
| 数据对象 | Data Object | Data object |
| 数据存储 | Data Store | BPMN data store |
| 池/通道 | Pool / Lane | Pool/lane container |

## Value Stream Mapping

Verified during the recovery-path test. Use when the user asks for VSM, value stream, kanban, supermarket, FIFO, shipment, or production control diagrams.

Stencil: `VSM_M.VSSX`

| Chinese name | NameU | Typical use |
| --- | --- | --- |
| 流程 | Process | Process |
| 库存 | Inventory | Inventory |
| 上推箭头 | Push arrow | Push flow |
| 价值流图 | VSM | VSM frame |
| 客户/供应商 | Customer/Supplier | Customer or supplier |
| 运输箭头 | Shipment arrow | Shipment flow |
| 运输卡车 | Shipment truck | Truck shipment |
| 生产控制 | Production control | Production control |
| 人工信息 | Manual information | Manual information flow |
| 电子信息 | Electronic information | Electronic information flow |
| 模拟运算表 | Data table | Data table |
| 生产看板 | Production kanban | Kanban |
| 超市 | Supermarket | Supermarket |
| FIFO 通道 | FIFO lane | FIFO lane |
| Kaizen 爆发 | Kaizen burst | Kaizen burst |

## Discovery Commands

Use these when a stencil or master is missing:

```powershell
. "$env:USERPROFILE\.codex\skills\visio-automation\scripts\visio_helpers.ps1"
Get-VisioMasters -StencilNameOrPath 'DATFLO_M.VSSX' | Format-Table -AutoSize
```

Search installed Visio content:

```powershell
Get-ChildItem -Path "$env:ProgramFiles\Microsoft Office\root\Office16\Visio Content" `
  -Recurse -Include '*.vssx','*.vss','*.vstx' |
  Select-Object FullName
```
