# Diagram Selection

Use this file when choosing a Visio diagram family, stencil, or built-in master.

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

If the request does not fit one row, combine families intentionally. For example, a cloud architecture diagram can use network/cloud masters for systems and flowchart connectors for process steps.

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
