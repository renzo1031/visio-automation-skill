# Visio Automation Skill

这是一个用于让 Codex 通过 Microsoft Visio COM 自动化直接控制 Visio 的 skill。它适合需要 `.vsdx` 可编辑文件、Visio 自带 stencil/master、动态连接线和 ShapeSheet 参数验证的场景。

## 能力

- 使用可见模式控制真实 Visio 窗口。
- 从 Visio 内置 stencil 中拖放 master shape，而不是手写 `.vsdx` XML。
- 创建原生动态连接线，并用 `GlueToPos` 胶合到 shape。
- 默认使用直角/正交动态连接线，只有用户明确要求直线时才设置 `ShapeRouteStyle = 2` 和 `ConLineRouteExt = 1`。
- 搜索本机 Visio stencil/master，解决 skill 文档里没有列出的图形。
- 生成最小 proof `.vsdx` 并导出 PNG 预览。

## 目录结构

```text
skills/
└── visio-automation/
    ├── SKILL.md
    ├── evals/
    │   └── evals.json
    ├── references/
    │   ├── master-catalog.md
    │   └── official-docs.md
    └── scripts/
        ├── test_visio_automation.ps1
        └── visio_helpers.ps1
```

## 环境要求

- Windows
- Microsoft Visio Desktop
- PowerShell 7 (`pwsh`) 推荐
- Visio COM automation 可用

## 安装

将 `skills/visio-automation` 复制到 Codex skills 目录：

```powershell
Copy-Item -Recurse -Force .\skills\visio-automation "$env:USERPROFILE\.codex\skills\visio-automation"
```

## 使用场景

当用户提出类似请求时应触发该 skill：

- “用 Visio 画这个流程图”
- “不要生成死的文件，要用 Visio 自带形状和连接线”
- “把截图复刻成可编辑 `.vsdx`”
- “把 Visio 里的连接线改成直线”
- “画普通流程图，连接线按 Visio 默认直角走线”
- “查找 Visio 里有没有某个 master”

## 自测

运行：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\skills\visio-automation\scripts\test_visio_automation.ps1 -OutputDirectory .\visio-automation-skill-test
```

测试覆盖：

- 已知 DFD master 生成 `.vsdx`
- 动态连接线是否为 `OneD`
- 默认直角/正交连接线 ShapeSheet 参数是否正确
- 显式直线连接线 ShapeSheet 参数是否正确
- 连接端点是否胶合
- 文档中未列出的 BPMN `Gateway` master 是否可通过本机 stencil 搜索发现并落图

## 已验证结果

在本机验证中：

```text
helperLoaded: True
known_passed: True
explicit_straight_passed: True
recovery_passed: True
known_connector_route: 0
known_connector_ext: 0
straight_connector_route: 2
straight_connector_ext: 1
recovery_master: Gateway
```

## 参考资料

官方文档入口见：

- `skills/visio-automation/references/official-docs.md`
- `skills/visio-automation/references/master-catalog.md`

## 说明

这个 skill 优先使用 Visio 原生能力。只有在本机没有合适 stencil/master 时，才建议退回到基础 shape 拼装，并且应在结果中说明该 fallback。
