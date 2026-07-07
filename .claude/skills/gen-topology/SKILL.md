---
name: gen-topology
description: 将云服务 Excel 清单转换为交互式 HTML 拓扑图。当用户提供 Excel 文件路径、或要求生成拓扑图、云架构图时触发。
user-invocable: true
allowed-tools:
  - Bash(python *)
  - Read
---

# /gen-topology — Excel 转云服务拓扑图

根据用户提供的 Excel 文件，调用 `converter.py` 生成可交互的 HTML 拓扑图。

参数：`$ARGUMENTS`

---

## 参数解析

从 `$ARGUMENTS` 中提取以下信息：

- `excel`：Excel 文件路径（必填）。若用户未提供，询问文件路径后再继续。
- `output`：输出 HTML 路径（可选，默认与 Excel 同目录同名，扩展名改为 `.html`）
- `--title`：图表标题（可选，默认 `云服务拓扑图`）

若 `$ARGUMENTS` 为空，询问用户："请提供 Excel 文件路径，例如：`/gen-topology C:/data/services.xlsx`"

---

## 执行步骤

1. 确认 Excel 文件存在（用 Read 工具或 Bash 检查）
2. 运行转换命令：

```bash
python E:/draw-picture/ourresult/converter.py <excel路径> [output路径] [--title "标题"]
```

3. 报告生成结果：输出 HTML 文件路径、节点数、边数

---

## Excel 表格格式说明

若用户询问格式，告知以下列定义：

| 列名 | 是否必填 | 说明 |
|------|----------|------|
| 服务名称 | 必填 | 节点唯一标识 |
| 服务类型 | 推荐 | ecs/elb/cce/rds/waf/cdn/vpc 等 |
| 区域 | 可选 | 按区域自动生成容器分组 |
| 资源分组 | 可选 | 区域内二级分组 |
| 所属业务 | 可选 | 业务分组容器（绿色框） |
| 下游服务 | 可选 | 逗号分隔的目标服务名，画实线箭头 |
| 下游服务(推断) | 可选 | 逗号分隔，画灰色虚线箭头 |
| 描述 | 可选 | 侧边栏详情显示 |
| 规格 | 可选 | 侧边栏详情显示 |

支持的服务类型（`服务类型` 列填写）：
`ecs` `bms` `cce` `elb` `waf` `cdn` `vpc` `nat` `vpn` `rds` `redis` `kafka` `obs` `sfs` `apig` `iam` 等

---

## 生成结果说明

生成的 HTML 无需服务器，浏览器直接打开，支持：
- 缩放/平移，自适应视图
- 点击节点查看详情
- 搜索节点名称或类型
- 切换布局（层次/网格/圆形）
- 业务分组视图开关
- 导出 PNG 图片

---

## 错误处理

| 错误 | 处理方式 |
|------|----------|
| Excel 文件不存在 | 提示用户检查路径 |
| 缺少 openpyxl | 提示运行 `pip install openpyxl` |
| 找不到服务名称列 | 提示检查表头，给出支持的列名列表 |
| cytoscape.min.js 缺失 | 提示确认 `E:/draw-picture/cloudmapper-main/web/js/` 目录完整 |
