# 云服务拓扑图生成器工作交接文档

## 1. 文档范围

本文档对应当前 Git 分支 `draw-picture-normally`，基线提交为 `f08aeb6`（2026-09-01，撤销“聚合所有 CCE 资源”改动）。核心产品代码位于 `ourresult/`；仓库中的 `cloudmapper-main/` 提供 Cytoscape 及布局插件资源，`cytoscape.js-unstable/` 是备用引擎目录。

当前工作区还存在若干未跟踪的压测/浏览器测试产物，交接时不要将它们误认为核心源码，也不要用清理命令覆盖用户已有改动。

## 2. Skill 的使用与上手

项目配套 skill 是 `.agents/skills/gen-topology/SKILL.md`，触发场景是“提供 Excel 清单、生成云架构图或拓扑图”。skill 的职责是解析参数、检查输入文件、调用转换器并汇报输出统计；实际业务逻辑仍在 `ourresult/converter.py`。

### 2.1 调用方式

```text
/gen-topology <Excel路径> [输出HTML路径] [--title "图表标题"]
```

等价命令：

```powershell
py ourresult\converter.py <Excel路径> [输出HTML路径] --title "图表标题"
```

`output` 省略时，输出文件与 Excel 同名、扩展名改为 `.html`；也可以使用 `--output-dir <目录>` 指定输出目录。生成结果是自包含 HTML，浏览器直接打开即可，不需要启动 Web 服务器或 Node.js。

### 2.2 环境准备

- Python 3.7+。
- `openpyxl`，安装：`pip install -r requirements.txt` 或 `pip install openpyxl`。
- 转换器需要读取以下前端资源：
  - `cloudmapper-main/web/js/cytoscape.min.js`（优先，必须与插件版本匹配）；
  - `cloudmapper-main/web/js/cytoscape-cose-bilkent.js`；
  - `cloudmapper-main/web/js/dagre.js`、`cytoscape-dagre.js`。
- 图标从 `ourresult/picture/*.png` 内嵌到 HTML；缺失图标时节点仍可渲染，只是不显示图标。

### 2.3 最小上手流程

1. 准备 Excel，首行是表头，至少包含“服务名称/resource name”。
2. 推荐同时填写“服务类型/resource_type”“所属业务”“资源分组”“区域”和“下游服务”。
3. 运行转换命令。
4. 打开 HTML，默认使用 BDAT 预设布局；先点击“自适应”检查全图，再按需搜索、查看详情或调整布局。

仓库内可用 `ourresult/create_sample.py` 生成示例清单：

```powershell
py ourresult\create_sample.py
py ourresult\converter.py ourresult\sample.xlsx
```

## 3. Excel 输入契约

列名大小写不敏感，并支持中英文及别名；完整映射见 `converter.py` 的 `COL_ALIASES`。关键列如下：

| 规范列 | 必填 | 作用 |
| --- | --- | --- |
| 服务名称 | 是 | 资源唯一名称，也是调用关系引用键。 |
| 服务类型 | 否（推荐） | 决定服务分类、图标、颜色、架构层；为空时尝试从名称识别。 |
| 资源ID | 否 | 详情展示；包含 `ccaas` 时可辅助识别为 CC。 |
| 企业项目 / 企业项目ID | 否 | CCE 部署的项目标识；名称优先，ID 作为回退显示。 |
| 区域 | 否 | 相同区域的资源放入 Region 容器；为空则不挂区域容器。 |
| 资源分组 | 否 | 资源组聚合和 CCE 二级分组框的依据。 |
| 所属业务 | 否 | 业务容器主键；大小写、空格、连字符、下划线差异会归一化。 |
| 是否核心业务 | 否 | 只要该列存在，就只绘制值严格等于“是”的行。 |
| 下游服务 | 否 | 已确认调用关系，支持逗号、中文逗号、分号和换行分隔。 |
| 下游服务(推断) | 否 | 保留到详情，不生成拓扑边。 |
| 推断原因 | 否 | 详情侧栏中的分析依据；原因文本中的链路不会生成边。 |
| 描述、规格、来源 | 否 | 详情和聚合明细表展示。 |

名称为 `N/A`、`NA`、`None`、`Null`、`无` 或包含“暂无/不适用/客户未提供”等占位语义的行会被过滤。

## 4. 节点与聚合规则

### 4.1 服务识别

`get_service_key()` 先看类型字段，再看资源 ID、资源分组提示和资源名称关键词。支持 ECS、BMS、CCE、CCE_Deployment、CCI、ELB/SLB、WAF、EIP、VPC、NAT、VPN、CDN、RDS、DCS/Redis、DMS/Kafka、OBS/OSS、SFS、APIG、CSS、DDS、GaussDB、DWS、SWR/ECR、IAM/KMS、Nginx、Tomcat、K8s、MaaS 等。类型为空、`default`、`Other` 且无法识别时归入 `default`（界面名为 Other），该类资源会被过滤，不生成 Other 容器或外部目标。

### 4.2 业务归属

- 有“所属业务”：按归一化后的业务键建业务框；例如 `lingee-prod`、`LINGEE_prod`、`lingee prod` 会视为同一业务。
- 无“所属业务”：按服务类型创建虚拟业务框，例如 CCE、RDS；虚拟业务框不会再套同名服务框。
- 未填业务的 CCE_Deployment 可回退并入已有业务：必须同时满足企业项目相同、`resource_name` 前缀相同。前缀默认取名称按 `-` 分隔的前 3 段，由 `CCE_BUSINESS_PREFIX_SEGMENTS` 控制；多个候选业务时选择已有 CCE 节点数最多者。无法匹配时留在虚拟 CCE 业务框。

### 4.3 架构层归类

业务内部固定四层（从上到下）：

1. 渠道接入与安全边界：CDN、WAF、EIP、DNS、CNAD。
2. 网络与负载均衡层：ELB、SLB、VPN、ER；VPC、NAT、CC、DCAAS 保持网络逻辑分类，但在业务框右侧独立轨道展示。
3. 计算、容器与应用服务层：ECS、BMS、CCE、CCI、FunctionGraph、APIG、Nginx、Tomcat、SWR/ECR、K8s、MaaS 等。
4. 中间件、数据与存储层：RDS、DCS、Redis、DDS、OBS/OSS、Kafka/DMS、CSS、SFS、EVS、CBR、GaussDB、DWS、ZooKeeper、RabbitMQ 等。

存在已确认 CCE 下游关系时，名称含 `db`/`tidb` 的 ECS 会被提升到数据层；CCE 指向名称含 `db`/`tidb` 的 ELB 时，该 ELB 进入计算层，其下游 ECS 进入数据层。这些上下文重分类只使用“下游服务”确认边，不使用推断字段。

### 4.4 聚合键（当前分支最重要的行为）

聚合元组为：

```text
(业务键, 资源分组, 架构层, 服务键, 聚合签名)
```

- 聚合先按业务、架构层、服务键和资源分组隔离，再应用不同服务类型的“聚合签名”。因此相同名称但业务、层、服务类型或资源分组不同，永远不会合并。
- 普通资源（包括 ECS、BMS、ELB/SLB、WAF、RDS、DCS/Redis、OBS/OSS、NAT、VPC 等）的签名是 `resource_name` 的非数字结构：连续数字统一替换为 `#`，其他字符必须完全一致。`ecs-web-01` 与 `ecs-web-02` 会合并；`ecs-web-01` 与 `ecs-db-01` 不会合并；`elb-api-1` 与 `elb-api-2` 的处理方式与 ECS 相同。
- CCE_Deployment（包括 `CCE_Deployment`、`cce_deploym` 等类型别名，或资源分组明确写成 `CCE_Deployment` 的行）的签名固定为 `__cce_resource_group__`。同一业务、同一架构层、同一服务键、同一资源分组下，`deploy-api-001` 和 `deploy-worker-999` 也会合并为一个 `...+N` 聚合节点，名称结构不会再拆分该组。
- 普通 CCE 集群必须与 Deployment 分开理解：类型为 `cce` 且不是 Deployment 别名时，按普通资源的名称签名聚合。例如同组的 `cce-cluster-api-01/02` 可合并；名称非数字结构不同的两个集群会得到两个聚合节点。普通 CCE 不创建 Deployment 专用的资源分组框。
- CCE 与普通资源的展示层级也不同：CCE_Deployment 会创建“服务框 → 资源分组框 → `...+N` 聚合节点”；ECS、ELB 等通常是“服务框 → `...+N` 聚合节点”。如果普通服务框内存在多个资源分组值，系统只为布局隔离创建分组小框，不改变普通资源按名称签名拆分的规则。
- 一个服务框中混有普通 CCE 和 CCE_Deployment 时，代码按行独立判定：Deployment 仍进入对应资源分组框，普通 CCE 仍挂在 CCE 服务框下，不会因为混合数据而关闭 Deployment 分组框。
- 企业项目名称/ID和 CCE 名称前缀不是当前聚合元组的一部分。它们只用于“未填写所属业务的 CCE_Deployment”回退业务匹配（相同企业项目 + 相同前缀），并保留在明细详情中；一旦业务归属确定，同一业务内不同企业项目不会自动拆成不同聚合节点，除非资源分组、层、服务键或名称签名不同。
- 所有受支持服务类型都会生成 `__agg_compute__` 聚合节点和数字徽标；原始资源完整保存在 `summary_resources` 中。`default/Other` 记录会在聚合前过滤，不会产生 Other 聚合。
- 空资源分组统一显示为“(未设置资源分组)”。不同业务、架构层或服务类型之间不会合并同名分组。

#### 按类型对照

| 资源类型 | 聚合签名 | 资源分组框 | 示例结果 |
| --- | --- | --- | --- |
| `CCE_Deployment` / Deployment 别名 | 固定值 `__cce_resource_group__`，不看名称结构 | 保留，按资源分组一组一个框 | 同组 20 个不同名称 → 一个 `...+20` |
| 普通 `CCE` 集群 | `resource_name` 替换数字后的签名 | 不使用 Deployment 专用框 | `cce-api-01/02` → `...+2`；`cce-api-*` 与 `cce-worker-*` → 两组 |
| `ECS`、`ELB/SLB`、`RDS` 等普通服务 | 同上，按非数字名称结构 | 默认不保留 CCE 式分组框；多分组时可生成布局隔离小框 | 同组同结构合并，不同结构拆分 |

聚合节点承接调用关系：同业务调用落到对应聚合节点；跨业务调用落到源/目标服务组的代表聚合节点；多条相同端点合并并递增 `call_count`。

### 4.5 调用边归并

- “下游服务”解析为确认边；推断目标只写入详情，不画边。
- 同业务边连接到资源对应的聚合节点；跨业务边连接到源/目标服务组聚合节点。
- 多条相同源、目标、确认/推断状态的边合并为一条并增加 `call_count`，标签显示 `×N`。
- 确认边按关系着色；跨业务确认边为橙色；推断边为紫色虚线（当前 Excel 推断字段本身不创建边，保留该样式以兼容动态编辑/历史数据）。
- 被核心业务过滤掉的资源不会生成外部节点或边；无法解析的外部目标仅在能识别为 K8s/MaaS 等服务时创建外部聚合节点。

## 5. BDAT 布局与前端功能

### 5.1 默认布局效果

`compute_bdat_positions()` 按业务将图排成最多 4 列网格；业务内部按四层、服务类型和依赖关系排列。服务级依赖会形成纵向流水线；无依赖的服务按紧凑网格排布。每个资源分组块每行最多 7 个节点，数据层服务最多分 3 行；VPC/NAT/CC/DCAAS 放到右侧独立轨道，避免挤压主架构层。容器位置由子节点中心自底向上补算。

### 5.2 工具栏操作

| 操作 | 效果 |
| --- | --- |
| 自适应 | 按当前可见元素缩放并居中。 |
| 复位 | 回到当前布局的初始视图。 |
| 导出图片 | 导出 2× 分辨率 PNG。 |
| 业务分组：开/关 | 显示或隐藏业务、层、服务、资源分组容器；关闭时临时解除父子关系，便于自由查看节点。 |
| BDAT 架构：开/关 | 仅隐藏/恢复四层架构框和固定层级边，业务框、服务框及资源组框仍保留。 |
| 锁定容器：开/关 | 开启后业务框可整体拖动，内部层/服务框不可单独拖动；资源节点仍可调整。 |
| 保存布局 | 将节点坐标、缩放和视角写入浏览器 `localStorage`；刷新同一 HTML 自动恢复。未点击保存时，拖拽不会覆盖生成的默认布局。 |
| 新增资源 | 打开编辑器，创建资源节点、自动创建业务/层/服务容器并重建其确认下游边。仅支持非 Other 服务类型。 |
| 导出编辑表 | 导出 `topology-edited.csv` 和带自适应列宽的 Excel 兼容 `topology-edited.xls`；支持安全上下文下选择目录，否则走浏览器下载。 |
| 搜索 | 按资源名、类型、ID、资源分组和业务过滤；匹配节点及其容器会保留显示。 |

### 5.3 点击行为

- 点击资源节点：右侧显示区域、类型、ID、项目、业务、资源组、规格、描述、推断原因等详情，并高亮完整调用链，其余元素淡化。
- 点击 `...+N` 摘要节点或数字徽标：打开右侧抽屉，按每页 10 条展示该聚合组全部明细；每行可编辑。
- 在抽屉中编辑：更新内存中的聚合明细和摘要统计，立即刷新详情；通过“导出编辑表”落盘。
- 点击连线：显示确认/推断状态、源节点、目标节点和调用次数。
- 点击空白：清除高亮；滚轮缩放，拖拽平移。

## 6. 代码地图与维护入口

- `ourresult/converter.py`：唯一主入口；包括 Excel 读取、分类、聚合、边构建、BDAT 坐标、样式和 HTML 模板。
- `ourresult/test_converter.py`：`unittest` 测试，覆盖聚合、CCE 回退、四层布局、核心业务过滤、外部 K8s/MaaS、图标和 HTML 生成等。
- `ourresult/create_sample.py`：生成约千级节点的示例 Excel，适合手工回归。
- `ourresult/README.md`：面向使用者的说明；若改动列名、聚合语义或按钮行为，应同步更新。
- `.agents/skills/gen-topology/SKILL.md`：skill 调用约定和错误处理提示。
- `cloudmapper-main/web/js/`：前端库；不要随意替换 `cytoscape.min.js`，否则 cose-bilkent/dagre 可能因版本不兼容导致空白图。
- `ourresult/picture/`：PNG 图标目录；新增服务类型时通常需要同时更新 `SERVICE_STYLE`、`SERVICE_DISPLAY_NAMES`、`SERVICE_ICON_FILES`/别名及图例。

推荐修改顺序：先补/改 Python 纯函数和单元测试，再更新 HTML 交互模板，最后用示例清单做浏览器回归。

## 7. 验证与已知边界

当前环境没有可用的 Python 解释器：`python` 不在 PATH，`py` 未发现安装，`uv` 也无法访问受管控的解释器目录；因此本次交接未能现场执行测试或重新生成 HTML。仓库已有 `ourresult/benchmark_100k_report.md` 记录：100,000 条 RDS 资源、10,000 条确认边可完成图构建、BDAT 布局和自包含 HTML 生成，总耗时约 30.1 秒；该记录不等同于浏览器一次性流畅渲染 10 万节点的承诺。

接手后建议首先执行：

```powershell
py -3 -m unittest discover -s ourresult -p "test_*.py" -v
py -3 ourresult\converter.py ourresult\sample.xlsx ourresult\handover-smoke.html --title "交接验证拓扑"
```

并在浏览器检查：默认 BDAT 布局、CCE_Deployment 分组框、摘要抽屉、节点编辑、布局保存/刷新恢复、业务/BDAT 开关、PNG 与编辑表导出。

## 8. 交接检查清单

- [ ] Python 3.7+ 与 `openpyxl` 可用。
- [ ] `cloudmapper-main/web/js/` 中 Cytoscape 与布局插件文件齐全且版本匹配。
- [ ] `ourresult/picture/` 图标目录存在。
- [ ] 输入 Excel 表头能被 `COL_ALIASES` 识别，资源名称无重复或空值。
- [ ] CCE_Deployment 的“资源分组”、企业项目和名称前缀符合预期。
- [ ] 核心业务列是否存在已确认；该列一旦存在，非“是”行会被完全过滤。
- [ ] 下游服务只填写已确认关系；推断信息放入“下游服务(推断)”和“推断原因”。
- [ ] HTML 能直接打开，且保存布局/导出文件的浏览器权限符合使用场景。
