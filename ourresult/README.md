# 云服务拓扑图生成器

从 Excel 清单一键生成可交互的云服务拓扑图（HTML 单文件，无需服务器）。

底层渲染引擎使用 [Cytoscape.js](https://js.cytoscape.org/)，布局算法参考 [CloudMapper](https://github.com/duo-labs/cloudmapper) 的工程实践。

---

## 快速开始

```bash
# 1. 安装依赖
pip install openpyxl

# 2. 生成示例 Excel（可选）
py create_sample.py

# 3. 转换为拓扑图
py converter.py sample.xlsx

# 4. 用浏览器打开生成的 sample.html
```

---

## Excel 格式说明

支持的列名（大小写不敏感，允许中英混用）：

| 列名 | 必填 | 说明 | 示例 |
|------|------|------|------|
| `resource name` | **是** | 资源唯一名称（作为节点标签和边的引用依据） | `ecs-web-01` |
| `resource_type` | **是** | 服务类型（见下方支持列表） | `ECS` |
| `resource id` | 否 | 资源实例 ID | `ecs-11111111` |
| `enterprise_project_id` | 否 | 企业项目 ID / MAC / 其他标识 | `ep-prod-001` |
| `region` | 否 | 资源所在区域；相同 region 的资源会被分组到同一容器内 | `cn-north-4` |
| `资源分组` | 否 | 功能分组标签；相同分组的资源在 region 内聚合显示 | `K8s工作节点` |
| `下游服务` | 否 | **已确认**的下游服务，逗号分隔，值为 `resource name` | `rds-prod,redis-01` |
| `下游服务(推断)` | 否 | **推断**的下游服务，逗号分隔；渲染为虚线箭头 | `apig-01` |
| `推断原因` | 否 | 推断依据的详细描述；点击节点时在侧边栏展示 | 配置文件中 DB_HOST 指向… |

> **兼容性**：列名支持多种别名，例如 `resource name` 也可写成 `资源名称`、`name`；`resource_type` 可写成 `服务类型`、`type`。完整别名见 `converter.py` 中的 `COL_ALIASES`。

---

## 支持的服务类型

| 类型值 | 云服务 | 颜色 |
|--------|--------|------|
| `ECS` | 弹性云服务器 | 橙色 |
| `ELB` / `SLB` | 弹性负载均衡 | 青色 |
| `EIP` | 弹性公网 IP | 绿色 |
| `WAF` | Web 应用防火墙 | 红色菱形 |
| `NAT` | NAT 网关 | 灰色 |
| `VPN` | VPN 网关 | 灰色 |
| `CDN` | 内容分发网络 | 蓝色 |
| `RDS` | 云数据库 | 蓝色 |
| `DCS` / `Redis` | 分布式缓存 | 红色圆形 |
| `DMS` / `Kafka` | 分布式消息 | 黄色 |
| `OBS` / `OSS` | 对象存储 | 紫色 |
| `CCE` | 云容器引擎 | 蓝色 |
| `APIG` | API 网关 | 青绿色 |
| `CSS` | 云搜索服务 | 黄色 |
| `VPC` | 虚拟私有云 | 浅蓝容器 |

> 不在列表中的类型将显示为默认灰蓝色方块。

---

## 拓扑图层次结构

```
Region（区域容器，蓝色虚线框）
  └── 资源分组（功能容器，灰色虚线框）
        └── 资源节点（叶节点）
```

- 未指定 `region` 的节点不属于任何区域容器
- 未指定 `资源分组` 的节点直接属于 region（或独立存在）

---

## 连线类型

| 类型 | 来源列 | 渲染样式 |
|------|--------|----------|
| 确认连接 | `下游服务` | 实线箭头，彩色 |
| 推断连接 | `下游服务(推断)` | 虚线箭头，灰色 |

---

## 交互功能

| 操作 | 说明 |
|------|------|
| 点击节点 | 侧边栏显示资源名称、类型、ID、区域、分组、推断原因；高亮相邻节点 |
| 点击连线 | 显示连接类型（确认/推断）及源/目标节点 |
| 点击空白 | 取消高亮 |
| 滚轮 | 缩放 |
| 拖拽 | 平移 |
| 搜索框 | 按名称/类型/ID/分组过滤节点 |
| 布局切换 | cose（默认内置）/ cose-bilkent（可选插件）/ 层次 / 网格 / 圆形 |
| 导出图片 | 导出为 PNG（2× 分辨率） |

---

## 命令行参数

```
py converter.py <excel文件> [输出HTML] [--title 标题]

参数:
  excel       输入 Excel 文件路径
  output      输出 HTML 路径（可选，默认同名 .html）
  --title     页面标题（默认：云服务拓扑图）

示例:
  py converter.py services.xlsx
  py converter.py services.xlsx topology.html
  py converter.py services.xlsx output.html --title "生产环境拓扑"
```

---

## 文件结构

```
ourresult/
├── converter.py        # 主程序：Excel → 独立 HTML
├── create_sample.py    # 生成示例 Excel
├── README.md           # 本文档
├── sample.xlsx         # 示例数据（运行 create_sample.py 生成）
└── sample.html         # 示例拓扑图（运行 converter.py 生成）
```

引擎文件（由 `converter.py` 自动读取，无需手动操作）：
```
../cloudmapper-main/web/js/cytoscape.min.js           # 核心渲染引擎（优先，与插件版本兼容）
../cloudmapper-main/web/js/cytoscape-cose-bilkent.js  # cose-bilkent 布局算法
../cloudmapper-main/web/js/dagre.js                   # 层次布局依赖
../cloudmapper-main/web/js/cytoscape-dagre.js         # Dagre 布局
../cytoscape.js-unstable/dist/cytoscape.min.js        # 备用渲染引擎
```

> **注意**：`cloudmapper-main/web/js/cytoscape.min.js` 与同目录的插件（cose-bilkent、dagre）版本匹配。
> 若擅自替换为其他版本的 cytoscape.min.js，可能导致拓扑图节点不显示（布局插件不兼容）。

---

## 常见问题

### 网页打开后没有节点显示

**原因**：cytoscape.js 与布局插件（cose-bilkent）版本不兼容，布局计算失败导致节点无坐标。

**已内置的解决方案**：`converter.py` 已按以下策略处理：
1. 优先加载与插件版本匹配的 `cloudmapper-main/web/js/cytoscape.min.js`
2. 生成的 HTML 中布局调用带有 `try/catch`，若 `cose-bilkent` 失败会自动回退到内置 `cose` 布局

**如果仍然不显示**：切换布局选择框中的其他选项（如"层次布局"或"网格布局"），确认节点是否存在。

---

## 依赖

- Python 3.7+
- `openpyxl`（`pip install openpyxl`）
- 无需 Node.js 或其他前端工具链
