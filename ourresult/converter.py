#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
云服务拓扑图生成器（支持华为云/阿里云 Excel 格式）
用法: python converter.py <excel文件> [输出HTML]
示例: python converter.py sample.xlsx topology.html
"""

import json
import sys
import os
import re
import argparse

try:
    import openpyxl
except ImportError:
    sys.exit("请先安装 openpyxl: pip install openpyxl")

# ── 服务类型样式 (背景色, 边框色, 形状) ─────────────────────────────────────

SERVICE_STYLE = {
    # 计算
    "ecs":          ("#FF6B35", "#C0441F", "roundrectangle"),
    "bms":          ("#FF8C42", "#BB5A1A", "roundrectangle"),
    "cce":          ("#4A90D9", "#2C5F8A", "roundrectangle"),
    "cci":          ("#5BA3E0", "#2C5F8A", "ellipse"),
    "functiongraph":("#E67E22", "#9A4A00", "roundrectangle"),
    # 网络
    "vpc":          ("#F0F3FA", "#3B6FB6", "roundrectangle"),
    "eip":          ("#27AE60", "#1A6B3A", "ellipse"),
    "elb":          ("#16A085", "#0E6655", "roundrectangle"),
    "slb":          ("#16A085", "#0E6655", "roundrectangle"),
    "nat":          ("#5D6D7E", "#2E4057", "roundrectangle"),
    "vpn":          ("#7F8C8D", "#424949", "roundrectangle"),
    "cdn":          ("#2980B9", "#1A5276", "ellipse"),
    "waf":          ("#C0392B", "#7B241C", "diamond"),
    "er":           ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "dli":          ("#2C3E50", "#1A252F", "roundrectangle"),
    # 存储
    "oss":          ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "obs":          ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "sfs":          ("#9B59B6", "#6C3483", "roundrectangle"),
    "evs":          ("#A569BD", "#7D3C98", "roundrectangle"),
    "cbr":          ("#7FB3D3", "#2C7BB6", "roundrectangle"),
    # 数据库
    "rds":          ("#3498DB", "#1A5276", "roundrectangle"),
    "dcs":          ("#E74C3C", "#922B21", "ellipse"),
    "redis":        ("#E74C3C", "#922B21", "ellipse"),
    "dds":          ("#27AE60", "#1A7A43", "roundrectangle"),
    "mongodb":      ("#27AE60", "#1A7A43", "roundrectangle"),
    "gaussdb":      ("#3498DB", "#1A5276", "roundrectangle"),
    "dws":          ("#1F618D", "#154360", "roundrectangle"),
    "css":          ("#F1C40F", "#9A7D0A", "roundrectangle"),
    "elasticsearch":("#F1C40F", "#9A7D0A", "roundrectangle"),
    # 消息 / 中间件
    "kafka":        ("#F39C12", "#9A6007", "roundrectangle"),
    "mq":           ("#F39C12", "#9A6007", "roundrectangle"),
    "dms":          ("#F39C12", "#9A6007", "roundrectangle"),
    "smn":          ("#D35400", "#7E5109", "roundrectangle"),
    # 应用 / API
    "apigateway":   ("#1ABC9C", "#0E8A72", "roundrectangle"),
    "apig":         ("#1ABC9C", "#0E8A72", "roundrectangle"),
    # 容器镜像
    "swr":          ("#3D9BDC", "#1D6A9E", "roundrectangle"),
    "ecr":          ("#3D9BDC", "#1D6A9E", "roundrectangle"),
    # 安全
    "iam":          ("#E67E22", "#9A4A00", "pentagon"),
    "kms":          ("#F39C12", "#9A6007", "pentagon"),
    # 默认
    "default":      ("#AED6F1", "#5D8AA8", "roundrectangle"),
}

# 这些类型作为 compound 容器节点（自动生成，不直接来自行数据）
_VIRTUAL_CONTAINER_TYPES = {"__region__", "__group__"}

# 连接关系 → 确认边颜色
RELATION_COLORS = {
    "调用":     "#3498DB",
    "访问":     "#27AE60",
    "依赖":     "#E67E22",
    "发布":     "#9B59B6",
    "消费":     "#E74C3C",
    "转发":     "#1ABC9C",
    "路由":     "#F39C12",
    "分发流量": "#2980B9",
    "主从同步": "#7F8C8D",
    "default":  "#2196F3",
}

# ── Excel 列名别名（兼容中英文、大小写、有无空格）────────────────────────────

COL_ALIASES = {
    "服务名称":     ["resource name", "resource_name", "资源名称", "服务名称",
                     "name", "名称", "服务"],
    "服务类型":     ["resource_type", "resourcetype", "服务类型", "type",
                     "类型", "service_type"],
    "资源ID":       ["resource id", "resource_id", "resourceid", "资源id",
                     "资源ID", "id", "服务id", "instance_id"],
    "企业项目ID":   ["enterprise_project_id", "enterprise_project",
                     "企业项目id", "企业项目ID", "project_id"],
    "区域":         ["region", "区域", "地域", "availability_zone", "az"],
    "资源分组":     ["资源分组", "分组", "resource_group", "group",
                     "resource group", "标签分组"],
    "描述":         ["描述", "description", "备注", "说明", "desc", "remark"],
    "规格":         ["规格", "spec", "specification", "配置", "instance_type",
                     "flavor"],
    "下游服务":     ["下游服务", "downstream", "downstream_service",
                     "连接目标", "target", "targets", "依赖服务",
                     "下游服务(逗号分隔)", "下游服务（逗号分隔）"],
    "下游服务推断": ["下游服务(推断)", "下游服务（推断）", "inferred_downstream",
                     "推断下游", "inferred downstream"],
    "推断原因":     ["推断原因", "inference_reason", "reason", "推断依据",
                     "infer_reason"],
}


def normalize(s):
    return re.sub(r"[\s\(\（\)\）_-]", "", str(s)).lower()


def detect_columns(header_row):
    header_map = {}
    for i, cell in enumerate(header_row, 1):
        if cell.value is not None:
            header_map[normalize(cell.value)] = i
    result = {}
    for canonical, aliases in COL_ALIASES.items():
        for alias in aliases:
            if normalize(alias) in header_map:
                result[canonical] = header_map[normalize(alias)]
                break
    return result


def read_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    rows_iter = ws.iter_rows()
    header_row = next(rows_iter, None)
    if header_row is None:
        sys.exit("Excel 文件为空")

    col_map = detect_columns(header_row)

    if "服务名称" not in col_map:
        print(f"[警告] 未找到 '服务名称/resource name' 列，已检测到的列: "
              f"{[c.value for c in header_row if c.value]}")
        sys.exit("请检查 Excel 表头是否匹配（见 README.md）")

    records = []
    for row in rows_iter:
        def get(field, default=""):
            idx = col_map.get(field)
            if idx is None:
                return default
            v = row[idx - 1].value
            return str(v).strip() if v is not None else default

        name = get("服务名称")
        if not name or name.lower() == "none":
            continue

        records.append({
            "name":            name,
            "type":            get("服务类型", "default").lower().strip(),
            "resource_id":     get("资源ID"),
            "enterprise_id":   get("企业项目ID"),
            "region":          get("区域"),
            "group":           get("资源分组"),
            "desc":            get("描述"),
            "spec":            get("规格"),
            "targets":         get("下游服务"),
            "inferred_targets":get("下游服务推断"),
            "reason":          get("推断原因"),
        })

    return records


# ── 图数据构建 ────────────────────────────────────────────────────────────────

def build_graph(records):
    """构建 cytoscape elements（nodes + edges）"""

    def safe_id(s):
        return re.sub(r"[^a-zA-Z0-9_\-]", "_", str(s))

    nodes = []
    edges = []
    edge_set = set()
    node_id_map = {}   # resource_name → node_id
    edge_count = [0]

    # ── 1. 收集所有 region / group 组合，生成虚拟容器节点 ──────────────────

    regions = {}   # region_label → region_node_id
    groups  = {}   # (region_label, group_label) → group_node_id

    for r in records:
        reg = (r["region"] or "").strip()
        grp = (r["group"]  or "").strip()
        if reg and reg not in regions:
            rid = f"__region__{safe_id(reg)}"
            regions[reg] = rid
            nodes.append({"group": "nodes", "data": {
                "id": rid, "name": reg,
                "type": "__region__", "is_container": 1,
                "bg_color": "#EBF5FB", "border_color": "#2E86C1",
                "shape": "roundrectangle",
                "spec": "", "desc": "区域",
                "resource_id": "", "enterprise_id": "",
                "region": "", "group_label": "",
            }})
        if grp:
            key = (reg, grp)
            if key not in groups:
                gid = f"__group__{safe_id(reg)}_{safe_id(grp)}"
                groups[key] = gid
                g_data = {
                    "id": gid, "name": grp,
                    "type": "__group__", "is_container": 1,
                    "bg_color": "#F9F9F9", "border_color": "#AAB7B8",
                    "shape": "roundrectangle",
                    "spec": "", "desc": "资源分组",
                    "resource_id": "", "enterprise_id": "",
                    "region": reg, "group_label": grp,
                }
                if reg in regions:
                    g_data["parent"] = regions[reg]
                nodes.append({"group": "nodes", "data": g_data})

    # ── 2. 创建资源叶节点 ──────────────────────────────────────────────────

    for i, r in enumerate(records):
        nid = f"n_{safe_id(r['name'])}_{i}"
        # 同名资源只保留首个 id（边引用时查名字）
        if r["name"] not in node_id_map:
            node_id_map[r["name"]] = nid

        stype = r["type"] or "default"
        bg, border, shape = SERVICE_STYLE.get(stype, SERVICE_STYLE["default"])

        reg = (r["region"] or "").strip()
        grp = (r["group"]  or "").strip()

        d = {
            "id":            nid,
            "name":          r["name"],
            "type":          stype,
            "spec":          r["spec"],
            "desc":          r["desc"],
            "resource_id":   r["resource_id"],
            "enterprise_id": r["enterprise_id"],
            "region":        reg,
            "group_label":   grp,
            "reason":        r["reason"],
            "bg_color":      bg,
            "border_color":  border,
            "shape":         shape,
            "is_container":  0,
        }

        # 设置父节点：优先 group，其次 region
        if grp and (reg, grp) in groups:
            d["parent"] = groups[(reg, grp)]
        elif reg in regions:
            d["parent"] = regions[reg]

        nodes.append({"group": "nodes", "data": d})

    # ── 3. 解析外部节点（targets 中不存在的名称）─────────────────────────

    all_node_ids = {n["data"]["id"] for n in nodes}

    def ensure_ext_node(name):
        """目标名不在记录中时，创建外部虚拟节点"""
        if name in node_id_map:
            return node_id_map[name]
        ext_id = f"n_ext_{safe_id(name)}"
        if ext_id not in all_node_ids:
            bg, border, shape = SERVICE_STYLE["default"]
            nodes.append({"group": "nodes", "data": {
                "id": ext_id, "name": name,
                "type": "default", "is_container": False,
                "bg_color": bg, "border_color": border, "shape": shape,
                "spec": "", "desc": "（外部/未列出服务）",
                "resource_id": "", "enterprise_id": "",
                "region": "", "group_label": "", "reason": "",
            }})
            all_node_ids.add(ext_id)
        node_id_map[name] = ext_id
        return ext_id

    # ── 4. 创建边 ─────────────────────────────────────────────────────────

    def add_edge(src_id, tgt_name, relation, inferred=False):
        tgt_id = ensure_ext_node(tgt_name)
        key = (src_id, tgt_id, inferred)
        if key in edge_set:
            return
        edge_set.add(key)
        color = RELATION_COLORS.get(relation, RELATION_COLORS["default"])
        if inferred:
            color = "#AAAAAA"
        eid = f"e_{edge_count[0]}"
        edge_count[0] += 1
        edges.append({"group": "edges", "data": {
            "id":       eid,
            "source":   src_id,
            "target":   tgt_id,
            "relation": relation + ("（推断）" if inferred else ""),
            "color":    color,
            "inferred": 1 if inferred else 0,
        }})

    for r in records:
        src_id = node_id_map.get(r["name"])
        if not src_id:
            continue

        # 确认下游
        if r["targets"]:
            for t in [x.strip() for x in r["targets"].split(",") if x.strip()]:
                add_edge(src_id, t, "调用", inferred=False)

        # 推断下游
        if r["inferred_targets"]:
            for t in [x.strip() for x in r["inferred_targets"].split(",") if x.strip()]:
                add_edge(src_id, t, "调用", inferred=True)

    return nodes + edges


# ── HTML 生成辅助 ─────────────────────────────────────────────────────────────

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_CYTOSCAPE_PATHS = [
    os.path.join(_SCRIPT_DIR, "..", "cytoscape.js-unstable", "dist", "cytoscape.min.js"),
    os.path.join(_SCRIPT_DIR, "js", "cytoscape.min.js"),
]
_COSE_BILKENT_PATHS = [
    os.path.join(_SCRIPT_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape-cose-bilkent.js"),
    os.path.join(_SCRIPT_DIR, "js", "cytoscape-cose-bilkent.js"),
]
_DAGRE_PATHS = [
    os.path.join(_SCRIPT_DIR, "..", "cloudmapper-main", "web", "js", "dagre.js"),
    os.path.join(_SCRIPT_DIR, "js", "dagre.js"),
]
_CYTOSCAPE_DAGRE_PATHS = [
    os.path.join(_SCRIPT_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape-dagre.js"),
    os.path.join(_SCRIPT_DIR, "js", "cytoscape-dagre.js"),
]


def _read_first(paths):
    for p in paths:
        try:
            with open(p, "r", encoding="utf-8") as f:
                return f.read()
        except FileNotFoundError:
            continue
    return ""


def _make_cytoscape_style():
    return [
        {
            "selector": "node",
            "style": {
                "label": "data(name)",
                "font-size": 11,
                "text-valign": "bottom",
                "text-halign": "center",
                "text-margin-y": 4,
                "width": 56,
                "height": 56,
                "background-color": "data(bg_color)",
                "border-color": "data(border_color)",
                "border-width": 2,
                "shape": "data(shape)",
                "color": "#2C3E50",
                "text-background-color": "#fff",
                "text-background-opacity": 0.8,
                "text-background-padding": "2px",
                "text-max-width": "100px",
                "text-wrap": "ellipsis",
            }
        },
        {
            "selector": "node[is_container = 1]",
            "style": {
                "label": "data(name)",
                "text-valign": "top",
                "font-size": 13,
                "font-weight": "bold",
                "background-opacity": 0.06,
                "border-style": "dashed",
                "border-width": 2,
                "padding": "28px",
                "color": "#1A5276",
                "text-background-color": "#fff",
                "text-background-opacity": 0.9,
            }
        },
        # region 容器加重边框
        {
            "selector": "node[type = '__region__']",
            "style": {
                "border-color": "#2E86C1",
                "border-width": 3,
                "background-color": "#EBF5FB",
                "background-opacity": 0.12,
                "font-size": 15,
                "font-weight": "bold",
            }
        },
        # 资源分组容器
        {
            "selector": "node[type = '__group__']",
            "style": {
                "border-color": "#AAB7B8",
                "border-width": 1,
                "background-color": "#FDFEFE",
                "background-opacity": 0.08,
                "font-size": 12,
            }
        },
        # 确认边
        {
            "selector": "edge[inferred = 0]",
            "style": {
                "label": "data(relation)",
                "width": 2,
                "line-color": "data(color)",
                "target-arrow-color": "data(color)",
                "target-arrow-shape": "triangle",
                "curve-style": "bezier",
                "font-size": 10,
                "color": "#555",
                "text-background-color": "#fff",
                "text-background-opacity": 0.8,
                "text-background-padding": "2px",
                "opacity": 0.9,
            }
        },
        # 推断边（虚线）
        {
            "selector": "edge[inferred = 1]",
            "style": {
                "label": "data(relation)",
                "width": 1.5,
                "line-style": "dashed",
                "line-dash-pattern": [6, 4],
                "line-color": "#AAAAAA",
                "target-arrow-color": "#AAAAAA",
                "target-arrow-shape": "triangle",
                "curve-style": "bezier",
                "font-size": 9,
                "color": "#888",
                "text-background-color": "#fff",
                "text-background-opacity": 0.7,
                "text-background-padding": "2px",
                "opacity": 0.65,
            }
        },
        {"selector": ":selected", "style": {
            "border-color": "#FFD700", "border-width": 4,
            "background-color": "#FFF9C4",
        }},
        {"selector": ".faded",       "style": {"opacity": 0.15}},
        {"selector": ".highlighted", "style": {
            "border-color": "#E74C3C", "border-width": 3,
        }},
    ]


def _make_legend_html():
    exclude = {"__region__", "__group__"}
    items = []
    for stype, (bg, border, shape) in SERVICE_STYLE.items():
        if stype in exclude:
            continue
        radius = "50%" if shape == "ellipse" else ("50% 0" if shape == "diamond" else "4px")
        items.append(
            f'<div class="legend-item">'
            f'<span class="legend-dot" style="background:{bg};border-color:{border};'
            f'border-radius:{radius}"></span>'
            f'<span>{stype.upper()}</span>'
            f'</div>'
        )
    return "\n".join(items)


def generate_html(elements, title="云服务拓扑图", output_path="topology.html"):
    cytoscape_js    = _read_first(_CYTOSCAPE_PATHS)
    cose_bilkent    = _read_first(_COSE_BILKENT_PATHS)
    dagre_js        = _read_first(_DAGRE_PATHS)
    cytoscape_dagre = _read_first(_CYTOSCAPE_DAGRE_PATHS)

    if not cytoscape_js:
        sys.exit("找不到 cytoscape.min.js，请确认路径:\n" + "\n".join(_CYTOSCAPE_PATHS))

    elements_json = json.dumps(elements, ensure_ascii=False, indent=2)
    style_json    = json.dumps(_make_cytoscape_style(), ensure_ascii=False)
    legend_html   = _make_legend_html()
    has_cose      = bool(cose_bilkent)
    has_dagre     = bool(dagre_js and cytoscape_dagre)
    default_layout = "cose-bilkent" if has_cose else "cose"

    # 各布局 option 标签
    layout_options = f'<option value="{default_layout}">{default_layout} 布局（推荐）</option>\n'
    layout_options += '    <option value="breadthfirst">层次布局</option>\n'
    layout_options += '    <option value="grid">网格布局</option>\n'
    layout_options += '    <option value="circle">圆形布局</option>\n'
    if has_dagre:
        layout_options += '    <option value="dagre">Dagre 布局</option>\n'

    extra_scripts = ""
    if has_cose:
        extra_scripts += f"<script>{cose_bilkent}</script>\n"
    if has_dagre:
        extra_scripts += f"<script>{dagre_js}</script>\n"
        extra_scripts += f"<script>{cytoscape_dagre}</script>\n"

    # 统计数字
    n_nodes = sum(1 for e in elements if e.get("group") == "nodes"
                  and e["data"].get("type") not in ("__region__", "__group__"))
    n_edges = sum(1 for e in elements if e.get("group") == "edges")
    n_infer = sum(1 for e in elements if e.get("group") == "edges"
                  and e["data"].get("inferred") == 1)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>{title}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f4f6fb;
     height:100vh;overflow:hidden;display:flex;flex-direction:column}}
/* ── 工具栏 ── */
#toolbar{{display:flex;align-items:center;gap:8px;padding:6px 14px;
  background:linear-gradient(90deg,#1A237E 0%,#283593 100%);
  color:#fff;flex-shrink:0;height:50px;box-shadow:0 2px 8px rgba(0,0,0,.35)}}
#toolbar h1{{font-size:15px;font-weight:700;letter-spacing:1px;white-space:nowrap;margin-right:8px}}
.sep{{width:1px;height:26px;background:rgba(255,255,255,.25);margin:0 4px}}
.tb-btn{{cursor:pointer;padding:4px 10px;border-radius:4px;font-size:12px;
  background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
  color:#fff;transition:background .2s;white-space:nowrap}}
.tb-btn:hover{{background:rgba(255,255,255,.3)}}
#searchBox{{padding:4px 8px;border-radius:4px;border:1px solid rgba(255,255,255,.4);
  background:rgba(255,255,255,.15);color:#fff;font-size:12px;width:150px;outline:none}}
#searchBox::placeholder{{color:rgba(255,255,255,.55)}}
#layoutSelect{{padding:4px 8px;border-radius:4px;border:1px solid rgba(255,255,255,.3);
  background:rgba(255,255,255,.15);color:#fff;font-size:12px;cursor:pointer}}
#layoutSelect option{{background:#283593}}
.stats-bar{{font-size:11px;color:rgba(255,255,255,.75);margin-left:auto;
  display:flex;gap:10px;flex-shrink:0}}
.stat-item{{display:flex;align-items:center;gap:4px}}
.stat-dot{{width:8px;height:8px;border-radius:50%;display:inline-block}}
/* ── 主区域 ── */
#main{{display:flex;flex:1;overflow:hidden}}
#cy{{flex:1;background:#fff}}
/* ── 侧边栏 ── */
#sidebar{{width:280px;flex-shrink:0;background:#fff;border-left:1px solid #e0e0e0;
  display:flex;flex-direction:column;overflow:hidden}}
#sb-header{{padding:10px 14px;background:#E8EAF6;font-weight:700;font-size:13px;
  color:#283593;border-bottom:1px solid #c5cae9;display:flex;
  justify-content:space-between;align-items:center;flex-shrink:0}}
#detail-panel{{flex:1;overflow-y:auto;padding:10px 14px}}
.d-row{{margin-bottom:10px}}
.d-label{{font-size:10px;color:#999;text-transform:uppercase;letter-spacing:.6px;margin-bottom:2px}}
.d-value{{font-size:12px;color:#333;word-break:break-all;line-height:1.4}}
.tag{{display:inline-block;padding:2px 8px;border-radius:10px;
  font-size:11px;font-weight:700;color:#fff;margin-top:2px}}
.reason-box{{background:#FFF8E1;border:1px solid #FFE082;border-radius:4px;
  padding:6px 8px;font-size:11px;color:#5D4037;line-height:1.5;
  max-height:120px;overflow-y:auto;margin-top:4px;white-space:pre-wrap;word-break:break-all}}
/* ── 图例 ── */
#legend{{padding:8px 14px;border-top:1px solid #e0e0e0;background:#fafafa;
  max-height:180px;overflow-y:auto;flex-shrink:0}}
.legend-title{{font-size:10px;color:#999;text-transform:uppercase;
  letter-spacing:.6px;margin-bottom:6px}}
.legend-item{{display:flex;align-items:center;gap:6px;
  font-size:11px;color:#555;margin-bottom:3px}}
.legend-dot{{width:12px;height:12px;display:inline-block;border:2px solid;flex-shrink:0}}
.infer-legend{{display:flex;align-items:center;gap:6px;font-size:11px;
  color:#888;margin-bottom:3px}}
.infer-line{{width:24px;height:2px;border-top:2px dashed #AAAAAA}}
/* ── toast ── */
#toast{{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);
  background:#333;color:#fff;padding:8px 20px;border-radius:20px;
  font-size:12px;opacity:0;transition:opacity .3s;pointer-events:none;z-index:9999}}
</style>
</head>
<body>
<div id="toolbar">
  <h1>&#9729; {title}</h1>
  <div class="sep"></div>
  <button class="tb-btn" onclick="fitView()">&#9974; 自适应</button>
  <button class="tb-btn" onclick="resetView()">&#8635; 复位</button>
  <button class="tb-btn" onclick="exportPng()">&#128247; 导出图片</button>
  <div class="sep"></div>
  <select id="layoutSelect" onchange="changeLayout(this.value)">
    {layout_options}
  </select>
  <div class="sep"></div>
  <input id="searchBox" type="text" placeholder="&#128269; 搜索节点名/类型…"
         oninput="searchNodes(this.value)"/>
  <div class="stats-bar">
    <span class="stat-item"><span class="stat-dot" style="background:#3498DB"></span>{n_nodes} 节点</span>
    <span class="stat-item"><span class="stat-dot" style="background:#27AE60"></span>{n_edges - n_infer} 确认边</span>
    <span class="stat-item"><span class="stat-dot" style="background:#AAA"></span>{n_infer} 推断边</span>
  </div>
</div>

<div id="main">
  <div id="cy"></div>
  <div id="sidebar">
    <div id="sb-header">
      <span>节点 / 边 详情</span>
      <span id="node-type-tag"></span>
    </div>
    <div id="detail-panel">
      <p style="color:#bbb;font-size:12px;margin-top:30px;text-align:center;">
        点击节点或连线查看详情
      </p>
    </div>
    <div id="legend">
      <div class="legend-title">图例</div>
      {legend_html}
      <div class="infer-legend"><span class="infer-line"></span><span>推断连接（虚线）</span></div>
    </div>
  </div>
</div>
<div id="toast"></div>

<script>{cytoscape_js}</script>
{extra_scripts}
<script>
var elements   = {elements_json};
var styleRules = {style_json};

var cy = cytoscape({{
  container: document.getElementById('cy'),
  elements:  elements,
  style:     styleRules,
  layout: {{
    name: '{default_layout}',
    nodeDimensionsIncludeLabels: true,
    animate: false,
    padding: 50,
  }},
  wheelSensitivity: 0.2,
  minZoom: 0.05,
  maxZoom: 6,
}});

// ── 事件：点击节点 ─────────────────────────────────────────────────────────
cy.on('tap', 'node', function(evt) {{
  var d  = evt.target.data();
  var bg = d.bg_color || '#888';
  if (d.type === '__region__' || d.type === '__group__') {{
    showContainerDetail(d);
  }} else {{
    showNodeDetail(d, bg);
  }}
  cy.elements().addClass('faded');
  evt.target.removeClass('faded').addClass('highlighted');
  evt.target.neighborhood().removeClass('faded');
}});

// ── 事件：点击边 ──────────────────────────────────────────────────────────
cy.on('tap', 'edge', function(evt) {{
  var d = evt.target.data();
  var panel = document.getElementById('detail-panel');
  var inferred = d.inferred === 1;
  panel.innerHTML =
    '<div class="d-row"><div class="d-label">连接类型</div>' +
    '<div class="d-value">' + (inferred
      ? '<span style="color:#E74C3C">推断连接（虚线）</span>'
      : '<span style="color:#27AE60">已确认连接</span>') + '</div></div>' +
    '<div class="d-row"><div class="d-label">关系</div>' +
    '<div class="d-value">' + (d.relation || '-') + '</div></div>' +
    '<div class="d-row"><div class="d-label">来源</div>' +
    '<div class="d-value">' + (cy.getElementById(d.source).data('name') || d.source) + '</div></div>' +
    '<div class="d-row"><div class="d-label">目标</div>' +
    '<div class="d-value">' + (cy.getElementById(d.target).data('name') || d.target) + '</div></div>';
  document.getElementById('node-type-tag').innerHTML = inferred
    ? '<span class="tag" style="background:#AAA">推断</span>'
    : '<span class="tag" style="background:#27AE60">确认</span>';
  cy.elements().addClass('faded');
  evt.target.removeClass('faded').addClass('highlighted');
  evt.target.source().removeClass('faded');
  evt.target.target().removeClass('faded');
}});

// ── 点击空白 ──────────────────────────────────────────────────────────────
cy.on('tap', function(evt) {{
  if (evt.target === cy) {{
    cy.elements().removeClass('faded highlighted');
    document.getElementById('detail-panel').innerHTML =
      '<p style="color:#bbb;font-size:12px;margin-top:30px;text-align:center;">点击节点或连线查看详情</p>';
    document.getElementById('node-type-tag').innerHTML = '';
  }}
}});

function showNodeDetail(d, bg) {{
  var rows = [
    ['资源名称', d.name],
    ['资源类型', (d.type||'').toUpperCase()],
    ['资源ID',   d.resource_id  || '-'],
    ['企业项目', d.enterprise_id|| '-'],
    ['区域',     d.region       || '-'],
    ['资源分组', d.group_label  || '-'],
    ['规格',     d.spec         || '-'],
    ['描述',     d.desc         || '-'],
  ];
  var html = rows.map(function(r) {{
    return '<div class="d-row"><div class="d-label">'+r[0]+'</div>' +
           '<div class="d-value">'+escHtml(r[1])+'</div></div>';
  }}).join('');

  // 推断原因
  if (d.reason) {{
    html += '<div class="d-row"><div class="d-label">&#128270; 推断原因</div>' +
            '<div class="reason-box">'+escHtml(d.reason)+'</div></div>';
  }}

  // 相邻节点
  var src   = cy.getElementById(d.id);
  var nbrs  = src.neighborhood().nodes().filter(function(n) {{
    return n.data('type') !== '__region__' && n.data('type') !== '__group__';
  }});
  if (nbrs.length > 0) {{
    html += '<div class="d-row"><div class="d-label">相邻节点</div><div class="d-value">';
    nbrs.forEach(function(n) {{
      html += '<span style="display:inline-block;margin:2px 2px 0 0;padding:2px 6px;' +
              'background:#E8EAF6;border-radius:4px;font-size:11px;cursor:pointer" ' +
              'onclick="focusNode(\''+n.id()+'\')">'+escHtml(n.data('name'))+'</span>';
    }});
    html += '</div></div>';
  }}

  document.getElementById('detail-panel').innerHTML = html;
  document.getElementById('node-type-tag').innerHTML =
    '<span class="tag" style="background:'+bg+'">'+((d.type||'').toUpperCase())+'</span>';
}}

function showContainerDetail(d) {{
  var label = d.type === '__region__' ? '区域' : '资源分组';
  var html = '<div class="d-row"><div class="d-label">'+label+'</div>' +
             '<div class="d-value">'+escHtml(d.name)+'</div></div>';
  var children = cy.getElementById(d.id).descendants().filter(function(n) {{
    return !n.data('is_container');
  }});
  html += '<div class="d-row"><div class="d-label">包含服务 ('+children.length+')</div><div class="d-value">';
  children.forEach(function(n) {{
    html += '<span style="display:inline-block;margin:2px 2px 0 0;padding:2px 6px;' +
            'background:#E8EAF6;border-radius:4px;font-size:11px;cursor:pointer" ' +
            'onclick="focusNode(\''+n.id()+'\')">'+escHtml(n.data('name'))+'</span>';
  }});
  html += '</div></div>';
  document.getElementById('detail-panel').innerHTML = html;
  document.getElementById('node-type-tag').innerHTML =
    '<span class="tag" style="background:#2E86C1">'+label+'</span>';
}}

// ── 工具函数 ──────────────────────────────────────────────────────────────
function escHtml(s) {{
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

function fitView()   {{ cy.fit(undefined, 50); }}
function resetView() {{ cy.reset(); }}

function exportPng() {{
  var a = document.createElement('a');
  a.download = 'topology.png';
  a.href = cy.png({{ output:'base64uri', bg:'#ffffff', full:true, scale:2 }});
  a.click();
  showToast('图片已导出');
}}

function changeLayout(name) {{
  var opts = {{ name:name, nodeDimensionsIncludeLabels:true,
               animate:true, animationDuration:600, padding:50 }};
  if (name === 'cose' || name === 'cose-bilkent') opts.randomize = false;
  cy.layout(opts).run();
}}

function searchNodes(q) {{
  q = q.trim().toLowerCase();
  if (!q) {{ cy.elements().removeClass('faded highlighted'); return; }}
  cy.elements().addClass('faded');
  var matched = cy.nodes().filter(function(n) {{
    var d = n.data();
    return (d.name||'').toLowerCase().includes(q)
        || (d.type||'').toLowerCase().includes(q)
        || (d.resource_id||'').toLowerCase().includes(q)
        || (d.group_label||'').toLowerCase().includes(q);
  }});
  matched.removeClass('faded').addClass('highlighted');
  matched.neighborhood().removeClass('faded');
  if (matched.length > 0) {{
    cy.fit(matched, 80);
    showToast('找到 ' + matched.length + ' 个匹配节点');
  }} else {{
    showToast('未找到匹配节点');
  }}
}}

function focusNode(id) {{
  var n = cy.getElementById(id);
  if (!n.empty()) {{
    cy.elements().addClass('faded');
    n.removeClass('faded').addClass('highlighted');
    n.neighborhood().removeClass('faded');
    cy.animate({{ fit:{{ eles:n.closedNeighborhood(), padding:80 }}, duration:400 }});
    n.trigger('tap');
  }}
}}

function showToast(msg) {{
  var t = document.getElementById('toast');
  t.textContent = msg; t.style.opacity = 1;
  setTimeout(function(){{ t.style.opacity = 0; }}, 2200);
}}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(output_path) // 1024
    print(f"已生成: {output_path}  ({size_kb} KB)")
    print(f"  资源节点: {n_nodes}")
    print(f"  确认连线: {n_edges - n_infer}")
    print(f"  推断连线: {n_infer}")
    print("  用浏览器直接打开即可（无需启动服务器）。")


# ── 入口 ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="将云服务 Excel 清单转换为交互式拓扑图 HTML",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  py converter.py sample.xlsx
  py converter.py my_services.xlsx output.html
  py converter.py my_services.xlsx --title "生产环境拓扑"
"""
    )
    parser.add_argument("excel", help="输入 Excel 文件路径")
    parser.add_argument("output", nargs="?", default=None,
                        help="输出 HTML 文件路径（默认同名 .html）")
    parser.add_argument("--title", default="云服务拓扑图", help="页面标题")
    args = parser.parse_args()

    if not os.path.exists(args.excel):
        sys.exit(f"文件不存在: {args.excel}")

    output = args.output or (os.path.splitext(args.excel)[0] + ".html")

    print(f"读取 Excel: {args.excel}")
    records = read_excel(args.excel)
    print(f"  读取到 {len(records)} 条记录")

    elements = build_graph(records)
    generate_html(elements, title=args.title, output_path=output)


if __name__ == "__main__":
    main()
