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
from collections import defaultdict

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
    # 应用服务
    "nginx":        ("#27AE60", "#1A6B3A", "roundrectangle"),
    "k8s":          ("#326CE5", "#1A3A8F", "roundrectangle"),
    "kubernetes":   ("#326CE5", "#1A3A8F", "roundrectangle"),
    "tomcat":       ("#FF6B35", "#C0441F", "roundrectangle"),
    "mysql":        ("#3498DB", "#1A5276", "roundrectangle"),
    # 默认
    "default":      ("#AED6F1", "#5D8AA8", "roundrectangle"),
}

# 服务类型显示名称（用于子容器标签）
SERVICE_DISPLAY_NAMES = {
    "nginx": "Nginx", "k8s": "K8s", "kubernetes": "K8s",
    "tomcat": "Tomcat", "mysql": "MySQL",
    "ecs": "ECS", "bms": "BMS", "cce": "CCE", "cci": "CCI",
    "functiongraph": "FunctionGraph",
    "vpc": "VPC", "eip": "EIP", "elb": "ELB", "slb": "SLB",
    "nat": "NAT", "vpn": "VPN", "cdn": "CDN", "waf": "WAF",
    "er": "ER", "dli": "DLI",
    "oss": "OSS", "obs": "OBS", "sfs": "SFS", "evs": "EVS", "cbr": "CBR",
    "rds": "RDS", "dcs": "DCS", "redis": "Redis", "dds": "DDS",
    "mongodb": "MongoDB", "gaussdb": "GaussDB", "dws": "DWS",
    "css": "CSS", "elasticsearch": "Elasticsearch",
    "kafka": "Kafka", "mq": "MQ", "dms": "DMS", "smn": "SMN",
    "apigateway": "API Gateway", "apig": "APIG",
    "swr": "SWR", "ecr": "ECR",
    "iam": "IAM", "kms": "KMS",
    "default": "Other",
}

# 从节点名称中识别服务类型的关键词（顺序从精确到宽泛）
NAME_PATTERNS = [
    ("kubernetes", "k8s"), ("k8s", "k8s"),
    ("nginx",      "nginx"),
    ("tomcat",     "tomcat"),
    ("redis",      "redis"),
    ("kafka",      "kafka"),
    ("mysql",      "mysql"),
    ("mongodb",    "mongodb"),
    ("elasticsearch", "elasticsearch"),
    ("gaussdb",    "gaussdb"),
    ("dws",        "dws"),
    ("rds",        "rds"),
    ("dcs",        "dcs"),
    ("obs",        "obs"),
    ("oss",        "oss"),
    ("sfs",        "sfs"),
    ("evs",        "evs"),
    ("cbr",        "cbr"),
    ("elb",        "elb"),
    ("slb",        "slb"),
    ("nat",        "nat"),
    ("waf",        "waf"),
    ("cdn",        "cdn"),
    ("vpn",        "vpn"),
    ("vpc",        "vpc"),
    ("cce",        "cce"),
    ("cci",        "cci"),
    ("bms",        "bms"),
    ("ecs",        "ecs"),
    ("dms",        "dms"),
    ("smn",        "smn"),
    ("apig",       "apig"),
    ("apigateway", "apigateway"),
    ("swr",        "swr"),
    ("iam",        "iam"),
    ("kms",        "kms"),
]


def get_service_key(name, type_field):
    """从 type 字段或节点名称中提取规范化服务类型键。"""
    if type_field and type_field != "default":
        # 统一 kubernetes → k8s
        if type_field == "kubernetes":
            return "k8s"
        return type_field
    name_lower = (name or "").lower()
    for pattern, key in NAME_PATTERNS:
        if pattern in name_lower:
            return key
    return "default"


def get_service_display_name(service_key):
    return SERVICE_DISPLAY_NAMES.get(service_key, service_key.upper())


def compact_resource_name(name, max_units=26):
    """按近似显示宽度截短资源名，完整名称仍保留在节点详情中。"""
    text = str(name or "")
    units = 0
    result = []
    for char in text:
        char_units = 1 if ord(char) < 128 else 2
        if units + char_units > max_units - 3:
            return "".join(result) + "..."
        result.append(char)
        units += char_units
    return text


def make_node_display_label(name, service_key):
    """节点框固定显示两行：服务统称 + resource_name。"""
    service_name = compact_resource_name(get_service_display_name(service_key))
    return f"{service_name}\n{compact_resource_name(name)}"


# 这些类型作为 compound 容器节点（自动生成，不直接来自行数据）
_VIRTUAL_CONTAINER_TYPES = {"__region__", "__group__", "__business__"}

# 同一业务、同一服务类型最多渲染的真实资源数。其余资源合并为一个摘要节点。
MAX_VISIBLE_SERVICE_NODES = 5

# 业务容器的调色板（背景, 边框, 标题色）——按顺序循环分配给各业务
BUSINESS_PALETTE = [
    ("#E8F8F5", "#1ABC9C", "#0E6655"),   # 青绿
    ("#FEF9E7", "#F1C40F", "#7D6608"),   # 明黄
    ("#FDEDEC", "#E74C3C", "#922B21"),   # 珊瑚红
    ("#EBF5FB", "#3498DB", "#1A5276"),   # 天蓝
    ("#F4ECF7", "#9B59B6", "#5B2C6F"),   # 紫罗兰
    ("#FDF2E9", "#E67E22", "#7E5109"),   # 琥珀橙
    ("#EAFAF1", "#27AE60", "#186A3B"),   # 森林绿
    ("#F2F3F4", "#5D6D7E", "#2E4053"),   # 石墨灰
]

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
    "所属业务":     ["所属业务", "业务", "business", "business_name",
                     "app_name", "application", "所属应用", "业务域", "app"],
}


# 无意义占位节点名称关键词（大小写不敏感，包含即过滤）
FILTER_KEYWORDS = [
    "不涉及", "客户未提供", "暂不涉及", "暂无", "无需", "不需要",
    "n/a", "na", "none", "null", "无", "不适用",
]


def is_filtered_node(name: str) -> bool:
    """节点名称含无意义占位词时返回 True。"""
    n = name.strip().lower()
    return any(kw in n for kw in FILTER_KEYWORDS)


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
        if not name or is_filtered_node(name):
            continue

        records.append({
            "name":            name,
            "type":            get("服务类型", "default").lower().strip(),
            "resource_id":     get("资源ID"),
            "enterprise_id":   get("企业项目ID"),
            "region":          get("区域"),
            "group":           get("资源分组"),
            "business":        get("所属业务"),
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
    edge_lookup = {}
    edge_count = [0]

    # ── 1. 规划资源可见性 ──────────────────────────────────────────────────
    # 聚合键是 (业务键, 服务类型)。未填写业务的资源按服务类型生成虚拟业务，
    # 与普通业务执行相同的 5 节点聚合规则。

    entries = []
    service_groups = defaultdict(list)
    name_entry_map = {}

    for i, r in enumerate(records):
        nid = f"n_{safe_id(r['name'])}_{i}"
        stype = r["type"] or "default"
        bg, border, shape = SERVICE_STYLE.get(stype, SERVICE_STYLE["default"])
        reg = (r["region"]   or "").strip()
        grp = (r["group"]    or "").strip()
        source_biz = (r["business"] or "").strip()
        svc_key = get_service_key(r["name"], stype)
        is_virtual_business = not bool(source_biz)
        biz_key = source_biz or f"__service_business__:{svc_key}"
        biz_label = source_biz or get_service_display_name(svc_key)
        group_key = (biz_key, svc_key)

        data = {
            "id":            nid,
            "name":          r["name"],
            "display_label": make_node_display_label(r["name"], svc_key),
            "type":          stype,
            "service_key":   svc_key,
            "spec":          r["spec"],
            "desc":          r["desc"],
            "resource_id":   r["resource_id"],
            "enterprise_id": r["enterprise_id"],
            "region":        reg,
            "group_label":   grp,
            "business":      biz_label,
            "business_key":  biz_key,
            "is_virtual_business": 1 if is_virtual_business else 0,
            "reason":        r["reason"],
            "bg_color":      bg,
            "border_color":  border,
            "shape":         shape,
            "is_container":  0,
            "is_summary":    0,
        }
        entry = {
            "index": i,
            "record": r,
            "data": data,
            "group_key": group_key,
        }
        entries.append(entry)
        name_entry_map.setdefault(r["name"], entry)
        service_groups[group_key].append(entry)

    visible_indexes = set()
    group_info = {}
    for group_index, (group_key, group_entries) in enumerate(service_groups.items()):
        visible_entries = group_entries[:MAX_VISIBLE_SERVICE_NODES]
        hidden_entries = group_entries[MAX_VISIBLE_SERVICE_NODES:]
        visible_indexes.update(entry["index"] for entry in visible_entries)
        group_info[group_key] = {
            "entries": group_entries,
            "visible": visible_entries,
            "hidden": hidden_entries,
            "first_id": visible_entries[0]["data"]["id"],
            "summary_id": "",
            "group_index": group_index,
        }

    for entry in entries:
        if entry["index"] in visible_indexes:
            nodes.append({"group": "nodes", "data": entry["data"]})

    # 每个超限组只增加一个摘要节点；摘要节点承接组内隐藏资源的连线。
    for group_key, info in group_info.items():
        hidden_count = len(info["hidden"])
        if not hidden_count:
            continue
        biz_key, svc_key = group_key
        sample_data = info["entries"][0]["data"]
        summary_id = f"summary_{info['group_index']}"
        info["summary_id"] = summary_id
        bg, border, _ = SERVICE_STYLE.get(svc_key, SERVICE_STYLE["default"])
        nodes.append({"group": "nodes", "data": {
            "id":              summary_id,
            "name":            f"...+{hidden_count}",
            "type":            "__summary__",
            "service_key":     svc_key,
            "service_name":    get_service_display_name(svc_key),
            "business":        sample_data["business"],
            "business_key":    biz_key,
            "is_virtual_business": sample_data["is_virtual_business"],
            "is_grouped_resource": 1,
            "collapsed_count": hidden_count,
            "resource_total":  len(info["entries"]),
            "spec":            "",
            "desc":            "其余同类资源已折叠显示",
            "resource_id":     "",
            "enterprise_id":   "",
            "region":          "",
            "group_label":     "",
            "reason":          "",
            "bg_color":        bg,
            "border_color":    border,
            "shape":           "roundrectangle",
            "is_container":    0,
            "is_summary":      1,
        }})

    # ── 2. 创建业务容器节点，为叶节点设置 parent ─────────────────────────

    business_map = {}   # business_key → biz_id
    business_meta = {}
    business_totals = defaultdict(int)
    for entry in entries:
        business_totals[entry["data"]["business_key"]] += 1
    _biz_counter = 0
    for n in nodes:
        biz_key = n["data"].get("business_key", "").strip()
        if not biz_key:
            continue
        if biz_key not in business_map:
            business_map[biz_key] = f"biz_{_biz_counter}"
            business_meta[biz_key] = {
                "name": n["data"].get("business", ""),
                "is_virtual": n["data"].get("is_virtual_business", 0),
            }
            _biz_counter += 1
        n["data"]["parent"] = business_map[biz_key]
        n["data"]["is_grouped_resource"] = 1

    biz_containers = []
    for i, (biz_key, biz_id) in enumerate(business_map.items()):
        meta = business_meta[biz_key]
        bg, border, text = BUSINESS_PALETTE[i % len(BUSINESS_PALETTE)]
        biz_containers.append({"group": "nodes", "data": {
            "id":           biz_id,
            "name":         meta["name"],
            "business_key": biz_key,
            "resource_total": business_totals[biz_key],
            "is_virtual_business": meta["is_virtual"],
            "type":         "__business__",
            "is_container": 1,
            "bg_color":     bg,
            "border_color": border,
            "text_color":   text,
            "shape":        "roundrectangle",
        }})
    # ── 2.6 在业务容器内按服务类型创建子容器 ──────────────────────────────
    service_containers = []
    node_by_id = {node["data"]["id"]: node for node in nodes}
    for (biz_key, svc_key), info in service_groups.items():
        biz_id = business_map[biz_key]
        group_meta = group_info[(biz_key, svc_key)]
        rendered_ids = {entry["data"]["id"] for entry in group_meta["visible"]}
        if group_meta["summary_id"]:
            rendered_ids.add(group_meta["summary_id"])

        # 虚拟业务框本身已经代表服务类型，不再重复嵌套同名服务框。
        if business_meta[biz_key]["is_virtual"]:
            continue

        sc_id = f"sc_{biz_id}_{safe_id(svc_key)}"
        service_containers.append({"group": "nodes", "data": {
            "id":              sc_id,
            "name":            get_service_display_name(svc_key),
            "type":            "__service__",
            "service_key":     svc_key,
            "resource_total":  len(info),
            "visible_count":   len(group_meta["visible"]),
            "collapsed_count": len(group_meta["hidden"]),
            "is_container":    1,
            "parent":          biz_id,
            "bg_color":        "#F2F3F4",
            "border_color":    "#7F8C8D",
            "shape":           "roundrectangle",
        }})
        for node_id in rendered_ids:
            node = node_by_id.get(node_id)
            if node:
                node["data"]["parent"] = sc_id
                node["data"]["service_container"] = sc_id

    nodes = biz_containers + service_containers + nodes

    # ── 3. 解析外部节点（targets 中不存在的名称）─────────────────────────

    all_node_ids = {n["data"]["id"] for n in nodes}

    def ensure_ext_node(name):
        """目标名不在记录中时，创建外部虚拟节点"""
        ext_id = f"n_ext_{safe_id(name)}"
        if ext_id not in all_node_ids:
            bg, border, shape = SERVICE_STYLE["default"]
            nodes.append({"group": "nodes", "data": {
                "id": ext_id, "name": name,
                "display_label": make_node_display_label(name, "default"),
                "type": "default", "service_key": "default",
                "business": "", "business_key": "",
                "is_virtual_business": 0, "is_grouped_resource": 0,
                "is_container": 0, "is_summary": 0,
                "bg_color": bg, "border_color": border, "shape": shape,
                "spec": "", "desc": "（外部/未列出服务）",
                "resource_id": "", "enterprise_id": "",
                "region": "", "group_label": "", "reason": "",
            }})
            all_node_ids.add(ext_id)
        return ext_id

    # ── 4. 创建边 ─────────────────────────────────────────────────────────

    def internal_endpoint(entry):
        """同业务边保留可见节点；隐藏节点由摘要节点承接。"""
        if entry["index"] in visible_indexes:
            return entry["data"]["id"]
        return group_info[entry["group_key"]]["summary_id"]

    def representative_endpoint(entry):
        """跨业务边统一落到对应服务组的第一个真实资源节点。"""
        if entry["group_key"]:
            return group_info[entry["group_key"]]["first_id"]
        return entry["data"]["id"]

    def add_edge(src_id, tgt_id, source_name, target_name, cross_biz,
                 relation="调用", inferred=False):
        key = (src_id, tgt_id, inferred)
        if key in edge_lookup:
            edge_data = edge_lookup[key]
            edge_data["call_count"] += 1
            edge_data["relation"] = f"{relation} ×{edge_data['call_count']}"
            return
        color = RELATION_COLORS.get(relation, RELATION_COLORS["default"])
        if inferred:
            color = "#AAAAAA"
        if cross_biz and not inferred:
            color = "#E67E22"
        eid = f"e_{edge_count[0]}"
        edge_count[0] += 1
        edge_data = {
            "id":            eid,
            "source":        src_id,
            "target":        tgt_id,
            "source_name":   source_name,
            "target_name":   target_name,
            "relation":      relation + ("（推断）" if inferred else ""),
            "call_count":    1,
            "color":         color,
            "inferred":      1 if inferred else 0,
            "cross_business": 1 if cross_biz else 0,
        }
        edge_lookup[key] = edge_data
        edges.append({"group": "edges", "data": edge_data})

    for source_entry in entries:
        r = source_entry["record"]
        # 确认下游
        if r["targets"]:
            for t in [x.strip() for x in r["targets"].split(",") if x.strip() and not is_filtered_node(x.strip())]:
                target_entry = name_entry_map.get(t)
                if target_entry:
                    source_biz = source_entry["data"]["business_key"]
                    target_biz = target_entry["data"]["business_key"]
                    cross_biz = bool(source_biz and target_biz and source_biz != target_biz)
                    if cross_biz:
                        src_id = representative_endpoint(source_entry)
                        tgt_id = representative_endpoint(target_entry)
                    else:
                        src_id = internal_endpoint(source_entry)
                        tgt_id = internal_endpoint(target_entry)
                else:
                    cross_biz = False
                    src_id = internal_endpoint(source_entry)
                    tgt_id = ensure_ext_node(t)
                add_edge(src_id, tgt_id, r["name"], t, cross_biz)

    return nodes + edges


# ── BDAT 分组层次布局位置计算 ─────────────────────────────────────────────────

def compute_bdat_positions(elements):
    """
    按业务分组（BDAT）计算节点的预设坐标：
    - 各业务组横向排列成网格
    - 组内以服务类型为不可拆分的连续矩形块
    - 服务块按依赖关系纵向排序，块内每行最多 5 个节点
    返回 {node_id: {'x': float, 'y': float}}
    """
    NODE_W        = 195
    ROW_H         = 105
    SERVICE_GAP   = 90
    MAX_ROW_NODES = MAX_VISIBLE_SERVICE_NODES
    GROUP_PAD     = 110
    GAP_X         = 450
    GAP_Y         = 380
    MAX_COLS      = 2

    leaf_nodes = {}
    edges_list = []
    for e in elements:
        if e["group"] == "nodes" and not e["data"].get("is_container"):
            leaf_nodes[e["data"]["id"]] = e["data"]
        elif e["group"] == "edges":
            edges_list.append((e["data"]["source"], e["data"]["target"]))

    biz_groups = defaultdict(list)
    for nid, data in leaf_nodes.items():
        biz_key = (data.get("business_key") or "").strip() or "__external__"
        biz_groups[biz_key].append(nid)

    sorted_bizs = list(biz_groups.keys())
    n_groups = len(sorted_bizs)
    cols = min(MAX_COLS, n_groups) if n_groups else 1

    group_services = {}
    service_orders = {}
    service_offsets = {}
    group_sizes = {}

    for biz_key in sorted_bizs:
        nodes_in_group = set(biz_groups[biz_key])
        services = defaultdict(list)
        service_first_index = {}
        for index, nid in enumerate(biz_groups[biz_key]):
            service_key = leaf_nodes[nid].get("service_key", "default")
            services[service_key].append(nid)
            service_first_index.setdefault(service_key, index)

        # 将节点级依赖提升为服务级依赖，用于排列完整服务块。
        in_deg = {service_key: 0 for service_key in services}
        adjacency = defaultdict(set)
        for src, tgt in edges_list:
            if src not in nodes_in_group or tgt not in nodes_in_group:
                continue
            src_service = leaf_nodes[src].get("service_key", "default")
            tgt_service = leaf_nodes[tgt].get("service_key", "default")
            if src_service == tgt_service or tgt_service in adjacency[src_service]:
                continue
            adjacency[src_service].add(tgt_service)
            in_deg[tgt_service] += 1

        ready = sorted(
            (key for key, degree in in_deg.items() if degree == 0),
            key=lambda key: service_first_index[key],
        )
        service_order = []
        while ready:
            service_key = ready.pop(0)
            service_order.append(service_key)
            for target_key in sorted(adjacency[service_key],
                                     key=lambda key: service_first_index[key]):
                in_deg[target_key] -= 1
                if in_deg[target_key] == 0:
                    ready.append(target_key)
                    ready.sort(key=lambda key: service_first_index[key])

        # 服务间存在环路时，剩余服务按首次出现顺序连续排在末尾。
        service_order.extend(sorted(
            (key for key in services if key not in service_order),
            key=lambda key: service_first_index[key],
        ))

        max_row_cap = max(min(len(node_ids), MAX_ROW_NODES)
                          for node_ids in services.values())
        group_w = max_row_cap * NODE_W + 2 * GROUP_PAD

        y_acc = 0
        offsets = {}
        for index, service_key in enumerate(service_order):
            offsets[service_key] = y_acc
            row_count = ((len(services[service_key]) + MAX_ROW_NODES - 1)
                         // MAX_ROW_NODES)
            y_acc += row_count * ROW_H
            if index < len(service_order) - 1:
                y_acc += SERVICE_GAP
        group_h = y_acc + 2 * GROUP_PAD

        group_services[biz_key] = services
        service_orders[biz_key] = service_order
        service_offsets[biz_key] = offsets
        group_sizes[biz_key] = (group_w, group_h)

    # 计算网格各列最大宽度、各行最大高度
    col_w = defaultdict(int)
    row_h = defaultdict(int)
    grid_pos = {}
    for i, biz_key in enumerate(sorted_bizs):
        c, r = i % cols, i // cols
        grid_pos[biz_key] = (c, r)
        w, h = group_sizes[biz_key]
        col_w[c] = max(col_w[c], w)
        row_h[r] = max(row_h[r], h)

    max_rows = (n_groups + cols - 1) // cols if n_groups else 1

    col_x = {}
    x = 0
    for c in range(cols):
        col_x[c] = x
        x += col_w[c] + GAP_X

    row_y = {}
    y = 0
    for r in range(max_rows):
        row_y[r] = y
        y += row_h.get(r, 0) + GAP_Y

    # 逐节点赋坐标
    positions = {}
    for biz_key in sorted_bizs:
        c, r       = grid_pos[biz_key]
        base_x     = col_x[c] + GROUP_PAD
        base_y     = row_y[r] + GROUP_PAD
        content_w  = group_sizes[biz_key][0] - 2 * GROUP_PAD

        for service_key in service_orders[biz_key]:
            service_nodes = group_services[biz_key][service_key]
            service_y = base_y + service_offsets[biz_key][service_key]
            row_count = ((len(service_nodes) + MAX_ROW_NODES - 1)
                         // MAX_ROW_NODES)
            for row_index in range(row_count):
                row_nodes = service_nodes[
                    row_index * MAX_ROW_NODES:(row_index + 1) * MAX_ROW_NODES
                ]
                span = (len(row_nodes) - 1) * NODE_W
                start_x = base_x + (content_w - span) / 2
                y_pos = service_y + row_index * ROW_H
                for column_index, node_id in enumerate(row_nodes):
                    positions[node_id] = {
                        "x": round(start_x + column_index * NODE_W, 1),
                        "y": round(y_pos, 1),
                    }

    return positions


# ── HTML 生成辅助 ─────────────────────────────────────────────────────────────

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_CYTOSCAPE_PATHS = [
    # 优先使用与 cose-bilkent 插件版本匹配的稳定版
    os.path.join(_SCRIPT_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape.min.js"),
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
                "font-size": 9,
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
        # 服务类型子容器（业务内按服务名分组）
        {
            "selector": "node[type = '__service__']",
            "style": {
                "border-color": "#95A5A6",
                "border-width": 1.5,
                "border-style": "dashed",
                "background-color": "#F8F9FA",
                "background-opacity": 0.55,
                "font-size": 11,
                "font-weight": "bold",
                "color": "#2C3E50",
                "text-valign": "top",
                "text-background-color": "#fff",
                "text-background-opacity": 0.9,
                "text-background-padding": "3px",
                "padding": "20px",
                "shape": "roundrectangle",
            }
        },
        # 聚合资源固定显示两行：服务统称 + 截短后的 resource_name
        {
            "selector": "node[is_grouped_resource = 1][is_summary = 0]",
            "style": {
                "label": "data(display_label)",
                "width": 170,
                "height": 58,
                "shape": "roundrectangle",
                "background-opacity": 0.14,
                "font-size": 10,
                "font-weight": "bold",
                "text-valign": "center",
                "text-halign": "center",
                "text-margin-y": 0,
                "text-background-opacity": 0,
                "text-max-width": "158px",
                "text-wrap": "wrap",
                "color": "#2C3E50",
            }
        },
        # 被折叠资源的摘要节点
        {
            "selector": "node[is_summary = 1]",
            "style": {
                "label": "data(name)",
                "width": 72,
                "height": 46,
                "shape": "roundrectangle",
                "background-opacity": 0.14,
                "border-style": "dashed",
                "border-width": 2,
                "font-size": 12,
                "font-weight": "bold",
                "text-valign": "center",
                "text-halign": "center",
                "text-margin-y": 0,
                "text-background-opacity": 0,
                "color": "#566573",
            }
        },
        # 业务容器节点：颜色/文字从 data 属性取，各业务组配色互不相同
        {
            "selector": "node[type = '__business__']",
            "style": {
                "border-color": "data(border_color)",
                "border-width": 4,
                "border-style": "solid",
                "background-color": "data(bg_color)",
                "background-opacity": 0.28,
                "font-size": 16,
                "font-weight": "bold",
                "color": "data(text_color)",
                "text-background-color": "#fff",
                "text-background-opacity": 0.95,
                "text-background-padding": "4px",
                "text-margin-y": -6,
                "padding": "50px",
                "shape": "roundrectangle",
                "corner-radius": "12",
            }
        },
        # 跨业务确认边：橙色粗实线
        {
            "selector": "edge[cross_business = 1][inferred = 0]",
            "style": {
                "line-color": "#E67E22",
                "target-arrow-color": "#E67E22",
                "width": 2.5,
                "line-style": "dashed",
                "line-dash-pattern": [10, 4],
                "opacity": 0.9,
            }
        },
    ]


def _make_legend_html():
    exclude = {"__region__", "__group__", "__business__"}
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
    cytoscape_js = _read_first(_CYTOSCAPE_PATHS)

    if not cytoscape_js:
        sys.exit("找不到 cytoscape.min.js，请确认路径:\n" + "\n".join(_CYTOSCAPE_PATHS))

    # 计算 BDAT 分组层次布局坐标并注入节点数据
    bdat_pos = compute_bdat_positions(elements)
    for e in elements:
        if e["group"] == "nodes" and not e["data"].get("is_container"):
            nid = e["data"]["id"]
            if nid in bdat_pos:
                e["data"]["bdat_x"] = bdat_pos[nid]["x"]
                e["data"]["bdat_y"] = bdat_pos[nid]["y"]

    # ensure_ascii=True：将中文转为 \uXXXX 转义，避免内联 JS 中出现非 ASCII 字符导致的解析问题
    elements_json = json.dumps(elements, ensure_ascii=True)
    style_json    = json.dumps(_make_cytoscape_style(), ensure_ascii=True)
    legend_html   = _make_legend_html()

    layout_options  = '<option value="bdat">BDAT分组层次布局（推荐）</option>\n'
    layout_options += '    <option value="breadthfirst">层次布局</option>\n'
    layout_options += '    <option value="grid">网格布局</option>\n'
    layout_options += '    <option value="circle">圆形布局</option>\n'
    layout_options += '    <option value="concentric">同心圆布局</option>\n'
    layout_options += '    <option value="random">随机布局</option>\n'
    default_layout  = "bdat"

    # 统计数字
    visible_resources = sum(1 for e in elements if e.get("group") == "nodes"
                            and not e["data"].get("is_container")
                            and not e["data"].get("is_summary"))
    summary_nodes = sum(1 for e in elements if e.get("group") == "nodes"
                        and e["data"].get("is_summary"))
    collapsed_resources = sum(e["data"].get("collapsed_count", 0) for e in elements
                              if e.get("group") == "nodes" and e["data"].get("is_summary"))
    total_resources = visible_resources + collapsed_resources
    resource_stat = (f"{visible_resources + summary_nodes} 可见 / {total_resources} 资源"
                     if collapsed_resources else f"{total_resources} 节点")
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
#cy{{position:fixed;top:50px;left:0;right:280px;bottom:0;background:#fff}}
/* ── 侧边栏 ── */
#sidebar{{position:fixed;top:50px;right:0;width:280px;bottom:0;background:#fff;
  border-left:1px solid #e0e0e0;display:flex;flex-direction:column;overflow:hidden}}
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
  <button id="bizToggleBtn" class="tb-btn" onclick="toggleBusinessView()" style="background:rgba(26,188,156,.35)">&#127968; 业务分组:开</button>
  <div class="sep"></div>
  <select id="layoutSelect" onchange="changeLayout(this.value)">
    {layout_options}
  </select>
  <div class="sep"></div>
  <input id="searchBox" type="text" placeholder="&#128269; 搜索节点名/类型…"
         oninput="searchNodes(this.value)"/>
  <div class="stats-bar">
    <span class="stat-item"><span class="stat-dot" style="background:#3498DB"></span>{resource_stat}</span>
    <span class="stat-item"><span class="stat-dot" style="background:#27AE60"></span>{n_edges - n_infer} 确认边</span>
    <span class="stat-item"><span class="stat-dot" style="background:#AAA"></span>{n_infer} 推断边</span>
  </div>
</div>

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
      <div class="infer-legend">
        <span style="display:inline-block;width:24px;height:2px;border-top:2px dashed #E67E22;margin-right:6px"></span>
        <span style="color:#E67E22">跨业务调用</span>
      </div>
      <div class="infer-legend">
        <span style="display:inline-block;width:16px;height:10px;border:2px solid #1ABC9C;background:rgba(26,188,156,.15);border-radius:3px;margin-right:6px"></span>
        <span>业务分组</span>
      </div>
    </div>
  </div>
</div>
<div id="toast"></div>

<script>{cytoscape_js}</script>
<script>
var elements   = {elements_json};
var styleRules = {style_json};

var cy = cytoscape({{
  container: document.getElementById('cy'),
  elements:  elements,
  style:     styleRules,
  layout: {{
    name: 'preset',
    positions: function(node) {{
      var bx = node.data('bdat_x'), by = node.data('bdat_y');
      return (bx != null) ? {{x: +bx, y: +by}} : undefined;
    }},
    animate: false,
    padding: 80,
  }},
  wheelSensitivity: 0.2,
  minZoom: 0.02,
  maxZoom: 6,
}});

setTimeout(function() {{ cy.resize(); cy.fit(undefined, 50); }}, 50);

// ── 事件：点击节点 ─────────────────────────────────────────────────────────
cy.on('tap', 'node', function(evt) {{
  var d  = evt.target.data();
  var bg = d.bg_color || '#888';
  if (d.type === '__region__' || d.type === '__group__' || d.type === '__business__' || d.type === '__service__') {{
    showContainerDetail(d);
  }} else if (d.is_summary === 1) {{
    showSummaryDetail(d);
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
    '<div class="d-value">' + escHtml(d.source_name || cy.getElementById(d.source).data('name') || d.source) + '</div></div>' +
    '<div class="d-row"><div class="d-label">目标</div>' +
    '<div class="d-value">' + escHtml(d.target_name || cy.getElementById(d.target).data('name') || d.target) + '</div></div>' +
    (d.call_count > 1
      ? '<div class="d-row"><div class="d-label">合并调用</div><div class="d-value">'+d.call_count+' 条</div></div>'
      : '');
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
    ['所属业务', d.business  || '-'],
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
    var t = n.data('type');
    return t !== '__region__' && t !== '__group__' && t !== '__business__';
  }});
  if (nbrs.length > 0) {{
    html += '<div class="d-row"><div class="d-label">相邻节点</div><div class="d-value">';
    nbrs.forEach(function(n) {{
      html += '<span style="display:inline-block;margin:2px 2px 0 0;padding:2px 6px;' +
              'background:#E8EAF6;border-radius:4px;font-size:11px;cursor:pointer" ' +
              'onclick="focusNode(\\''+n.id()+'\\')" >'+escHtml(n.data('name'))+'</span>';
    }});
    html += '</div></div>';
  }}

  document.getElementById('detail-panel').innerHTML = html;
  document.getElementById('node-type-tag').innerHTML =
    '<span class="tag" style="background:'+bg+'">'+((d.type||'').toUpperCase())+'</span>';
}}

function showSummaryDetail(d) {{
  var rows = [
    ['服务类型', d.service_name || d.service_key || '-'],
    ['所属业务', d.business || '-'],
    ['省略节点', d.collapsed_count || 0],
    ['资源总数', d.resource_total || 0],
  ];
  document.getElementById('detail-panel').innerHTML = rows.map(function(r) {{
    return '<div class="d-row"><div class="d-label">'+r[0]+'</div>' +
           '<div class="d-value">'+escHtml(r[1])+'</div></div>';
  }}).join('');
  document.getElementById('node-type-tag').innerHTML =
    '<span class="tag" style="background:#7F8C8D">省略节点</span>';
}}

function showContainerDetail(d) {{
  var label = d.type === '__region__' ? '区域'
            : d.type === '__business__' ? '所属业务'
            : d.type === '__service__' ? '服务类型'
            : '资源分组';
  var tagColor = d.type === '__business__' ? '#1ABC9C'
               : d.type === '__service__'  ? '#7F8C8D'
               : '#2E86C1';
  var html = '<div class="d-row"><div class="d-label">'+label+'</div>' +
             '<div class="d-value">'+escHtml(d.name)+'</div></div>';
  var children = cy.getElementById(d.id).descendants().filter(function(n) {{
    return !n.data('is_container');
  }});
  var total = d.resource_total || children.filter(function(n) {{ return !n.data('is_summary'); }}).length;
  html += '<div class="d-row"><div class="d-label">包含资源 ('+total+')</div><div class="d-value">';
  children.forEach(function(n) {{
    html += '<span style="display:inline-block;margin:2px 2px 0 0;padding:2px 6px;' +
            'background:#E8EAF6;border-radius:4px;font-size:11px;cursor:pointer" ' +
            'onclick="focusNode(\\''+n.id()+'\\')" >'+escHtml(n.data('name'))+'</span>';
  }});
  html += '</div></div>';
  document.getElementById('detail-panel').innerHTML = html;
  document.getElementById('node-type-tag').innerHTML =
    '<span class="tag" style="background:'+tagColor+'">'+label+'</span>';
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
  var opts;
  if (name === 'bdat') {{
    opts = {{
      name: 'preset',
      positions: function(node) {{
        var bx = node.data('bdat_x'), by = node.data('bdat_y');
        return (bx != null) ? {{x: +bx, y: +by}} : undefined;
      }},
      animate: true, animationDuration: 800, padding: 80
    }};
  }} else {{
    opts = {{ name: name, nodeDimensionsIncludeLabels: true,
             animate: true, animationDuration: 600, padding: 50 }};
    if (name === 'breadthfirst') opts.directed = true;
  }}
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
        || (d.group_label||'').toLowerCase().includes(q)
        || (d.business||'').toLowerCase().includes(q);
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

var _bizViewOn = true;
var _origParents = {{}};

function toggleBusinessView() {{
  _bizViewOn = !_bizViewOn;
  var btn = document.getElementById('bizToggleBtn');
  if (!_bizViewOn) {{
    // 先摘除叶节点与服务子容器的父子关系
    cy.nodes().forEach(function(n) {{
      var p = n.data('parent');
      if (p && (p.indexOf('biz_') === 0 || p.indexOf('sc_') === 0)) {{
        _origParents[n.id()] = p;
        n.move({{ parent: null }});
      }}
    }});
    cy.nodes('[type = "__business__"],[type = "__service__"]').style('display', 'none');
    btn.innerHTML = '&#127968; 业务分组:关';
    btn.style.background = 'rgba(255,255,255,.15)';
  }} else {{
    cy.nodes('[type = "__business__"],[type = "__service__"]').style('display', 'element');
    // 先恢复业务容器直属子节点（含服务子容器），再恢复叶节点到服务子容器
    var bizChildren = {{}}, scChildren = {{}};
    Object.keys(_origParents).forEach(function(nid) {{
      var p = _origParents[nid];
      if (p.indexOf('biz_') === 0) bizChildren[nid] = p;
      else scChildren[nid] = p;
    }});
    Object.keys(bizChildren).forEach(function(nid) {{
      cy.getElementById(nid).move({{ parent: bizChildren[nid] }});
    }});
    Object.keys(scChildren).forEach(function(nid) {{
      cy.getElementById(nid).move({{ parent: scChildren[nid] }});
    }});
    _origParents = {{}};
    btn.innerHTML = '&#127968; 业务分组:开';
    btn.style.background = 'rgba(26,188,156,.35)';
  }}
  cy.fit(undefined, 50);
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
    print(f"  可见节点: {visible_resources + summary_nodes}")
    print(f"  资源总数: {total_resources}")
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
