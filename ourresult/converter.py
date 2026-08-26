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
import base64
from collections import defaultdict

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
    "dns":          ("#5DADE2", "#2874A6", "ellipse"),
    "nat":          ("#5D6D7E", "#2E4057", "roundrectangle"),
    "vpn":          ("#7F8C8D", "#424949", "roundrectangle"),
    "cdn":          ("#2980B9", "#1A5276", "ellipse"),
    "waf":          ("#C0392B", "#7B241C", "diamond"),
    "cc":           ("#566573", "#273746", "roundrectangle"),
    "cnad":         ("#117A65", "#0B5345", "roundrectangle"),
    "er":           ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "dli":          ("#2C3E50", "#1A252F", "roundrectangle"),
    # 存储
    "oss":          ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "obs":          ("#8E44AD", "#5B2C6F", "roundrectangle"),
    "sfs":          ("#9B59B6", "#6C3483", "roundrectangle"),
    "sfs3":         ("#9B59B6", "#6C3483", "roundrectangle"),
    "evs":          ("#A569BD", "#7D3C98", "roundrectangle"),
    "cbr":          ("#7FB3D3", "#2C7BB6", "roundrectangle"),
    # 数据库
    "rds":          ("#3498DB", "#1A5276", "roundrectangle"),
    "dcs":          ("#E74C3C", "#922B21", "ellipse"),
    "redis":        ("#E74C3C", "#922B21", "ellipse"),
    "geminidb":     ("#2ECC71", "#1E8449", "roundrectangle"),
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
    "zookeeper":    ("#F39C12", "#9A6007", "roundrectangle"),
    "rabbitmq":     ("#F39C12", "#9A6007", "roundrectangle"),
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
    "maas":         ("#00897B", "#005B4F", "roundrectangle"),
    "tomcat":       ("#FF6B35", "#C0441F", "roundrectangle"),
    "mysql":        ("#3498DB", "#1A5276", "roundrectangle"),
    # 默认
    "default":      ("#AED6F1", "#5D8AA8", "roundrectangle"),
}

# 服务类型显示名称（用于子容器标签）
SERVICE_DISPLAY_NAMES = {
    "nginx": "Nginx", "k8s": "K8s", "kubernetes": "K8s", "maas": "MaaS",
    "tomcat": "Tomcat", "mysql": "MySQL",
    "ecs": "ECS", "bms": "BMS", "cce": "CCE", "cci": "CCI",
    "functiongraph": "FunctionGraph",
    "vpc": "VPC", "eip": "EIP", "elb": "ELB", "slb": "SLB", "dns": "DNS",
    "nat": "NAT", "vpn": "VPN", "cdn": "CDN", "waf": "WAF",
    "cc": "CC", "cnad": "CNAD",
    "er": "ER", "dli": "DLI",
    "oss": "OSS", "obs": "OBS", "sfs": "SFS", "sfs3": "SFS3", "evs": "EVS", "cbr": "CBR",
    "rds": "RDS", "dcs": "DCS", "redis": "Redis", "dds": "DDS",
    "geminidb": "GeminiDB", "mongodb": "MongoDB", "gaussdb": "GaussDB", "dws": "DWS",
    "css": "CSS", "elasticsearch": "Elasticsearch",
    "kafka": "Kafka", "mq": "MQ", "dms": "DMS",
    "zookeeper": "ZooKeeper", "rabbitmq": "RabbitMQ", "smn": "SMN",
    "apigateway": "API Gateway", "apig": "APIG",
    "swr": "SWR", "ecr": "ECR",
    "iam": "IAM", "kms": "KMS",
    "default": "Other",
}

# PNG icon configuration.
ICON_DIR = os.path.join(_SCRIPT_DIR, "picture")
ICON_SIZE = 48

# Expected PNG files under ourresult/picture. Aliases reuse the closest icon so
# existing spreadsheets do not need to change their service type values.
SERVICE_ICON_FILES = {
    "cce": "CCE_Deployment.png",
    "cce_deployment": "CCE_Deployment.png",
    "cce_deploym": "CCE_Deployment.png",
    "cdn": "CDN.png",
    "cc": "CC.png",
    "cnad": "CNAD.png",
    "dcs": "DCS.png",
    "redis": "DCS.png",
    "dns": "DNS.png",
    "ecs": "ECS.png",
    "bms": "ECS.png",
    "elb": "ELB.png",
    "slb": "ELB.png",
    "mq": "MQ.png",
    "dms": "MQ.png",
    "kafka": "MQ.png",
    "rabbitmq": "MQ.png",
    "obs": "OBS.png",
    "oss": "OBS.png",
    "rds": "RDS.png",
    "geminidb": "GeminiDB.png",
    "sfs": "SFS.png",
    "sfs3": "SFS.png",
    "vpc": "VPC.png",
    "waf": "WAF.png",
    "zookeeper": "ZooKeeper.png",
}

SERVICE_ICON_ALIASES = {
    "cce": ("cce_deployment", "cce deploym", "cce deploy", "cce deployment"),
    "cce_deployment": ("cce", "cce_deploym", "cce deployment"),
    "cce_deploym": ("cce", "cce_deployment", "cce deployment"),
    "dcs": ("redis",),
    "redis": ("dcs",),
    "elb": ("slb",),
    "slb": ("elb",),
    "mq": ("dms", "kafka", "rabbitmq"),
    "dms": ("mq",),
    "kafka": ("mq",),
    "rabbitmq": ("mq",),
    "obs": ("oss",),
    "oss": ("obs",),
    "sfs": ("sfs3",),
    "sfs3": ("sfs",),
}

_ICON_DATA_URI_CACHE = None


def _normalize_icon_key(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _read_png_data_uri(path):
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _load_service_icon_data_uris():
    global _ICON_DATA_URI_CACHE
    if _ICON_DATA_URI_CACHE is not None:
        return _ICON_DATA_URI_CACHE

    icons = {}
    normalized_icons = {}

    def remember_icon(key, data_uri):
        key = str(key or "").casefold()
        if not key:
            return
        icons.setdefault(key, data_uri)
        icons.setdefault(_normalize_icon_key(key), data_uri)
        normalized_icons.setdefault(_normalize_icon_key(key), data_uri)

    for service_key, filename in SERVICE_ICON_FILES.items():
        path = os.path.join(ICON_DIR, filename)
        try:
            data_uri = _read_png_data_uri(path)
        except FileNotFoundError:
            continue
        remember_icon(service_key, data_uri)
        remember_icon(os.path.splitext(filename)[0], data_uri)

    try:
        filenames = os.listdir(ICON_DIR)
    except FileNotFoundError:
        filenames = []

    for filename in filenames:
        if not filename.casefold().endswith(".png"):
            continue
        path = os.path.join(ICON_DIR, filename)
        if not os.path.isfile(path):
            continue
        data_uri = _read_png_data_uri(path)
        remember_icon(os.path.splitext(filename)[0], data_uri)

    for service_key, display_name in SERVICE_DISPLAY_NAMES.items():
        candidates = [
            service_key,
            display_name,
            os.path.splitext(SERVICE_ICON_FILES.get(service_key, ""))[0],
            *SERVICE_ICON_ALIASES.get(service_key, ()),
        ]
        for candidate in candidates:
            data_uri = (icons.get(str(candidate or "").casefold())
                        or normalized_icons.get(_normalize_icon_key(candidate)))
            if data_uri:
                remember_icon(service_key, data_uri)
                break

    _ICON_DATA_URI_CACHE = icons
    return icons


def get_service_icon_data_uri(service_key):
    icons = _load_service_icon_data_uris()
    return (icons.get(str(service_key or "").casefold())
            or icons.get(_normalize_icon_key(service_key))
            or "")


# 从节点名称中识别服务类型的关键词（顺序从精确到宽泛）
NAME_PATTERNS = [
    ("kubernetes", "k8s"), ("k8s", "k8s"),
    ("maas",       "maas"),
    ("nginx",      "nginx"),
    ("tomcat",     "tomcat"),
    ("redis",      "redis"),
    ("zookeeper",  "zookeeper"),
    ("zoo-keeper", "zookeeper"),
    ("rabbitmq",   "rabbitmq"),
    ("rabbit-mq",  "rabbitmq"),
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
    ("sfs3",       "sfs3"),
    ("evs",        "evs"),
    ("cbr",        "cbr"),
    ("elb",        "elb"),
    ("slb",        "slb"),
    ("dns",        "dns"),
    ("nat",        "nat"),
    ("waf",        "waf"),
    ("cdn",        "cdn"),
    ("vpn",        "vpn"),
    ("vpc",        "vpc"),
    ("geminidb",   "geminidb"),
    ("cnad",       "cnad"),
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


# CCE_Deployment resource-grouping helpers

# CCE_Deployment 业务前缀取 resource_name 的前 N 段（按 - 分隔）。
# 一般取前 3 段；如需按前 4 段分组，将本常量改为 4 即可。
CCE_BUSINESS_PREFIX_SEGMENTS = 3


def _compact_lower(value):
    """Strip whitespace, underscores and dashes, then casefold."""
    return re.sub(r"[\s_\-]+", "", str(value or "")).casefold()


def is_cce_deployment_resource(record):
    """True when the resource-group column marks a CCE_Deployment, or the
    service-type is exactly CCE_Deployment (not the generic cce_deploym)."""
    group_text = _compact_lower(record.get("group", ""))
    type_text = _compact_lower(record.get("type", ""))
    return ("ccedeployment" in group_text
            or type_text == "ccedeployment")


def extract_cce_business_prefix(name):
    """Return the first CCE_BUSINESS_PREFIX_SEGMENTS dash-separated segments."""
    parts = [part.strip() for part in str(name or "").split("-") if part.strip()]
    return ("-".join(parts[:CCE_BUSINESS_PREFIX_SEGMENTS])
            if len(parts) >= CCE_BUSINESS_PREFIX_SEGMENTS else "")


def split_group_key(group_key):
    """Unpack a 3-tuple (plain), 4-tuple (CCE project only) or 5-tuple
    (CCE prefix + enterprise project) grouping key."""
    if len(group_key) == 5:
        biz_key, layer_key, svc_key, prefix_key, project_key = group_key
        return biz_key, layer_key, svc_key, prefix_key, project_key
    if len(group_key) == 4:
        biz_key, layer_key, svc_key, project_key = group_key
        return biz_key, layer_key, svc_key, None, project_key
    biz_key, layer_key, svc_key = group_key
    return biz_key, layer_key, svc_key, None, None


def _resource_id_hint(resource_id):
    text = str(resource_id or "").casefold()
    if "ccaas" in text:
        return "cc"
    return ""


def _resource_group_hint(group):
    text = str(group or "").casefold()
    if re.match(r"^cc(?:[-_]|$)", text):
        return "cc"
    return ""


def get_service_key(name, type_field, resource_id="", group=""):
    """从 type 字段或节点名称中提取规范化服务类型键。"""
    resource_hint = _resource_id_hint(resource_id)
    if resource_hint:
        return resource_hint
    group_hint = _resource_group_hint(group)
    if group_hint:
        return group_hint
    type_text = str(type_field or "").lower().strip()
    # “Other” is the UI label for the fallback service type. Treat an
    # explicit value the same as an empty/default type so a recognizable
    # resource name can still be classified, otherwise it will be filtered.
    if type_text in {"default", "other"}:
        type_text = ""
    if type_text:
        # ECS/BMS-hosted middleware should be grouped as middleware, not app compute.
        if type_text in {"ecs", "bms"}:
            name_lower = (name or "").lower()
            for pattern, key in (
                    ("zookeeper", "zookeeper"), ("zoo-keeper", "zookeeper"),
                    ("rabbitmq", "rabbitmq"), ("rabbit-mq", "rabbitmq"),
                    ("rabbit_mq", "rabbitmq")):
                if pattern in name_lower:
                    return key
        normalized_type = re.split(r"[-_./:\s]+", type_text, maxsplit=1)[0]
        if normalized_type == "kubernetes":
            return "k8s"
        if normalized_type == "ccaas":
            return "cc"
        if normalized_type == "dcaas":
            return "dcaas"
        if normalized_type in SERVICE_DISPLAY_NAMES and normalized_type != "default":
            return normalized_type
        for pattern, key in NAME_PATTERNS:
            if pattern in type_text:
                return "k8s" if key == "kubernetes" else key
        # 统一 kubernetes → k8s
        if type_text == "kubernetes":
            return "k8s"
        return type_text
    name_lower = (name or "").lower()
    if re.match(r"^cc(?:[-_]|$)", name_lower):
        return "cc"
    for pattern, key in NAME_PATTERNS:
        if pattern in name_lower:
            return key
    return "default"


def get_service_display_name(service_key):
    return SERVICE_DISPLAY_NAMES.get(service_key, service_key.upper())


def is_other_service_record(record):
    """Return whether a resource would be placed in the fallback Other group."""
    return get_service_key(
        record.get("name", ""),
        record.get("type", "default"),
        record.get("resource_id", ""),
        record.get("group", ""),
    ) == "default"


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


def wrap_display_name(name, max_units=18):
    """按近似显示宽度把长名称折成多行（显式插入换行），避免一行溢出重叠。"""
    lines = []
    current = []
    units = 0
    for char in str(name or ""):
        char_units = 1 if ord(char) < 128 else 2
        if units + char_units > max_units and current:
            lines.append("".join(current))
            current = []
            units = 0
        current.append(char)
        units += char_units
    if current:
        lines.append("".join(current))
    return "\n".join(lines) if lines else ""


def make_node_display_label(name, service_key):
    """节点框固定显示两行：服务统称 + resource_name。"""
    service_name = compact_resource_name(get_service_display_name(service_key))
    return f"{service_name}\n{wrap_display_name(name)}"


def container_min_width(text, font_size=22, padding=48):
    """按名称长度估算容器最小宽度，保证框内文字完整显示、不缩写。"""
    units = 0
    for char in str(text or ""):
        units += font_size if ord(char) > 127 else font_size * 0.62
    return int(units + padding)


TOPOLOGY_LAYERS = [
    ("access", "渠道接入与安全边界"),
    ("network_lb", "网络与负载均衡层"),
    ("compute_app", "计算、容器与应用服务层"),
    ("data_middleware_storage", "中间件、数据与存储层"),
]
TOPOLOGY_LAYER_ORDER = {key: index for index, (key, _) in enumerate(TOPOLOGY_LAYERS)}
TOPOLOGY_LAYER_NAMES = dict(TOPOLOGY_LAYERS)

ACCESS_LAYER_TYPES = {"cdn", "waf", "eip", "dns", "cnad"}
NETWORK_LB_LAYER_TYPES = {"elb", "slb", "vpn", "er"}
# These network resources remain inside the business compound node, but are
# intentionally rendered outside the second architecture-layer container.
EXTRACTED_FROM_LAYER_TYPES = {"vpc", "dcaas", "nat", "cc"}
COMPUTE_APP_LAYER_TYPES = {
    "ecs", "bms", "cce", "cci", "k8s", "kubernetes", "functiongraph",
    "apigateway", "apig", "nginx", "tomcat", "maas", "swr", "ecr",
}
DATA_MIDDLEWARE_STORAGE_LAYER_TYPES = {
    "dcs", "redis", "rds", "dds", "mongodb", "gaussdb", "dws", "css",
    "elasticsearch", "oss", "obs", "sfs", "sfs3", "evs", "cbr", "kafka", "mq",
    "dms", "smn", "mysql", "zookeeper", "rabbitmq", "geminidb",
}
DB_RESOURCE_NAME_RE = re.compile(r"(?:tidb|db)", re.IGNORECASE)
MIDDLEWARE_NAME_RE = re.compile(
    r"(^|[^a-z0-9])(zookeeper|zoo-keeper|zk|rabbitmq|rabbit-mq|rabbit_mq|rabbit)"
    r"([^a-z0-9]|$)",
    re.IGNORECASE,
)


def get_topology_layer(service_key, name="", type_field=""):
    """Return the fixed business-internal layer key for a resource."""
    service = str(service_key or "default").casefold()
    text = f"{name or ''} {type_field or ''} {service}".casefold()

    if service in DATA_MIDDLEWARE_STORAGE_LAYER_TYPES or MIDDLEWARE_NAME_RE.search(text):
        return "data_middleware_storage"
    if service in ACCESS_LAYER_TYPES:
        return "access"
    if service in EXTRACTED_FROM_LAYER_TYPES:
        # Keep the logical network classification for relationships and
        # ordering, while the parent assignment renders it outside the layer.
        return "network_lb"
    if service in NETWORK_LB_LAYER_TYPES:
        return "network_lb"
    if service in COMPUTE_APP_LAYER_TYPES:
        return "compute_app"
    return "compute_app"


def get_topology_layer_name(layer_key):
    return TOPOLOGY_LAYER_NAMES.get(layer_key, TOPOLOGY_LAYER_NAMES["compute_app"])


def _is_cce_service_record(record):
    service_key = get_service_key(
        record.get("name", ""), record.get("type", "default"),
        record.get("resource_id", ""), record.get("group", ""),
    )
    return service_key == "cce"


def _resource_name_matches_db(name):
    return bool(DB_RESOURCE_NAME_RE.search(str(name or "")))


def get_contextual_layer_overrides(records):
    """Return resource identities whose layer depends on CCE downstream links.

    A DB-named ECS directly referenced by CCE belongs to the data layer.  A
    DB-named ELB directly referenced by CCE belongs to the compute layer, and
    ECS resources directly referenced by that ELB belong to the data layer.
    Only confirmed downstream links are used here; inferred-link behavior is
    intentionally left unchanged.
    """
    records = list(records)
    records_by_name = defaultdict(list)
    for record in records:
        records_by_name[str(record.get("name", "")).strip().casefold()].append(
            record
        )

    ecs_data_records = set()
    db_elb_compute_records = set()

    for source in records:
        if not _is_cce_service_record(source):
            continue
        for target_name in split_target_names(source.get("targets", "")):
            targets = records_by_name.get(str(target_name).strip().casefold(), [])
            for target in targets:
                target_service = get_service_key(
                    target.get("name", ""), target.get("type", "default"),
                    target.get("resource_id", ""), target.get("group", ""),
                )
                if target_service == "ecs":
                    if _resource_name_matches_db(target.get("name", "")):
                        ecs_data_records.add(id(target))
                    continue

                if target_service not in {"elb", "slb"}:
                    continue
                if not _resource_name_matches_db(target.get("name", "")):
                    continue

                db_elb_compute_records.add(id(target))
                for downstream_name in split_target_names(
                        target.get("targets", "")
                ):
                    for downstream in records_by_name.get(
                            str(downstream_name).strip().casefold(), []
                    ):
                        downstream_service = get_service_key(
                            downstream.get("name", ""),
                            downstream.get("type", "default"),
                            downstream.get("resource_id", ""),
                            downstream.get("group", ""),
                        )
                        if downstream_service == "ecs":
                            ecs_data_records.add(id(downstream))

    return ecs_data_records, db_elb_compute_records


# 这些类型作为 compound 容器节点（自动生成，不直接来自行数据）
_VIRTUAL_CONTAINER_TYPES = {
    "__region__", "__group__", "__business__", "__layer__", "__service__",
}

# 同一业务、同一服务类型最多渲染的真实资源数。其余资源合并为一个摘要节点。
# 布局：每行最多摆放的云资源节点数
MAX_NODES_PER_ROW = 7
DATA_LAYER_SERVICE_ROWS = 3
EXTRACTED_RIGHT_GAP = 520

# 业务关键字过滤：只绘制“所属业务”包含该关键字（不区分大小写）的业务及其节点；
# 设为 None 表示不过滤。
# 业务过滤默认关闭；拓扑按 Excel 的“所属业务”原值分别分组。
# 保留配置项以兼容调用方显式设置过滤关键字的场景。
BUSINESS_FILTER_KEYWORD = None

# 需要聚合为一个缩略节点的服务类型（按业务聚合）
AGGREGATE_SERVICE_KEYS = {"cce", "ecs"}

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
    "企业项目":     ["企业项目", "enterprise project", "enterprise_project",
                     "enterprise_project_name", "enterpriseprojectname",
                     "ep_name", "企业项目名称"],
    "企业项目ID":   ["enterprise_project_id", "enterpriseprojectid",
                     "企业项目id", "企业项目ID", "project_id"],
    "区域":         ["region", "区域", "地域", "availability_zone", "az"],
    "资源分组":     ["资源分组", "分组", "resource_group", "group",
                     "resource group", "标签分组"],
    "描述":         ["描述", "description", "备注", "说明", "desc", "remark"],
    "规格":         ["规格", "spec", "specification", "配置", "instance_type",
                     "flavor"],
    "来源":         ["来源", "source", "origin", "来源系统", "source_system"],
    "下游服务":     ["下游服务", "downstream", "downstream_service",
                     "连接目标", "target", "targets", "依赖服务",
                     "下游服务(逗号分隔)", "下游服务（逗号分隔）"],
    "下游服务推断": ["下游服务(推断)", "下游服务（推断）", "inferred_downstream",
                     "推断下游", "inferred downstream"],
    "推断原因":     ["推断原因", "inference_reason", "reason", "推断依据",
                     "infer_reason"],
    "所属业务":     ["所属业务", "业务", "business", "business_name",
                     "app_name", "application", "所属应用", "业务域", "app"],
    "是否核心业务": ["是否核心业务", "核心业务", "is_core_business",
                     "core_business", "is core business"],
}


# 无意义占位节点名称：中文短语按包含匹配，短英文值仅做整值匹配。
FILTER_PHRASES = [
    "不涉及", "客户未提供", "暂不涉及", "暂无", "无需", "不需要", "不适用",
]
FILTER_EXACT_VALUES = {"n/a", "na", "none", "null", "无"}


def is_filtered_node(name: str) -> bool:
    """节点名称是无意义占位值时返回 True。"""
    n = name.strip().lower()
    return n in FILTER_EXACT_VALUES or any(phrase in n for phrase in FILTER_PHRASES)


def split_target_names(value):
    """兼容中英文逗号、分号和换行分隔的下游资源引用。"""
    return [part.strip() for part in re.split(r"[,，;；\r\n]+", str(value or ""))
            if part.strip() and not is_filtered_node(part.strip())]


def extract_inferred_reason_chains(value):
    """提取推断原因中“下游(推断)：A->B”形式的链路。"""
    text = str(value or "")
    marker = re.compile(
        r"下游\s*[（(]\s*推断\s*[）)]\s*[:：]\s*",
        re.IGNORECASE,
    )
    matches = list(marker.finditer(text))
    chains = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        segment = text[match.end():end]
        segment = re.split(r"[\r\n;；。]", segment, maxsplit=1)[0].strip()
        if not re.search(r"(?:-+>|=+>|→|⇒|➜)", segment):
            continue
        parts = [
            part.strip()
            for part in re.split(r"\s*(?:-+>|=+>|→|⇒|➜)\s*", segment)
            if part.strip()
        ]
        if len(parts) >= 2:
            chains.append(parts)
    return chains


def merged_business_name(value):
    """Return the visible business label after optional keyword-based merging."""
    text = str(value or "").strip()
    if BUSINESS_FILTER_KEYWORD and BUSINESS_FILTER_KEYWORD.casefold() in text.casefold():
        return BUSINESS_FILTER_KEYWORD
    return text


def canonical_business_name(value):
    """生成业务身份键，合并大小写及常见分隔符差异。"""
    text = merged_business_name(value)
    return re.sub(r"[\s_-]+", "-", text).strip("-").casefold()


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

    has_core_col = "是否核心业务" in col_map
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
            "enterprise_project": get("企业项目"),
            "region":          get("区域"),
            "group":           get("资源分组"),
            "business":        get("所属业务"),
            "desc":            get("描述"),
            "spec":            get("规格"),
            "source":          get("来源"),
            "targets":         get("下游服务"),
            "inferred_targets":get("下游服务推断"),
            "reason":          get("推断原因"),
            "is_core_business": (
                get("是否核心业务") if has_core_col else None
            ),
        })

    return records


# ── 图数据构建 ────────────────────────────────────────────────────────────────

def build_graph(records):
    """构建 cytoscape elements（nodes + edges）"""

    def safe_id(s):
        return re.sub(r"[^a-zA-Z0-9_\-]", "_", str(s))

    records = list(records)

    # 核心业务过滤：只要表里存在“是否核心业务”列，仅保留值为“是”的节点；
    # 被过滤掉的节点不出现在图中，对它们的引用也不会生成外部节点。
    def _name_key(name):
        return str(name or "").strip().casefold()

    core_column_present = any(
        r.get("is_core_business") is not None for r in records
    )
    filtered_names = set()
    if core_column_present:
        for r in records:
            if (r.get("is_core_business") or "").strip() != "是":
                filtered_names.add(_name_key(r["name"]))
        records = [
            r for r in records if _name_key(r["name"]) not in filtered_names
        ]

    # 可选业务关键字过滤：仅在调用方显式设置 BUSINESS_FILTER_KEYWORD 时生效。
    # 未填写所属业务的记录暂时保留，待回退归并后再按最终业务过滤。
    if BUSINESS_FILTER_KEYWORD:
        keyword = BUSINESS_FILTER_KEYWORD.casefold()
        for r in records:
            biz_text = (r.get("business") or "").strip()
            if biz_text and keyword not in biz_text.casefold():
                filtered_names.add(_name_key(r["name"]))
        records = [
            r for r in records
            if not ((r.get("business") or "").strip()
                    and keyword
                    not in (r.get("business") or "").strip().casefold())
        ]

    # The fallback service type is displayed as "Other". Remove those
    # records before planning containers so neither resources nor references
    # to them can create an Other business/service group.
    other_names = {
        _name_key(r["name"])
        for r in records
        if is_other_service_record(r)
    }
    if other_names:
        filtered_names.update(other_names)
        records = [
            r for r in records if _name_key(r["name"]) not in other_names
        ]

    nodes = []
    edges = []
    edge_lookup = {}
    edge_count = [0]
    contextual_ecs_records, contextual_elb_records = (
        get_contextual_layer_overrides(records)
    )

    # ── 1. 规划资源可见性 ──────────────────────────────────────────────────
    # 聚合键是 (业务键, 服务类型)。未填写业务的资源按服务类型生成虚拟业务，
    # 与普通业务执行相同的 5 节点聚合规则。

    entries = []

    # 预扫描：收集“已填写所属业务”的 CCE_Deployment 节点所归属的企业项目，
    # 供未填写所属业务的 CCE 节点按相同企业项目回退并入对应业务分组。
    project_biz_candidates = defaultdict(list)  # 企业项目(casefold) -> [候选业务]
    for r in records:
        if not is_cce_deployment_resource(r):
            continue
        biz_text = (r["business"] or "").strip()
        project_text = (
            (r.get("enterprise_project") or r["enterprise_id"] or "").strip()
        )
        if not biz_text or not project_text:
            continue
        svc_key = get_service_key(
            r["name"], r["type"], r.get("resource_id"), r.get("group")
        )
        normalized_biz = canonical_business_name(biz_text)
        biz_key = f"__business__:{normalized_biz or biz_text.casefold()}"
        prefix = extract_cce_business_prefix(r["name"])
        candidates = project_biz_candidates[project_text.casefold()]
        cand = next(
            (item for item in candidates if item["biz_key"] == biz_key),
            None,
        )
        if cand is None:
            cand = {
                "biz_key": biz_key,
                "biz_label": merged_business_name(biz_text),
                "prefix_keys": set(),
                "count": 0,
            }
            candidates.append(cand)
        if prefix:
            cand["prefix_keys"].add(prefix.casefold())
        cand["count"] += 1

    def make_entry(r, index, is_external=False, forced_service_key=None):
        nid = f"n_{safe_id(r['name'])}_{index}"
        stype = r["type"] or "default"
        reg = (r["region"]   or "").strip()
        grp = (r["group"]    or "").strip()
        source_biz = (r["business"] or "").strip()
        svc_key = forced_service_key or get_service_key(
            r["name"], stype, r.get("resource_id"), r.get("group")
        )
        layer_key = get_topology_layer(svc_key, r["name"], stype)
        if (svc_key == "ecs"
                and "tidb" in str(r["name"] or "").casefold()):
            layer_key = "data_middleware_storage"
        elif id(r) in contextual_ecs_records:
            layer_key = "data_middleware_storage"
        elif id(r) in contextual_elb_records:
            layer_key = "compute_app"
        bg, border, shape = SERVICE_STYLE.get(
            svc_key, SERVICE_STYLE.get(stype, SERVICE_STYLE["default"])
        )

        # CCE_Deployment 资源：业务框仍按“所属业务”字段划分，组内再按
        # resource_name 前 3 段（前缀）与企业项目（enterprise_id）逐级细分。
        is_cce = is_cce_deployment_resource(r)
        cce_prefix = extract_cce_business_prefix(r["name"]) if is_cce else ""
        enterprise_project = (
            (r.get("enterprise_project") or r["enterprise_id"] or "").strip()
        )
        project_label = enterprise_project or "(未设置企业项目)"

        # 未填写所属业务时，退而求其次并入“相同企业项目 + 相同前缀”所在的
        # 业务分组；多个候选时按该业务内匹配的 CCE 资源数取最多者。
        fallback_biz_key = ""
        fallback_biz_label = ""
        if is_cce and not source_biz and enterprise_project and cce_prefix:
            prefix_key_hint = cce_prefix.casefold()
            candidates = [
                item for item in project_biz_candidates.get(
                    enterprise_project.casefold(), []
                )
                if prefix_key_hint in item["prefix_keys"]
            ]
            if candidates:
                best = max(candidates, key=lambda item: item["count"])
                fallback_biz_key = best["biz_key"]
                fallback_biz_label = best["biz_label"]
                source_biz = fallback_biz_label

        is_virtual_business = not bool(source_biz)
        if fallback_biz_key:
            biz_key = fallback_biz_key
        elif source_biz:
            normalized_biz = canonical_business_name(source_biz)
            biz_key = f"__business__:{normalized_biz or source_biz.casefold()}"
        else:
            biz_key = f"__service_business__:{svc_key}"
        biz_label = merged_business_name(source_biz) or get_service_display_name(svc_key)
        group_key = (biz_key, layer_key, svc_key)

        data = {
            "id":            nid,
            "name":          r["name"],
            "display_label": make_node_display_label(r["name"], svc_key),
            "type":          stype,
            "service_key":   svc_key,
            "layer_key":     layer_key,
            "layer_name":    get_topology_layer_name(layer_key),
            "spec":          r["spec"],
            "desc":          r["desc"],
            "resource_id":   r["resource_id"],
            "enterprise_id": r["enterprise_id"],
            "region":        reg,
            "group_label":   grp,
            "source":        r.get("source", ""),
            # Keep the original downstream field on each rendered resource so
            # the in-page editor can prefill it from the Excel source.
            "targets":       r.get("targets", "") or "",
            "business":      biz_label,
            "business_key":  biz_key,
            "source_business": (r["business"] or "").strip(),
            "enterprise_project": enterprise_project,
            "enterprise_project_name": (r.get("enterprise_project") or "").strip(),
            "cce_business_prefix": cce_prefix,
            "is_cce_deployment": 1 if is_cce else 0,
            "business_fallback": 1 if fallback_biz_key else 0,
            "is_virtual_business": 1 if is_virtual_business else 0,
            "reason":        r["reason"],
            "bg_color":      bg,
            "border_color":  border,
            "shape":         shape,
            "is_container":  0,
            "is_summary":    0,
            "is_external":   1 if is_external else 0,
        }
        return {
            "index": index,
            "record": r,
            "data": data,
            "group_key": group_key,
        }

    for i, r in enumerate(records):
        entry = make_entry(r, i)
        entries.append(entry)

    # 回退归并后的最终业务若不含关键字，同样丢弃（含虚拟业务）。
    if BUSINESS_FILTER_KEYWORD:
        keyword = BUSINESS_FILTER_KEYWORD.casefold()
        kept_entries = []
        for entry in entries:
            biz_label = (entry["data"].get("business") or "").strip()
            if keyword in biz_label.casefold():
                kept_entries.append(entry)
            else:
                filtered_names.add(_name_key(entry["data"].get("name", "")))
        entries = kept_entries

    def name_key(name):
        return str(name or "").strip().casefold()

    def build_name_map(candidate_entries):
        result = defaultdict(list)
        for candidate in candidate_entries:
            result[name_key(candidate["data"]["name"])].append(candidate)
        return result

    service_aliases = {}
    for key, display_name in SERVICE_DISPLAY_NAMES.items():
        service_aliases[name_key(key)] = key
        service_aliases[name_key(display_name)] = key

    def service_key_matches(candidate_key, requested_key):
        candidate_key = name_key(candidate_key)
        requested_key = name_key(requested_key)
        if requested_key in {"dcs", "redis"}:
            return any(
                candidate_key == key
                or candidate_key.startswith(key + "_")
                or candidate_key.startswith(key + "-")
                for key in ("dcs", "redis")
            )
        return (candidate_key == requested_key
                or candidate_key.startswith(requested_key + "_")
                or candidate_key.startswith(requested_key + "-"))

    def resolve_target_entry(source_entry, target_name, candidate_entries,
                             candidate_name_map):
        """优先解析同业务同名资源，也支持 RDS/DCS 等服务统称引用。"""
        candidates = candidate_name_map.get(name_key(target_name), [])
        source_biz = source_entry["data"]["business_key"]
        for candidate in candidates:
            if candidate["data"]["business_key"] == source_biz:
                return candidate
        if candidates:
            return candidates[0]

        service_key = service_aliases.get(name_key(target_name))
        if service_key:
            for candidate in candidate_entries:
                if (candidate["data"]["business_key"] == source_biz
                        and service_key_matches(
                            candidate["data"]["service_key"], service_key
                        )):
                    return candidate
        return None

    def parse_inferred_reference(value):
        """解析 Service:名称、服务类型-业务名及名称中的服务类型字段。"""
        text = str(value or "").strip()
        service_match = re.match(r"^service\s*[:：]\s*(.+)$", text, re.IGNORECASE)
        if service_match:
            return {
                "kind": "resource_name",
                "target_name": service_match.group(1).strip(),
            }

        typed_match = re.match(
            r"^(dcs|dsc|rds|dds|redis)[-_](.+)$", text, re.IGNORECASE
        )
        if typed_match:
            service_key = typed_match.group(1).casefold()
            return {
                "kind": "business_service",
                "service_key": "dcs" if service_key == "dsc" else service_key,
                "business": typed_match.group(2).strip(),
                "target_name": text,
            }

        parts = [part for part in re.split(r"[-_.:/\s]+", text) if part]
        for index, part in enumerate(parts):
            service_key = part.casefold()
            if service_key not in {"dcs", "dsc", "rds", "dds", "redis"}:
                continue
            normalized_service = "dcs" if service_key == "dsc" else service_key
            remaining_parts = parts[:index] + parts[index + 1:]
            hints = []
            for hint_parts in (
                    remaining_parts, parts[index + 1:], parts[:index]):
                hint = "-".join(hint_parts).strip("-")
                if hint and hint not in hints:
                    hints.append(hint)
            return {
                "kind": "service_marker",
                "service_key": normalized_service,
                "business_hints": hints,
                "target_name": text,
            }
        return {"kind": "resource_name", "target_name": text}

    def business_reference_key(value):
        return canonical_business_name(value)

    def resolve_inferred_target_entry(source_entry, target_name,
                                      candidate_entries, candidate_name_map):
        reference = parse_inferred_reference(target_name)
        # 资源名恰好含 rds/dcs 等字段时，精确名称仍然优先。
        direct_target = resolve_target_entry(
            source_entry, reference["target_name"],
            candidate_entries, candidate_name_map
        )
        if direct_target:
            return direct_target
        if reference["kind"] == "resource_name":
            return None

        requested_service = reference["service_key"]
        if reference["kind"] == "service_marker":
            source_biz = source_entry["data"]["business_key"]
            for candidate in candidate_entries:
                if (candidate["data"]["business_key"] == source_biz
                        and service_key_matches(
                            candidate["data"]["service_key"], requested_service
                        )):
                    return candidate
            business_hints = reference["business_hints"]
        else:
            business_hints = [reference["business"]]

        requested_businesses = {
            business_reference_key(hint) for hint in business_hints if hint
        }
        for candidate in candidate_entries:
            candidate_business = candidate["data"].get("business", "")
            if (service_key_matches(
                        candidate["data"]["service_key"], requested_service
                    )
                    and business_reference_key(candidate_business)
                    in requested_businesses):
                return candidate
        return None

    def resolve_reason_chain_entry(source_entry, token, candidate_entries,
                                   candidate_name_map):
        """解析原因链路节点；描述性中间词无法命中时返回 None。"""
        cleaned_token = str(token or "").strip().strip("-_ ")
        if not cleaned_token:
            return None

        marker_parts = [
            part.casefold()
            for part in re.split(r"[-_.:/,，、;；()（）\[\]\s]+", cleaned_token)
            if part
        ]
        marker_key = ""
        for part in marker_parts:
            if part == "dsc":
                marker_key = "dcs"
                break
            service_key = service_aliases.get(part)
            if service_key and service_key != "default":
                marker_key = service_key
                break

        # 链路首段若是当前资源的服务统称，应锚定当前行而非同组首节点。
        if marker_key and service_key_matches(
                source_entry["data"]["service_key"], marker_key):
            exact_candidates = candidate_name_map.get(name_key(cleaned_token), [])
            if not exact_candidates:
                return source_entry

        target_entry = resolve_inferred_target_entry(
            source_entry, cleaned_token, candidate_entries, candidate_name_map
        )
        if target_entry:
            return target_entry
        if not marker_key:
            return None

        source_biz = source_entry["data"]["business_key"]
        for candidate in candidate_entries:
            if (candidate["data"]["business_key"] == source_biz
                    and service_key_matches(
                        candidate["data"]["service_key"], marker_key
                    )):
                return candidate
        return None

    def infer_external_service_key(source_entry, target_name):
        service_key = get_service_key(target_name, "default")
        if service_key != "default":
            return service_key
        source_service = source_entry["data"]["service_key"]
        if source_service == "k8s" or source_service.startswith("cce"):
            return "k8s"
        if re.search(
                r"(^|[/_.:\-])(pod|deployment|deploy|statefulset|daemonset|namespace)"
                r"($|[/_.:\-])", target_name, re.IGNORECASE):
            return "k8s"
        return "default"

    def infer_external_business(target_name, service_key):
        """从标准外部节点名的第 3、4 段提取业务名。"""
        if service_key not in {"k8s", "maas"}:
            return ""
        parts = [part.strip() for part in str(target_name).split("-") if part.strip()]
        service_tokens = {"k8s", "kubernetes", "maas"}
        if len(parts) < 5 or not any(
                part.casefold() in service_tokens for part in parts[4:]):
            return ""
        business_parts = parts[2:4]
        if any(part.casefold() in service_tokens for part in business_parts):
            return ""
        return "-".join(business_parts)

    # 外部目标也必须先进入分组规划，否则它们会绕过“前 5 个 + 摘要”规则。
    real_entries = list(entries)
    real_name_map = build_name_map(real_entries)
    external_specs = {}

    def register_external_target(source_entry, target_name):
        key = name_key(target_name)
        inferred_key = infer_external_service_key(source_entry, target_name)
        if inferred_key == "default":
            return
        spec = external_specs.setdefault(key, {
            "name": target_name,
            "service_key": "default",
        })
        spec["service_key"] = inferred_key

    for source_entry in real_entries:
        for target_name in split_target_names(source_entry["record"]["targets"]):
            if _name_key(target_name) in filtered_names:
                continue
            if resolve_target_entry(source_entry, target_name,
                                    real_entries, real_name_map):
                continue
            register_external_target(source_entry, target_name)

        # 推断下游不参与外部节点注册；服务类型-业务名仅作为现有服务组引用。
    for spec in external_specs.values():
        inferred_business = infer_external_business(
            spec["name"], spec["service_key"]
        )
        external_record = {
            "name": spec["name"],
            "type": spec["service_key"],
            "resource_id": "",
            "enterprise_id": "",
            "region": "",
            "group": "",
            "business": inferred_business,
            "desc": "（外部/未列出服务）",
            "spec": "",
            "source": "",
            "targets": "",
            "inferred_targets": "",
            "reason": "",
        }
        entries.append(make_entry(
            external_record, len(entries), is_external=True,
            forced_service_key=spec["service_key"],
        ))

    # 外部节点同样遵循调用方显式设置的业务关键字过滤。
    if BUSINESS_FILTER_KEYWORD:
        keyword = BUSINESS_FILTER_KEYWORD.casefold()
        entries = [
            entry for entry in entries
            if keyword in (entry["data"].get("business") or "").casefold()
        ]

    name_entry_map = build_name_map(entries)
    service_groups = defaultdict(list)
    for entry in entries:
        service_groups[entry["group_key"]].append(entry)

    visible_indexes = set()
    group_info = {}
    for group_index, (group_key, group_entries) in enumerate(service_groups.items()):
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        # CCE / CCE_Deployment / ECS 聚合为单个缩略节点，不逐个渲染；
        # 其他服务全部展示。
        if svc_key in AGGREGATE_SERVICE_KEYS:
            visible_entries = []
            hidden_entries = []
        else:
            visible_entries = list(group_entries)
            hidden_entries = []
        visible_indexes.update(entry["index"] for entry in visible_entries)
        first_id = (
            visible_entries[0]["data"]["id"]
            if visible_entries else group_entries[0]["data"]["id"]
        )
        group_info[group_key] = {
            "entries": group_entries,
            "visible": visible_entries,
            "hidden": hidden_entries,
            "first_id": first_id,
            "summary_id": "",
            "group_index": group_index,
        }

    for entry in entries:
        if entry["index"] in visible_indexes:
            nodes.append({"group": "nodes", "data": entry["data"]})

    def make_summary_resource(entry):
        r = entry["record"]
        d = entry["data"]
        return {
            "name": d.get("name", ""),
            "type": (r.get("type") or d.get("type") or "").upper(),
            "service_name": get_service_display_name(d.get("service_key", "default")),
            "service_key": d.get("service_key", ""),
            "resource_id": d.get("resource_id", ""),
            "spec": d.get("spec", ""),
            "project": d.get("enterprise_id", ""),
            "enterprise_project": (
                d.get("enterprise_project") or d.get("enterprise_id", "")
            ),
            "business": d.get("business", ""),
            "group": d.get("group_label", ""),
            "source": d.get("source") or d.get("region", ""),
            "region": d.get("region", ""),
            "desc": d.get("desc", ""),
            "downstream": split_target_names(r.get("targets", "")),
            "inferred_downstream": split_target_names(r.get("inferred_targets", "")),
            "downstream_text": r.get("targets", "") or "",
        }

    # 摘要节点：隐藏资源按“资源分组”值拆分，每个资源分组一个摘要节点，
    # 以便体现摘要节点归属的资源分组；摘要节点承接组内隐藏资源的连线。
    summary_counter = [0]
    for group_key, info in group_info.items():
        hidden_entries = info["hidden"]
        if not hidden_entries:
            continue
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        sample_data = info["entries"][0]["data"]
        bg, border, _ = SERVICE_STYLE.get(svc_key, SERVICE_STYLE["default"])
        hidden_by_group = defaultdict(list)
        for entry in hidden_entries:
            gkey = (
                (entry["data"].get("group_label") or "").strip().casefold()
                or "__none__"
            )
            hidden_by_group[gkey].append(entry)
        summary_ids = {}
        for gkey, g_entries in hidden_by_group.items():
            glabel = (
                g_entries[0]["data"].get("group_label", "")
                or "(未设置资源分组)"
            )
            summary_id = f"summary_{summary_counter[0]}"
            summary_counter[0] += 1
            summary_ids[gkey] = summary_id
            nodes.append({"group": "nodes", "data": {
                "id":              summary_id,
                "name":            f"...+{len(g_entries)}",
                "type":            "__summary__",
                "service_key":     svc_key,
                "service_name":    get_service_display_name(svc_key),
                "layer_key":       layer_key,
                "layer_name":      get_topology_layer_name(layer_key),
                "business":        sample_data["business"],
                "business_key":    biz_key,
                "is_virtual_business": sample_data["is_virtual_business"],
                "is_grouped_resource": 1,
                "collapsed_count": len(g_entries),
                "resource_total":  len(g_entries),
                "summary_resources": [
                    make_summary_resource(entry) for entry in g_entries
                ],
                "spec":            "",
                "desc":            "其余同类资源已折叠显示",
                "resource_id":     "",
                "enterprise_id":   "",
                "enterprise_project": "",
                "cce_business_prefix": sample_data.get("cce_business_prefix", ""),
                "region":          "",
                "group_label":     glabel,
                "reason":          "",
                "bg_color":        bg,
                "border_color":    border,
                "shape":           "roundrectangle",
                "is_container":    0,
                "is_summary":      1,
            }})
        info["summary_ids"] = summary_ids
        info["summary_id"] = next(iter(summary_ids.values()), "")

    # ── 2. 创建业务容器节点，为叶节点设置 parent ─────────────────────────

    business_map = {}   # business_key → biz_id
    business_meta = {}
    business_totals = defaultdict(int)
    for entry in entries:
        business_totals[entry["data"]["business_key"]] += 1
    _biz_counter = 0
    # 从全部 entries 建立业务容器映射。
    for entry in entries:
        biz_key = entry["data"].get("business_key", "").strip()
        if not biz_key:
            continue
        if biz_key not in business_map:
            business_map[biz_key] = f"biz_{_biz_counter}"
            business_meta[biz_key] = {
                "name": entry["data"].get("business", ""),
                "is_virtual": entry["data"].get("is_virtual_business", 0),
            }
            _biz_counter += 1
    for n in nodes:
        biz_key = n["data"].get("business_key", "").strip()
        if biz_key and biz_key in business_map:
            n["data"]["parent"] = business_map[biz_key]
            n["data"]["is_grouped_resource"] = 1

    biz_containers = []
    for i, (biz_key, biz_id) in enumerate(business_map.items()):
        meta = business_meta[biz_key]
        bg, border, text = BUSINESS_PALETTE[i % len(BUSINESS_PALETTE)]
        biz_containers.append({"group": "nodes", "data": {
            "id":           biz_id,
            "name":         meta["name"],
            "min_width":    container_min_width(meta["name"], font_size=26),
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
    # ── 2.6 在业务容器内按固定架构层和服务类型创建子容器 ─────────────────
    layer_containers = []
    service_containers = []
    node_by_id = {node["data"]["id"]: node for node in nodes}

    layers_by_business = defaultdict(set)
    layer_totals = defaultdict(int)
    for group_key, info in service_groups.items():
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        if business_meta[biz_key]["is_virtual"]:
            continue
        if svc_key in EXTRACTED_FROM_LAYER_TYPES:
            continue
        layers_by_business[biz_key].add(layer_key)
        layer_totals[(biz_key, layer_key)] += len(info)

    layer_map = {}
    for biz_key, layer_keys in layers_by_business.items():
        biz_id = business_map[biz_key]
        for layer_key in sorted(
                layer_keys,
                key=lambda key: TOPOLOGY_LAYER_ORDER.get(key, 999)):
            layer_id = f"layer_{biz_id}_{safe_id(layer_key)}"
            layer_map[(biz_key, layer_key)] = layer_id
            layer_containers.append({"group": "nodes", "data": {
                "id":              layer_id,
                "name":            get_topology_layer_name(layer_key),
                "min_width":       container_min_width(
                    get_topology_layer_name(layer_key), font_size=22
                ),
                "type":            "__layer__",
                "layer_key":       layer_key,
                "layer_name":      get_topology_layer_name(layer_key),
                "resource_total":  layer_totals[(biz_key, layer_key)],
                "is_container":    1,
                "parent":          biz_id,
                "bg_color":        "#F7F9FC",
                "border_color":    "#5D6D7E",
                "shape":           "roundrectangle",
            }})

    # 服务容器按 (业务, 层, 服务) 去重创建；CCE/CCE_Deployment 与其他非 ECS
    # 资源一样全部展示，不再聚合或缩略。
    service_agg = {}
    for group_key, info in service_groups.items():
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        sc_key = (biz_key, layer_key, svc_key)
        if not business_meta[biz_key]["is_virtual"]:
            agg = service_agg.setdefault(sc_key, {
                "resource_total": 0, "visible_count": 0, "collapsed_count": 0,
            })
            agg["resource_total"] += len(info)
            agg["visible_count"] += len(group_info[group_key]["visible"])
            agg["collapsed_count"] += len(group_info[group_key]["hidden"])

    # 按“业务 × 资源分组”聚合 CCE / CCE_Deployment / ECS：
    # 每个资源分组生成一个缩略节点。
    agg_meta = {}
    for group_key, info in service_groups.items():
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        if svc_key not in AGGREGATE_SERVICE_KEYS:
            continue
        for entry in info:
            glabel = (
                (entry["data"].get("group_label") or "").strip()
                or "(未设置资源分组)"
            )
            gkey = glabel.casefold()
            meta = agg_meta.setdefault((biz_key, gkey, layer_key), {
                "count": 0, "resources": [], "layer_key": layer_key,
                "label": glabel, "svc_keys": set(),
            })
            meta["count"] += 1
            meta["resources"].append(make_summary_resource(entry))
            meta["svc_keys"].add(svc_key)

    agg_compute_ids = {}
    agg_compute_nodes = []
    agg_count_nodes = []
    service_container_ids = {}
    group_containers = []
    group_container_counter = [0]
    for group_key, info in service_groups.items():
        biz_key, layer_key, svc_key, prefix_key, project_key = split_group_key(
            group_key
        )
        biz_id = business_map[biz_key]
        group_meta = group_info[group_key]
        rendered_ids = {entry["data"]["id"] for entry in group_meta["visible"]}
        for sid in group_meta.get("summary_ids", {}).values():
            rendered_ids.add(sid)

        is_virtual = business_meta[biz_key]["is_virtual"]
        if svc_key in AGGREGATE_SERVICE_KEYS:
            # Aggregated ECS/CCE resources still get a visible service frame so
            # the architecture diagram shows their ownership explicitly.
            if not is_virtual:
                sc_key = (biz_key, layer_key, svc_key)
                if sc_key not in service_container_ids:
                    sc_id = f"sc_{biz_id}_{safe_id(layer_key)}_{safe_id(svc_key)}"
                    service_container_ids[sc_key] = sc_id
                    agg = service_agg[sc_key]
                    service_containers.append({"group": "nodes", "data": {
                        "id": sc_id,
                        "name": get_service_display_name(svc_key),
                        "min_width": container_min_width(
                            get_service_display_name(svc_key), font_size=22
                        ),
                        "type": "__service__",
                        "service_key": svc_key,
                        "layer_key": layer_key,
                        "layer_name": get_topology_layer_name(layer_key),
                        "resource_total": agg["resource_total"],
                        "visible_count": agg["visible_count"],
                        "collapsed_count": agg["collapsed_count"],
                        "is_container": 1,
                        "parent": (biz_id if svc_key in EXTRACTED_FROM_LAYER_TYPES
                                   else layer_map[(biz_key, layer_key)]),
                        "bg_color": "#F2F3F4",
                        "border_color": "#7F8C8D",
                        "shape": "roundrectangle",
                    }})
            for entry in info:
                glabel = (
                    (entry["data"].get("group_label") or "").strip()
                    or "(未设置资源分组)"
                )
                agg_key = (biz_key, glabel.casefold(), layer_key)
                if agg_key in agg_compute_ids:
                    continue
                meta = agg_meta[agg_key]
                primary_svc = (
                    "cce" if "cce" in meta["svc_keys"]
                    else next(iter(meta["svc_keys"]))
                )
                agg_id = f"aggc_{biz_id}_{len(agg_compute_ids)}"
                agg_compute_ids[agg_key] = agg_id
                agg_parent = (
                    biz_id if is_virtual or meta["layer_key"] in EXTRACTED_FROM_LAYER_TYPES
                    else service_container_ids[(biz_key, meta["layer_key"], primary_svc)]
                )
                service_label = "/".join(
                    get_service_display_name(s) for s in sorted(meta["svc_keys"])
                )
                agg_compute_nodes.append({"group": "nodes", "data": {
                    "id":              agg_id,
                    "name":            f"...+{meta['count']}",
                    "display_label":   f"{meta['label']}\n...+{meta['count']}",
                    "type":            "__agg_compute__",
                    "service_key":     primary_svc,
                    "service_name":    service_label,
                    "layer_key":       meta["layer_key"],
                    "layer_name":      get_topology_layer_name(meta["layer_key"]),
                    "business":        business_meta[biz_key]["name"],
                    "business_key":    biz_key,
                    "group_label":     meta["label"],
                    "resource_total":  meta["count"],
                    "summary_resources": meta["resources"],
                    "is_container":    0,
                    "is_summary":      0,
                    "parent":          agg_parent,
                }})
                agg_count_nodes.append({"group": "nodes", "data": {
                    "id":              f"aggcnt_{len(agg_count_nodes)}",
                    "name":            str(meta["count"]),
                    "type":            "__agg_count__",
                    "anchor":          agg_id,
                    "resource_total":  meta["count"],
                    "parent":          agg_parent,
                    "is_container":    0,
                    "is_summary":      0,
                }})
            continue

        if is_virtual:
            continue

        layer_id = layer_map.get((biz_key, layer_key))
        sc_key = (biz_key, layer_key, svc_key)
        sc_id = None
        # 虚拟业务框本身已经代表服务类型，不再重复嵌套同名服务框。
        if not is_virtual:
            if sc_key not in service_container_ids:
                sc_id = f"sc_{biz_id}_{safe_id(layer_key)}_{safe_id(svc_key)}"
                service_container_ids[sc_key] = sc_id
                agg = service_agg[sc_key]
                service_containers.append({"group": "nodes", "data": {
                    "id":              sc_id,
                    "name":            get_service_display_name(svc_key),
                    "min_width":       container_min_width(
                        get_service_display_name(svc_key), font_size=22
                    ),
                    "type":            "__service__",
                    "service_key":     svc_key,
                    "layer_key":       layer_key,
                    "layer_name":      get_topology_layer_name(layer_key),
                    "resource_total":  agg["resource_total"],
                    "visible_count":   agg["visible_count"],
                    "collapsed_count": agg["collapsed_count"],
                    "is_container":    1,
                    "parent":          (biz_id if svc_key in EXTRACTED_FROM_LAYER_TYPES
                                        else layer_id),
                    "bg_color":        "#F2F3F4",
                    "border_color":    "#7F8C8D",
                    "shape":           "roundrectangle",
                }})
            sc_id = service_container_ids[sc_key]

        if sc_id is not None:
            summary_ids = group_meta.get("summary_ids", {})
            summary_nodes = [
                node_by_id.get(sid) for sid in summary_ids.values()
            ]
            summary_nodes = [n for n in summary_nodes if n]

            leaf_ids = [
                node_id for node_id in rendered_ids
                if node_id not in set(summary_ids.values())
            ]

            # 按“资源分组”字段细分：同一服务组内存在多个不同资源分组值时，
            # 用 __group__ 容器把相同值的节点（含摘要节点）框在一起，
            # 框名即资源分组名。
            def node_group_key(node):
                return (
                    (node["data"].get("group_label") or "").strip().casefold()
                    or "__none__"
                )

            distinct_groups = set()
            for node_id in leaf_ids:
                node = node_by_id.get(node_id)
                if node:
                    distinct_groups.add(node_group_key(node))
            for snode in summary_nodes:
                distinct_groups.add(node_group_key(snode))

            if len(distinct_groups) >= 2:
                grouped = defaultdict(list)
                for node_id in leaf_ids:
                    node = node_by_id.get(node_id)
                    if node:
                        grouped[node_group_key(node)].append(node_id)
                group_box_ids = {}
                for gkey in distinct_groups:
                    sample_node = None
                    if grouped.get(gkey):
                        sample_node = node_by_id[grouped[gkey][0]]
                    else:
                        sample_node = next(
                            (n for n in summary_nodes
                             if node_group_key(n) == gkey), None
                        )
                    glabel = (
                        sample_node["data"].get("group_label", "")
                        or "(未设置资源分组)"
                    )
                    gid = f"grp_{group_container_counter[0]}"
                    group_container_counter[0] += 1
                    group_box_ids[gkey] = gid
                    group_containers.append({"group": "nodes", "data": {
                        "id":              gid,
                        "name":            glabel,
                        "min_width":       container_min_width(
                            glabel, font_size=22
                        ),
                        "type":            "__group__",
                        "service_key":     svc_key,
                        "layer_key":       layer_key,
                        "layer_name":      get_topology_layer_name(layer_key),
                        "resource_total":  len(grouped.get(gkey, []))
                                         + sum(
                                             1 for n in summary_nodes
                                             if node_group_key(n) == gkey
                                         ),
                        "is_container":    1,
                        "parent":          sc_id,
                        "bg_color":        "#FDFEFE",
                        "border_color":    "#AAB7B8",
                        "shape":           "roundrectangle",
                    }})
                for gkey, node_ids in grouped.items():
                    gid = group_box_ids[gkey]
                    for node_id in node_ids:
                        node = node_by_id.get(node_id)
                        if node:
                            node["data"]["parent"] = gid
                            node["data"]["group_container"] = gid
                            node["data"]["service_container"] = sc_id
                            node["data"]["layer_container"] = layer_id
                for snode in summary_nodes:
                    gid = group_box_ids.get(node_group_key(snode), sc_id)
                    snode["data"]["parent"] = gid
                    snode["data"]["group_container"] = gid
                    snode["data"]["service_container"] = sc_id
                    snode["data"]["layer_container"] = layer_id
            else:
                for node_id in leaf_ids:
                    node = node_by_id.get(node_id)
                    if node:
                        node["data"]["parent"] = sc_id
                        node["data"]["service_container"] = sc_id
                        node["data"]["layer_container"] = layer_id
                for snode in summary_nodes:
                    snode["data"]["parent"] = sc_id
                    snode["data"]["service_container"] = sc_id
                    snode["data"]["layer_container"] = layer_id

    nodes = (biz_containers + layer_containers + service_containers
             + group_containers + agg_compute_nodes + agg_count_nodes + nodes)

    # ── 3. 创建边 ─────────────────────────────────────────────────────────

    def internal_endpoint(entry):
        """同业务边保留可见节点；隐藏节点由摘要节点承接。"""
        if entry["data"]["service_key"] in AGGREGATE_SERVICE_KEYS:
            gkey = (
                (entry["data"].get("group_label") or "").strip().casefold()
                or "(未设置资源分组)".casefold()
            )
            return agg_compute_ids.get(
                (entry["data"]["business_key"], gkey,
                 entry["data"].get("layer_key", "compute_app")), ""
            )
        if entry["index"] in visible_indexes:
            return entry["data"]["id"]
        gkey = (
            (entry["data"].get("group_label") or "").strip().casefold()
            or "__none__"
        )
        sid = group_info[entry["group_key"]].get("summary_ids", {}).get(gkey)
        return sid or group_info[entry["group_key"]]["summary_id"]

    def representative_endpoint(entry):
        """跨业务边统一落到对应服务组的第一个真实资源节点。"""
        if entry["data"]["service_key"] in AGGREGATE_SERVICE_KEYS:
            gkey = (
                (entry["data"].get("group_label") or "").strip().casefold()
                or "(未设置资源分组)".casefold()
            )
            return agg_compute_ids.get(
                (entry["data"]["business_key"], gkey,
                 entry["data"].get("layer_key", "compute_app")), ""
            )
        if entry["group_key"]:
            return group_info[entry["group_key"]]["first_id"]
        return entry["data"]["id"]

    def add_edge(src_id, tgt_id, source_name, target_name, cross_biz,
                 relation="调用", inferred=False):
        if src_id == tgt_id:
            return
        key = (src_id, tgt_id, inferred)
        if key in edge_lookup:
            edge_data = edge_lookup[key]
            edge_data["call_count"] += 1
            relation_label = relation + ("（推断）" if inferred else "")
            edge_data["relation"] = (
                f"{relation_label} ×{edge_data['call_count']}"
            )
            return
        color = RELATION_COLORS.get(relation, RELATION_COLORS["default"])
        if inferred:
            color = "#8E44AD"
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

    def connect_entries(source_entry, target_entry, source_name, target_name,
                        inferred=False):
        source_biz = source_entry["data"]["business_key"]
        target_biz = target_entry["data"]["business_key"]
        cross_biz = bool(source_biz and target_biz and source_biz != target_biz)
        if cross_biz:
            src_id = representative_endpoint(source_entry)
            tgt_id = representative_endpoint(target_entry)
        else:
            src_id = internal_endpoint(source_entry)
            tgt_id = internal_endpoint(target_entry)
        add_edge(
            src_id, tgt_id, source_name, target_name, cross_biz,
            inferred=inferred,
        )

    for source_entry in real_entries:
        r = source_entry["record"]

        def connect_target(target_entry, target_name):
            connect_entries(
                source_entry, target_entry, r["name"], target_name,
            )

        # 确认下游
        if r["targets"]:
            for t in split_target_names(r["targets"]):
                target_entry = resolve_target_entry(
                    source_entry, t, entries, name_entry_map
                )
                if not target_entry:
                    continue
                connect_target(target_entry, t)

    # ── 四层架构固定层级关系 ──────────────────────────────────────────────
    # 同一业务内，按固定层级顺序（接入 → 网络负载 → 计算容器 → 中间件数据
    # 存储）连接相邻的层容器，生成独立且更粗的固定关系箭头。
    layer_edge_count = [0]
    layer_order = [key for key, _ in TOPOLOGY_LAYERS]
    for biz_key, biz_id in business_map.items():
        present_layers = [
            layer_key for layer_key in layer_order
            if (biz_key, layer_key) in layer_map
        ]
        for src_key, tgt_key in zip(present_layers, present_layers[1:]):
            edges.append({"group": "edges", "data": {
                "id":            f"elayer_{layer_edge_count[0]}",
                "source":        layer_map[(biz_key, src_key)],
                "target":        layer_map[(biz_key, tgt_key)],
                "source_name":   get_topology_layer_name(src_key),
                "target_name":   get_topology_layer_name(tgt_key),
                "relation":      "固定层级关系",
                "layer_relation": 1,
                "inferred":      0,
                "cross_business": 0,
                "color":         "#5D6D7E",
            }})
            layer_edge_count[0] += 1

    return nodes + edges


# ── BDAT 分组层次布局位置计算 ─────────────────────────────────────────────────

def compute_bdat_positions(elements):
    """
    按业务分组（BDAT）计算节点的预设坐标：
    - 各业务组横向排列成网格
    - 组内以服务类型为不可拆分的连续矩形块
    - 服务块按依赖关系纵向排序，块内每行最多 MAX_NODES_PER_ROW 个节点
    - 同一服务块内按“资源分组”再拆成独立小方块，块与块之间留有间隔
    返回 {node_id: {'x': float, 'y': float}}
    """
    NODE_W        = 190
    ROW_H         = 120
    SERVICE_GAP   = 120
    GROUP_GAP     = 170
    SERVICES_PER_ROW = 100
    SERVICE_ROW_GAP = 160
    LAYER_GAP     = 260
    MAX_ROW_NODES = MAX_NODES_PER_ROW
    GROUP_PAD     = 110
    GAP_X         = 320
    GAP_Y         = 300
    MAX_COLS      = 4

    leaf_nodes = {}
    badge_nodes = {}
    edges_list = []
    for e in elements:
        if e["group"] == "nodes":
            data = e["data"]
            if data.get("type") == "__agg_count__":
                badge_nodes[data["id"]] = data
            elif not data.get("is_container"):
                leaf_nodes[data["id"]] = data
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
            layer_key = leaf_nodes[nid].get("layer_key") or get_topology_layer(
                service_key, leaf_nodes[nid].get("name", ""),
                leaf_nodes[nid].get("type", "")
            )
            slot_key = (layer_key, service_key)
            services[slot_key].append(nid)
            service_first_index.setdefault(slot_key, index)

        # 将节点级依赖提升为同层服务级依赖，用于排列完整服务块。
        in_deg = {slot_key: 0 for slot_key in services}
        adjacency = defaultdict(set)
        for src, tgt in edges_list:
            if src not in nodes_in_group or tgt not in nodes_in_group:
                continue
            src_service = leaf_nodes[src].get("service_key", "default")
            tgt_service = leaf_nodes[tgt].get("service_key", "default")
            src_layer = leaf_nodes[src].get("layer_key") or get_topology_layer(
                src_service, leaf_nodes[src].get("name", ""),
                leaf_nodes[src].get("type", "")
            )
            tgt_layer = leaf_nodes[tgt].get("layer_key") or get_topology_layer(
                tgt_service, leaf_nodes[tgt].get("name", ""),
                leaf_nodes[tgt].get("type", "")
            )
            src_slot = (src_layer, src_service)
            tgt_slot = (tgt_layer, tgt_service)
            if src_slot == tgt_slot or src_layer != tgt_layer:
                continue
            if tgt_slot in adjacency[src_slot]:
                continue
            adjacency[src_slot].add(tgt_slot)
            in_deg[tgt_slot] += 1

        layer_service_orders = {}
        ordered_layer_keys = [
            key for key, _ in TOPOLOGY_LAYERS
            if any(slot[0] == key for slot in services)
        ]
        ordered_layer_keys.extend(sorted(
            {slot[0] for slot in services if slot[0] not in TOPOLOGY_LAYER_ORDER},
            key=lambda key: min(service_first_index[slot]
                                for slot in services if slot[0] == key),
        ))
        for layer_key in ordered_layer_keys:
            layer_slots = [slot for slot in services if slot[0] == layer_key]
            ordered_slots = []
            ready = sorted(
                (slot for slot in layer_slots if in_deg[slot] == 0),
                key=lambda slot: service_first_index[slot],
            )
            while ready:
                slot_key = ready.pop(0)
                ordered_slots.append(slot_key)
                for target_key in sorted(
                        adjacency[slot_key],
                        key=lambda slot: service_first_index[slot]):
                    in_deg[target_key] -= 1
                    if in_deg[target_key] == 0:
                        ready.append(target_key)
                        ready.sort(key=lambda slot: service_first_index[slot])

            # Keep cyclic services in first-seen order at the end of the layer.
            ordered_slots.extend(sorted(
                (slot for slot in layer_slots if slot not in ordered_slots),
                key=lambda slot: service_first_index[slot],
            ))
            layer_service_orders[layer_key] = ordered_slots

        service_blocks = {}
        layer_sizes = {}
        layer_rows = {}
        layer_row_offsets = {}
        max_layer_w = NODE_W
        for layer_key, ordered_slots in layer_service_orders.items():
            has_in_layer_chain = any(
                target_key in adjacency[source_key]
                for source_key in ordered_slots
                for target_key in ordered_slots
            )
            if has_in_layer_chain:
                # A service-level dependency chain reads more clearly as a
                # vertical pipeline: source service at the top, downstream
                # service below it. Independent slots keep the compact grid.
                slot_rows = [[slot] for slot in ordered_slots]
            elif layer_key == "data_middleware_storage":
                row_count = min(DATA_LAYER_SERVICE_ROWS, len(ordered_slots))
                base_size, remainder = divmod(len(ordered_slots), row_count)
                slot_rows = []
                cursor = 0
                for row_index in range(row_count):
                    current_size = base_size + (1 if row_index < remainder else 0)
                    slot_rows.append(
                        ordered_slots[cursor:cursor + current_size]
                    )
                    cursor += current_size
            else:
                slot_rows = [
                    ordered_slots[i:i + SERVICES_PER_ROW]
                    for i in range(0, len(ordered_slots), SERVICES_PER_ROW)
                ]
            layer_rows[layer_key] = slot_rows
            layer_w = 0
            layer_h = 0
            y_acc = 0
            row_offsets = {}
            for row_index, row_slots in enumerate(slot_rows):
                row_w = 0
                row_h = 0
                for slot_key in row_slots:
                    slot_nodes = services[slot_key]
                    by_group = defaultdict(list)
                    for nid in slot_nodes:
                        gkey = leaf_nodes[nid].get("group_container") or ""
                        by_group[gkey].append(nid)
                    group_blocks = []
                    block_w = 0
                    block_h = 0
                    for gkey, g_nodes in by_group.items():
                        g_rows = ((len(g_nodes) + MAX_ROW_NODES - 1)
                                  // MAX_ROW_NODES)
                        g_w = min(len(g_nodes), MAX_ROW_NODES) * NODE_W
                        g_h = g_rows * ROW_H
                        group_blocks.append({
                            "group_key": gkey,
                            "nodes": g_nodes,
                            "w": g_w,
                            "h": g_h,
                            "row_count": g_rows,
                        })
                        block_w += g_w
                        block_h = max(block_h, g_h)
                    if len(group_blocks) > 1:
                        block_w += GROUP_GAP * (len(group_blocks) - 1)
                    service_blocks[slot_key] = {
                        "w": block_w,
                        "h": block_h,
                        "group_blocks": group_blocks,
                    }
                    row_w += block_w
                    row_h = max(row_h, block_h)
                if len(row_slots) > 1:
                    row_w += SERVICE_GAP * (len(row_slots) - 1)
                row_offsets[row_index] = (y_acc, row_w, row_h)
                layer_w = max(layer_w, row_w)
                layer_h = max(layer_h, y_acc + row_h)
                y_acc += row_h + SERVICE_ROW_GAP
            layer_sizes[layer_key] = (layer_w, layer_h)
            layer_row_offsets[layer_key] = row_offsets
            max_layer_w = max(max_layer_w, layer_w)

        y_acc = 0
        layer_offsets = {}
        for index, layer_key in enumerate(layer_service_orders.keys()):
            if index > 0:
                y_acc += LAYER_GAP
            layer_offsets[layer_key] = y_acc
            y_acc += layer_sizes[layer_key][1]
        group_h = y_acc + 2 * GROUP_PAD

        group_services[biz_key] = services
        service_orders[biz_key] = layer_service_orders
        service_offsets[biz_key] = {
            "layers": layer_offsets,
            "blocks": service_blocks,
            "layer_sizes": layer_sizes,
            "layer_rows": layer_rows,
            "row_offsets": layer_row_offsets,
        }
        extracted_count = sum(
            len(nodes) for (layer_key, service_key), nodes in services.items()
            if service_key in EXTRACTED_FROM_LAYER_TYPES
        )
        if extracted_count:
            # Reserve room for the right-hand extracted-resource rail so
            # neighboring business frames cannot overlap it.
            max_layer_w += EXTRACTED_RIGHT_GAP + min(
                extracted_count, MAX_ROW_NODES
            ) * NODE_W
        group_sizes[biz_key] = (max_layer_w + 2 * GROUP_PAD, group_h)

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

        offset_info = service_offsets[biz_key]
        for layer_key in service_orders[biz_key].keys():
            layer_y = base_y + offset_info["layers"][layer_key]
            layer_w, layer_h = offset_info["layer_sizes"][layer_key]
            layer_x = base_x + (content_w - layer_w) / 2
            layer_rows = offset_info["layer_rows"][layer_key]
            row_offsets = offset_info["row_offsets"][layer_key]
            for row_index, row_slots in enumerate(layer_rows):
                row_y_off, row_w, row_h = row_offsets[row_index]
                row_x = layer_x + (layer_w - row_w) / 2
                service_x = row_x
                for slot_key in row_slots:
                    block = offset_info["blocks"][slot_key]
                    service_y = layer_y + row_y_off + (row_h - block["h"]) / 2
                    next_service_x = service_x + block["w"] + SERVICE_GAP
                    group_x = service_x
                    for gblock in block["group_blocks"]:
                        g_nodes = gblock["nodes"]
                        g_w = min(len(g_nodes), MAX_ROW_NODES) * NODE_W
                        y_pos = service_y + (block["h"] - gblock["h"]) / 2
                        for g_row_index in range(gblock["row_count"]):
                            row_nodes = g_nodes[
                                g_row_index * MAX_ROW_NODES:
                                (g_row_index + 1) * MAX_ROW_NODES
                            ]
                            row_span = (len(row_nodes) - 1) * NODE_W
                            row_start_x = group_x + (g_w - row_span) / 2
                            y_row = y_pos + g_row_index * ROW_H
                            for column_index, node_id in enumerate(row_nodes):
                                positions[node_id] = {
                                    "x": round(
                                        row_start_x + column_index * NODE_W, 1
                                    ),
                                    "y": round(y_row, 1),
                                }
                        group_x += g_w + GROUP_GAP
                    service_x = next_service_x

        # VPC/CC/NAT/DCAAS are business-owned but intentionally outside the
        # architecture-layer containers. Move their resources into a clearly
        # separated right-hand rail after the normal layer layout is computed.
        extracted_ids = [
            nid for nid in biz_groups[biz_key]
            if leaf_nodes[nid].get("service_key") in EXTRACTED_FROM_LAYER_TYPES
        ]
        if extracted_ids:
            normal_points = [
                positions[nid] for nid in biz_groups[biz_key]
                if nid in positions and nid not in extracted_ids
            ]
            normal_max_x = max((point["x"] for point in normal_points), default=base_x)
            normal_min_y = min((point["y"] for point in normal_points), default=base_y)
            lane_x = normal_max_x + EXTRACTED_RIGHT_GAP
            lane_y = normal_min_y
            lane_cursor = lane_x
            lane_slots = []
            for service_key in dict.fromkeys(
                    leaf_nodes[nid].get("service_key") for nid in extracted_ids):
                slot_nodes = [
                    nid for nid in extracted_ids
                    if leaf_nodes[nid].get("service_key") == service_key
                ]
                lane_slots.append(slot_nodes)
            for slot_nodes in lane_slots:
                for index, node_id in enumerate(slot_nodes):
                    row, column = divmod(index, MAX_ROW_NODES)
                    positions[node_id] = {
                        "x": round(lane_cursor + column * NODE_W, 1),
                        "y": round(lane_y + row * ROW_H, 1),
                    }
                lane_cursor += (
                    min(len(slot_nodes), MAX_ROW_NODES) * NODE_W
                    + SERVICE_GAP
                )

    # 数字徽标跟随其锚点（聚合缩略节点），显示在其右上方。
    for nid, data in badge_nodes.items():
        anchor_id = data.get("anchor")
        if anchor_id in positions:
            pos = positions[anchor_id]
            positions[nid] = {
                "x": round(pos["x"] + 64, 1),
                "y": round(pos["y"] - 70, 1),
            }

    # 给容器节点补坐标：自底向上按直接子节点中心定位，保证多级嵌套容器
    # （资源分组框 / 服务框 / 层框 / 业务框）之间互不重叠。
    node_data = {e["data"]["id"]: e["data"] for e in elements
                 if e["group"] == "nodes"}
    children_of = defaultdict(list)
    for nid, data in node_data.items():
        parent = data.get("parent")
        if parent:
            children_of[parent].append(nid)

    def container_depth(nid):
        depth = 0
        cur = node_data[nid].get("parent")
        while cur:
            depth += 1
            cur = node_data[cur].get("parent")
        return depth

    containers = sorted(
        (nid for nid, data in node_data.items() if data.get("is_container")),
        key=lambda nid: -container_depth(nid),
    )
    for cid in containers:
        child_points = [
            positions[kid] for kid in children_of.get(cid, [])
            if kid in positions
        ]
        if child_points:
            positions[cid] = {
                "x": round(sum(p["x"] for p in child_points)
                           / len(child_points), 1),
                "y": round(sum(p["y"] for p in child_points)
                           / len(child_points), 1),
            }

    return positions


# ── HTML 生成辅助 ─────────────────────────────────────────────────────────────

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


def _make_cytoscape_style(icon_data_uris=None):
    icon_data_uris = icon_data_uris or {}
    style = [
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
                "text-max-width": "170px",
                "text-wrap": "wrap",
            }
        },
        {
            "selector": "node[is_container = 1]",
            "style": {
                "label": "data(name)",
                "min-width": "data(min_width)",
                "text-valign": "top",
                "font-size": 22,
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
                "font-size": 30,
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
                "font-size": 24,
            }
        },
        {
            "selector": "node[type = '__layer__']",
            "style": {
                "border-color": "#5D6D7E",
                "border-width": 2,
                "border-style": "dashed",
                "background-color": "#F7F9FC",
                "background-opacity": 0.34,
                "font-size": 24,
                "font-weight": "bold",
                "color": "#34495E",
                "text-wrap": "none",
                "min-width": "300px",
                "text-valign": "top",
                "text-background-color": "#fff",
                "text-background-opacity": 0.92,
                "text-background-padding": "3px",
                "padding": "34px",
                "shape": "roundrectangle",
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
                "font-size": 12,
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
                "line-color": "data(color)",
                "target-arrow-color": "data(color)",
                "target-arrow-shape": "triangle",
                "curve-style": "bezier",
                "font-size": 11,
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
                "font-size": 24,
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
                "font-size": 13,
                "font-weight": "bold",
                "text-valign": "center",
                "text-halign": "center",
                "text-margin-y": 0,
                "text-background-opacity": 0,
                "text-max-width": "165px",
                "text-wrap": "wrap",
                "color": "#2C3E50",
            }
        },
        # Real resources use same-size PNG icons when matching assets exist.
        {
            "selector": "node[has_icon = 1][is_summary = 0]",
            "style": {
                "label": "data(display_label)",
                "width": ICON_SIZE,
                "height": ICON_SIZE,
                "shape": "rectangle",
                "background-fit": "contain",
                "background-clip": "none",
                "background-width": ICON_SIZE,
                "background-height": ICON_SIZE,
                "background-opacity": 1,
                "background-color": "#FFFFFF",
                "border-width": 0,
                "text-valign": "bottom",
                "text-halign": "center",
                "text-margin-y": 6,
                "text-background-color": "#fff",
                "text-background-opacity": 0.86,
                "text-background-padding": "2px",
                "text-max-width": "160px",
                "text-wrap": "wrap",
                "font-size": 13,
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
                "font-size": 14,
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
                "font-size": 28,
                "font-weight": "bold",
                "color": "data(text_color)",
                "text-background-color": "#fff",
                "text-background-opacity": 0.95,
                "text-background-padding": "4px",
                "text-wrap": "wrap",
                "text-max-width": "320px",
                "text-overflow-wrap": "anywhere",
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
        # 四层架构固定层级关系：独立、加粗箭头
        {
            "selector": "edge[layer_relation = 1]",
            "style": {
                "label": "",
                "width": 6,
                "line-color": "#5D6D7E",
                "line-style": "solid",
                "target-arrow-color": "#5D6D7E",
                "target-arrow-shape": "triangle",
                "target-arrow-fill": "filled",
                "curve-style": "bezier",
                "font-size": 14,
                "font-weight": "bold",
                "color": "#34495E",
                "text-background-color": "#fff",
                "text-background-opacity": 0.95,
                "text-background-padding": "3px",
                "opacity": 0.95,
            }
        },
    ]

    for service_key, icon_data_uri in sorted(icon_data_uris.items()):
        style.append({
            "selector": f"node[service_key = '{service_key}'][has_icon = 1]",
            "style": {"background-image": f'url("{icon_data_uri}")'},
        })

    # CCE/ECS 聚合缩略节点：优先走标准图标机制（保留图标原样），
    # 无图标时使用此回退样式；数字徽标为独立红色圆圈。
    style.append({
        "selector": "node[type = '__agg_compute__']",
        "style": {
            "label": "data(display_label)",
            "width": 64,
            "height": 64,
            "shape": "roundrectangle",
            "background-color": "#EBF5FB",
            "border-color": "#2C5F8A",
            "border-width": 2,
            "font-size": 13,
            "font-weight": "bold",
            "color": "#2C3E50",
            "text-valign": "bottom",
            "text-halign": "center",
            "text-margin-y": 6,
            "text-wrap": "wrap",
            "text-max-width": "170px",
        }
    })
    style.append({
        "selector": "node[type = '__agg_count__']",
        "style": {
            "label": "data(name)",
            "width": 36,
            "height": 36,
            "shape": "ellipse",
            "background-color": "#FFFFFF",
            "border-color": "#2C5F8A",
            "border-width": 3,
            "font-size": 16,
            "font-weight": "bold",
            "color": "#000000",
            "text-valign": "center",
            "text-halign": "center",
            "text-background-opacity": 0,
        }
    })

    style.extend([
        {"selector": "node[has_icon = 1]:selected", "style": {
            "border-color": "#FFD700",
            "border-width": 4,
        }},
        {"selector": "node[has_icon = 1].highlighted", "style": {
            "border-color": "#E74C3C",
            "border-width": 3,
        }},
    ])

    return style


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
        if e["group"] == "nodes":
            nid = e["data"]["id"]
            if nid in bdat_pos:
                e["data"]["bdat_x"] = bdat_pos[nid]["x"]
                e["data"]["bdat_y"] = bdat_pos[nid]["y"]

    # Mark real resources that have matching PNG icons.
    _load_service_icon_data_uris()
    icon_data_uris = {}
    for e in elements:
        if e.get("group") != "nodes":
            continue
        data = e.get("data", {})
        if data.get("is_container") or data.get("is_summary"):
            data["has_icon"] = 0
            continue
        service_key = str(data.get("service_key") or "").casefold()
        icon_data_uri = get_service_icon_data_uri(service_key)
        if icon_data_uri and service_key:
            icon_data_uris[service_key] = icon_data_uri
        data["has_icon"] = 1 if icon_data_uri else 0

    # ensure_ascii=True：将中文转为 \uXXXX 转义，避免内联 JS 中出现非 ASCII 字符导致的解析问题
    elements_json = json.dumps(elements, ensure_ascii=True)
    style_json    = json.dumps(_make_cytoscape_style(icon_data_uris), ensure_ascii=True)
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
                            and not e["data"].get("is_summary")
                            and e["data"].get("type") != "__agg_count__")
    agg_compute_nodes = [e for e in elements if e.get("group") == "nodes"
                         and e["data"].get("type") == "__agg_compute__"]
    summary_nodes = sum(1 for e in elements if e.get("group") == "nodes"
                        and e["data"].get("is_summary"))
    collapsed_resources = sum(e["data"].get("collapsed_count", 0) for e in elements
                              if e.get("group") == "nodes" and e["data"].get("is_summary"))
    agg_total = sum(
        node["data"].get("resource_total", 0) for node in agg_compute_nodes
    )
    total_resources = visible_resources + collapsed_resources + agg_total
    resource_stat = (
        f"{visible_resources + summary_nodes + len(agg_compute_nodes)} 可见 / "
        f"{total_resources} 资源"
        if (collapsed_resources or agg_compute_nodes)
        else f"{total_resources} 节点"
    )
    n_edges = sum(1 for e in elements if e.get("group") == "edges"
                  and not e["data"].get("layer_relation"))
    n_infer = sum(1 for e in elements if e.get("group") == "edges"
                  and e["data"].get("inferred") == 1
                  and not e["data"].get("layer_relation"))

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
#lockContainersBtn.locked{{background:rgba(231,76,60,.45);border-color:#ffb3aa}}
#lockContainersBtn{{font-size:14px;font-weight:700;padding:6px 12px}}
#layoutSelect{{padding:4px 8px;border-radius:4px;border:1px solid rgba(255,255,255,.3);
  background:rgba(255,255,255,.15);color:#fff;font-size:12px;cursor:pointer}}
#layoutSelect option{{background:#283593}}
.stats-bar{{font-size:11px;color:rgba(255,255,255,.75);margin-left:auto;
  display:flex;gap:10px;flex-shrink:0}}
.stat-item{{display:flex;align-items:center;gap:4px}}
.stat-dot{{width:8px;height:8px;border-radius:50%;display:inline-block}}
/* ── 主区域 ── */
#cy{{position:fixed;top:50px;left:0;right:280px;bottom:0;background:#fff}}
#clear-highlight-float{{position:fixed;display:none;z-index:40;cursor:pointer;
  padding:3px 8px;border:1px solid #c0392b;border-radius:12px;background:#E74C3C;
  color:#fff;font-size:11px;box-shadow:0 2px 6px rgba(0,0,0,.25)}}
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
.infer-line{{width:24px;height:2px;border-top:2px dashed #8E44AD}}
/* ── toast ── */
#toast{{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);
  background:#333;color:#fff;padding:8px 20px;border-radius:20px;
  font-size:12px;opacity:0;transition:opacity .3s;pointer-events:none;z-index:9999}}
/* ── 右侧滑出明细抽屉（折叠资源表格） ── */
#drawer-backdrop{{position:fixed;top:50px;left:0;right:0;bottom:0;
  background:rgba(0,0,0,.28);opacity:0;pointer-events:none;
  transition:opacity .25s;z-index:90}}
#drawer-backdrop.open{{opacity:1;pointer-events:auto}}
#summary-drawer{{position:fixed;top:50px;right:0;bottom:0;
  width:min(76vw,1250px);background:#fff;
  box-shadow:-4px 0 18px rgba(0,0,0,.28);
  transform:translateX(105%);transition:transform .28s ease;
  z-index:100;display:flex;flex-direction:column}}
#summary-drawer.open{{transform:translateX(0)}}
#drawer-header{{display:flex;justify-content:space-between;align-items:center;
  padding:10px 16px;background:#283593;color:#fff;font-weight:700;
  font-size:15px;flex-shrink:0}}
#drawer-body{{flex:1;overflow:auto;padding:14px 16px}}
.sum-table{{border-collapse:collapse;width:100%;table-layout:auto;
  white-space:nowrap;font-size:13px}}
.sum-table th,.sum-table td{{border:1px solid #d5d8dc;padding:8px 12px;
  text-align:left;vertical-align:top;overflow-wrap:normal}}
.sum-table thead th{{background:#E8EAF6;color:#283593;position:sticky;top:0;
  z-index:1}}
.sum-table tbody tr{{background:transparent}}
.sum-table tbody tr:hover{{background:#F4F6F7}}
.sum-pager{{display:flex;gap:10px;align-items:center;padding:12px 2px;
  font-size:13px;color:#555}}
.sum-pager .pg-btn{{cursor:pointer;padding:4px 14px;border:1px solid #9AA7B8;
  border-radius:4px;background:#fff;color:#283593}}
.sum-pager .pg-btn:hover{{background:#E8EAF6}}
.sum-pager .pg-btn:disabled{{opacity:.4;cursor:not-allowed}}
.sum-hint{{font-size:12px;color:#888;margin:0 0 8px 2px}}
.sum-info{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px}}
.sum-info-item{{display:flex;flex-direction:column;min-width:130px;
  padding:8px 14px;border:1px solid #d5d8dc;border-radius:6px;
  background:#FAFBFC}}
.sum-info-label{{font-size:11px;color:#888;text-transform:uppercase;
  letter-spacing:.5px;margin-bottom:4px}}
.sum-info-value{{font-size:15px;font-weight:700;color:#2C3E50;
  word-break:break-all}}
/* ── 节点编辑器 ── */
#editor-modal{{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;
  align-items:center;justify-content:center;z-index:200}}
#editor-modal.open{{display:flex}}
#editor-card{{width:min(560px,calc(100vw - 32px));max-height:calc(100vh - 48px);overflow:auto;
  background:#fff;border-radius:8px;box-shadow:0 12px 40px rgba(0,0,0,.28);padding:18px}}
#editor-card h2{{font-size:17px;color:#283593;margin-bottom:14px}}
.edit-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px 14px}}
.edit-field{{display:flex;flex-direction:column;gap:4px}}
.edit-field.full{{grid-column:1 / -1}}
.edit-field label{{font-size:11px;color:#64748b}}
.edit-field input,.edit-field textarea{{border:1px solid #cbd5e1;border-radius:4px;padding:7px 8px;font:12px inherit;color:#1e293b}}
.edit-field textarea{{min-height:64px;resize:vertical}}
.editor-actions{{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}}
.editor-actions button{{cursor:pointer;padding:6px 14px;border-radius:4px;border:1px solid #cbd5e1;background:#fff;color:#334155}}
.editor-actions .primary{{border-color:#283593;background:#283593;color:#fff}}
#addResourceBtn{{font-size:14px;font-weight:700;background:rgba(52,152,219,.45);padding:6px 12px}}
.edit-hint{{font-size:11px;color:#64748b;margin-top:8px;line-height:1.4}}
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
  <button id="lockContainersBtn" class="tb-btn" onclick="toggleContainerLock()" title="锁定或解锁所有业务/层/服务框">&#128274; 锁定容器:关</button>
  <button id="addResourceBtn" class="tb-btn" onclick="openAddNodeEditor()" title="添加资源并生成新节点">&#43; 新增资源</button>
  <button class="tb-btn" onclick="downloadEditedCsv()" title="选择目录导出 CSV 和自适应 Excel 编辑表">&#128190; 导出编辑表</button>
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
    <span class="stat-item"><span class="stat-dot" style="background:#8E44AD"></span>{n_infer} 推断边</span>
  </div>
</div>

<div id="cy"></div>
<button id="clear-highlight-float" onclick="clearHighlight()">&#10005; 取消高亮</button>
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
<div id="drawer-backdrop"></div>
<div id="summary-drawer">
  <div id="drawer-header">
    <span id="drawer-title">折叠资源明细</span>
    <button class="tb-btn" onclick="closeSummaryDrawer()">&#10005; 关闭</button>
  </div>
  <div id="drawer-body"></div>
</div>
<div id="editor-modal" role="dialog" aria-modal="true" aria-labelledby="editor-title">
  <div id="editor-card">
    <h2 id="editor-title">编辑节点</h2>
    <form id="node-editor-form" onsubmit="saveNodeEdit(event)">
      <div id="editor-fields" class="edit-grid"></div>
      <p class="edit-hint">保存后会立即更新节点标签、详情和当前拓扑；“导出编辑表”会生成 CSV 和带自适应列宽的 Excel 兼容表。</p>
      <div class="editor-actions">
        <button type="button" onclick="closeNodeEditor()">取消</button>
        <button type="submit" class="primary">保存并重新渲染</button>
      </div>
    </form>
  </div>
</div>

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
  window._activeNodeId = evt.target.id();
  showFloatingClear(evt.target);
  var bg = d.bg_color || '#888';
  if (d.type === '__region__' || d.type === '__group__' || d.type === '__business__' || d.type === '__layer__' || d.type === '__service__') {{
    showContainerDetail(d);
  }} else if (d.type === '__agg_compute__' || d.type === '__agg_count__') {{
    showSummaryDetail(d);
  }} else if (d.is_summary === 1) {{
    showSummaryDetail(d);
  }} else {{
    showNodeDetail(d, bg);
  }}
  cy.elements().addClass('faded');
  var isContainerType = d.type === '__region__' || d.type === '__group__' ||
                        d.type === '__business__' || d.type === '__layer__' ||
                        d.type === '__service__';
  if (!isContainerType) {{
    // 沿调用方向展开，只保留与当前节点有直接/间接调用关系的节点：
    // 上游 = 能调用当前节点的所有节点（沿边反向遍历）；
    // 下游 = 当前节点能调用的所有节点（沿边正向遍历）。
    // 只高亮这些节点之间处于链路上的调用线，不包含仅与旁支相连的节点。
    var downVisited = evt.target.collection();
    var frontier = evt.target.collection();
    while (frontier.length > 0) {{
      var next = cy.collection();
      frontier.forEach(function(n) {{
        n.outgoers('edge').targets().forEach(function(t) {{
          if (!downVisited.has(t)) {{
            downVisited.merge(t);
            next.merge(t);
          }}
        }});
      }});
      frontier = next;
    }}
    var upVisited = evt.target.collection();
    frontier = evt.target.collection();
    while (frontier.length > 0) {{
      var next = cy.collection();
      frontier.forEach(function(n) {{
        n.incomers('edge').sources().forEach(function(s) {{
          if (!upVisited.has(s)) {{
            upVisited.merge(s);
            next.merge(s);
          }}
        }});
      }});
      frontier = next;
    }}
    var chainNodes = upVisited.union(downVisited).filter(function(n) {{
      var t = n.data('type');
      return t !== '__region__' && t !== '__group__' && t !== '__business__' &&
             t !== '__layer__' && t !== '__service__';
    }});
    var chainEdges = cy.collection();
    chainNodes.edgesWith(chainNodes).forEach(function(e) {{
      var s = e.source(), t = e.target();
      if ((upVisited.has(s) && upVisited.has(t)) ||
          (downVisited.has(s) && downVisited.has(t))) {{
        chainEdges.merge(e);
      }}
    }});
    chainEdges.removeClass('faded');
    chainNodes.removeClass('faded').addClass('highlighted');
    // 父容器（业务/层/服务框）也取消淡化：cytoscape 渲染复合父节点时会把
    // 父级透明度作用于子节点，否则链路节点的图标和名称仍会被父框压暗。
    chainNodes.ancestors().removeClass('faded');
  }} else {{
    evt.target.neighborhood().removeClass('faded');
  }}
  evt.target.removeClass('faded').addClass('highlighted');
}});

// ── 事件：点击边 ──────────────────────────────────────────────────────────
cy.on('tap', 'edge', function(evt) {{
  var d = evt.target.data();
  var panel = document.getElementById('detail-panel');
  var inferred = d.inferred === 1;
  panel.innerHTML =
    '<div class="d-row"><div class="d-label">连接类型</div>' +
    '<div class="d-value">' + (inferred
      ? '<span style="color:#8E44AD">推断连接（紫色虚线）</span>'
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
  window._activeNodeId = d.id;
  var rows = [
    ['资源名称', d.name],
    ['所属业务', d.business  || '-'],
    ['所属层',   d.layer_name || '-'],
    ['资源类型', (d.type||'').toUpperCase()],
    ['资源ID',   d.resource_id  || '-'],
    ['企业项目', d.enterprise_project || d.enterprise_id || '-'],
    ['企业项目ID', d.enterprise_id || '-'],
    ['区域',     d.region       || '-'],
    ['资源分组', d.group_label  || '-'],
    ['规格',     d.spec         || '-'],
    ['描述',     d.desc         || '-'],
  ];
  if (d.cce_business_prefix) {{
    rows.splice(1, 0, ['CCE业务前缀', d.cce_business_prefix]);
  }}
  if (d.source_business && d.source_business !== d.business) {{
    rows.push(['原始所属业务', d.source_business]);
  }}
  var html = '<div style="display:flex;gap:6px;margin-bottom:10px">' +
             '<button class="tb-btn" style="background:#E74C3C;color:#fff;border:0" ' +
             'onclick="clearHighlight()">&#10005; 取消高亮</button>' +
             '<button class="tb-btn" style="background:#283593;color:#fff;border:0" ' +
             'onclick="openNodeEditor()">&#9998; 编辑节点</button></div>';
  html += rows.map(function(r) {{
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
    return t !== '__region__' && t !== '__group__' && t !== '__business__' &&
           t !== '__layer__' && t !== '__service__';
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

// ── 折叠资源明细：右侧滑出 xlsx 样式表格 ────────────────────────────────────
var SUMMARY_PAGE_SIZE = 10;
var SUMMARY_COLUMNS = [
  ['region', '区域'],
  ['name', '资源名称'],
  ['type', '资源类型'],
  ['ep', '企业项目'],
  ['business', '所属业务'],
  ['group', '资源分组'],
  ['downstream', '下游服务'],
];

function summaryRow(item) {{
  return {{
    region: item.region || '-',
    name: item.name || '-',
    type: item.type || item.service_name || '-',
    ep: item.enterprise_project || item.project || '-',
    business: item.business || '-',
    group: item.group || '-',
    downstream: item.downstream_text || (item.downstream || []).join(', ') || '-',
  }};
}}

function showSummaryDetail(d) {{
  if (d.type === '__agg_count__' && d.anchor) {{
    d = cy.getElementById(d.anchor).data();
  }}
  var drawer = document.getElementById('summary-drawer');
  drawer._summary = d;
  drawer._resources = d.summary_resources || [];
  drawer._page = 0;
  renderSummaryTable();
  drawer.classList.add('open');
  document.getElementById('drawer-backdrop').classList.add('open');
}}

function renderSummaryTable() {{
  var drawer = document.getElementById('summary-drawer');
  var d = drawer._summary || {{}};
  var resources = drawer._resources || [];
  var pageSize = SUMMARY_PAGE_SIZE;
  var pages = Math.max(1, Math.ceil(resources.length / pageSize));
  var page = Math.min(Math.max(drawer._page || 0, 0), pages - 1);
  drawer._page = page;
  var start = page * pageSize;
  var pageItems = resources.slice(start, start + pageSize);

  var infoItems = [
    ['所属业务', d.business || '-'],
    ['服务类型', d.service_name || d.service_key || '-'],
    ['所属层', d.layer_name || '-'],
    ['资源总数', d.resource_total || resources.length],
    ['省略节点', d.collapsed_count || 0],
  ];
  var html = '<div class="sum-info">';
  infoItems.forEach(function(item) {{
    html += '<div class="sum-info-item"><span class="sum-info-label">' +
            escHtml(item[0]) + '</span><span class="sum-info-value">' +
            escHtml(item[1]) + '</span></div>';
  }});
  html += '</div>';
  html += '<div class="sum-hint">共 ' + resources.length +
          ' 个节点，每页 ' + pageSize + ' 条；表格可左右拖动查看完整内容。</div>';
  html += '<table class="sum-table"><thead><tr>';
  SUMMARY_COLUMNS.forEach(function(col) {{
    html += '<th>' + escHtml(col[1]) + '</th>';
  }});
  html += '</tr></thead><tbody>';
  pageItems.forEach(function(item) {{
    var row = summaryRow(item);
    html += '<tr>';
    SUMMARY_COLUMNS.forEach(function(col) {{
      html += '<td>' + escHtml(row[col[0]]) + '</td>';
    }});
    html += '</tr>';
  }});
  if (pageItems.length === 0) {{
    html += '<tr><td colspan="' + SUMMARY_COLUMNS.length +
            '" style="text-align:center;color:#999">暂无数据</td></tr>';
  }}
  html += '</tbody></table>';
  html += '<div class="sum-pager">' +
          '<button class="pg-btn" ' + (page === 0 ? 'disabled' : '') +
          ' onclick="summaryPage(-1)">&#9664; 上一页</button>' +
          '<span>第 ' + (page + 1) + ' / ' + pages + ' 页</span>' +
          '<button class="pg-btn" ' + (page >= pages - 1 ? 'disabled' : '') +
          ' onclick="summaryPage(1)">下一页 &#9654;</button>' +
          '</div>';
  document.getElementById('drawer-body').innerHTML = html;
}}

function summaryPage(delta) {{
  var drawer = document.getElementById('summary-drawer');
  var pages = Math.max(1, Math.ceil((drawer._resources || []).length /
                                    SUMMARY_PAGE_SIZE));
  var next = Math.min(Math.max((drawer._page || 0) + delta, 0), pages - 1);
  drawer._page = next;
  renderSummaryTable();
}}

function closeSummaryDrawer() {{
  document.getElementById('summary-drawer').classList.remove('open');
  document.getElementById('drawer-backdrop').classList.remove('open');
}}

document.getElementById('drawer-backdrop').addEventListener('click',
  closeSummaryDrawer);

function showContainerDetail(d) {{
  var label = d.type === '__region__' ? '区域'
            : d.type === '__business__' ? '所属业务'
            : d.type === '__layer__' ? '业务层'
            : d.type === '__service__' ? '服务类型'
            : '资源分组';
  var tagColor = d.type === '__business__' ? '#1ABC9C'
               : d.type === '__layer__' ? '#5D6D7E'
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

function clearHighlight() {{
  cy.elements().removeClass('faded highlighted');
  window._activeNodeId = null;
  document.getElementById('clear-highlight-float').style.display = 'none';
  document.getElementById('detail-panel').innerHTML =
    '<p style="color:#bbb;font-size:12px;margin-top:30px;text-align:center;">点击节点或连线查看详情</p>';
  document.getElementById('node-type-tag').innerHTML = '';
}}

function showFloatingClear(node) {{
  var button = document.getElementById('clear-highlight-float');
  var pos = node.renderedPosition();
  button.style.left = Math.max(8, Math.min(window.innerWidth - 140, pos.x + 14)) + 'px';
  button.style.top = Math.max(56, pos.y + 42) + 'px';
  button.style.display = 'block';
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
  // 数字徽标依赖 BDAT 相对定位，切换其他布局时隐藏。
  cy.nodes('[type = "__agg_count__"]').style(
    'display', name === 'bdat' ? 'element' : 'none'
  );
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
        || (d.layer_name||'').toLowerCase().includes(q)
        || (d.business||'').toLowerCase().includes(q)
        || (d.enterprise_project||'').toLowerCase().includes(q);
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

var EDITABLE_FIELDS = [
  ['name', '资源名称', 'input'],
  ['type', '资源类型', 'input'],
  ['business', '所属业务', 'input'],
  ['group_label', '资源分组', 'input'],
  ['resource_id', '资源ID', 'input'],
  ['enterprise_project', '企业项目', 'input'],
  ['region', '区域', 'input'],
  ['spec', '规格', 'input'],
  ['targets', '下游服务（逗号分隔）', 'textarea'],
  ['desc', '描述', 'textarea']
];
var _editorMode = 'edit';

function openAddNodeEditor() {{
  _editorMode = 'add';
  window._editingNodeId = null;
  var fields = document.getElementById('editor-fields');
  var defaults = {{name:'', type:'ecs', business:'', group_label:'',
                   resource_id:'', enterprise_project:'', region:'', spec:'',
                   targets:'', desc:''}};
  fields.innerHTML = EDITABLE_FIELDS.map(function(field) {{
    var key = field[0], label = field[1], kind = field[2];
    var control = kind === 'textarea'
      ? '<textarea id="edit-' + key + '">' + escHtml(defaults[key]) + '</textarea>'
      : '<input id="edit-' + key + '" value="' + escAttr(defaults[key]) + '" />';
    return '<div class="edit-field ' + (kind === 'textarea' ? 'full' : '') + '">' +
      '<label for="edit-' + key + '">' + label + '</label>' + control + '</div>';
  }}).join('');
  document.getElementById('editor-title').textContent = '新增资源节点';
  document.getElementById('editor-modal').classList.add('open');
}}

function openNodeEditor() {{
  var id = window._activeNodeId;
  var node = id ? cy.getElementById(id) : cy.$(':selected').first();
  if (!node || node.empty() || node.data('is_container')) {{
    showToast('请先选择一个资源节点');
    return;
  }}
  window._editingNodeId = node.id();
  var d = node.data();
  var fields = document.getElementById('editor-fields');
  fields.innerHTML = EDITABLE_FIELDS.map(function(field) {{
    var key = field[0], label = field[1], kind = field[2];
    var value = d[key] || '';
    var control = kind === 'textarea'
      ? '<textarea id="edit-' + key + '">' + escHtml(value) + '</textarea>'
      : '<input id="edit-' + key + '" value="' + escAttr(value) + '" />';
    return '<div class="edit-field ' + (kind === 'textarea' ? 'full' : '') + '">' +
      '<label for="edit-' + key + '">' + label + '</label>' + control + '</div>';
  }}).join('');
  document.getElementById('editor-title').textContent = '编辑节点：' + (d.name || '未命名');
  document.getElementById('editor-modal').classList.add('open');
}}

function closeNodeEditor() {{
  document.getElementById('editor-modal').classList.remove('open');
  window._editingNodeId = null;
  _editorMode = 'edit';
}}

function saveNodeEdit(event) {{
  event.preventDefault();
  if (_editorMode === 'add') {{
    addResourceNode();
    return;
  }}
  var node = window._editingNodeId ? cy.getElementById(window._editingNodeId) : cy.collection();
  if (!node || node.empty()) return closeNodeEditor();
  var updated = {{}};
  EDITABLE_FIELDS.forEach(function(field) {{
    var key = field[0], input = document.getElementById('edit-' + key);
    if (input) updated[key] = input.value.trim();
  }});
  var editedService = (updated.type || node.data('service_key') || 'default')
    .toLowerCase().split(/[-_./:\\s]+/)[0];
  if (editedService === 'default' || editedService === 'other') {{
    showToast('服务类型不能为空或 Other');
    return;
  }}
  node.data(updated);
  var type = (updated.type || node.data('service_key') || 'default').toUpperCase();
  node.data('display_label', type + '\\n' + (updated.name || node.data('name') || ''));
  node.data('service_key', (updated.type || node.data('service_key') || 'default').toLowerCase());
  node.data('source_business', updated.business || node.data('source_business') || '');
  rebuildOutgoingEdges(node, updated.targets);
  try {{ localStorage.setItem('topology-node-' + node.id(), JSON.stringify(node.data())); }} catch (ignore) {{}}
  closeNodeEditor();
  showNodeDetail(node.data(), node.data('bg_color') || '#888');
  showToast('节点已更新，拓扑已重新渲染');
}}

var SERVICE_COLORS = {{
  ecs:['#FF6B35','#C0441F'], cce:['#4A90D9','#2C5F8A'], cci:['#5BA3E0','#2C5F8A'],
  vpc:['#F0F3FA','#3B6FB6'], nat:['#5D6D7E','#2E4057'], cc:['#566573','#273746'],
  dcaas:['#AED6F1','#5D8AA8'], cnad:['#117A65','#0B5345'], waf:['#C0392B','#7B241C'],
  elb:['#16A085','#0E6655'], default:['#AED6F1','#5D8AA8']
}};
var SERVICE_LABELS = {{ecs:'ECS', cce:'CCE', cci:'CCI', vpc:'VPC', nat:'NAT', cc:'CC',
  dcaas:'DCAAS', cnad:'CNAD', waf:'WAF', elb:'ELB'}};

function isOtherService(service) {{
  return service === 'default' || service === 'other';
}}

function getLayerForService(service) {{
  if (['cnad','cdn','waf','eip','dns'].indexOf(service) >= 0) return 'access';
  if (['vpc','dcaas','nat','cc','elb','slb','vpn','er'].indexOf(service) >= 0) return 'network_lb';
  if (['rds','dcs','redis','dds','obs','oss','kafka','mq','dms','sfs','evs'].indexOf(service) >= 0) return 'data_middleware_storage';
  return 'compute_app';
}}
var LAYER_LABELS = {{access:'渠道接入与安全边界', network_lb:'网络与负载均衡层',
  compute_app:'计算、容器与应用服务层', data_middleware_storage:'中间件、数据与存储层'}};

function findNodeByData(key, value) {{
  var result = cy.collection();
  cy.nodes().forEach(function(n) {{ if (n.data(key) === value) result.merge(n); }});
  return result.first();
}}

function rebuildOutgoingEdges(sourceNode, targetText) {{
  // Rebuild only user/resource call edges; fixed architecture edges remain.
  sourceNode.outgoers('edge').filter(function(edge) {{
    return !edge.data('layer_relation');
  }}).remove();
  String(targetText || '').split(/[,，;；\\r\\n]+/).forEach(function(targetName, index) {{
    targetName = targetName.trim();
    if (!targetName) return;
    var target = findNodeByData('name', targetName);
    if (target.empty() || target.id() === sourceNode.id()) return;
    cy.add({{group:'edges', data: {{
      id:'e_dynamic_' + Date.now() + '_' + index + '_' + Math.random().toString(36).slice(2),
      source:sourceNode.id(), target:target.id(), source_name:sourceNode.data('name'),
      target_name:targetName, relation:'调用', call_count:1, color:'#3498DB',
      inferred:0, cross_business:target.data('business_key') !== sourceNode.data('business_key')
    }}}});
  }});
}}

function addResourceNode() {{
  var values = {{}};
  EDITABLE_FIELDS.forEach(function(field) {{
    var input = document.getElementById('edit-' + field[0]);
    values[field[0]] = input ? input.value.trim() : '';
  }});
  if (!values.name) {{ showToast('资源名称不能为空'); return; }}
  var service = (values.type || 'default').toLowerCase().split(/[-_./:\\s]+/)[0];
  if (isOtherService(service)) {{
    showToast('服务类型不能为空或 Other');
    return;
  }}
  var layer = getLayerForService(service);
  var business = (values.business || '').trim();
  var businessKey = business
    ? '__business__:' + business.toLowerCase().replace(/[\\s_-]+/g, '-')
    : '__service_business__:' + service;
  var businessLabel = business || (SERVICE_LABELS[service] || service.toUpperCase());
  var biz = findNodeByData('business_key', businessKey);
  if (biz.empty()) {{
    var bizId = 'biz_dynamic_' + Date.now();
    biz = cy.add({{ group:'nodes', data: {{id:bizId, name:businessLabel, business_key:businessKey,
      type:'__business__', is_container:1, resource_total:0, bg_color:'#E8F8F5',
      border_color:'#1ABC9C', text_color:'#0E6655', shape:'roundrectangle'}} }});
    biz.position({{x: cy.nodes().length * 80, y: 160}});
  }}
  var expectedParent = biz;
  if (['vpc','dcaas','nat','cc'].indexOf(service) < 0) {{
    var layerNode = findNodeByData('layer_key', layer);
    if (!layerNode.empty() && layerNode.data('business_key') === businessKey) expectedParent = layerNode;
    if (expectedParent === biz) {{
      var layerId = 'layer_dynamic_' + layer + '_' + Date.now();
      expectedParent = cy.add({{group:'nodes', data: {{id:layerId, name:LAYER_LABELS[layer],
        layer_key:layer, layer_name:LAYER_LABELS[layer], business_key:businessKey,
        type:'__layer__', is_container:1, resource_total:0, min_width:240,
        bg_color:'#F7F9FC', border_color:'#5D6D7E', shape:'roundrectangle', parent:biz.id()}}}});
      expectedParent.position({{x:biz.position('x'), y:biz.position('y') + 140}});
    }}
  }}
  var serviceContainer = cy.collection();
  cy.nodes().forEach(function(n) {{
    if (n.data('type') === '__service__' && n.data('service_key') === service &&
        n.parent().id() === expectedParent.id()) serviceContainer = n;
  }});
  if (serviceContainer.empty() && service !== 'default') {{
    var parent = expectedParent;
    var serviceId = 'sc_dynamic_' + service + '_' + Date.now();
    serviceContainer = cy.add({{group:'nodes', data: {{id:serviceId, name:SERVICE_LABELS[service] || service.toUpperCase(),
      type:'__service__', service_key:service, layer_key:layer, layer_name:LAYER_LABELS[layer],
      business_key:businessKey, is_container:1, parent:parent.id(), min_width:88,
      bg_color:'#F2F3F4', border_color:'#7F8C8D', shape:'roundrectangle'}}}});
    serviceContainer.position({{x:biz.position('x') + 120, y:biz.position('y') + 120}});
  }}
  var colors = SERVICE_COLORS[service] || SERVICE_COLORS.default;
  var nodeId = 'n_dynamic_' + Date.now();
  var parentId = serviceContainer.empty() ? biz.id() : serviceContainer.id();
  var node = cy.add({{group:'nodes', data: {{id:nodeId, name:values.name, display_label:(SERVICE_LABELS[service] || service.toUpperCase()) + '\\n' + values.name,
    type:values.type || 'default', service_key:service, layer_key:layer, layer_name:LAYER_LABELS[layer],
    business:businessLabel, business_key:businessKey, group_label:values.group_label, resource_id:values.resource_id,
    enterprise_project:values.enterprise_project, region:values.region, spec:values.spec, desc:values.desc,
    targets:values.targets, bg_color:colors[0], border_color:colors[1], shape:'roundrectangle',
    is_container:0, is_summary:0, is_grouped_resource:1, parent:parentId}}}});
  var parentPosition = serviceContainer.empty() ? biz.position() : serviceContainer.position();
  node.position({{x:parentPosition.x + 120, y:parentPosition.y + 80}});
  rebuildOutgoingEdges(node, values.targets);
  biz.data('resource_total', (biz.data('resource_total') || 0) + 1);
  if (_containersLocked) {{
    cy.nodes('[is_container = 1]').forEach(function(container) {{
      var isBusinessFrame = container.data('type') === '__business__';
      if (isBusinessFrame) {{
        container.grabify();
        container.style('events', 'yes');
      }} else {{
        container.ungrabify();
        container.style('events', 'no');
      }}
    }});
  }}
  closeNodeEditor();
  cy.resize();
  cy.fit(undefined, 50);
  showToast('资源已添加并重新渲染到拓扑图');
}}

function escAttr(s) {{
  return String(s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;')
    .replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

function csvEscape(value) {{
  return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
}}

function displayWidth(value) {{
  var text = String(value == null ? '' : value);
  var width = 0;
  for (var i = 0; i < text.length; i++) {{
    width += text.charCodeAt(i) > 255 ? 2 : 1;
  }}
  return width;
}}

function xmlEscape(value) {{
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}}

function buildAdaptiveExcelXml(rows) {{
  // CSV cannot carry widths; Excel 2003 XML can, so export a companion file
  // with measured, clamped column widths for spreadsheet users.
  var widths = rows[0].map(function(_, columnIndex) {{
    var maxWidth = 0;
    rows.forEach(function(row) {{
      maxWidth = Math.max(maxWidth, displayWidth(row[columnIndex]));
    }});
    return Math.max(72, Math.min(360, (maxWidth + 2) * 7));
  }});
  var xml = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<?mso-application progid="Excel.Sheet"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' +
    'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
    'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    '<Styles><Style ss:ID="Header"><Font ss:Bold="1"/><Interior ' +
    'ss:Color="#E8EEF7" ss:Pattern="Solid"/></Style></Styles>' +
    '<Worksheet ss:Name="Edited Resources"><Table>';
  widths.forEach(function(width) {{
    xml += '<Column ss:Width="' + width.toFixed(1) + '"/>';
  }});
  rows.forEach(function(row, rowIndex) {{
    xml += '<Row>';
    row.forEach(function(value) {{
      var style = rowIndex === 0 ? ' ss:StyleID="Header"' : '';
      xml += '<Cell' + style + '><Data ss:Type="String">' +
             xmlEscape(value) + '</Data></Cell>';
    }});
    xml += '</Row>';
  }});
  return xml + '</Table></Worksheet></Workbook>';
}}

function triggerBlobDownload(blob, filename) {{
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  setTimeout(function() {{ URL.revokeObjectURL(link.href); }}, 1000);
}}

async function writeBlobToDirectory(directory, filename, blob) {{
  var handle = await directory.getFileHandle(filename, {{create: true}});
  var writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();
}}

async function downloadEditedCsv() {{
  var columns = ['name','type','business','group_label','resource_id',
                 'enterprise_project','region','spec','desc','targets'];
  var labels = ['服务名称','服务类型','所属业务','资源分组','资源ID',
                '企业项目','区域','规格','描述','下游服务'];
  var rows = [labels];
  cy.nodes().filter(function(n) {{ return !n.data('is_container') && !n.data('is_summary') &&
    n.data('type') !== '__agg_count__' && n.data('type') !== '__agg_compute__';
  }}).forEach(function(n) {{
    var d = n.data();
    rows.push(columns.map(function(key) {{ return d[key] || ''; }}));
  }});
  var csv = rows.map(function(row) {{
    return row.map(csvEscape).join(',');
  }}).join('\\r\\n');
  var csvBlob = new Blob(['\\ufeff' + csv], {{type:'text/csv;charset=utf-8'}});
  var excelBlob = new Blob([buildAdaptiveExcelXml(rows)],
    {{type:'application/vnd.ms-excel;charset=utf-8'}});
  var files = [
    {{name:'topology-edited.csv', blob:csvBlob}},
    {{name:'topology-edited.xls', blob:excelBlob}}
  ];

  if (window.showDirectoryPicker && window.isSecureContext) {{
    try {{
      var directory = await window.showDirectoryPicker({{mode:'readwrite'}});
      for (var i = 0; i < files.length; i++) {{
        await writeBlobToDirectory(directory, files[i].name, files[i].blob);
      }}
      showToast('CSV 和自适应 Excel 表已保存到指定目录');
      return;
    }} catch (error) {{
      if (error && error.name === 'AbortError') {{
        showToast('已取消选择导出目录');
        return;
      }}
      console.warn('目录写入失败，改用浏览器下载:', error);
    }}
  }}

  files.forEach(function(file) {{ triggerBlobDownload(file.blob, file.name); }});
  showToast('CSV 和自适应 Excel 表已下载');
}}

var _containersLocked = false;
function toggleContainerLock() {{
  _containersLocked = !_containersLocked;
  var containers = cy.nodes('[is_container = 1]');
  containers.forEach(function(n) {{
    if (_containersLocked) {{
      var isBusinessFrame = n.data('type') === '__business__';
      if (isBusinessFrame) {{
        // The business frame is the movable viewport proxy. Its descendants
        // remain fixed relative to it, while inner architecture frames stay
        // locked and cannot be grabbed independently.
        n.grabify();
        n.style('events', 'yes');
      }} else {{
        n.ungrabify();
        n.style('events', 'no');
      }}
    }} else {{
      n.grabify();
      n.style('events', 'yes');
    }}
  }});
  cy.userPanningEnabled(true);
  var btn = document.getElementById('lockContainersBtn');
  btn.classList.toggle('locked', _containersLocked);
  btn.innerHTML = _containersLocked ? '&#128274; 锁定容器:开' : '&#128275; 锁定容器:关';
  showToast(_containersLocked ? '业务/层/服务框已锁定，可拖动视角和资源节点' : '容器已解锁');
}}


// ── 数字徽标与缩略节点图标绑定：拖动图标时数字跟随移动 ──────────────────────
cy.on('position', 'node[type = "__agg_compute__"]', function(evt) {{
  var anchor = evt.target;
  cy.nodes('[type = "__agg_count__"]').forEach(function(badge) {{
    if (badge.data('anchor') === anchor.id()) {{
      var p = anchor.position();
      badge.position({{ x: p.x + 64, y: p.y - 70 }});
    }}
  }});
}});

var _bizViewOn = true;
var _origParents = {{}};

function toggleBusinessView() {{
  _bizViewOn = !_bizViewOn;
  var btn = document.getElementById('bizToggleBtn');
  if (!_bizViewOn) {{
    // 先摘除叶节点与各级子容器的父子关系
    cy.nodes().forEach(function(n) {{
      var p = n.data('parent');
      if (p && (p.indexOf('biz_') === 0 || p.indexOf('layer_') === 0 || p.indexOf('sc_') === 0 || p.indexOf('grp_') === 0)) {{
        _origParents[n.id()] = p;
        n.move({{ parent: null }});
      }}
    }});
    cy.nodes('[type = "__business__"],[type = "__layer__"],[type = "__service__"],[type = "__group__"]').style('display', 'none');
    btn.innerHTML = '&#127968; 业务分组:关';
    btn.style.background = 'rgba(255,255,255,.15)';
  }} else {{
    cy.nodes('[type = "__business__"],[type = "__layer__"],[type = "__service__"],[type = "__group__"]').style('display', 'element');
    Object.keys(_origParents).forEach(function(nid) {{
      cy.getElementById(nid).move({{ parent: _origParents[nid] }});
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
    print(f"  可见节点: {visible_resources + summary_nodes + len(agg_compute_nodes)}")
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
  py converter.py my_services.xlsx --output-dir ./dist
"""
    )
    parser.add_argument("excel", help="输入 Excel 文件路径")
    parser.add_argument("output", nargs="?", default=None,
                        help="输出 HTML 文件路径（默认同名 .html）")
    parser.add_argument("--title", default="云服务拓扑图", help="页面标题")
    parser.add_argument("--output-dir", default=None,
                        help="输出目录；未指定时沿用输入文件目录或 output 参数目录")
    args = parser.parse_args()

    if not os.path.exists(args.excel):
        sys.exit(f"文件不存在: {args.excel}")

    if args.output:
        output = args.output
    else:
        output = os.path.splitext(args.excel)[0] + ".html"
    if args.output_dir:
        os.makedirs(args.output_dir, exist_ok=True)
        output = os.path.join(
            args.output_dir,
            os.path.basename(output),
        )

    print(f"读取 Excel: {args.excel}")
    records = read_excel(args.excel)
    print(f"  读取到 {len(records)} 条记录")

    elements = build_graph(records)
    generate_html(elements, title=args.title, output_path=output)


if __name__ == "__main__":
    main()
