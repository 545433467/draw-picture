#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成示例 Excel，展示 converter.py 所需的列格式。
运行: py create_sample.py
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def create_sample():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "云服务清单"

    headers = [
        "region", "resource name", "resource id", "resource_type",
        "enterprise_project_id", "资源分组",
        "下游服务", "下游服务(推断)", "推断原因",
    ]

    hfill  = PatternFill("solid", fgColor="1A237E")
    hfont  = Font(color="FFFFFF", bold=True, size=11)
    thin   = Side(style="thin", color="C0C0C0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = hfill; c.font = hfont
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws.row_dimensions[1].height = 28

    # ── 示例数据（模拟一个华为云环境下的典型 Web 应用架构）──────────────────
    # region | resource name | resource id | resource_type | enterprise_project_id
    # | 资源分组 | 下游服务 | 下游服务(推断) | 推断原因
    rows = [
        # ── 公网接入层 ──────────────────────────────────────────────────────
        ["cn-north-4", "waf-prod-01",     "waf-a1b2c3", "WAF",
         "ep-prod-001", "公网接入层",
         "elb-prod-01", "",
         "WAF 检测到请求后转发至 ELB，直接调用关系，来源：安全组规则配置"],

        ["cn-north-4", "elb-prod-01",     "elb-d4e5f6", "ELB",
         "ep-prod-001", "公网接入层",
         "ecs-web-01,ecs-web-02", "nat-prod-01",
         "ELB 后端服务器组包含 ecs-web-01/02；NAT 网关与 ELB 同子网，推断存在出口流量关系"],

        # ── Web 应用层 ───────────────────────────────────────────────────────
        ["cn-north-4", "ecs-web-01",      "ecs-11111111", "ECS",
         "ep-prod-001", "Web应用层",
         "rds-prod-master,dcs-redis-01,obs-prod-01", "apig-prod-01",
         "配置文件中 DB_HOST 指向 rds-prod-master；缓存地址指向 dcs-redis-01；"
         "应用日志写入 OBS；同子网存在 APIG，推断存在 API 路由"],

        ["cn-north-4", "ecs-web-02",      "ecs-22222222", "ECS",
         "ep-prod-001", "Web应用层",
         "rds-prod-master,dcs-redis-01,obs-prod-01", "apig-prod-01",
         "同 ecs-web-01 配置相同，为水平扩展副本"],

        # ── API 网关 ─────────────────────────────────────────────────────────
        ["cn-north-4", "apig-prod-01",    "apig-aabbcc", "APIG",
         "ep-prod-001", "API网关",
         "ecs-web-01,ecs-web-02", "",
         "APIG 后端服务配置指向 ecs-web-01/02"],

        # ── 缓存 / 数据库层 ───────────────────────────────────────────────────
        ["cn-north-4", "dcs-redis-01",    "dcs-cc3344", "DCS",
         "ep-prod-001", "缓存层",
         "", "dcs-redis-slave",
         "Redis 主从架构，主节点写，推断存在从节点同步"],

        ["cn-north-4", "dcs-redis-slave", "dcs-cc3345", "DCS",
         "ep-prod-001", "缓存层",
         "", "",
         ""],

        ["cn-north-4", "rds-prod-master", "rds-55667788", "RDS",
         "ep-prod-001", "数据库层",
         "rds-prod-slave", "",
         "RDS 高可用主备模式，主库向备库同步"],

        ["cn-north-4", "rds-prod-slave",  "rds-55667799", "RDS",
         "ep-prod-001", "数据库层",
         "", "",
         ""],

        # ── 消息队列 ─────────────────────────────────────────────────────────
        ["cn-north-4", "dms-kafka-01",    "dms-aac001", "DMS",
         "ep-prod-001", "消息队列",
         "", "ecs-worker-01",
         "Kafka topic 订阅关系在配置中心配置，ecs-worker-01 为消费者，推断连接"],

        ["cn-north-4", "ecs-worker-01",   "ecs-33333333", "ECS",
         "ep-prod-001", "后台任务层",
         "rds-prod-master,obs-prod-01", "dms-kafka-01",
         "Worker 消费 Kafka 消息，写库写 OBS；Kafka 来源为推断，需人工核实"],

        # ── 存储 ──────────────────────────────────────────────────────────────
        ["cn-north-4", "obs-prod-01",     "obs-bucket-prod", "OBS",
         "ep-prod-001", "对象存储",
         "", "",
         ""],

        # ── 网络出口 ──────────────────────────────────────────────────────────
        ["cn-north-4", "nat-prod-01",     "nat-ff0011", "NAT",
         "ep-prod-001", "网络出口",
         "", "obs-prod-01",
         "NAT 网关出口规则中包含 OBS 域名，推断服务通过 NAT 访问 OBS"],

        # ── CCE（K8s）集群 ─────────────────────────────────────────────────────
        ["cn-north-4", "cce-cluster-01",  "cce-99001122", "CCE",
         "ep-prod-001", "K8s工作节点",
         "rds-prod-master,dcs-redis-01,dms-kafka-01", "obs-prod-01",
         "K8s ConfigMap 中包含数据库、缓存连接串；OBS 作为持久卷，属推断关系"],
    ]

    alt_fill = [PatternFill("solid", fgColor="EBF5FB"),
                PatternFill("solid", fgColor="FFFFFF")]
    for i, row_data in enumerate(rows, 2):
        for col, val in enumerate(row_data, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill   = alt_fill[i % 2]
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=(col >= 7))
        ws.row_dimensions[i].height = 40 if row_data[7] or row_data[8] else 20

    # 列宽
    widths = [14, 22, 20, 12, 18, 14, 36, 28, 55]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 说明 Sheet ────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("列说明")
    notes = [
        ["列名",                   "必填", "说明",                                        "示例"],
        ["region",                 "否",   "资源所在区域/地域",                           "cn-north-4"],
        ["resource name",          "是",   "资源名称（唯一标识）",                        "ecs-web-01"],
        ["resource id",            "否",   "资源实例 ID",                                 "ecs-11111111"],
        ["resource_type",          "是",   "服务类型（ECS/ELB/RDS/DCS/OBS/CCE 等）",      "ECS"],
        ["enterprise_project_id",  "否",   "企业项目 ID / MAC / 其他地址",                "ep-prod-001"],
        ["资源分组",               "否",   "功能分组标签（如 K8s工作节点、反向代理）",     "Web应用层"],
        ["下游服务",               "否",   "已确认下游服务，逗号分隔（值为 resource name）","rds-prod-master,dcs-redis-01"],
        ["下游服务(推断)",         "否",   "推断的下游服务，逗号分隔（显示为虚线）",        "apig-prod-01"],
        ["推断原因",               "否",   "推断依据（详细描述分析过程）",                 "配置文件中 DB_HOST 指向…"],
    ]
    h2fill = PatternFill("solid", fgColor="283593")
    for i, row_data in enumerate(notes, 1):
        for j, val in enumerate(row_data, 1):
            c = ws2.cell(row=i, column=j, value=val)
            if i == 1:
                c.fill = h2fill; c.font = Font(color="FFFFFF", bold=True)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws2.row_dimensions[i].height = 36
    for col, w in zip("ABCD", [22, 6, 45, 32]):
        ws2.column_dimensions[col].width = w

    wb.save("sample.xlsx")
    print("已生成 sample.xlsx")
    print("运行: py converter.py sample.xlsx")


if __name__ == "__main__":
    create_sample()
