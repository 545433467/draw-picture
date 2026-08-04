#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成示例 Excel（约 1000 个节点），展示 converter.py 所需的列格式。
运行: py create_sample.py
"""

import random

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def create_sample():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "云服务清单"

    headers = [
        "region", "resource name", "resource id", "resource_type",
        "enterprise_project_id", "企业项目", "资源分组", "所属业务",
        "是否核心业务", "下游服务", "下游服务(推断)", "推断原因",
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

    # ── 生成约 1000 个节点的模拟数据 ────────────────────────────────────────
    random.seed(20260803)

    businesses = ["订单业务", "支付业务", "用户中心", "风控业务", "数据平台"]
    biz_keys   = {"订单业务": "order", "支付业务": "pay", "用户中心": "user",
                  "风控业务": "risk", "数据平台": "data"}
    projects = [("生产项目", "ep-prod-001"),
                ("测试项目", "ep-test-002"),
                ("预发项目", "ep-stage-003")]
    regions = ["cn-north-4", "cn-east-3"]

    def pick_project():
        r = random.random()
        if r < 0.6:
            return projects[0]
        if r < 0.85:
            return projects[1]
        return projects[2]

    def make_row(name, rtype, group, biz, region=None, down="",
                 inferred="", reason=""):
        proj_name, proj_id = pick_project()
        is_core = "是" if random.random() < 0.5 else "否"
        return [
            region or random.choice(regions), name, f"id-{name}", rtype,
            proj_id, proj_name, group, biz, is_core, down, inferred, reason,
        ]

    rows = []

    # 各业务的通用服务（每业务 54 个基础资源）
    for biz in businesses:
        bk = biz_keys[biz]
        r = lambda name, rtype, group, **kw: make_row(
            name, rtype, group, biz, **kw)

        for i in range(1, 3):
            rows.append(r(f"waf-{bk}-{i:02d}", "WAF", "公网接入层"))
        for i in range(1, 5):
            rows.append(r(f"elb-{bk}-{i:02d}", "ELB", "接入与负载层"))
        for i in range(1, 13):
            rows.append(r(f"ecs-{bk}-web-{i:02d}", "ECS", "Web应用层"))
        for i in range(1, 4):
            rows.append(r(f"apig-{bk}-{i:02d}", "APIG", "API网关"))
        for i in range(1, 5):
            rows.append(r(f"dcs-{bk}-redis-{i:02d}", "DCS", "缓存层"))
        rows.append(r(f"rds-{bk}-master", "RDS", "数据库层"))
        for i in range(1, 6):
            rows.append(r(f"rds-{bk}-slave-{i:02d}", "RDS", "数据库层"))
        for i in range(1, 5):
            rows.append(r(f"dms-{bk}-kafka-{i:02d}", "DMS", "消息队列"))
        for i in range(1, 6):
            rows.append(r(f"obs-{bk}-{i:02d}", "OBS", "对象存储"))
        for i in range(1, 4):
            rows.append(r(f"nat-{bk}-{i:02d}", "NAT", "网络出口"))
        for i in range(1, 5):
            rows.append(r(f"vpc-{bk}-{i:02d}", "VPC", "网络"))
        for i in range(1, 6):
            rows.append(r(f"eip-{bk}-{i:02d}", "EIP", "公网IP"))
        for i in range(1, 3):
            rows.append(r(f"cdn-{bk}-{i:02d}", "CDN", "接入层"))

        # CCE_Deployment（每个业务 132 个，资源分组字段取多个值，
        # 用于展示“资源分组”次级框）
        cce_groups = ["CCE_Deployment", "CCE工作负载", "K8s集群"]
        for i in range(1, 133):
            rows.append(r(
                f"bj-{bk}-deploy-{i:03d}", "CCE_Deployment",
                cce_groups[i % len(cce_groups)],
                region="cn-north-4",
                down=(f"rds-{bk}-master,dcs-{bk}-redis-01"
                      if i % 25 == 0 else ""),
                reason=("ConfigMap 包含数据库、缓存连接串"
                        if i % 25 == 0 else ""),
            ))

        # CCE 集群节点（非 CCE_Deployment，正常展示）
        for i in range(1, 5):
            rows.append(r(f"cce-{bk}-cluster-{i:02d}", "CCE", "K8s集群"))

        # 数据平台类扩展服务
        for i in range(1, 3):
            rows.append(r(f"css-{bk}-{i:02d}", "CSS", "检索服务"))
            rows.append(r(f"dds-{bk}-{i:02d}", "DDS", "数据库层"))
            rows.append(r(f"swr-{bk}-{i:02d}", "SWR", "容器镜像"))
            rows.append(r(f"cci-{bk}-{i:02d}", "CCI", "容器实例"))
        for i in range(1, 4):
            rows.append(r(f"functiongraph-{bk}-{i:02d}", "FunctionGraph",
                          "函数计算"))

    # ── 关键调用关系（确认边 + 推断边）──────────────────────────────────────
    for biz in businesses:
        bk = biz_keys[biz]
        set_fields = {}

        def patch(name, **fields):
            set_fields[name] = fields

        patch(f"elb-{bk}-01",
              down=f"ecs-{bk}-web-01,ecs-{bk}-web-02",
              inferred=f"nat-{bk}-01",
              reason="ELB 后端服务器组包含 ecs-web-01/02；NAT 网关与 ELB "
                     "同子网，推断存在出口流量关系")
        patch(f"apig-{bk}-01",
              down=f"ecs-{bk}-web-01,ecs-{bk}-web-02",
              reason="APIG 后端服务配置指向 ecs-web-01/02")
        patch(f"ecs-{bk}-web-01",
              down=f"rds-{bk}-master,dcs-{bk}-redis-01,obs-{bk}-01",
              inferred=f"apig-{bk}-01",
              reason="配置文件中 DB_HOST 指向 RDS；缓存地址指向 DCS；"
                     "日志写入 OBS；同子网存在 APIG，推断存在 API 路由")
        patch(f"ecs-{bk}-web-02",
              down=f"rds-{bk}-master,dcs-{bk}-redis-01",
              reason="同 ecs-web-01 配置相同，为水平扩展副本")
        patch(f"rds-{bk}-master",
              down=f"rds-{bk}-slave-01",
              reason="RDS 高可用主备模式，主库向备库同步")
        patch(f"dcs-{bk}-redis-01",
              inferred=f"dcs-{bk}-redis-02",
              reason="Redis 主从架构，主节点写，推断存在从节点同步")
        patch(f"dms-{bk}-kafka-01",
              inferred=f"ecs-{bk}-web-02",
              reason="Kafka topic 订阅关系在配置中心配置，推断存在消费者")
        patch(f"nat-{bk}-01",
              inferred=f"obs-{bk}-01",
              reason="NAT 网关出口规则中包含 OBS 域名，推断服务经 NAT 访问 OBS")

    # 将调用关系回填到对应行
    name_to_row = {row[1]: row for row in rows}
    for name, fields in set_fields.items():
        row = name_to_row.get(name)
        if not row:
            continue
        if "down" in fields:
            row[8] = fields["down"]
        if "inferred" in fields:
            row[9] = fields["inferred"]
        if "reason" in fields:
            row[10] = fields["reason"]

    # ── 写回 Excel ───────────────────────────────────────────────────────────
    alt_fill = [PatternFill("solid", fgColor="EBF5FB"),
                PatternFill("solid", fgColor="FFFFFF")]
    for i, row_data in enumerate(rows, 2):
        for col, val in enumerate(row_data, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill   = alt_fill[i % 2]
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=(col >= 9))
        ws.row_dimensions[i].height = 40 if row_data[9] or row_data[10] else 20

    widths = [12, 24, 20, 14, 18, 12, 14, 12, 14, 42, 30, 50]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 说明 Sheet ────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("列说明")
    notes = [
        ["列名",                   "必填", "说明",                                        "示例"],
        ["region",                 "否",   "资源所在区域/地域",                           "cn-north-4"],
        ["resource name",          "是",   "资源名称（唯一标识）",                        "ecs-order-web-01"],
        ["resource id",            "否",   "资源实例 ID",                                 "ecs-11111111"],
        ["resource_type",          "是",   "服务类型（ECS/ELB/RDS/DCS/CCE_Deployment 等）", "ECS"],
        ["enterprise_project_id",  "否",   "企业项目 ID",                                 "ep-prod-001"],
        ["企业项目",               "否",   "企业项目名称（CCE_Deployment 聚合节点名）",     "生产项目"],
        ["资源分组",               "否",   "功能分组标签（CCE_Deployment 触发聚合）",       "CCE_Deployment"],
        ["所属业务",               "否",   "业务分组依据",                                 "订单业务"],
        ["是否核心业务",           "否",   "仅“是”的节点会绘制（存在该列时生效）",          "是"],
        ["下游服务",               "否",   "已确认下游服务，逗号分隔（值为 resource name）", "rds-order-master,dcs-order-redis-01"],
        ["下游服务(推断)",         "否",   "推断的下游服务，逗号分隔（显示为虚线）",        "apig-order-01"],
        ["推断原因",               "否",   "推断依据（详细描述分析过程）",                  "配置文件中 DB_HOST 指向…"],
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
    for col, w in zip("ABCD", [22, 6, 55, 42]):
        ws2.column_dimensions[col].width = w

    wb.save("sample.xlsx")
    print(f"已生成 sample.xlsx（数据行数: {len(rows)}）")
    print("运行: py converter.py sample.xlsx")


if __name__ == "__main__":
    create_sample()
