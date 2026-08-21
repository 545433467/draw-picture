import os
import tempfile
import unittest

import converter
import openpyxl

from converter import (
    ICON_SIZE,
    MAX_NODES_PER_ROW,
    _make_cytoscape_style,
    build_graph,
    compute_bdat_positions,
    get_service_key,
    get_topology_layer,
    read_excel,
)


def make_record(name, business="业务A", service_type="cce_deploym", targets="",
                inferred_targets="", reason="", source="", enterprise_id="",
                enterprise_project="", group="", is_core_business=None):
    return {
        "name": name,
        "type": service_type,
        "resource_id": f"id-{name}",
        "enterprise_id": enterprise_id,
        "enterprise_project": enterprise_project,
        "region": "cn-test-1",
        "group": group,
        "business": business,
        "is_core_business": is_core_business,
        "desc": "",
        "spec": "",
        "source": source,
        "targets": targets,
        "inferred_targets": inferred_targets,
        "reason": reason,
    }


def element_data(elements, group):
    return [element["data"] for element in elements if element["group"] == group]


def call_edges(elements):
    """仅返回普通调用边（排除四层架构固定层级关系边）。"""
    return [
        element["data"] for element in elements
        if element["group"] == "edges"
        and not element["data"].get("layer_relation")
    ]


class BuildGraphAggregationTests(unittest.TestCase):
    def setUp(self):
        # 默认关闭业务关键字过滤，避免影响既有用例；专门测试中再开启。
        self._old_biz_filter = converter.BUSINESS_FILTER_KEYWORD
        converter.BUSINESS_FILTER_KEYWORD = None

    def tearDown(self):
        converter.BUSINESS_FILTER_KEYWORD = self._old_biz_filter

    def test_long_resource_name_uses_two_line_full_label(self):
        long_name = "ecs-production-payment-service-with-a-very-long-instance-name"

        elements = build_graph([
            make_record(long_name, service_type="rds")
        ])
        node = next(data for data in element_data(elements, "nodes")
                    if not data.get("is_container"))

        label_lines = node["display_label"].splitlines()
        self.assertEqual(label_lines[0], "RDS")
        # 资源名称完整显示并折成多行，不做“...”缩写
        self.assertGreaterEqual(len(label_lines), 3)
        self.assertEqual("".join(label_lines[1:]), long_name)
        self.assertEqual(node["name"], long_name)

        custom_elements = build_graph([
            make_record("node-1", service_type="extremely_long_custom_cloud_service")
        ])
        custom_node = next(data for data in element_data(custom_elements, "nodes")
                           if not data.get("is_container"))
        self.assertTrue(custom_node["display_label"].splitlines()[0].endswith("..."))

    def test_cce_ecs_aggregated_per_business(self):
        records = [
            make_record("ecs-1", service_type="ecs", targets="rds-1"),
            make_record("ecs-2", service_type="ecs"),
            make_record("cce-1", service_type="cce", group="CCE_Deployment"),
            make_record("rds-1", service_type="rds"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]
        cce_ecs_leaves = [
            node for node in nodes
            if node.get("service_key") in {"cce", "ecs"}
            and not node.get("is_container") and not node.get("is_summary")
            and node.get("type") != "__agg_compute__"
        ]

        # 按资源分组聚合：空分组 2 个（ecs-1/ecs-2），CCE_Deployment 分组 1 个
        self.assertEqual(len(aggs), 2)
        by_label = {node["group_label"]: node for node in aggs}
        self.assertEqual(by_label["(未设置资源分组)"]["resource_total"], 2)
        self.assertEqual(by_label["CCE_Deployment"]["resource_total"], 1)
        self.assertEqual(
            len(by_label["(未设置资源分组)"]["summary_resources"]), 2
        )
        self.assertEqual(cce_ecs_leaves, [])
        self.assertTrue(any(node.get("name") == "rds-1" for node in nodes))

    def test_elb_points_to_aggregate_with_single_edge(self):
        records = [
            make_record("elb-1", service_type="elb", targets="ecs-1,ecs-2"),
            make_record("ecs-1", service_type="ecs"),
            make_record("ecs-2", service_type="ecs"),
        ]

        elements = build_graph(records)
        edges = call_edges(elements)
        agg = next(node for node in element_data(elements, "nodes")
                   if node.get("type") == "__agg_compute__")

        self.assertEqual(len(edges), 1)
        self.assertEqual(edges[0]["target"], agg["id"])
        self.assertEqual(edges[0]["call_count"], 2)

    def test_cross_business_call_targets_aggregate_node(self):
        records = [make_record("api-1", business="业务A", service_type="apig",
                               targets="cce-node-6")]
        records.extend(make_record(f"cce-node-{i}", business="业务B")
                       for i in range(1, 7))

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        agg_b = next(
            node for node in nodes
            if node.get("type") == "__agg_compute__"
            and node.get("business") == "业务B"
        )
        edge = next(edge for edge in edges if edge["target_name"] == "cce-node-6")

        self.assertEqual(edge["target"], agg_b["id"])
        self.assertEqual(edge["cross_business"], 1)

    def test_max_nodes_per_row_stay_on_one_layout_row(self):
        records = [make_record(f"rds-node-{i}", service_type="rds")
                   for i in range(1, MAX_NODES_PER_ROW + 1)]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        positions = compute_bdat_positions(elements)
        resource_ids = [node["id"] for node in nodes if not node.get("is_container")]

        self.assertFalse(any(node.get("is_summary") for node in nodes))
        self.assertEqual(len({positions[node_id]["y"] for node_id in resource_ids}), 1)
        self.assertEqual(len({positions[node_id]["x"] for node_id in resource_ids}),
                         MAX_NODES_PER_ROW)

    def test_records_without_calls_do_not_create_edges(self):
        elements = build_graph([make_record(f"cce-node-{i}") for i in range(1, 8)])

        self.assertEqual(element_data(elements, "edges"), [])

    def test_unassigned_resources_get_service_virtual_business(self):
        records = [make_record(f"waf-node-{i}", business="", service_type="waf")
                   for i in range(1, 9)]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        business = next(node for node in nodes if node.get("type") == "__business__")
        resources = [node for node in nodes if not node.get("is_container")]

        self.assertEqual(business["name"], "WAF")
        self.assertEqual(business["is_virtual_business"], 1)
        self.assertEqual(business["resource_total"], 8)
        self.assertFalse(any(node.get("type") == "__service__" for node in nodes))
        self.assertEqual(len(resources), 8)
        self.assertFalse(any(node.get("is_summary") for node in resources))
        self.assertTrue(all(node.get("parent") == business["id"]
                            for node in resources))

    def test_equivalent_business_names_share_one_container(self):
        external_name = "gz-hw-lingee-prod-k8s-api"
        records = [
            make_record(
                "ecs-app", business="lingee-prod", service_type="ecs",
                targets=external_name,
            ),
            make_record("rds-app", business="LINGEE_prod", service_type="rds"),
            make_record("obs-app", business="lingee prod", service_type="obs"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        businesses = [node for node in nodes if node.get("type") == "__business__"]
        service_containers = [
            node for node in nodes if node.get("type") == "__service__"
        ]
        node_by_id = {node["id"]: node for node in nodes}

        self.assertEqual(len(businesses), 1)
        self.assertEqual(businesses[0]["name"], "lingee-prod")
        self.assertEqual(businesses[0]["resource_total"], 4)
        self.assertTrue(all(
            node_by_id[node["parent"]]["parent"] == businesses[0]["id"]
            for node in service_containers
        ))
        self.assertTrue(all(edge["cross_business"] == 0 for edge in edges))

    def test_business_style_wraps_full_name_instead_of_ellipsis(self):
        business_style = next(
            item["style"] for item in _make_cytoscape_style()
            if item["selector"] == "node[type = '__business__']"
        )

        self.assertEqual(business_style["text-wrap"], "wrap")
        self.assertEqual(business_style["text-max-width"], "320px")
        self.assertEqual(business_style["text-overflow-wrap"], "anywhere")

    def test_container_labels_use_larger_fonts(self):
        style = _make_cytoscape_style()

        def font_size(selector):
            return next(
                item["style"]["font-size"] for item in style
                if item["selector"] == selector
            )

        self.assertEqual(font_size("node[type = '__service__']"), 24)
        self.assertEqual(font_size("node[type = '__layer__']"), 24)
        self.assertEqual(font_size("node[type = '__business__']"), 28)

    def test_layer_relation_edges_connect_consecutive_layers(self):
        records = [
            make_record("waf-1", business="业务A", service_type="waf"),
            make_record("elb-1", business="业务A", service_type="elb"),
            make_record("ecs-1", business="业务A", service_type="ecs",
                        targets="rds-1"),
            make_record("rds-1", business="业务A", service_type="rds"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        layers = {
            node["id"]: node for node in nodes
            if node.get("type") == "__layer__"
        }
        layer_edges = [
            edge for edge in edges if edge.get("layer_relation") == 1
        ]

        self.assertEqual(len(layer_edges), 3)
        for edge in layer_edges:
            self.assertIn(edge["source"], layers)
            self.assertIn(edge["target"], layers)
            self.assertEqual(edge["relation"], "固定层级关系")
        order = ["access", "network_lb", "compute_app",
                 "data_middleware_storage"]
        chained = [
            layers[edge["source"]]["layer_key"] for edge in layer_edges
        ] + [layers[layer_edges[-1]["target"]]["layer_key"]]
        self.assertEqual(chained, order)

    def test_resource_group_subcontainers_separate_mixed_groups(self):
        records = [
            make_record("dcs-1", service_type="dcs", group="缓存A"),
            make_record("dcs-2", service_type="dcs", group="缓存A"),
            make_record("dcs-3", service_type="dcs", group="缓存B"),
            make_record("dcs-4", service_type="dcs", group=""),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        node_by_id = {node["id"]: node for node in nodes}
        group_boxes = [node for node in nodes if node.get("type") == "__group__"]

        # 同一服务组内有 3 个不同资源分组值 → 生成 3 个次级框，框名为资源分组
        self.assertEqual({node["name"] for node in group_boxes},
                         {"缓存A", "缓存B", "(未设置资源分组)"})
        dcs1 = next(node for node in nodes if node.get("name") == "dcs-1")
        box = node_by_id[dcs1["parent"]]
        self.assertEqual(box["type"], "__group__")
        self.assertEqual(box["name"], "缓存A")
        self.assertEqual(box["parent"], dcs1["service_container"])
        self.assertEqual(box["resource_total"], 2)

    def test_single_resource_group_value_keeps_flat_layout(self):
        records = [
            make_record("dcs-1", service_type="dcs", group="缓存A"),
            make_record("dcs-2", service_type="dcs", group="缓存A"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")

        self.assertFalse(any(node.get("type") == "__group__" for node in nodes))
        dcs1 = next(node for node in nodes if node.get("name") == "dcs-1")
        self.assertEqual(dcs1["parent"], dcs1["service_container"])

    def test_service_containers_have_distinct_positions(self):
        records = [
            make_record("rds-1", service_type="rds"),
            make_record("dcs-1", service_type="dcs"),
        ]

        elements = build_graph(records)
        positions = compute_bdat_positions(elements)
        service_nodes = [
            e["data"] for e in elements
            if e["group"] == "nodes" and e["data"]["type"] == "__service__"
        ]
        pos = {node["id"]: positions[node["id"]] for node in service_nodes}

        # 每个服务框都有独立坐标，不同资源类型不会重叠
        self.assertTrue(all(pos[node["id"]] is not None
                            for node in service_nodes))
        self.assertEqual(
            len({(round(p["x"], 1), round(p["y"], 1)) for p in pos.values()}),
            len(service_nodes),
        )

    def test_icon_style_uses_fixed_dimensions(self):
        style = _make_cytoscape_style({"ecs": "data:image/png;base64,abc"})
        icon_style = next(
            item["style"] for item in style
            if item["selector"] == "node[has_icon = 1][is_summary = 0]"
        )
        image_style = next(
            item["style"] for item in style
            if item["selector"] == "node[service_key = 'ecs'][has_icon = 1]"
        )

        self.assertEqual(icon_style["width"], ICON_SIZE)
        self.assertEqual(icon_style["height"], ICON_SIZE)
        self.assertEqual(icon_style["background-width"], ICON_SIZE)
        self.assertEqual(icon_style["background-height"], ICON_SIZE)
        self.assertEqual(
            image_style["background-image"],
            'url("data:image/png;base64,abc")',
        )

    def test_dns_service_is_recognized_as_access_layer(self):
        self.assertEqual(get_service_key("prod-dns-entry", "default"), "dns")
        self.assertEqual(get_topology_layer("dns"), "access")

    def test_cc_prefixed_resource_name_maps_to_cc(self):
        self.assertEqual(get_service_key("CC-bandwidth-main", "default"), "cc")

    def test_cc_resource_id_hint_maps_to_cc(self):
        self.assertEqual(
            converter.get_service_key("anything", "default", "foo-CCAAS-bar"),
            "cc",
        )

    def test_dcaas_resource_type_maps_to_network_layer(self):
        self.assertEqual(
            converter.get_service_key("anything", "DCAAS_service"),
            "dcaas",
        )
        elements = build_graph([
            make_record(
                "dcaas-node",
                service_type="DCAAS_service",
                resource_id="id-dcaas",
            )
        ])
        node = next(
            data for data in element_data(elements, "nodes")
            if not data.get("is_container")
        )
        self.assertEqual(node["service_key"], "dcaas")
        self.assertEqual(node["layer_key"], "network_lb")

    def test_cc_group_and_resource_id_map_to_network_layer(self):
        records = [
            make_record(
                "svc-1",
                service_type="default",
                group="CC_bandwidth",
                resource_id="id-CCAAS-001",
            ),
        ]

        elements = build_graph(records)
        nodes = {
            node["name"]: node for node in element_data(elements, "nodes")
            if not node.get("is_container")
        }

        self.assertEqual(nodes["svc-1"]["service_key"], "cc")
        self.assertEqual(nodes["svc-1"]["layer_key"], "network_lb")

    def test_new_resources_map_to_expected_layers(self):
        self.assertEqual(get_topology_layer("nat"), "network_lb")
        self.assertEqual(get_topology_layer("cc"), "network_lb")
        self.assertEqual(get_topology_layer("dcaas"), "network_lb")
        self.assertEqual(get_topology_layer("cnad"), "network_lb")
        self.assertEqual(get_topology_layer("geminidb"), "data_middleware_storage")

    def test_cce_db_named_ecs_is_placed_in_data_layer(self):
        elements = build_graph([
            make_record("cce-app", service_type="cce", targets="ecs-db"),
            make_record("ecs-db", service_type="ecs"),
            make_record("ecs-db-unrelated", service_type="ecs"),
        ])

        aggregate_nodes = [
            node for node in element_data(elements, "nodes")
            if node.get("type") == "__agg_compute__"
        ]
        data_node = next(
            node for node in aggregate_nodes
            if node.get("layer_key") == "data_middleware_storage"
        )
        compute_node = next(
            node for node in aggregate_nodes
            if node.get("layer_key") == "compute_app"
        )

        self.assertEqual(
            [resource["name"] for resource in data_node["summary_resources"]],
            ["ecs-db"],
        )
        self.assertIn(
            "ecs-db-unrelated",
            [resource["name"] for resource in compute_node["summary_resources"]],
        )

    def test_cce_db_named_elb_and_its_ecs_are_relayered(self):
        elements = build_graph([
            make_record("cce-app", service_type="cce", targets="elb-tidb"),
            make_record("elb-tidb", service_type="elb", targets="ecs-backend"),
            make_record("ecs-backend", service_type="ecs"),
        ])

        nodes = element_data(elements, "nodes")
        elb_node = next(node for node in nodes if node.get("name") == "elb-tidb")
        ecs_aggregate = next(
            node for node in nodes
            if node.get("type") == "__agg_compute__"
            and node.get("layer_key") == "data_middleware_storage"
        )
        edges = call_edges(elements)

        self.assertEqual(elb_node["layer_key"], "compute_app")
        self.assertEqual(
            [resource["name"] for resource in ecs_aggregate["summary_resources"]],
            ["ecs-backend"],
        )
        self.assertTrue(any(
            edge["source"] == elb_node["id"]
            and edge["target"] == ecs_aggregate["id"]
            for edge in edges
        ))

    def test_tidb_ecs_without_links_is_in_data_layer(self):
        records = [
            make_record(
                "cce-app", service_type="cce",
                targets="elb-tidb"
            ),
            make_record("ecs-tidb", service_type="ecs"),
            make_record("elb-tidb", service_type="elb", targets="ecs-tidb-backend"),
            make_record("ecs-tidb-backend", service_type="ecs"),
            make_record("ecs-tidb-standalone", service_type="ecs"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        data_aggregates = [
            node for node in nodes
            if node.get("type") == "__agg_compute__"
            and node.get("layer_key") == "data_middleware_storage"
        ]
        self.assertTrue(any(
            resource["name"] == "ecs-tidb"
            for node in data_aggregates
            for resource in node["summary_resources"]
        ))
        self.assertTrue(any(
            resource["name"] == "ecs-tidb-backend"
            for node in data_aggregates
            for resource in node["summary_resources"]
        ))
        self.assertTrue(any(
            resource["name"] == "ecs-tidb-standalone"
            for node in data_aggregates
            for resource in node["summary_resources"]
        ))

    def test_icon_loader_matches_aliases_and_future_service_pngs(self):
        png_bytes = b"\x89PNG\r\n\x1a\n"
        old_icon_dir = converter.ICON_DIR
        old_cache = converter._ICON_DATA_URI_CACHE

        with tempfile.TemporaryDirectory() as icon_dir:
            for filename in (
                "CCE_Deployment.png",
                "RDS.png",
                "VPC.png",
                "CC.png",
                "CNAD.png",
                "GeminiDB.png",
            ):
                with open(os.path.join(icon_dir, filename), "wb") as f:
                    f.write(png_bytes + filename.encode("ascii"))

            converter.ICON_DIR = icon_dir
            converter._ICON_DATA_URI_CACHE = None
            try:
                cce_icon = converter.get_service_icon_data_uri("cce")

                self.assertTrue(cce_icon.startswith("data:image/png;base64,"))
                self.assertEqual(
                    converter.get_service_icon_data_uri("cce_deployment"),
                    cce_icon,
                )
                self.assertEqual(
                    converter.get_service_icon_data_uri("cce-deployment"),
                    cce_icon,
                )
                self.assertTrue(converter.get_service_icon_data_uri("rds"))
                self.assertTrue(converter.get_service_icon_data_uri("vpc"))
                self.assertTrue(converter.get_service_icon_data_uri("cc"))
                self.assertTrue(converter.get_service_icon_data_uri("cnad"))
                self.assertTrue(converter.get_service_icon_data_uri("geminidb"))
            finally:
                converter.ICON_DIR = old_icon_dir
                converter._ICON_DATA_URI_CACHE = old_cache

    def test_external_k8s_targets_are_aggregated_in_virtual_business(self):
        target_names = [f"checkout-service-{i}" for i in range(1, 9)]
        records = [make_record("cce-cluster", targets="，".join(target_names))]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        k8s_business = next(
            node for node in nodes
            if node.get("type") == "__business__" and node.get("name") == "K8s"
        )
        external_nodes = [
            node for node in nodes
            if node.get("is_external") and not node.get("is_container")
        ]

        self.assertEqual(k8s_business["resource_total"], 8)
        self.assertEqual(len(external_nodes), 8)
        self.assertTrue(all(node["service_key"] == "k8s" for node in external_nodes))
        self.assertTrue(all(node["desc"] == "（外部/未列出服务）"
                            for node in external_nodes))
        self.assertFalse(any(node.get("is_summary") for node in nodes))
        self.assertEqual(len(edges), 1)
        self.assertEqual(edges[0]["call_count"], 8)
        self.assertEqual(edges[0]["cross_business"], 1)

    def test_external_k8s_and_maas_names_infer_shared_business(self):
        k8s_names = [f"gz-hw-lingee-prod-k8s-api-{i}" for i in range(1, 8)]
        maas_names = [f"gz-hw-lingee-prod-maas-model-{i}" for i in range(1, 8)]
        records = [make_record(
            "api-entry", business="入口业务", service_type="apig",
            targets=",".join(k8s_names + maas_names),
        )]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        business = next(
            node for node in nodes
            if node.get("type") == "__business__"
            and node.get("name") == "lingee-prod"
        )
        service_nodes = [
            node for node in nodes
            if node.get("type") == "__service__"
        ]
        node_by_id = {node["id"]: node for node in nodes}
        business_service_nodes = [
            node for node in service_nodes
            if node_by_id[node["parent"]]["parent"] == business["id"]
        ]
        external_nodes = [
            node for node in nodes
            if node.get("is_external") and node.get("business") == "lingee-prod"
        ]

        self.assertEqual(business["resource_total"], 14)
        self.assertEqual({node["service_key"] for node in business_service_nodes},
                         {"k8s", "maas"})
        self.assertEqual(len(external_nodes), 14)
        self.assertFalse(any(node.get("is_summary") for node in nodes))
        self.assertEqual(len(edges), 2)
        self.assertEqual({edge["call_count"] for edge in edges}, {7})

    def test_same_business_service_chain_keeps_confirmed_calls(self):
        records = [
            make_record("cce-prod", service_type="cce", targets="ECS"),
            make_record("ecs-prod", service_type="ecs",
                        targets="rds-prod，DCS；obs-prod"),
            make_record("rds-prod", service_type="rds"),
            make_record("dcs-prod", service_type="dcs"),
            make_record("obs-prod", service_type="obs", targets="elb-prod"),
            make_record("elb-prod", service_type="elb"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")
        node_ids = {
            node["name"]: node["id"] for node in nodes
            if not node.get("is_container") and node.get("type") != "__agg_compute__"
        }
        endpoints = {(edge["source"], edge["target"]) for edge in edges}

        self.assertIn((agg["id"], node_ids["rds-prod"]), endpoints)
        self.assertIn((agg["id"], node_ids["dcs-prod"]), endpoints)
        self.assertIn((agg["id"], node_ids["obs-prod"]), endpoints)
        self.assertIn((node_ids["obs-prod"], node_ids["elb-prod"]), endpoints)
        self.assertTrue(all(edge["cross_business"] == 0 for edge in edges))

    def test_inferred_type_and_business_references_connect_service_groups(self):
        records = [
            make_record(
                "ecs-billing", business="billing-prod", service_type="ecs",
                inferred_targets=(
                    "rds-billing-prod,dcs_billing_prod,"
                    "redis-billing-prod,dds_billing-prod"
                ),
            ),
            make_record("rds-main", business="billing-prod", service_type="rds_mysql"),
            make_record("dcs-main", business="billing-prod", service_type="dcs_redis"),
            make_record("dds-main", business="billing-prod", service_type="dds"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        node_ids = {node["name"]: node["id"] for node in nodes
                    if not node.get("is_container")}
        edge_by_target = {edge["target"]: edge for edge in edges}

        self.assertEqual(len(edges), 3)
        self.assertEqual(edge_by_target[node_ids["rds-main"]]["call_count"], 1)
        self.assertEqual(edge_by_target[node_ids["dcs-main"]]["call_count"], 2)
        self.assertEqual(edge_by_target[node_ids["dds-main"]]["call_count"], 1)
        self.assertTrue(all(edge["inferred"] == 1 for edge in edges))
        self.assertTrue(all(edge["cross_business"] == 0 for edge in edges))
        self.assertTrue(all(edge["color"] == "#8E44AD" for edge in edges))
        self.assertIn("（推断）", edge_by_target[node_ids["dcs-main"]]["relation"])

    def test_inferred_service_prefix_resolves_resource_and_external_target(self):
        records = [
            make_record(
                "api-entry", business="业务A", service_type="apig",
                inferred_targets="Service:payment-api；Service：external-k8s-api",
            ),
            make_record("payment-api", business="业务A", service_type="ecs"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")
        external = next(node for node in nodes
                        if node.get("name") == "external-k8s-api")
        edge_by_name = {edge["target_name"]: edge for edge in edges}

        self.assertEqual(len(edges), 2)
        self.assertEqual(edge_by_name["payment-api"]["target"], agg["id"])
        self.assertEqual(edge_by_name["external-k8s-api"]["target"], external["id"])
        self.assertTrue(external["is_external"])
        self.assertTrue(all(edge["inferred"] == 1 for edge in edges))

    def test_inferred_service_marker_anywhere_uses_source_business(self):
        records = [
            make_record(
                "ecs-ai", business="ai-prod", service_type="ecs",
                inferred_targets="ai-rds-primary,cache-dsc-main,platform_dds_node",
            ),
            make_record("rds-ai", business="ai-prod", service_type="rds"),
            make_record("dcs-ai", business="ai-prod", service_type="dcs"),
            make_record("dds-ai", business="ai-prod", service_type="dds"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        node_ids = {node["name"]: node["id"] for node in nodes
                    if not node.get("is_container")}

        self.assertEqual(len(edges), 3)
        self.assertEqual(
            {edge["target"] for edge in edges},
            {node_ids["rds-ai"], node_ids["dcs-ai"], node_ids["dds-ai"]},
        )
        self.assertTrue(all(edge["inferred"] == 1 for edge in edges))
        self.assertTrue(all(edge["cross_business"] == 0 for edge in edges))

    def test_inferred_reason_chain_connects_existing_business_services(self):
        records = [
            make_record(
                "vpc-main", service_type="vpc",
                reason=(
                    "网络分析；下游（推断）：VPC->数据->CCE=>消息=>DMS"
                    "→缓存->dcs-业务A->RDS。其他说明"
                ),
            ),
            make_record("cce-main", service_type="cce"),
            make_record("dms-main", service_type="dms"),
            make_record("dcs-main", service_type="dcs"),
            make_record("rds-main", service_type="rds"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = call_edges(elements)
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")
        node_ids = {
            node["name"]: node["id"] for node in nodes
            if not node.get("is_container") and node.get("type") != "__agg_compute__"
        }
        endpoints = {(edge["source"], edge["target"]) for edge in edges}

        self.assertEqual(len(edges), 4)
        self.assertIn((node_ids["vpc-main"], agg["id"]), endpoints)
        self.assertIn((agg["id"], node_ids["dms-main"]), endpoints)
        self.assertIn((node_ids["dms-main"], node_ids["dcs-main"]), endpoints)
        self.assertIn((node_ids["dcs-main"], node_ids["rds-main"]), endpoints)
        self.assertTrue(all(edge["inferred"] == 1 for edge in edges))
        self.assertFalse(any(node.get("name") in {"数据", "消息", "缓存"}
                             for node in nodes))

    def test_placeholder_filter_does_not_remove_real_names_containing_na(self):
        records = [
            make_record("ecs-prod", service_type="ecs", targets="finance-rds"),
            make_record("finance-rds", service_type="rds"),
        ]

        edges = call_edges(build_graph(records))

        self.assertEqual(len(edges), 1)
        self.assertEqual(edges[0]["target_name"], "finance-rds")

    def test_same_service_nodes_form_one_contiguous_layout_block(self):
        records = [
            make_record("cci-1", service_type="cci", targets="rds-1"),
            make_record("rds-1", service_type="rds", targets="cci-2"),
            make_record("cci-2", service_type="cci"),
            make_record("cci-3", service_type="cci", targets="rds-1"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        positions = compute_bdat_positions(elements)
        cci_nodes = [node for node in nodes if node.get("service_key") == "cci"
                     and not node.get("is_container")]
        rds_nodes = [node for node in nodes if node.get("service_key") == "rds"
                     and not node.get("is_container")]

        self.assertEqual(len({node["parent"] for node in cci_nodes}), 1)
        cci_y = [positions[node["id"]]["y"] for node in cci_nodes]
        rds_y = [positions[node["id"]]["y"] for node in rds_nodes]
        self.assertTrue(max(cci_y) < min(rds_y) or max(rds_y) < min(cci_y))

    def test_data_layer_services_are_split_into_three_rows(self):
        records = [
            make_record("ecs-rabbitmq-1", service_type="ecs"),
            make_record("rds-main", service_type="rds"),
            make_record("dcs-main", service_type="dcs"),
            make_record("obs-bucket", service_type="obs"),
        ]

        elements = build_graph(records)
        nodes = {
            node["name"]: node for node in element_data(elements, "nodes")
            if not node.get("is_container")
        }
        positions = compute_bdat_positions(elements)
        names = ["ecs-rabbitmq-1", "rds-main", "dcs-main", "obs-bucket"]

        self.assertTrue(all(nodes[name]["layer_key"] == "data_middleware_storage"
                            for name in names))
        ys = [positions[nodes[name]["id"]]["y"] for name in names]
        # 同一层内服务保持一行排开，互不重叠
        self.assertEqual(len(set(ys)), 3)

    def test_data_service_block_keeps_its_nodes_on_one_row(self):
        records = [
            make_record(f"obs-{i}", service_type="obs")
            for i in range(1, MAX_NODES_PER_ROW + 1)
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        positions = compute_bdat_positions(elements)
        obs_nodes = [
            node for node in nodes
            if node.get("service_key") == "obs" and not node.get("is_container")
        ]

        self.assertEqual(
            len({positions[node["id"]]["y"] for node in obs_nodes}), 1
        )

    def test_read_excel_keeps_rows_even_when_core_column_exists(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["服务名称", "服务类型", "所属业务", "是否核心业务"])
        ws.append(["core-ecs", "ecs", "业务A", "是"])
        ws.append(["non-core-rds", "rds", "业务A", ""])
        ws.append(["blank-core-cce", "cce", "业务A", None])

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "core-filter.xlsx")
            wb.save(path)
            records = read_excel(path)

        self.assertEqual([record["name"] for record in records],
                         ["core-ecs", "non-core-rds", "blank-core-cce"])
        self.assertEqual([record["is_core_business"] for record in records],
                         ["是", "", ""])

    def test_cce_deployment_and_sfs3_map_to_expected_layers(self):
        self.assertEqual(
            get_topology_layer("cce", "cce_Deployment-01", "cce_Deployment"),
            "compute_app",
        )
        self.assertEqual(
            get_topology_layer("sfs3", "sfs3-share", "sfs3"),
            "data_middleware_storage",
        )

    def test_topology_layer_names_use_new_text_without_prefix(self):
        self.assertEqual(
            converter.get_topology_layer_name("access"),
            "渠道接入与安全边界",
        )
        self.assertEqual(
            converter.get_topology_layer_name("network_lb"),
            "网络与负载均衡层",
        )
        self.assertEqual(
            converter.get_topology_layer_name("compute_app"),
            "计算、容器与应用服务层",
        )
        self.assertEqual(
            converter.get_topology_layer_name("data_middleware_storage"),
            "中间件、数据与存储层",
        )

    def test_bdat_business_grid_uses_four_columns(self):
        records = [
            make_record(f"ecs-{i}", business=f"业务{i}", service_type="ecs",
                        targets=f"ecs-{i + 1 if i < 5 else 1}")
            for i in range(1, 6)
        ]

        elements = build_graph(records)
        nodes = [
            node for node in element_data(elements, "nodes")
            if not node.get("is_container")
        ]
        positions = compute_bdat_positions(elements)
        ordered_positions = [positions[node["id"]] for node in nodes]

        self.assertEqual(len({pos["y"] for pos in ordered_positions[:4]}), 1)
        self.assertEqual(len({pos["x"] for pos in ordered_positions[:4]}), 4)
        self.assertGreater(ordered_positions[4]["y"], ordered_positions[0]["y"])

    def test_cce_deployment_aggregated_into_one_node(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-2", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj4-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-B"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]
        cce_leaves = [
            node for node in nodes
            if node.get("service_key") == "cce"
            and not node.get("is_container") and not node.get("is_summary")
            and node.get("type") != "__agg_compute__"
        ]

        # CCE_Deployment 全部聚合为一个缩略节点
        self.assertEqual(len(aggs), 1)
        self.assertEqual(aggs[0]["resource_total"], 3)
        self.assertEqual(cce_leaves, [])
        self.assertFalse(any(node.get("is_summary") for node in nodes))

    def test_cce_deployment_without_business_stays_in_virtual_business(self):
        records = [
            make_record("bj-prod-app-deploy-1", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj4-prod-app-deploy-1", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        business = next(node for node in nodes if node.get("type") == "__business__")
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")

        self.assertEqual(business["name"], "CCE")
        self.assertEqual(business["is_virtual_business"], 1)
        self.assertEqual(agg["resource_total"], 2)
        self.assertEqual(agg["parent"], business["id"])

    def test_cce_deployment_without_business_falls_back_to_project_business(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-2", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("gz-prod-app-deploy-1", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-Z"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        businesses = [node for node in nodes if node.get("type") == "__business__"]
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]
        agg_by_biz = {node["business"]: node for node in aggs}

        self.assertEqual({node["name"] for node in businesses}, {"业务A", "CCE"})
        self.assertEqual(agg_by_biz["业务A"]["resource_total"], 2)
        self.assertEqual(agg_by_biz["CCE"]["resource_total"], 1)

    def test_cce_fallback_requires_same_prefix(self):
        records = [
            make_record("bj-prod-app-deploy-1", business="业务A", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj4-prod-app-deploy-1", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]

        # 企业项目相同但前缀不同（bj4-prod-app ≠ bj-prod-app），不合并，
        # 无业务节点留在虚拟 CCE 业务。
        by_biz = {node["business"]: node for node in aggs}
        self.assertEqual(by_biz["业务A"]["resource_total"], 1)
        self.assertEqual(by_biz["CCE"]["resource_total"], 1)

    def test_cce_fallback_prefers_business_with_more_same_prefix_nodes(self):
        records = [
            make_record("bj-prod-app-deploy-1", business="业务A", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-2", business="业务A", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-web-1", business="业务B", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-3", business="", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]
        by_biz = {node["business"]: node for node in aggs}

        # 业务A 有 2 个同前缀节点，业务B 只有 1 个 → 无业务节点并入业务A。
        self.assertEqual(by_biz["业务A"]["resource_total"], 3)
        self.assertEqual(by_biz["业务B"]["resource_total"], 1)

    def test_cce_prefix_uses_three_segments_by_default(self):
        self.assertEqual(
            converter.extract_cce_business_prefix("bj-prod-app-deploy-01"),
            "bj-prod-app",
        )

    def test_cce_deployment_keeps_enterprise_project_field(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-id-1",
                        enterprise_project="生产项目"),
            make_record("bj-prod-app-deploy-2", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-id-1",
                        enterprise_project="生产项目"),
            make_record("bj-prod-app-deploy-3", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-id-2",
                        enterprise_project="测试项目"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")
        resources = agg["summary_resources"]
        first = next(r for r in resources if r["name"] == "bj-prod-app-deploy-1")

        self.assertEqual(first["enterprise_project"], "生产项目")
        self.assertEqual(first["project"], "ep-id-1")

    def test_core_business_only_nodes_are_drawn(self):
        records = [
            make_record("ecs-core-1", service_type="ecs", is_core_business="是",
                        targets="rds-core"),
            make_record("ecs-core-2", service_type="ecs", is_core_business="是",
                        targets="rds-core"),
            make_record("rds-core", service_type="rds", is_core_business="是"),
            make_record("rds-non-core", service_type="rds", is_core_business="否"),
            make_record("obs-empty-core", service_type="obs", is_core_business=""),
            make_record("elb-no-flag", service_type="elb", is_core_business=None),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        leaves = [
            node["name"] for node in nodes
            if not node.get("is_container") and not node.get("is_summary")
            and node.get("type") != "__agg_compute__"
        ]
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]

        # 存在“是否核心业务”列时，只有“是”的节点被绘制；
        # 无该列标记的行（None）等同于不存在该列，不参与。
        self.assertEqual(len(aggs), 1)
        self.assertEqual(aggs[0]["resource_total"], 2)
        self.assertIn("rds-core", leaves)
        self.assertNotIn("rds-non-core", leaves)
        self.assertNotIn("obs-empty-core", leaves)
        self.assertNotIn("elb-no-flag", leaves)

    def test_core_filter_drops_edges_and_external_nodes_for_filtered_names(self):
        records = [
            make_record("ecs-core-1", service_type="ecs", is_core_business="是",
                        targets="rds-non-core,elb-no-flag"),
            make_record("rds-non-core", service_type="rds", is_core_business="否"),
            make_record("elb-no-flag", service_type="elb", is_core_business=None),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        leaf_names = {
            node["name"] for node in nodes
            if not node.get("is_container") and not node.get("is_summary")
            and node.get("type") not in {"__agg_compute__", "__agg_count__"}
        }
        aggs = [node for node in nodes if node.get("type") == "__agg_compute__"]

        self.assertEqual(leaf_names, set())
        self.assertEqual(len(aggs), 1)
        self.assertEqual(aggs[0]["resource_total"], 1)
        # 被过滤节点不生成外部节点，也不产生连线
        self.assertEqual(edges, [])

    def test_cce_core_filter_draws_only_core_deployments(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A",
                        is_core_business="是"),
            make_record("bj-prod-app-deploy-2", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A",
                        is_core_business="是"),
            make_record("bj-prod-app-deploy-3", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A",
                        is_core_business="否"),
            make_record("bj4-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-B",
                        is_core_business="否"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        agg = next(node for node in nodes if node.get("type") == "__agg_compute__")

        # 聚合节点只统计核心部署
        self.assertEqual(agg["resource_total"], 2)

    def test_business_keyword_filter_merges_chat_businesses(self):
        records = [
            make_record("ecs-1", business="chat", service_type="ecs",
                        targets="rds-1,rds-2"),
            make_record("rds-1", business="chat", service_type="rds"),
            make_record("ecs-2", business="chat&推荐", service_type="ecs"),
            make_record("ecs-3", business="支付业务", service_type="ecs"),
            make_record("rds-2", business="支付业务", service_type="rds"),
        ]
        converter.BUSINESS_FILTER_KEYWORD = "chat"
        try:
            elements = build_graph(records)
        finally:
            converter.BUSINESS_FILTER_KEYWORD = self._old_biz_filter

        nodes = element_data(elements, "nodes")
        businesses = [node for node in nodes if node.get("type") == "__business__"]
        names = {node["name"] for node in nodes}
        chat_nodes = [
            node for node in nodes
            if not node.get("is_container")
            and node.get("name") in {"ecs-1", "rds-1", "ecs-2"}
        ]

        self.assertEqual([node["name"] for node in businesses], ["chat"])
        self.assertEqual(businesses[0]["business_key"], "__business__:chat")
        self.assertEqual(businesses[0]["resource_total"], 3)
        self.assertTrue(all(node["business"] == "chat" for node in chat_nodes))
        self.assertEqual(
            next(node for node in chat_nodes if node["name"] == "ecs-2")["source_business"],
            "chat&推荐",
        )
        self.assertIn("rds-1", names)
        # 非 chat 业务的节点不绘制，也不生成外部节点
        self.assertNotIn("rds-2", names)
        self.assertNotIn("ecs-3", names)

    def test_read_excel_detects_enterprise_project_name_column(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["服务名称", "服务类型", "企业项目", "企业项目ID"])
        ws.append(["cce-1", "cce", "生产项目", "ep-id-1"])
        ws.append(["cce-2", "cce", "测试项目", "ep-id-2"])

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "ep-name.xlsx")
            wb.save(path)
            records = read_excel(path)

        self.assertEqual([record["enterprise_project"] for record in records],
                         ["生产项目", "测试项目"])
        self.assertEqual([record["enterprise_id"] for record in records],
                         ["ep-id-1", "ep-id-2"])

    def test_non_cce_deployment_resources_keep_business_column_grouping(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        business="业务A", enterprise_id="ep-A"),
            make_record("bj4-prod-app-deploy-2", service_type="cce",
                        business="业务A", enterprise_id="ep-B"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        businesses = [node for node in nodes if node.get("type") == "__business__"]
        projects = [node for node in nodes if node.get("type") == "__cce_project__"]

        self.assertEqual([node["name"] for node in businesses], ["业务A"])
        self.assertEqual(projects, [])


if __name__ == "__main__":
    unittest.main()
