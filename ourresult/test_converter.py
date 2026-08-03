import os
import tempfile
import unittest

import converter
import openpyxl

from converter import (
    ICON_SIZE,
    MAX_VISIBLE_SERVICE_NODES,
    _make_cytoscape_style,
    build_graph,
    compute_bdat_positions,
    get_service_key,
    get_topology_layer,
    read_excel,
)


def make_record(name, business="业务A", service_type="cce_deploym", targets="",
                inferred_targets="", reason="", source="", enterprise_id="",
                group=""):
    return {
        "name": name,
        "type": service_type,
        "resource_id": f"id-{name}",
        "enterprise_id": enterprise_id,
        "region": "cn-test-1",
        "group": group,
        "business": business,
        "desc": "",
        "spec": "",
        "source": source,
        "targets": targets,
        "inferred_targets": inferred_targets,
        "reason": reason,
    }


def element_data(elements, group):
    return [element["data"] for element in elements if element["group"] == group]


class BuildGraphAggregationTests(unittest.TestCase):
    def test_long_resource_name_uses_two_line_compact_label(self):
        long_name = "ecs-production-payment-service-with-a-very-long-instance-name"

        elements = build_graph([
            make_record(long_name, service_type="ecs")
        ])
        node = next(data for data in element_data(elements, "nodes")
                    if not data.get("is_container"))

        label_lines = node["display_label"].splitlines()
        self.assertEqual(label_lines[0], "ECS")
        self.assertEqual(len(label_lines), 2)
        self.assertTrue(label_lines[1].endswith("..."))
        self.assertEqual(node["name"], long_name)

        custom_elements = build_graph([
            make_record("node-1", service_type="extremely_long_custom_cloud_service")
        ])
        custom_node = next(data for data in element_data(custom_elements, "nodes")
                           if not data.get("is_container"))
        self.assertTrue(custom_node["display_label"].splitlines()[0].endswith("..."))

    def test_collapses_resources_after_first_five(self):
        records = [make_record(f"cce-node-{i}") for i in range(1, 1004)]
        records[5]["enterprise_id"] = "ep-hidden"
        records[5]["group"] = "hidden-group"
        records[5]["source"] = "inventory-sheet"
        records[5]["targets"] = "cce-node-1,cce-node-2"

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        resource_nodes = [node for node in nodes if not node.get("is_container")]
        service = next(node for node in nodes if node.get("type") == "__service__")
        summary = next(node for node in nodes if node.get("is_summary"))

        self.assertEqual(service["name"], "CCE")
        self.assertEqual(service["resource_total"], 1003)
        self.assertEqual(len(resource_nodes), MAX_VISIBLE_SERVICE_NODES + 1)
        self.assertEqual(summary["name"], "...+998")
        self.assertEqual(summary["collapsed_count"], 998)
        self.assertEqual(summary["parent"], service["id"])
        self.assertEqual(len(summary["summary_resources"]), 998)
        first_hidden = summary["summary_resources"][0]
        self.assertEqual(first_hidden["name"], "cce-node-6")
        self.assertEqual(first_hidden["type"], "CCE_DEPLOYM")
        self.assertEqual(first_hidden["project"], "ep-hidden")
        self.assertEqual(first_hidden["business"], "业务A")
        self.assertEqual(first_hidden["group"], "hidden-group")
        self.assertEqual(first_hidden["source"], "inventory-sheet")
        self.assertEqual(first_hidden["downstream"], ["cce-node-1", "cce-node-2"])

    def test_cross_business_call_targets_first_resource_in_service_group(self):
        records = [make_record("api-1", business="业务A", service_type="apig",
                               targets="cce-node-6")]
        records.extend(make_record(f"cce-node-{i}", business="业务B")
                       for i in range(1, 7))

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edge = element_data(elements, "edges")[0]
        first_cce = next(node for node in nodes if node.get("name") == "cce-node-1")

        self.assertEqual(edge["target"], first_cce["id"])
        self.assertEqual(edge["target_name"], "cce-node-6")
        self.assertEqual(edge["cross_business"], 1)

    def test_same_business_calls_from_and_to_hidden_nodes_use_summary(self):
        records = [make_record("cce-node-1", targets="cce-node-6")]
        records.extend(make_record(f"cce-node-{i}") for i in range(2, 6))
        records.append(make_record("cce-node-6", targets="cce-node-1"))

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        edges = element_data(elements, "edges")
        summary = next(node for node in nodes if node.get("is_summary"))
        first_cce = next(node for node in nodes if node.get("name") == "cce-node-1")

        endpoints = {(edge["source"], edge["target"]) for edge in edges}
        self.assertIn((first_cce["id"], summary["id"]), endpoints)
        self.assertIn((summary["id"], first_cce["id"]), endpoints)
        self.assertTrue(all(edge["cross_business"] == 0 for edge in edges))

    def test_five_nodes_stay_on_one_layout_row_without_summary(self):
        records = [make_record(f"cce-node-{i}")
                   for i in range(1, MAX_VISIBLE_SERVICE_NODES + 1)]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        positions = compute_bdat_positions(elements)
        resource_ids = [node["id"] for node in nodes if not node.get("is_container")]

        self.assertFalse(any(node.get("is_summary") for node in nodes))
        self.assertEqual(len({positions[node_id]["y"] for node_id in resource_ids}), 1)
        self.assertEqual(len({positions[node_id]["x"] for node_id in resource_ids}), 5)

    def test_records_without_calls_do_not_create_edges(self):
        elements = build_graph([make_record(f"cce-node-{i}") for i in range(1, 8)])

        self.assertEqual(element_data(elements, "edges"), [])

    def test_unassigned_resources_get_service_business_and_aggregation(self):
        records = [make_record(f"waf-node-{i}", business="", service_type="waf")
                   for i in range(1, 9)]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        business = next(node for node in nodes if node.get("type") == "__business__")
        resources = [node for node in nodes if not node.get("is_container")]
        summary = next(node for node in resources if node.get("is_summary"))

        self.assertEqual(business["name"], "WAF")
        self.assertEqual(business["is_virtual_business"], 1)
        self.assertEqual(business["resource_total"], 8)
        self.assertFalse(any(node.get("type") == "__service__" for node in nodes))
        self.assertEqual(len(resources), MAX_VISIBLE_SERVICE_NODES + 1)
        self.assertEqual(summary["name"], "...+3")
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

    def test_icon_loader_matches_aliases_and_future_service_pngs(self):
        png_bytes = b"\x89PNG\r\n\x1a\n"
        old_icon_dir = converter.ICON_DIR
        old_cache = converter._ICON_DATA_URI_CACHE

        with tempfile.TemporaryDirectory() as icon_dir:
            for filename in ("CCE_Deployment.png", "RDS.png", "VPC.png"):
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
        summary = next(
            node for node in nodes
            if node.get("is_summary") and node.get("business_key")
            == k8s_business["business_key"]
        )

        self.assertEqual(k8s_business["resource_total"], 8)
        self.assertEqual(len(external_nodes), MAX_VISIBLE_SERVICE_NODES)
        self.assertTrue(all(node["service_key"] == "k8s" for node in external_nodes))
        self.assertTrue(all(node["desc"] == "（外部/未列出服务）"
                            for node in external_nodes))
        self.assertEqual(summary["name"], "...+3")
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
        summaries = [
            node for node in nodes
            if node.get("is_summary") and node.get("business") == "lingee-prod"
        ]
        external_nodes = [
            node for node in nodes
            if node.get("is_external") and node.get("business") == "lingee-prod"
        ]

        self.assertEqual(business["resource_total"], 14)
        self.assertEqual({node["service_key"] for node in business_service_nodes},
                         {"k8s", "maas"})
        self.assertEqual(len(external_nodes), 10)
        self.assertEqual(len(summaries), 2)
        self.assertEqual({node["name"] for node in summaries}, {"...+2"})
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
        edges = element_data(elements, "edges")
        node_ids = {node["name"]: node["id"] for node in nodes
                    if not node.get("is_container")}
        endpoints = {(edge["source"], edge["target"]) for edge in edges}

        self.assertIn((node_ids["cce-prod"], node_ids["ecs-prod"]), endpoints)
        self.assertIn((node_ids["ecs-prod"], node_ids["rds-prod"]), endpoints)
        self.assertIn((node_ids["ecs-prod"], node_ids["dcs-prod"]), endpoints)
        self.assertIn((node_ids["ecs-prod"], node_ids["obs-prod"]), endpoints)
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
        edges = element_data(elements, "edges")
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
        edges = element_data(elements, "edges")
        payment = next(node for node in nodes if node.get("name") == "payment-api")
        external = next(node for node in nodes
                        if node.get("name") == "external-k8s-api")
        edge_by_name = {edge["target_name"]: edge for edge in edges}

        self.assertEqual(len(edges), 2)
        self.assertEqual(edge_by_name["payment-api"]["target"], payment["id"])
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
        edges = element_data(elements, "edges")
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
        edges = element_data(elements, "edges")
        node_ids = {node["name"]: node["id"] for node in nodes
                    if not node.get("is_container")}
        endpoints = {(edge["source"], edge["target"]) for edge in edges}

        self.assertEqual(len(edges), 4)
        self.assertIn((node_ids["vpc-main"], node_ids["cce-main"]), endpoints)
        self.assertIn((node_ids["cce-main"], node_ids["dms-main"]), endpoints)
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

        edges = element_data(build_graph(records), "edges")

        self.assertEqual(len(edges), 1)
        self.assertEqual(edges[0]["target_name"], "finance-rds")

    def test_same_service_nodes_form_one_contiguous_layout_block(self):
        records = [
            make_record("ecs-1", service_type="ecs", targets="rds-1"),
            make_record("rds-1", service_type="rds", targets="ecs-2"),
            make_record("ecs-2", service_type="ecs"),
            make_record("ecs-3", service_type="ecs"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        positions = compute_bdat_positions(elements)
        ecs_nodes = [node for node in nodes if node.get("service_key") == "ecs"
                     and not node.get("is_container")]
        rds_nodes = [node for node in nodes if node.get("service_key") == "rds"
                     and not node.get("is_container")]

        self.assertEqual(len({node["parent"] for node in ecs_nodes}), 1)
        ecs_y = [positions[node["id"]]["y"] for node in ecs_nodes]
        rds_y = [positions[node["id"]]["y"] for node in rds_nodes]
        self.assertTrue(max(ecs_y) < min(rds_y) or max(rds_y) < min(ecs_y))

    def test_services_in_same_layer_are_laid_out_horizontally(self):
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
        self.assertEqual(len({positions[nodes[name]["id"]]["y"]
                              for name in names}), 1)
        self.assertEqual(len({positions[nodes[name]["id"]]["x"]
                              for name in names}), 4)

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

    def test_bdat_business_grid_uses_four_columns(self):
        records = [
            make_record(f"ecs-{i}", business=f"业务{i}", service_type="ecs")
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

    def test_cce_deployment_resources_group_by_name_prefix(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-2", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-B"),
            make_record("bj4-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        businesses = [node for node in nodes if node.get("type") == "__business__"]
        leaf_by_name = {
            node["name"]: node for node in nodes
            if not node.get("is_container") and not node.get("is_summary")
        }

        self.assertEqual(
            {node["name"] for node in businesses},
            {"bj-prod-app", "bj4-prod-app"},
        )
        self.assertEqual(leaf_by_name["bj-prod-app-deploy-1"]["business"],
                         "bj-prod-app")
        self.assertEqual(leaf_by_name["bj-prod-app-deploy-1"]["source_business"],
                         "业务A")
        self.assertEqual(leaf_by_name["bj4-prod-app-deploy-1"]["business"],
                         "bj4-prod-app")

    def test_cce_deployment_resources_split_by_enterprise_project(self):
        records = [
            make_record("bj-prod-app-deploy-1", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-2", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-A"),
            make_record("bj-prod-app-deploy-3", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-B"),
            make_record("bj-prod-app-deploy-4", service_type="cce",
                        group="CCE_Deployment", enterprise_id="ep-B"),
            make_record("bj-prod-app-deploy-5", service_type="cce",
                        group="CCE_Deployment", enterprise_id=""),
        ]

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        node_by_id = {node["id"]: node for node in nodes}
        projects = [node for node in nodes if node.get("type") == "__project__"]
        leaves = [
            node for node in nodes
            if not node.get("is_container") and not node.get("is_summary")
        ]

        self.assertEqual(
            {node["name"] for node in projects},
            {"ep-A", "ep-B", "(未设置企业项目)"},
        )
        for leaf in leaves:
            parent = node_by_id[leaf["parent"]]
            self.assertEqual(parent["type"], "__project__")
            self.assertEqual(parent["name"], leaf["enterprise_project"]
                             or "(未设置企业项目)")
            self.assertEqual(parent["parent"],
                             node_by_id[leaf["service_container"]]["id"])

    def test_cce_deployment_summarizes_per_enterprise_project(self):
        records = []
        for index in range(8):
            records.append(make_record(
                f"bj-prod-app-deploy-a-{index}", service_type="cce",
                group="CCE_Deployment", enterprise_id="ep-A",
            ))
            records.append(make_record(
                f"bj-prod-app-deploy-b-{index}", service_type="cce",
                group="CCE_Deployment", enterprise_id="ep-B",
            ))

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        projects = {
            node["name"]: node for node in nodes
            if node.get("type") == "__project__"
        }
        summaries = [node for node in nodes if node.get("is_summary")]
        project_ids = {node["id"] for node in nodes
                       if node.get("type") == "__project__"}

        self.assertEqual(set(projects), {"ep-A", "ep-B"})
        self.assertEqual(projects["ep-A"]["visible_count"],
                         MAX_VISIBLE_SERVICE_NODES)
        self.assertEqual(projects["ep-A"]["collapsed_count"], 3)
        self.assertEqual(len(summaries), 2)
        self.assertTrue(all(summary["parent"] in project_ids
                            for summary in summaries))

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
        projects = [node for node in nodes if node.get("type") == "__project__"]

        self.assertEqual([node["name"] for node in businesses], ["业务A"])
        self.assertEqual(projects, [])


if __name__ == "__main__":
    unittest.main()
