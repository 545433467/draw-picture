import unittest

from converter import MAX_VISIBLE_SERVICE_NODES, build_graph, compute_bdat_positions


def make_record(name, business="业务A", service_type="cce_deploym", targets=""):
    return {
        "name": name,
        "type": service_type,
        "resource_id": f"id-{name}",
        "enterprise_id": "",
        "region": "cn-test-1",
        "group": "",
        "business": business,
        "desc": "",
        "spec": "",
        "targets": targets,
        "inferred_targets": "",
        "reason": "",
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

        elements = build_graph(records)
        nodes = element_data(elements, "nodes")
        resource_nodes = [node for node in nodes if not node.get("is_container")]
        service = next(node for node in nodes if node.get("type") == "__service__")
        summary = next(node for node in nodes if node.get("is_summary"))

        self.assertEqual(service["name"], "CCE_DEPLOYM")
        self.assertEqual(service["resource_total"], 1003)
        self.assertEqual(len(resource_nodes), MAX_VISIBLE_SERVICE_NODES + 1)
        self.assertEqual(summary["name"], "...+998")
        self.assertEqual(summary["collapsed_count"], 998)
        self.assertEqual(summary["parent"], service["id"])

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


if __name__ == "__main__":
    unittest.main()
