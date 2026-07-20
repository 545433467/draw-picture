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


if __name__ == "__main__":
    unittest.main()
