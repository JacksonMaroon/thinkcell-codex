"""Office-free unit checks for render-only carrier identity remapping."""
import ast
import hashlib
from pathlib import Path
import unittest


SOURCE = Path(__file__).with_name("merge_sequence_graphs.py")


def load_functions():
    tree = ast.parse(SOURCE.read_text(encoding="utf-8"))
    names = {"fresh_tag", "render_tag_map"}
    body = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    module = ast.Module(body=body, type_ignores=[])
    namespace = {"hashlib": hashlib, "fail": lambda message: (_ for _ in ()).throw(ValueError(message))}
    exec(compile(module, str(SOURCE), "exec"), namespace)
    return namespace


class RenderTagMapTests(unittest.TestCase):
    def setUp(self):
        self.functions = load_functions()

    def test_rewrites_render_only_collision(self):
        mapping = self.functions["render_tag_map"](
            {"receiver-chart", "shared-anchor"},
            ["donor-chart", "shared-anchor"],
            {},
        )
        self.assertEqual(set(mapping), {"shared-anchor"})
        self.assertNotIn(mapping["shared-anchor"], {"receiver-chart", "shared-anchor", "donor-chart"})

    def test_reuses_model_identity_remap_and_rejects_duplicate_donor_tags(self):
        mapping = self.functions["render_tag_map"]({"shared"}, ["shared"], {"shared": "model-remapped"})
        self.assertEqual(mapping, {"shared": "model-remapped"})
        with self.assertRaisesRegex(ValueError, "donor render tags are not unique"):
            self.functions["render_tag_map"]({"a"}, ["a", "a"], {})


if __name__ == "__main__":
    unittest.main()
