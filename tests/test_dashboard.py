import unittest
import os
from pathlib import Path
import tools.gen_dashboard as gd

class TestDashboard(unittest.TestCase):
    def test_extract_dashboard_data(self):
        data = gd.extract_dashboard_data()
        self.assertIn("generated_at", data)
        self.assertIn("rpg_stats", data)
        self.assertIn("daily_hours", data)
        self.assertIn("categories", data)
        self.assertIn("estimate_vs_actual", data)
        self.assertIn("projects_tree", data)

        rpg = data["rpg_stats"]
        self.assertIn("hero_rank", rpg)
        self.assertIn("level", rpg)
        self.assertIn("total_exp", rpg)

    def test_generate_html(self):
        output_path = gd.generate_html()
        self.assertTrue(output_path.exists())
        content = output_path.read_text(encoding="utf-8")
        self.assertIn("WorkLog — RPG Performance Dashboard", content)
        self.assertNotIn("__DATA_PLACEHOLDER__", content)

if __name__ == '__main__':
    unittest.main()
