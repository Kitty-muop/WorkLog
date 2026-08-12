import unittest
import tools.gamify as gm

class TestGamifyDynamicExp(unittest.TestCase):
    def test_level_thresholds_generation(self):
        """Test that generate_level_thresholds produces 101 levels (0 to 100)."""
        thresholds = gm.generate_level_thresholds()
        self.assertEqual(len(thresholds), 101)
        self.assertEqual(thresholds[0], 0)
        self.assertEqual(thresholds[20], 1000)
        self.assertEqual(thresholds[40], 3000)
        self.assertEqual(thresholds[60], 7000)
        self.assertEqual(thresholds[80], 15000)
        self.assertEqual(thresholds[100], 31000)

    def test_get_tier_name(self):
        """Test 5 level milestones mapping."""
        self.assertEqual(gm.get_tier_name(0), "Novice")
        self.assertEqual(gm.get_tier_name(15), "Novice")
        self.assertEqual(gm.get_tier_name(25), "Adventurer")
        self.assertEqual(gm.get_tier_name(50), "Expert")
        self.assertEqual(gm.get_tier_name(75), "Master")
        self.assertEqual(gm.get_tier_name(95), "Grandmaster Legend")

    def test_calculate_task_exp_on_time(self):
        """Test task finished on time (Actual <= Estimated)."""
        exp, is_debug = gm.calculate_task_exp(estimated_h=10.0, actual_h=8.0, category="Development")
        # Base: 10 * 20 = 200. Ratio = 0.8. Bonus factor = 1 + (1 - 0.8)*0.5 = 1.1. EXP = 220
        self.assertGreater(exp, 200)
        self.assertFalse(is_debug)

    def test_calculate_task_exp_overtime_penalty(self):
        """Test task over estimate (Actual > Estimated)."""
        exp, _ = gm.calculate_task_exp(estimated_h=10.0, actual_h=20.0, category="Development")
        # Base: 200. Penalty ratio: 10/20 = 0.5. EXP = 100
        self.assertEqual(exp, 100)

    def test_debug_category_on_time_bonus(self):
        """Test Debug category finished on time receives 30% bonus."""
        exp_dev, _ = gm.calculate_task_exp(estimated_h=10.0, actual_h=10.0, category="Development")
        exp_debug, is_debug = gm.calculate_task_exp(estimated_h=10.0, actual_h=10.0, category="Debug / Bug Fix")
        self.assertTrue(is_debug)
        self.assertAlmostEqual(exp_debug, exp_dev * 1.3, delta=1.0)

    def test_debug_category_overtime_heavy_penalty(self):
        """Test Debug category going over estimate suffers heavy penalty."""
        exp_dev, _ = gm.calculate_task_exp(estimated_h=10.0, actual_h=20.0, category="Development")
        exp_debug, is_debug = gm.calculate_task_exp(estimated_h=10.0, actual_h=20.0, category="Debug / Bug Fix")
        self.assertTrue(is_debug)
        self.assertLess(exp_debug, exp_dev)

if __name__ == '__main__':
    unittest.main()
