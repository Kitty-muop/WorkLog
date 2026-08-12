import unittest
from datetime import datetime, date
import tools.gamify as gm
import tools.timer as tm
import tools.discord_bot as db

class TestRemindersAndEstimate(unittest.TestCase):
    def test_estimate_default_fallback(self):
        """Test that estimate defaults to 7.5 when unspecified or invalid."""
        est_val = tm.parse_estimate(None)
        self.assertEqual(est_val, 7.5)
        est_str = tm.parse_estimate("invalid")
        self.assertEqual(est_str, 7.5)
        est_valid = tm.parse_estimate("3.5")
        self.assertEqual(est_valid, 3.5)

    def test_reminder_deduplication_key(self):
        """Test that reminders fire exactly once per user per window per day."""
        sent_dict = {}
        u_id = "123456789"
        d_str = "2026-08-12"
        w_key = "0830"

        # First trigger should return True (should send)
        can_send1 = db.check_reminder_dedup(sent_dict, u_id, d_str, w_key)
        self.assertTrue(can_send1)

        # Mark sent
        db.mark_reminder_sent(sent_dict, u_id, d_str, w_key)

        # Second trigger in same minute/window should return False (do not duplicate)
        can_send2 = db.check_reminder_dedup(sent_dict, u_id, d_str, w_key)
        self.assertFalse(can_send2)

    def test_actual_duration_summation(self):
        """Test that get_subtask_actual_hours sums past entries correctly."""
        entries = [
            {"task": "Task A", "duration": 2.5},
            {"task": "Task A", "duration": 1.5},
            {"task": "Task B", "duration": 4.0},
        ]
        total_a = tm.sum_actual_hours(entries, "Task A")
        self.assertEqual(total_a, 4.0)

if __name__ == '__main__':
    unittest.main()
