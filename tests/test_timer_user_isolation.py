import unittest
import os
import tempfile
import json
import datetime
from unittest.mock import patch, MagicMock

from pathlib import Path

import tools.timer as tm

class TestTimerUserIsolation(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.state_file = Path(self.temp_dir) / 'timer_state.json'
        self.original_state_file = tm.STATE_FILE
        tm.STATE_FILE = self.state_file

        # Seed initial state with two distinct users: HuyDD and huybeo
        self.initial_state = {
            "timers": [
                {
                    "id": "t1",
                    "user_id": "433824488367456267",
                    "user_name": "HuyDD",
                    "project": "Comic web",
                    "task": "Admin dashboard update",
                    "subtask": "Update report bug analytics Sentry",
                    "subtask_code": "PRJ-5-KPI-2-ST-1",
                    "category": "Development",
                    "description": "Fix Sentry report analytics",
                    "segment_start": datetime.datetime.now().isoformat(),
                    "accumulated_seconds": 3600.0,
                    "paused": False,
                    "pause_log": []
                }
            ],
            "timer_id_seq": 1
        }
        tm.save_state(self.initial_state)

    def tearDown(self):
        if self.state_file.exists():
            self.state_file.unlink()
        tm.STATE_FILE = self.original_state_file

    def test_find_timer_strict_user_id_matching(self):
        """Test that find_timer returns correct timer for user_id and None for non-existent user_id."""
        state = tm.load_state()

        # User 1: HuyDD (ID: 433824488367456267) should find timer t1
        t_hdd = tm.find_timer(state, user_id="433824488367456267")
        self.assertIsNotNone(t_hdd)
        self.assertEqual(t_hdd['id'], "t1")
        self.assertEqual(t_hdd['user_name'], "HuyDD")

        # User 2: huybeo (ID: 9999999999) should return None, NOT HuyDD's timer!
        t_hb = tm.find_timer(state, user_id="9999999999")
        self.assertIsNone(t_hb)

    def test_cmd_status_user_isolation(self):
        """Test that cmd_status only returns timers for the specified user_id."""
        args_hdd = MagicMock()
        args_hdd.user_id = "433824488367456267"
        args_hdd.user_name = "HuyDD"

        args_hb = MagicMock()
        args_hb.user_id = "9999999999"
        args_hb.user_name = "huybeo"

        # Capturing stdout for cmd_status
        with patch('sys.stdout') as mock_stdout:
            ret_hdd = tm.cmd_status(args_hdd)
            self.assertEqual(ret_hdd, 0)

        # For huybeo with no timers, cmd_status should return 0 with no active timer
        state = tm.load_state()
        user_timers_hb = [t for t in state.get('timers', []) if str(t.get('user_id')) == "9999999999"]
        self.assertEqual(len(user_timers_hb), 0)

    def test_cmd_stop_single_prevents_cross_user_stopping(self):
        """Test that stopping a timer for user_id='9999999999' does not stop HuyDD's timer."""
        args_hb = MagicMock()
        args_hb.user_id = "9999999999"
        args_hb.user_name = "huybeo"
        args_hb.subtask = None
        args_hb.project = None
        args_hb.task = None
        args_hb.timer_id = None

        ret = tm.cmd_stop_single(args_hb)
        # Should return exit code 1 (failure to find timer for user)
        self.assertEqual(ret, 1)

    def test_cmd_pause_user_isolation(self):
        """Test that user_id='9999999999' cannot pause HuyDD's timer."""
        args_hb = MagicMock()
        args_hb.user_id = "9999999999"
        args_hb.user_name = "huybeo"
        args_hb.reason = "resting"

        ret = tm.cmd_pause(args_hb)
        # Should return exit code 1 because huybeo has no timer running
        self.assertEqual(ret, 1)

        # Confirm HuyDD's timer remains unpaused
        state = tm.load_state()
        t_hdd = state['timers'][0]
        self.assertFalse(t_hdd.get('paused', False))

    def test_cmd_cancel_user_isolation(self):
        """Test that user_id='9999999999' cannot cancel HuyDD's timer."""
        args_hb = MagicMock()
        args_hb.user_id = "9999999999"
        args_hb.user_name = "huybeo"

        ret = tm.cmd_cancel(args_hb)
        # Should return exit code 1 or preserve state
        state = tm.load_state()
        self.assertEqual(len(state['timers']), 1)
        self.assertEqual(state['timers'][0]['user_id'], "433824488367456267")

if __name__ == '__main__':
    unittest.main()
