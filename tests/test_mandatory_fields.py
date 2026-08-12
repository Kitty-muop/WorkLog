import unittest
import tools.timer as tm

class TestMandatoryFields(unittest.TestCase):
    def test_cmd_start_missing_mandatory_fields(self):
        """Test that cmd_start fails when mandatory fields are missing."""
        # Dummy args with missing project/task/subtask/category/estimate
        class DummyArgs:
            project = None
            task = None
            subtask = None
            category = None
            estimate = None
            user_id = "12345"
            user_name = "TestUser"

        ret = tm.cmd_start(DummyArgs())
        self.assertEqual(ret, 1)

    def test_cmd_sub_add_missing_mandatory_fields(self):
        """Test that cmd_sub_add fails when project or estimate is missing."""
        class DummyArgs:
            subtask = "TestSub"
            task = "TestTask"
            project = None
            estimate = None

        ret = tm.cmd_sub_add(DummyArgs())
        self.assertEqual(ret, 1)

if __name__ == '__main__':
    unittest.main()
