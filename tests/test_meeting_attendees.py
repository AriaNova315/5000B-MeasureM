import unittest
from datetime import date

from app.services.meeting_narratives import meeting_defaults, task_resource_names


def resource(name):
    return {"name": name}


class MeetingAttendeeTests(unittest.TestCase):
    def setUp(self):
        self.project = {
            "manager": "负责人",
            "team_members": "项目成员甲,项目成员乙",
        }
        self.phase = {
            "start_date": date(2026, 1, 1),
            "end_date": date(2026, 1, 31),
        }

    def test_regular_meeting_uses_resources_of_matching_activity_and_date(self):
        tasks = [
            {
                "name": "双周例会",
                "start_date": date(2026, 1, 12),
                "end_date": date(2026, 1, 12),
                "resources": [resource("甲"), resource("乙"), resource("甲")],
            },
            {
                "name": "双周例会",
                "start_date": date(2026, 1, 26),
                "end_date": date(2026, 1, 26),
                "resources": [resource("丙")],
            },
        ]

        defaults = meeting_defaults(
            "meeting_regular",
            self.project,
            self.phase,
            {"date": date(2026, 1, 12), "type": "双周"},
            tasks,
        )

        self.assertEqual("甲,乙", defaults["attendees"])

    def test_phase_meeting_uses_phase_meeting_task_resources(self):
        tasks = [
            {
                "name": "阶段例会",
                "start_date": date(2026, 1, 31),
                "end_date": date(2026, 1, 31),
                "resources": [resource("质量人员"), resource("配置人员"), resource("负责人")],
            },
        ]

        defaults = meeting_defaults(
            "meeting_regular",
            self.project,
            self.phase,
            {"date": date(2026, 1, 31), "node_type": "阶段"},
            tasks,
        )

        self.assertEqual("质量人员,配置人员,负责人", defaults["attendees"])

    def test_special_meetings_keep_using_their_own_activity_resources(self):
        tasks = [
            {
                "name": "项目启动会",
                "start_date": date(2026, 1, 1),
                "end_date": date(2026, 1, 1),
                "resources": [resource("启动人员甲"), resource("启动人员乙")],
            },
            {
                "name": "项目策划会",
                "start_date": date(2026, 1, 2),
                "end_date": date(2026, 1, 2),
                "resources": [resource("策划人员")],
            },
        ]

        launch = meeting_defaults(
            "meeting_lh001", self.project, self.phase,
            {"date": date(2026, 1, 31), "type": "阶段"}, tasks,
        )
        planning = meeting_defaults(
            "meeting_planning", self.project, self.phase,
            {"date": date(2026, 1, 31), "type": "阶段"}, tasks,
        )

        self.assertEqual("启动人员甲,启动人员乙", launch["attendees"])
        self.assertEqual("策划人员", planning["attendees"])

    def test_missing_meeting_activity_keeps_compatible_project_fallback(self):
        defaults = meeting_defaults(
            "meeting_regular",
            self.project,
            self.phase,
            {"date": date(2026, 1, 31), "type": "阶段"},
            [],
        )

        self.assertEqual("项目成员甲,项目成员乙", defaults["attendees"])

    def test_task_resource_names_ignores_blank_names_and_deduplicates(self):
        names = task_resource_names(
            [{
                "name": "阶段例会",
                "start_date": date(2026, 1, 31),
                "end_date": date(2026, 1, 31),
                "resources": [
                    resource("甲"), resource(" "), resource("乙"), resource("甲"),
                ],
            }],
            "阶段例会",
            date(2026, 1, 31),
        )

        self.assertEqual("甲,乙", names)


if __name__ == "__main__":
    unittest.main()
