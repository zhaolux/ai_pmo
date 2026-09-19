from ai_pmo.config import load_json


def test_cost_rate_is_consistent():
    project = load_json("project.json")
    assert project["hourly_rate"] == project["person_day_rate"] / project["hours_per_day"]


def test_weekly_capacity_is_40_hours():
    assert load_json("project.json")["hours_per_week"] == 40

