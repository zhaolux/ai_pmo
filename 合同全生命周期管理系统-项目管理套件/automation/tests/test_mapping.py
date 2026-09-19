from ai_pmo.config import load_json


def test_every_sheet_is_mapped_once():
    mapping = load_json("sheet_mapping.json")
    sheets = [sheet for names in mapping.values() for sheet in names]
    assert len(sheets) == len(set(sheets))
    assert "项目总控台账" in sheets
    assert "资源计划" in sheets

