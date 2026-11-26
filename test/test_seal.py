"""
Tests for the 3 seal/stamp batch-approval flows (用章/用印批量申请).

- 补签派遣员工合同用章申请   parse_seal_reseal_excel / build_seal_reseal_form_widgets
- 解除（终止）劳动合同证明书用印申请 parse_seal_terminate_cert_excel /
  build_seal_terminate_cert_form_widgets
- 新签/续签派遣员工合同批量用章用印申请 parse_seal_contract_batch_excel /
  build_seal_contract_batch_form_widgets

Widget IDs and radio option values were verified against the live
approval definitions on 2026-11-25; the row-radio widget display name
(用印主体) and per-flow option text suffixes (合同章/公章) were
re-verified after the online rename on 2026-11-26.
"""

import datetime as dt
import json
from pathlib import Path

import openpyxl
import pytest

from app.excel_parser import (
    parse_seal_contract_batch_excel,
    parse_seal_reseal_excel,
    parse_seal_terminate_cert_excel,
)


class TestBatchWidgetStateCleanup:
    """换文件/换类型后，批量编辑控件的 widget 状态必须一并清理。

    否则 Streamlit 按控件 key 保留旧输入，「用章事由（自动生成，
    可修改）」等输入框会显示上一个文件的编辑值而非新默认值。
    """

    def test_clear_batch_items_pops_widget_keys(self):
        import streamlit as st

        from app.app import _clear_batch_items

        st.session_state.batch_items = []
        st.session_state["batch_reason_0"] = "旧文件的事由"
        st.session_state["batch_title_1"] = "旧文件的标题"
        st.session_state["batch_remark_2"] = "旧文件的备注"
        st.session_state["unrelated_key"] = "keep me"

        _clear_batch_items()

        assert "batch_reason_0" not in st.session_state
        assert "batch_title_1" not in st.session_state
        assert "batch_remark_2" not in st.session_state
        assert st.session_state["unrelated_key"] == "keep me"
        assert st.session_state.batch_items == []
from app.form_builder import (
    SEAL_ATTACHMENT,
    SEAL_DETAIL_LIST,
    SEAL_ENTITY_OPTION_VALUES,
    SEAL_REASON,
    SEAL_REMARK,
    SEAL_ROW_COPIES,
    SEAL_ROW_DATE_INTERVAL,
    SEAL_ROW_ID_COPIES,
    SEAL_ROW_NAME,
    SEAL_ROW_OPTION,
    SEAL_ROW_REASON,
    build_seal_contract_batch_form_widgets,
    build_seal_reseal_form_widgets,
    build_seal_terminate_cert_form_widgets,
    resolve_seal_option_value,
    seal_option_labels,
)

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
RESEAL_TEMPLATE = TEMPLATES_DIR / "补签派遣员工合同用章申请模板.xlsx"
TERM_TEMPLATE = TEMPLATES_DIR / "解除（终止）劳动合同证明书用印申请模板.xlsx"
BATCH_TEMPLATE = TEMPLATES_DIR / "新签／续签派遣员工合同批量用章用印申请模板.xlsx"

FAKE_NAMES = {"李四", "王五", "赵六", "孙七", "周八", "张三"}


# ── Workbook builders ────────────────────────────────────────────────

def _make_reseal_xlsx(
    path: Path,
    *,
    start: str = "2026-01-01",
    end: str = "2026-12-31",
    seal: str = "彩虹公司合同章",
    reason_cell: str = "合同到期未及时续签，现补签",
    copies=1,
    omit: str | None = None,
    preamble: bool = False,
    total_copies: int | None = None,
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["序号", "派遣员工姓名", "合同开始时间", "合同结束时间", "用印主体", "补签原因", "合同份数"]
    if omit:
        headers.remove(omit)
    ws.append(headers)
    people = [("1", "李四"), ("2", "王五"), ("3", "赵六")]
    for serial, name in people:
        row = [int(serial), name, start, end, seal, reason_cell, copies]
        if omit:
            idx = ["序号", "派遣员工姓名", "合同开始时间", "合同结束时间", "用印主体", "补签原因", "合同份数"].index(omit)
            row.pop(idx)
        ws.append(row)
    if total_copies is not None:
        ws.append(["合计", None, None, None, None, None, total_copies])
    if preamble:
        ws.insert_rows(1)
        ws["A1"] = "说明：本表由XX单位提供（表前说明行不影响解析）"
    wb.save(path)


def _make_term_xlsx(path: Path, *, seal: str = "吉林分公司公章", copies=2) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["序号", "派遣员工姓名", "证明份数", "用印主体"])
    for serial, name in [("1", "李四"), ("2", "王五")]:
        ws.append([int(serial), name, copies, seal])
    wb.save(path)


def _make_batch_xlsx(path: Path, *, entity: str = "彩虹公司合同章", id_copies="1") -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["序号", "派遣员工姓名", "用印主体", "合同份数", "身份证复印件盖公章份数"])
    for serial, name in [("1", "李四"), ("2", "王五"), ("3", "赵六")]:
        ws.append([int(serial), name, entity, 2, id_copies])
    wb.save(path)


@pytest.fixture
def reseal_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "reseal.xlsx"
    _make_reseal_xlsx(path)
    return path


@pytest.fixture
def term_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "term.xlsx"
    _make_term_xlsx(path)
    return path


@pytest.fixture
def batch_xlsx(tmp_path: Path) -> Path:
    path = tmp_path / "batch.xlsx"
    _make_batch_xlsx(path)
    return path


# ── Parser: 补签 ─────────────────────────────────────────────────────

class TestParseSealReseal:
    def test_rows(self, reseal_xlsx: Path):
        sheet = parse_seal_reseal_excel(reseal_xlsx)
        assert len(sheet.rows) == 3
        first = sheet.rows[0]
        assert first.serial == "1"
        assert first.name == "李四"
        assert first.copies == 1
        assert first.option_text == "彩虹公司合同章"
        assert first.option_value == SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        assert first.start_date == "2026-01-01T00:00:00+08:00"
        assert first.end_date == "2026-12-31T00:00:00+08:00"
        assert first.interval_days == "364"
        assert first.reason == "合同到期未及时续签，现补签"

    def test_auto_reason_and_empty_remark(self, reseal_xlsx: Path):
        """用章事由按明细自动生成；备注留空（界面填写）。"""
        sheet = parse_seal_reseal_excel(reseal_xlsx)
        assert sheet.reason == "补签派遣员工合同用章，共3人，合同份数合计3份"
        assert sheet.remark == ""
        assert sheet.copies_label == "合同份数"
        assert sheet.option_label == "用印主体"
        assert sheet.total_copies == 3
        assert sheet.warnings == []

    def test_preamble_row_above_header(self, tmp_path: Path):
        """表前说明行不影响解析（表头按列名识别）。"""
        path = tmp_path / "preamble.xlsx"
        _make_reseal_xlsx(path, preamble=True)
        sheet = parse_seal_reseal_excel(path)
        assert len(sheet.rows) == 3
        assert sheet.warnings == []

    def test_date_string_formats(self, tmp_path: Path):
        for start, end, days in [
            ("2026/1/1", "2026.12.31", 364),
            ("2026年1月1日", "2026-1-31", 30),
        ]:
            path = tmp_path / f"d{days}.xlsx"
            _make_reseal_xlsx(path, start=start, end=end)
            sheet = parse_seal_reseal_excel(path)
            assert sheet.rows[0].start_date == "2026-01-01T00:00:00+08:00"
            assert sheet.rows[0].interval_days == str(days)

    def test_seal_text_without_suffix(self, tmp_path: Path):
        path = tmp_path / "nosuffix.xlsx"
        _make_reseal_xlsx(path, seal="白城分公司")
        sheet = parse_seal_reseal_excel(path)
        assert sheet.rows[0].option_value == SEAL_ENTITY_OPTION_VALUES["白城分公司"]

    def test_end_before_start_rejected(self, tmp_path: Path):
        path = tmp_path / "reversed.xlsx"
        _make_reseal_xlsx(path, start="2026-12-31", end="2026-01-01")
        with pytest.raises(ValueError, match="结束时间需晚于合同开始时间"):
            parse_seal_reseal_excel(path)

    def test_bad_date_rejected(self, tmp_path: Path):
        path = tmp_path / "baddate.xlsx"
        _make_reseal_xlsx(path, start="不是日期")
        with pytest.raises(ValueError, match="合同开始时间"):
            parse_seal_reseal_excel(path)

    def test_bad_copies_rejected(self, tmp_path: Path):
        for bad in (0, -1, 1.5, "abc"):
            path = tmp_path / f"c{bad}.xlsx"
            _make_reseal_xlsx(path, copies=bad)
            with pytest.raises(ValueError, match="合同份数"):
                parse_seal_reseal_excel(path)

    def test_unknown_seal_rejected(self, tmp_path: Path):
        path = tmp_path / "unknown.xlsx"
        _make_reseal_xlsx(path, seal="海淀分公司公章")
        with pytest.raises(ValueError, match="用印主体无法识别"):
            parse_seal_reseal_excel(path)

    def test_empty_reason_column_rejected(self, tmp_path: Path):
        path = tmp_path / "noreason.xlsx"
        _make_reseal_xlsx(path, reason_cell="")
        with pytest.raises(ValueError, match="补签原因为空"):
            parse_seal_reseal_excel(path)

    def test_missing_column_rejected(self, tmp_path: Path):
        path = tmp_path / "nocol.xlsx"
        _make_reseal_xlsx(path, omit="合同结束时间")
        with pytest.raises(ValueError, match="未找到表头行"):
            parse_seal_reseal_excel(path)

    def test_interleaved_blank_row_warning(self, tmp_path: Path):
        path = tmp_path / "blank.xlsx"
        _make_reseal_xlsx(path)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        ws.insert_rows(3)  # 数据行之间（李四与王五之间）插入空行
        wb.save(path)
        sheet = parse_seal_reseal_excel(path)
        assert any("跳过 1 行" in w for w in sheet.warnings)

    def test_total_row_validation(self, tmp_path: Path):
        path = tmp_path / "total.xlsx"
        _make_reseal_xlsx(path, total_copies=5)  # 明细合计 3
        sheet = parse_seal_reseal_excel(path)
        assert any("合同份数明细合计" in w for w in sheet.warnings)

    def test_total_row_matches(self, tmp_path: Path):
        path = tmp_path / "total_ok.xlsx"
        _make_reseal_xlsx(path, total_copies=3)
        sheet = parse_seal_reseal_excel(path)
        assert sheet.warnings == []

    def test_no_title_in_preview_only(self, reseal_xlsx: Path):
        sheet = parse_seal_reseal_excel(reseal_xlsx)
        assert isinstance(sheet.title, str)


# ── Parser: 解除（终止）劳动合同证明书 ────────────────────────

class TestParseSealTerminateCert:
    def test_rows(self, term_xlsx: Path):
        sheet = parse_seal_terminate_cert_excel(term_xlsx)
        assert len(sheet.rows) == 2
        row = sheet.rows[0]
        assert row.name == "李四"
        assert row.copies == 2
        assert row.option_value == SEAL_ENTITY_OPTION_VALUES["吉林分公司"]
        assert row.start_date == "" and row.end_date == ""
        assert row.reason == "" and row.id_copy_copies == ""
        assert sheet.copies_label == "证明份数"
        assert sheet.warnings == []


# ── Parser: 新签／续签批量 ───────────────────────────────────────────

class TestParseSealContractBatch:
    def test_rows(self, batch_xlsx: Path):
        sheet = parse_seal_contract_batch_excel(batch_xlsx)
        assert len(sheet.rows) == 3
        row = sheet.rows[1]
        assert row.name == "王五"
        assert row.copies == 2
        assert row.option_text == "彩虹公司合同章"
        assert row.option_value == SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        assert row.id_copy_copies == "1"
        assert sheet.option_label == "用印主体"
        assert sheet.copies_label == "合同份数"
        assert sheet.total_copies == 6
        assert sheet.warnings == []

    def test_id_copies_kept_as_text(self, tmp_path: Path):
        path = tmp_path / "text.xlsx"
        _make_batch_xlsx(path, id_copies="2份")
        sheet = parse_seal_contract_batch_excel(path)
        assert sheet.rows[0].id_copy_copies == "2份"
        assert sheet.warnings == []

    def test_id_copies_empty_rejected(self, tmp_path: Path):
        path = tmp_path / "empty.xlsx"
        _make_batch_xlsx(path, id_copies="")
        with pytest.raises(ValueError, match="身份证复印件盖公章份数为空"):
            parse_seal_contract_batch_excel(path)

    def test_id_copies_non_digit_warning(self, tmp_path: Path):
        path = tmp_path / "nodigit.xlsx"
        _make_batch_xlsx(path, id_copies="按需")
        sheet = parse_seal_contract_batch_excel(path)
        assert any("不含数字" in w for w in sheet.warnings)

    def test_unknown_entity_rejected(self, tmp_path: Path):
        path = tmp_path / "unknown.xlsx"
        _make_batch_xlsx(path, entity="海淀分公司")
        with pytest.raises(ValueError, match="用印主体无法识别"):
            parse_seal_contract_batch_excel(path)


# ── Option value resolution ──────────────────────────────────────────

class TestResolveSealOptionValue:
    def test_all_flow_suffixes(self):
        """各流程选项文本（主体名+本流程后缀）均可解析。"""
        for suffix in ("合同章", "公章"):
            for entity, value in SEAL_ENTITY_OPTION_VALUES.items():
                assert resolve_seal_option_value(f"{entity}{suffix}", suffix) == value

    def test_bare_entity_and_cross_suffix_tolerant(self):
        """省略后缀、后缀写错（跨流程）均按主体名解析。"""
        value = SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        assert resolve_seal_option_value("彩虹公司", "合同章") == value
        assert resolve_seal_option_value("彩虹公司公章", "合同章") == value
        assert resolve_seal_option_value("彩虹公司合同章", "公章") == value

    def test_whitespace_tolerant(self):
        assert (
            resolve_seal_option_value(" 彩虹公司公章 ", "公章")
            == SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        )

    def test_error_lists_allowed(self):
        with pytest.raises(ValueError, match="公主岭分公司公章"):
            resolve_seal_option_value("不存在公章", "公章")

    def test_label_helper(self):
        assert seal_option_labels("公章") == [
            f"{entity}公章" for entity in SEAL_ENTITY_OPTION_VALUES
        ]
        assert seal_option_labels("合同章") == [
            f"{entity}合同章" for entity in SEAL_ENTITY_OPTION_VALUES
        ]


# ── Form builders ────────────────────────────────────────────────────

class TestBuildSealReseal:
    def test_structure(self, reseal_xlsx: Path):
        sheet = parse_seal_reseal_excel(reseal_xlsx)
        widgets = build_seal_reseal_form_widgets(sheet, file_code="FAKE")
        assert [w["id"] for w in widgets] == [
            SEAL_REASON, SEAL_DETAIL_LIST, SEAL_REMARK, SEAL_ATTACHMENT,
        ]
        assert widgets[0]["type"] == "textarea"
        assert widgets[0]["value"] == sheet.reason
        assert widgets[2]["value"] == sheet.remark
        assert widgets[3]["value"] == ["FAKE"]
        # 无标题控件
        assert all(w["id"] != "widget17811628753290001" for w in widgets)

        row = widgets[1]["value"][0]
        assert [c["id"] for c in row] == [
            SEAL_ROW_NAME, SEAL_ROW_DATE_INTERVAL, SEAL_ROW_COPIES,
            SEAL_ROW_OPTION, SEAL_ROW_REASON,
        ]
        interval = row[1]
        assert interval["type"] == "dateInterval"
        assert interval["value"] == {
            "start": "2026-01-01T00:00:00+08:00",
            "end": "2026-12-31T00:00:00+08:00",
            "interval": 364.0,
        }
        assert row[2]["type"] == "number"
        assert row[2]["value"] == 1
        assert row[3]["type"] == "radioV2"
        assert row[3]["value"] == SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        assert row[4]["value"] == "合同到期未及时续签，现补签"
        json.dumps(widgets, ensure_ascii=False)

    def test_reason_remark_override(self, reseal_xlsx: Path):
        sheet = parse_seal_reseal_excel(reseal_xlsx)
        widgets = build_seal_reseal_form_widgets(
            sheet, file_code="FAKE", reason="改后事由", remark="改后备注"
        )
        assert widgets[0]["value"] == "改后事由"
        assert widgets[2]["value"] == "改后备注"


class TestBuildSealTerminateCert:
    def test_structure(self, term_xlsx: Path):
        sheet = parse_seal_terminate_cert_excel(term_xlsx)
        widgets = build_seal_terminate_cert_form_widgets(sheet, file_code="FAKE")
        assert [w["id"] for w in widgets] == [
            SEAL_REASON, SEAL_DETAIL_LIST, SEAL_REMARK, SEAL_ATTACHMENT,
        ]
        row = widgets[1]["value"][0]
        assert [c["id"] for c in row] == [
            SEAL_ROW_NAME, SEAL_ROW_COPIES, SEAL_ROW_OPTION,
        ]
        assert row[1]["value"] == 2
        assert row[2]["value"] == SEAL_ENTITY_OPTION_VALUES["吉林分公司"]
        assert SEAL_ROW_DATE_INTERVAL not in [c["id"] for c in row]
        json.dumps(widgets, ensure_ascii=False)


class TestBuildSealContractBatch:
    def test_structure(self, batch_xlsx: Path):
        sheet = parse_seal_contract_batch_excel(batch_xlsx)
        widgets = build_seal_contract_batch_form_widgets(sheet, file_code="FAKE")
        row = widgets[1]["value"][0]
        assert [c["id"] for c in row] == [
            SEAL_ROW_NAME, SEAL_ROW_COPIES, SEAL_ROW_OPTION, SEAL_ROW_ID_COPIES,
        ]
        assert row[2]["value"] == SEAL_ENTITY_OPTION_VALUES["彩虹公司"]
        # 身份证复印件盖公章份数在表单中是单行文本控件
        assert row[3]["type"] == "input"
        assert row[3]["value"] == "1"
        json.dumps(widgets, ensure_ascii=False)


# ── Shipped templates ────────────────────────────────────────────────

class TestShippedSealTemplates:
    def test_templates_exist(self):
        for p in (RESEAL_TEMPLATE, TERM_TEMPLATE, BATCH_TEMPLATE):
            assert p.exists(), f"Missing: {p}"

    def test_reseal_template_parses_cleanly(self):
        sheet = parse_seal_reseal_excel(RESEAL_TEMPLATE)
        assert len(sheet.rows) == 5
        assert sheet.total_copies == 20  # 5 行 × 默认 4 份
        assert all(r.copies == 4 for r in sheet.rows)
        assert sheet.reason == "补签派遣员工合同用章，共5人，合同份数合计20份"
        assert sheet.remark == ""
        assert sheet.warnings == []

    def test_term_template_parses_cleanly(self):
        sheet = parse_seal_terminate_cert_excel(TERM_TEMPLATE)
        assert len(sheet.rows) == 4
        assert sheet.total_copies == 16  # 4 行 × 默认 4 份
        assert all(r.copies == 4 for r in sheet.rows)
        assert sheet.reason == "解除劳动合同证明书用印，共4人，证明份数合计16份"
        assert sheet.warnings == []

    def test_batch_template_parses_cleanly(self):
        sheet = parse_seal_contract_batch_excel(BATCH_TEMPLATE)
        assert len(sheet.rows) == 5
        assert sheet.total_copies == 20  # 5 行 × 默认 4 份
        assert all(r.copies == 4 for r in sheet.rows)
        assert sheet.reason == "新签/续签派遣员工劳动合同批量用章用印，共5人，合同份数合计20份"
        assert sheet.warnings == []

    def test_templates_no_real_pii(self):
        for parser, path in [
            (parse_seal_reseal_excel, RESEAL_TEMPLATE),
            (parse_seal_terminate_cert_excel, TERM_TEMPLATE),
            (parse_seal_contract_batch_excel, BATCH_TEMPLATE),
        ]:
            sheet = parser(path)
            for row in sheet.rows:
                assert row.name in FAKE_NAMES, f"Real name leaked: {row.name}"

    def test_builders_accept_templates(self):
        cases = [
            (build_seal_reseal_form_widgets, parse_seal_reseal_excel, RESEAL_TEMPLATE, SEAL_ROW_DATE_INTERVAL),
            (build_seal_terminate_cert_form_widgets, parse_seal_terminate_cert_excel, TERM_TEMPLATE, None),
            (build_seal_contract_batch_form_widgets, parse_seal_contract_batch_excel, BATCH_TEMPLATE, SEAL_ROW_ID_COPIES),
        ]
        for builder, parser, path, extra_id in cases:
            widgets = builder(parser(path), file_code="fake_code")
            assert len(widgets) == 4
            detail = widgets[1]["value"]
            if extra_id:
                assert extra_id in [c["id"] for c in detail[0]]
            json.dumps(widgets, ensure_ascii=False)
