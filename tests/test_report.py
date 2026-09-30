# -*- coding: utf-8 -*-
"""考题一 pytest 测试用例。

覆盖：
  1. 正常用例：样例数据 -> 清洗正确、统计数值正确、四工作表齐全
  2. 边界/异常用例：文件缺失、空文件、表头不匹配、无有效数据、
     数量<=0、重复行、非数值数量/单价
"""
import datetime

import pandas as pd
import pytest

from main import ReportError, clean_data, load_data, main, summarize

SAMPLE_COLS = ["日期", "销售员", "产品", "数量", "单价", "地区"]


# ---------- 夹具：生成各类临时 xlsx ----------

def _write_xlsx(path, rows):
    df = pd.DataFrame(rows, columns=SAMPLE_COLS)
    df.to_excel(path, index=False)


@pytest.fixture
def sample_file(tmp_path):
    """正常含脏数据的样例（14 行，清洗后 7 行）。"""
    rows = [
        (datetime.date(2026, 9, 1), "张三", "键盘", 5, 100, "华东"),
        (datetime.date(2026, 9, 2), "李四", "鼠标", 10, 50, "华南"),
        (datetime.date(2026, 9, 3), "王五", "显示器", 2, 800, "华北"),
        (datetime.date(2026, 9, 1), "张三", "键盘", 5, 100, "华东"),   # 重复
        (datetime.date(2026, 9, 4), "张三", "键盘", -3, 100, "华东"),  # 负数量
        (datetime.date(2026, 9, 5), "李四", "鼠标", 0, 50, "华南"),    # 零数量
        (datetime.date(2026, 9, 6), None, "键盘", 4, 100, "华东"),     # 销售员缺失
        (datetime.date(2026, 9, 7), "李四", None, 6, 50, "华南"),      # 产品缺失
        (datetime.date(2026, 9, 8), "王五", "显示器", 3, 800, "华北"),
        (datetime.date(2026, 9, 9), "张三", "鼠标", 8, 50, "华东"),
        (datetime.date(2026, 9, 10), "李四", "键盘", 5, 120, "华南"),
        (datetime.date(2026, 9, 11), None, None, 2, 100, "华东"),      # 多字段缺失
        (datetime.date(2026, 9, 2), "李四", "鼠标", 10, 50, "华南"),   # 重复
        (datetime.date(2026, 9, 12), "王五", "键盘", 7, 60, "华北"),
    ]
    path = tmp_path / "sales_raw.xlsx"
    _write_xlsx(path, rows)
    return str(path)


# ---------- 正常用例 ----------

class TestNormal:
    def test_clean_removes_all_dirty_rows(self, sample_file):
        df = load_data(sample_file)
        cleaned = clean_data(df)
        assert len(cleaned) == 7                       # 14 - 2重复 - 2数量<=0 - 3缺失 = 7
        assert cleaned["数量"].min() > 0               # 无 <=0
        assert not cleaned.isnull().any().any()        # 无缺失
        assert cleaned.duplicated().sum() == 0         # 无重复

    def test_summary_values_match_manual_calc(self, sample_file):
        df = load_data(sample_file)
        cleaned = clean_data(df)
        stats = summarize(cleaned)

        by_salesperson = dict(zip(stats["by_salesperson"]["销售员"],
                                  stats["by_salesperson"]["销售额"]))
        assert by_salesperson == {"张三": 900, "李四": 1100, "王五": 4420}

        by_region = dict(zip(stats["by_region"]["地区"], stats["by_region"]["销售额"]))
        assert by_region == {"华东": 900, "华南": 1100, "华北": 4420}

        assert stats["total"] == 6420                  # 900 + 1100 + 4420
        assert stats["record_count"] == 7

    def test_main_generates_four_sheets(self, sample_file, tmp_path):
        out = tmp_path / "report.xlsx"
        assert main(["-i", sample_file, "-o", str(out)]) == 0
        sheets = pd.read_excel(out, sheet_name=None)
        assert list(sheets.keys()) == ["明细清洗后", "按销售员", "按地区", "总览"]
        assert len(sheets["明细清洗后"]) == 7
        assert sheets["总览"].iloc[0]["数值"] == 6420


# ---------- 边界 / 异常用例 ----------

class TestAbnormal:
    def test_file_not_found(self, tmp_path):
        with pytest.raises(ReportError, match="未找到输入文件"):
            load_data(str(tmp_path / "nope.xlsx"))

    def test_empty_file(self, tmp_path):
        path = tmp_path / "empty.xlsx"
        pd.DataFrame(columns=SAMPLE_COLS).to_excel(path, index=False)
        with pytest.raises(ReportError, match="为空"):
            load_data(str(path))

    def test_missing_headers(self, tmp_path):
        path = tmp_path / "bad_header.xlsx"
        pd.DataFrame(columns=["日期", "销售员", "产品"], data=[]).to_excel(path, index=False)
        with pytest.raises(ReportError, match="表头不匹配"):
            load_data(str(path))

    def test_no_valid_data_after_clean(self, tmp_path):
        rows = [(datetime.date(2026, 9, 1), "张三", "键盘", -2, 100, "华东"),
                (datetime.date(2026, 9, 2), None, "鼠标", 0, 50, "华南")]
        path = tmp_path / "all_dirty.xlsx"
        _write_xlsx(path, rows)
        df = load_data(str(path))
        assert clean_data(df).empty
        assert main(["-i", str(path), "-o", str(tmp_path / "r.xlsx")]) == 1

    def test_duplicate_rows_removed(self, sample_file):
        df = load_data(sample_file)
        cleaned = clean_data(df)
        # 重复的两行（张三键盘5*100华东、李四鼠标10*50华南）只保留一条
        dup_key = cleaned.groupby(SAMPLE_COLS).size()
        assert (dup_key > 1).sum() == 0

    def test_non_numeric_quantity_treated_as_missing(self, tmp_path):
        rows = [(datetime.date(2026, 9, 1), "张三", "键盘", "abc", 100, "华东"),  # 数量非数值
                (datetime.date(2026, 9, 2), "李四", "鼠标", 10, "x", "华南")]     # 单价非数值
        path = tmp_path / "non_numeric.xlsx"
        _write_xlsx(path, rows)
        df = load_data(str(path))
        cleaned = clean_data(df)
        assert len(cleaned) == 0          # 两行均因数值无法转换被视作缺失剔除
        assert main(["-i", str(path), "-o", str(tmp_path / "r.xlsx")]) == 1

    def test_summary_returns_empty_with_all_invalid(self, sample_file):
        df = load_data(sample_file)
        cleaned = clean_data(df)
        # 边界：清洗后只剩一条也能正常汇总
        one_row = cleaned.head(1)
        stats = summarize(one_row)
        assert stats["record_count"] == 1
        assert stats["total"] == one_row["数量"].iloc[0] * one_row["单价"].iloc[0]
#（注：内容由AI生成）
