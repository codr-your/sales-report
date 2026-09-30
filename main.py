# -*- coding: utf-8 -*-
"""销售报表生成脚本（考题一）

读取 sales_raw.xlsx -> 清洗 -> 统计 -> 输出 report.xlsx（4 个工作表）。

用法：
    python main.py                 # 使用默认 sales_raw.xlsx / report.xlsx
    python main.py -i 输入文件 -o 输出文件

异常场景均给出明确中文提示并以非零码退出，不生成残缺报表。
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

REQUIRED_COLUMNS = ["日期", "销售员", "产品", "数量", "单价", "地区"]


class ReportError(Exception):
    """业务异常：携带用户可读的中文提示。"""


def load_data(input_file: str) -> pd.DataFrame:
    """读取原始表并校验文件、表头与空数据。"""
    path = Path(input_file)
    if not path.exists():
        raise ReportError(f"未找到输入文件：{input_file}")

    try:
        df = pd.read_excel(path)
    except Exception as exc:  # 文件损坏 / 非 Excel 等
        raise ReportError(f"无法读取文件 {input_file}：{exc}") from exc

    # 表头校验：必须包含全部必需字段
    actual = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in actual]
    if missing:
        raise ReportError(f"表头不匹配，缺少字段：{'、'.join(missing)}")

    if df.empty:
        raise ReportError(f"文件 {input_file} 为空，无任何数据")

    df = df.rename(columns={c: c.strip() for c in df.columns if isinstance(c, str)})
    df = df[REQUIRED_COLUMNS]

    # 数值列强制转数值：无法转换的置为 NaN，视作缺失
    df["数量"] = pd.to_numeric(df["数量"], errors="coerce")
    df["单价"] = pd.to_numeric(df["单价"], errors="coerce")
    try:
        df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    except Exception:
        df["日期"] = pd.to_datetime(df["日期"].astype(str), errors="coerce")

    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """清洗：删缺失行 -> 剔数量<=0 -> 全字段去重（保留首次出现）。"""
    cleaned = df.dropna(subset=REQUIRED_COLUMNS)          # 1. 任一关键字段缺失即删
    cleaned = cleaned[cleaned["数量"] > 0]                # 2. 剔除数量 <= 0
    cleaned = cleaned.drop_duplicates(subset=REQUIRED_COLUMNS)  # 3. 全字段去重
    return cleaned.reset_index(drop=True)


def summarize(df: pd.DataFrame) -> dict:
    """按销售员 / 按地区汇总销售额，并计算月度总额。"""
    sales = (df["数量"] * df["单价"]).round(2)
    by_salesperson = (
        sales.groupby(df["销售员"]).sum().round(2)
        .reset_index().rename(columns={0: "销售额"})
        .sort_values("销售额", ascending=False).reset_index(drop=True)
    )
    by_region = (
        sales.groupby(df["地区"]).sum().round(2)
        .reset_index().rename(columns={0: "销售额"})
        .sort_values("销售额", ascending=False).reset_index(drop=True)
    )
    total = sales.sum().round(2)
    return {
        "by_salesperson": by_salesperson,
        "by_region": by_region,
        "total": float(total),
        "record_count": int(len(df)),
    }


def write_report(df: pd.DataFrame, stats: dict, output_file: str) -> None:
    """写出 report.xlsx：明细清洗后 / 按销售员 / 按地区 / 总览。"""
    detail = df.copy()
    detail["销售额"] = (detail["数量"] * detail["单价"]).round(2)

    overview = pd.DataFrame(
        [{"指标": "月度总额", "数值": stats["total"]},
         {"指标": "有效记录条数", "数值": stats["record_count"]}]
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="明细清洗后", index=False)
        stats["by_salesperson"].to_excel(writer, sheet_name="按销售员", index=False)
        stats["by_region"].to_excel(writer, sheet_name="按地区", index=False)
        overview.to_excel(writer, sheet_name="总览", index=False)

        # 轻量样式：表头加粗浅底、金额列数字格式、列宽自适应
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(name=cell.font.name, bold=True)
            for col_cells in ws.columns:
                letter = col_cells[0].column_letter
                width = max(len(str(c.value)) for c in col_cells if c.value is not None)
                ws.column_dimensions[letter].width = max(width + 2, 10)
        for ws in writer.book.worksheets:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.column_letter in ("D", "E", "G") or cell.column_letter == "B" and ws.title == "总览":
                        cell.number_format = "0.00"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="销售报表生成脚本（清洗 + 统计 + 输出）")
    parser.add_argument("-i", "--input", default="sales_raw.xlsx", help="原始数据文件（默认 sales_raw.xlsx）")
    parser.add_argument("-o", "--output", default="report.xlsx", help="输出报表文件（默认 report.xlsx）")
    args = parser.parse_args(argv)

    try:
        df = load_data(args.input)
        cleaned = clean_data(df)
        if cleaned.empty:
            raise ReportError("清洗后无有效数据（所有行均缺失 / 数量<=0 / 重复），不生成报表")
        stats = summarize(cleaned)
        write_report(cleaned, stats, args.output)
    except ReportError as exc:
        print(f"[错误] {exc}")
        return 1

    print(f"清洗前 {len(df)} 行 -> 有效 {stats['record_count']} 行")
    print(f"月度总额：{stats['total']:.2f}")
    print(f"报表已生成：{args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
#（注：内容由AI生成）
