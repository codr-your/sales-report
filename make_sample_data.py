# -*- coding: utf-8 -*-
"""生成考核用样例数据 sales_raw.xlsx（含脏数据，可复现）。

原始表字段：日期、销售员、产品、数量、单价、地区
脏数据类型：
  1. 重复行（完全相同）        -> 行4、行13
  2. 负数量 / 零数量           -> 行5、行6
  3. 单字段缺失（销售员/产品） -> 行7、行8
  4. 多字段缺失                -> 行12

人工核对基准（清洗后 7 条有效记录）：
  销售员：张三 900 | 李四 1100 | 王五 4420
  地区：华东 900 | 华南 1100 | 华北 4420
  月度总额 = 6420
"""
import datetime

import pandas as pd

OUTPUT_FILE = "sales_raw.xlsx"

COLUMNS = ["日期", "销售员", "产品", "数量", "单价", "地区"]

ROWS = [
    # 正常数据
    (datetime.date(2026, 9, 1), "张三", "键盘", 5, 100, "华东"),
    (datetime.date(2026, 9, 2), "李四", "鼠标", 10, 50, "华南"),
    (datetime.date(2026, 9, 3), "王五", "显示器", 2, 800, "华北"),
    # 重复行（与第1行完全相同）
    (datetime.date(2026, 9, 1), "张三", "键盘", 5, 100, "华东"),
    # 负数量
    (datetime.date(2026, 9, 4), "张三", "键盘", -3, 100, "华东"),
    # 数量 = 0
    (datetime.date(2026, 9, 5), "李四", "鼠标", 0, 50, "华南"),
    # 销售员缺失
    (datetime.date(2026, 9, 6), None, "键盘", 4, 100, "华东"),
    # 产品缺失
    (datetime.date(2026, 9, 7), "李四", None, 6, 50, "华南"),
    # 正常数据
    (datetime.date(2026, 9, 8), "王五", "显示器", 3, 800, "华北"),
    (datetime.date(2026, 9, 9), "张三", "鼠标", 8, 50, "华东"),
    (datetime.date(2026, 9, 10), "李四", "键盘", 5, 120, "华南"),
    # 日期、销售员、产品三字段缺失
    (datetime.date(2026, 9, 11), None, None, 2, 100, "华东"),
    # 重复行（与第2行完全相同）
    (datetime.date(2026, 9, 2), "李四", "鼠标", 10, 50, "华南"),
    # 正常数据
    (datetime.date(2026, 9, 12), "王五", "键盘", 7, 60, "华北"),
]


def main() -> None:
    df = pd.DataFrame(ROWS, columns=COLUMNS)
    # 数量、单价强制数值类型；日期列保留 date 类型
    df["数量"] = pd.to_numeric(df["数量"], errors="coerce")
    df["单价"] = pd.to_numeric(df["单价"], errors="coerce")
    df.to_excel(OUTPUT_FILE, index=False, sheet_name="sales_raw")
    print(f"已生成 {OUTPUT_FILE}，共 {len(df)} 行（含 7 行脏数据）")


if __name__ == "__main__":
    main()
#（注：内容由AI生成）
