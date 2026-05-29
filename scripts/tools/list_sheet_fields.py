"""
列出 .xlsm/.xlsx 工作簿每个 sheet 的字段名及非空记录数。
第一行为字段名，后续为数据行；记录数 = 该列非空数据行数。

用法:
    uv run python scripts/tools/list_sheet_fields.py
    uv run python scripts/tools/list_sheet_fields.py "上层产品净值数据库.xlsm"
"""

import sys
import os

import pandas as pd
import openpyxl


# ---------------------------------------------------------------------------
# 项目约定：根目录推导
# ---------------------------------------------------------------------------

def _project_root():
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, "..", ".."))


# ---------------------------------------------------------------------------
# 列名规范化（与 analyze_excel_structure.py 保持一致）
# ---------------------------------------------------------------------------

def _normalize_column_name(name):
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name)
    s = s.replace("\u3000", " ").replace("\r", " ").replace("\n", " ")
    s = " ".join(s.split()).strip()
    return s


def _make_unique_columns(columns):
    seen = {}
    out = []
    for c in columns:
        base = _normalize_column_name(c)
        if base == "":
            base = "Unnamed"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out


# ---------------------------------------------------------------------------
# 核心逻辑
# ---------------------------------------------------------------------------

def list_sheet_fields(file_path: str) -> None:
    """逐 sheet 打印字段名及其非空记录数。"""
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        print(f"文件不存在: {abs_path}", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=False)

    for sheet_name in wb.sheetnames:
        df = pd.read_excel(
            abs_path,
            sheet_name=sheet_name,
            header=0,  # 第一行为字段名
            engine="openpyxl",
        )
        df.columns = _make_unique_columns(df.columns.tolist())

        non_null = df.notna().sum().to_dict()

        print(f"\n=== Sheet: {sheet_name} ===")
        print(f"{'字段名':<30} {'非空记录数':>10}")
        print("-" * 42)

        for col in df.columns:
            print(f"{col:<30} {non_null[col]:>10}")

        total_rows = df.shape[0]
        col_count = df.shape[1]
        print("-" * 42)
        print(f"总数据行: {total_rows}, 字段数: {col_count}")

    wb.close()


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main():
    root = _project_root()
    default_input = os.path.join(root, "上层产品净值数据库.xlsm")

    input_path = sys.argv[1] if len(sys.argv) >= 2 else default_input

    print(f"工作簿: {os.path.relpath(os.path.abspath(input_path), root)}")
    list_sheet_fields(input_path)


if __name__ == "__main__":
    main()
