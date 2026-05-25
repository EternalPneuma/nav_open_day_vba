import sys
import os
from datetime import datetime

import pandas as pd
import openpyxl
import xlrd


DEPENDENCIES = {
    "pandas": ">=1.0",
    "openpyxl": ">=3.0",
    "xlrd": ">=1.2",
}


def _parse_major_minor_patch(version_text: str):
    parts = []
    for p in str(version_text).split("."):
        try:
            parts.append(int(p))
        except Exception:
            break
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def _require_python():
    if sys.version_info < (3, 8):
        raise RuntimeError(f"Python版本过低: {sys.version.split()[0]}，需要 >= 3.8")


def _check_dependency_versions():
    installed = {
        "pandas": getattr(pd, "__version__", "unknown"),
        "openpyxl": getattr(openpyxl, "__version__", "unknown"),
        "xlrd": getattr(xlrd, "__version__", "unknown"),
    }
    problems = []
    for name, constraint in DEPENDENCIES.items():
        need = constraint.replace(">=", "").strip()
        if need and installed.get(name) not in (None, "unknown"):
            if _parse_major_minor_patch(installed[name]) < _parse_major_minor_patch(need):
                problems.append(f"{name}=={installed[name]} (需要 {constraint})")
    if problems:
        raise RuntimeError("依赖版本不满足: " + ", ".join(problems))
    return installed


def _project_root():
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, ".."))


def _now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalize_text(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x)
    s = s.replace("\u3000", " ").replace("\r", " ").replace("\n", " ")
    s = " ".join(s.split()).strip()
    return s


def _make_unique_columns(columns):
    seen = {}
    out = []
    for c in columns:
        base = _normalize_text(c)
        if base == "":
            base = "Unnamed"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out


def _detect_header_row(file_path, sheet_name, engine, max_scan_rows=30):
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        engine=engine,
        nrows=max_scan_rows,
    )
    best_row = 0
    best_score = -1.0
    for r in range(min(max_scan_rows, df.shape[0])):
        row = df.iloc[r]
        non_null = row.notna().sum()
        if non_null < 2:
            continue
        normalized = [_normalize_text(v) for v in row.tolist()]
        normalized = [x for x in normalized if x != ""]
        if len(normalized) == 0:
            continue
        unique_ratio = len(set(normalized)) / max(len(normalized), 1)
        score = float(non_null) + unique_ratio
        if score > best_score:
            best_score = score
            best_row = r
    return best_row


def _try_convert_datetime(series: pd.Series):
    non_null = series.dropna()
    if non_null.empty:
        return series
    parsed = pd.to_datetime(series, errors="coerce")
    ok = parsed.notna().sum()
    base = non_null.shape[0]
    if base > 0 and ok / base >= 0.8:
        return parsed
    return series


def _try_convert_numeric(series: pd.Series):
    non_null = series.dropna()
    if non_null.empty:
        return series
    parsed = pd.to_numeric(series, errors="coerce")
    ok = parsed.notna().sum()
    base = non_null.shape[0]
    if base > 0 and ok / base >= 0.8:
        return parsed
    return series


def _convert_types(df: pd.DataFrame):
    for col in df.columns:
        s = df[col]
        if s.dtype == "object":
            name = str(col)
            if ("日期" in name) or ("date" in name.lower()) or name.lower().endswith("dt"):
                df[col] = _try_convert_datetime(s)
            else:
                df[col] = _try_convert_numeric(s)
    return df


def _write_log_line(log_path, level, message):
    line = f"[{_now_text()}] [{level}] {message}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def merge_excels(input_dir_relpath: str, output_file_relpath: str):
    root = _project_root()
    input_dir = os.path.abspath(os.path.join(root, input_dir_relpath))
    if not os.path.isdir(input_dir):
        raise NotADirectoryError(f"目录不存在: {input_dir}")

    out_path = os.path.abspath(os.path.join(root, output_file_relpath))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    log_path = os.path.abspath(
        os.path.join(root, "outputs", "logs", f"merge_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    )
    os.makedirs(os.path.dirname(log_path), exist_ok=True)

    deps = _check_dependency_versions()
    _write_log_line(log_path, "INFO", f"Python={sys.version.split()[0]} deps={deps}")
    _write_log_line(log_path, "INFO", f"输入目录={os.path.relpath(input_dir, root)} 输出文件={os.path.relpath(out_path, root)}")

    files = []
    for name in sorted(os.listdir(input_dir)):
        ext = os.path.splitext(name)[1].lower()
        if ext in (".xlsx", ".xls"):
            files.append(os.path.join(input_dir, name))

    if not files:
        raise RuntimeError(f"目录下未找到Excel文件: {input_dir}")

    frames = []
    meta_cols = ["source_file", "source_sheet"]

    for file_path in files:
        file_name = os.path.basename(file_path)
        ext = os.path.splitext(file_name)[1].lower()
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        _write_log_line(log_path, "INFO", f"开始处理文件: {file_name} engine={engine}")
        try:
            xls = pd.ExcelFile(file_path, engine=engine)
            sheet_names = list(xls.sheet_names)
        except Exception as e:
            _write_log_line(log_path, "ERROR", f"读取工作表列表失败: {file_name} err={repr(e)}")
            continue

        for sheet in sheet_names:
            try:
                header_row = _detect_header_row(file_path, sheet, engine=engine)
                df = pd.read_excel(file_path, sheet_name=sheet, header=header_row, engine=engine)
                if df is None or df.empty:
                    continue
                df.columns = _make_unique_columns(df.columns.tolist())
                df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
                if df.empty:
                    continue
                df[meta_cols[0]] = file_name
                df[meta_cols[1]] = sheet
                df = _convert_types(df)
                frames.append(df)
            except Exception as e:
                _write_log_line(
                    log_path,
                    "ERROR",
                    f"读取工作表失败: file={file_name} sheet={sheet} err={repr(e)}",
                )
                continue

    if not frames:
        raise RuntimeError("未读取到任何有效数据（所有文件/工作表均为空或读取失败）")

    merged = pd.concat(frames, ignore_index=True, sort=False)
    merged = merged.dropna(axis=0, how="all").dropna(axis=1, how="all")

    business_cols = [c for c in merged.columns if c not in meta_cols]
    if not business_cols:
        raise RuntimeError("合并后未识别到业务字段列（仅剩source字段）")

    row_hash = pd.util.hash_pandas_object(merged[business_cols].fillna("").astype(str), index=False)
    merged["_row_hash"] = row_hash
    dup_mask = merged.duplicated(subset=["_row_hash"], keep="first")
    dup_count = int(dup_mask.sum())
    _write_log_line(log_path, "INFO", f"合并总行数={int(merged.shape[0])} 重复行数={dup_count}")

    merged_main = merged.loc[~dup_mask].drop(columns=["_row_hash"])
    merged_dups = merged.loc[dup_mask].drop(columns=["_row_hash"])

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged_main.to_excel(writer, index=False, sheet_name="合并数据")
        merged_dups.to_excel(writer, index=False, sheet_name="重复数据")

    _write_log_line(log_path, "INFO", f"写入完成: {os.path.relpath(out_path, root)}")
    _write_log_line(log_path, "INFO", f"日志文件: {os.path.relpath(log_path, root)}")
    return out_path, log_path, {"total_rows": int(merged.shape[0]), "dedup_rows": int(merged_main.shape[0]), "dup_rows": dup_count}


def main():
    _require_python()
    root = _project_root()

    default_input_dir = "181-固收上层"
    default_output = "多套账净值查询.xlsx"

    input_dir_rel = sys.argv[1] if len(sys.argv) >= 2 else default_input_dir
    output_rel = sys.argv[2] if len(sys.argv) >= 3 else default_output

    out_path, log_path, stats = merge_excels(input_dir_rel, output_rel)
    print(f"输出文件: {os.path.relpath(out_path, root)}")
    print(f"日志文件: {os.path.relpath(log_path, root)}")
    print(f"统计: {stats}")


if __name__ == "__main__":
    main()

