import argparse
import os
import re
import sys
import zipfile
from datetime import datetime
from numbers import Number

import pandas as pd


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).lower()


def detect_class_column(columns):
    candidates = [
        "班级",
        "班级名称",
        "班别",
        "班級",
        "class",
        "classname",
    ]
    norm_cols = {c: _norm(c) for c in columns}
    for cand in candidates:
        nc = _norm(cand)
        for col, ncol in norm_cols.items():
            if ncol == nc:
                return col
    for col, ncol in norm_cols.items():
        if any(k in ncol for k in ["班级", "班級", "class"]):
            return col
    return None


def safe_filename(name: str, used: dict, max_len: int = 80) -> str:
    v = name
    if isinstance(v, Number) and pd.notna(v):
        try:
            if float(v).is_integer():
                v = int(v)
        except Exception:
            pass
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none"}:
        s = "未分班"
    s = re.sub(r"[\/\\:\*\?\"<>\|\r\n\t]", "_", s)
    s = re.sub(r"\s+", " ", s).strip().strip(".")
    if not s:
        s = "未分班"
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    base = s
    n = used.get(base, 0) + 1
    used[base] = n
    if n == 1:
        return base
    return f"{base} ({n})"


def build_default_outdir(input_path: str) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.abspath(f"{base}_按班级拆分_{ts}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="输入Excel路径")
    ap.add_argument("--sheet", default="0", help="工作表：索引(0开始)或名称，默认0")
    ap.add_argument("--class-col", default="", help="班级列名（可选，默认自动识别）")
    ap.add_argument("--outdir", default="", help="输出目录（可选）")
    ap.add_argument("--zip", dest="zip_path", default="", help="输出zip路径（可选）")
    args = ap.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    df = pd.read_excel(input_path, sheet_name=sheet)
    if df.empty:
        raise ValueError("读取到空表，无法拆分")

    class_col = args.class_col.strip() or detect_class_column(df.columns)
    if not class_col:
        cols = ", ".join([str(c) for c in df.columns])
        raise ValueError(f"未找到班级列。当前列名：{cols}")
    if class_col not in df.columns:
        cols = ", ".join([str(c) for c in df.columns])
        raise ValueError(f"指定班级列不存在：{class_col}。当前列名：{cols}")

    df[class_col] = df[class_col].ffill()
    outdir = os.path.abspath(args.outdir.strip() or build_default_outdir(input_path))
    os.makedirs(outdir, exist_ok=True)

    used = {}
    written = []

    for class_value, g in df.groupby(class_col, dropna=False, sort=True):
        name = safe_filename(class_value, used)
        out_path = os.path.join(outdir, f"{name}.xlsx")
        g2 = g.copy()
        with pd.ExcelWriter(out_path, engine="openpyxl") as w:
            g2.to_excel(w, sheet_name="Sheet1", index=False)
        written.append((out_path, len(g2)))

    zip_path = os.path.abspath(args.zip_path.strip() or f"{outdir}.zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p, _ in written:
            zf.write(p, arcname=os.path.basename(p))

    sample = written[:3]
    summary = {
        "input": input_path,
        "sheet": args.sheet,
        "class_col": class_col,
        "outdir": outdir,
        "zip": zip_path,
        "classes": len(written),
        "sample": [{"file": os.path.basename(p), "rows": n} for p, n in sample],
    }
    print(pd.Series(summary).to_json(force_ascii=False))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)
