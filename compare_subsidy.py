import pathlib
import re
from decimal import Decimal, InvalidOperation

import pandas as pd


def norm_text(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return re.sub(r"\s+", "", s)


def norm_id(x) -> str:
    s = norm_text(x)
    if not s:
        return ""
    s = s.replace("．", ".")
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", s):
        try:
            s = format(Decimal(s).quantize(Decimal("1")), "f")
        except (InvalidOperation, ValueError):
            pass
    if s.endswith(".0") and re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"[^0-9Xx]", "", s).upper()
    return s


def detect_header_row(preview_df: pd.DataFrame) -> int:
    max_rows = min(len(preview_df), 30)
    for i in range(max_rows):
        row = preview_df.iloc[i].astype(str).fillna("")
        joined = "|".join(row.tolist())
        if "姓名" in joined and (row.ne("nan").sum() >= 2):
            return i
    return 0


def read_excel_robust(path: pathlib.Path, sheet_name: str) -> pd.DataFrame:
    header = 0
    try:
        preview = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=str, nrows=30)
        header = detect_header_row(preview)
    except Exception:
        header = 0
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=str)
    drop_cols = []
    for c in df.columns:
        if str(c).startswith("Unnamed") and df[c].isna().all():
            drop_cols.append(c)
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def extract_name_id_from_df(df: pd.DataFrame):
    cols = [str(c) for c in df.columns]
    name_col = None
    id_col = None

    for c in cols:
        if "姓名" in c and all(k not in c for k in ("签名", "姓名（签字")):
            name_col = c
            break

    for c in cols:
        if any(k in c for k in ("公民身份号码", "身份证", "证件号码", "身份证号", "身份证号码")):
            id_col = c
            break

    names = set()
    ids = set()
    if name_col is not None:
        for v in df[name_col].tolist():
            nv = norm_text(v)
            if nv:
                names.add(nv)
    if id_col is not None:
        for v in df[id_col].tolist():
            nid = norm_id(v)
            if nid:
                ids.add(nid)
    return names, ids, name_col, id_col


def convert_xls_to_xlsx(desktop: pathlib.Path, xls_path: pathlib.Path) -> pathlib.Path:
    import win32com.client

    out_dir = desktop / "_xls_converted_tmp"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / (xls_path.stem + ".xlsx")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xls_path))
        wb.SaveAs(str(out_path), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()

    return out_path


def main():
    desktop = pathlib.Path(r"C:\Users\P1368\Desktop")
    project_dir = pathlib.Path(__file__).resolve().parent
    output_dir = project_dir / "output"
    output_dir.mkdir(exist_ok=True)
    pop_path = desktop / "2021坪坝村常住人口信息统计表（最新最全）(2)(1).xlsx"
    subsidy_dir = desktop / "有政府补贴的名单"

    if not pop_path.exists():
        raise SystemExit(f"人口表不存在: {pop_path}")
    if not subsidy_dir.exists():
        raise SystemExit(f"补贴名单文件夹不存在: {subsidy_dir}")

    pop_df = read_excel_robust(pop_path, sheet_name="户籍信息总表")
    pop_df.columns = [norm_text(c) for c in pop_df.columns]
    pop_df = pop_df.dropna(axis=1, how="all")
    if "姓名" not in pop_df.columns:
        raise SystemExit("人口表未找到“姓名”列")

    pop_df["__name"] = pop_df["姓名"].map(norm_text)
    if "公民身份号码" in pop_df.columns:
        pop_df["__id"] = pop_df["公民身份号码"].map(norm_id)
    else:
        pop_df["__id"] = ""

    pop_df = pop_df[(pop_df["__name"] != "") | (pop_df["__id"] != "")].copy()

    files = []
    for p in subsidy_dir.iterdir():
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xls") and not p.name.startswith("~$"):
            files.append(p)

    sub_names = set()
    sub_ids = set()
    problems = []

    for f in files:
        path = f
        if f.suffix.lower() == ".xls":
            try:
                pd.ExcelFile(f)
            except Exception:
                try:
                    path = convert_xls_to_xlsx(desktop, f)
                except Exception as e2:
                    problems.append((str(f), "xls无法读取/转换", str(e2)))
                    continue

        try:
            xl = pd.ExcelFile(path)
            for sh in xl.sheet_names:
                try:
                    df = read_excel_robust(path, sh)
                    df.columns = [norm_text(c) for c in df.columns]
                    ns, ids, _, _ = extract_name_id_from_df(df)
                    sub_names |= ns
                    sub_ids |= ids
                except Exception as e:
                    problems.append((str(path), sh, str(e)))
        except Exception as e:
            problems.append((str(path), "(无法打开工作簿)", str(e)))

    pop_has_id = pop_df["__id"].map(lambda s: len(s) in (15, 18))
    match_by_id = pop_df["__id"].isin(sub_ids)
    match_by_name = pop_df["__name"].isin(sub_names)
    in_subsidy = match_by_id | (~pop_has_id & match_by_name)
    out_df = pop_df.loc[~in_subsidy].copy()

    prefer_cols = [
        "姓名",
        "公民身份号码",
        "性别",
        "年龄",
        "村组",
        "户籍地址",
        "电话",
        "是否常住",
        "与户主关系",
        "住户属性",
        "总人口数",
    ]
    cols = [c for c in prefer_cols if c in out_df.columns]
    out_df = out_df[cols].dropna(axis=1, how="all")

    out_path = output_dir / "坪坝村不在政府补贴名单人员名单.xlsx"
    out_df.to_excel(out_path, index=False, engine="openpyxl")

    print("补贴名单文件数:", len(files))
    print("补贴名单提取到姓名数:", len(sub_names))
    print("补贴名单提取到身份证号数:", len(sub_ids))
    print("人口表有效记录数:", len(pop_df))
    print("不在补贴名单人数:", len(out_df))
    print("输出:", out_path)

    if problems:
        log_path = desktop / "补贴名单读取问题日志.txt"
        with open(log_path, "w", encoding="utf-8") as w:
            for a, b, c in problems:
                w.write(f"{a}\t{b}\t{c}\n")
        print("问题日志:", log_path)


if __name__ == "__main__":
    main()

