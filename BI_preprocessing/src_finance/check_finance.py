"""
檢查家庭共帳年結檔的資料內容，發現資料集有問題時輸出異常報告

- 使用方法 (在本程式所在資料夾執行)：
    python check_finance.py                             檢查整份檔案
    python check_finance.py --year 2026                 只檢查月結年月屬於 2026 年的資料
    python check_finance.py --month 202609              只檢查 202609
    python check_finance.py --month 202607 202609       檢查 202607 到 202609
    python check_finance.py --file 家庭共帳_年結2026.xlsx 指定要讀取的檔案，可搭配 --year 或 --month
    python check_finance.py -h                          顯示參數說明

    --year 與 --month 擇一使用；用 --year/--month 時，月結年月空白或格式錯誤的列會先被篩掉。
    指定範圍超出資料集的月結年月範圍時，即時停止執行。

0. 安裝相依 py 套件：見 requirements.txt (pandas、openpyxl)
    - 安裝方式：python -m pip install -r requirements.txt

1. 讀取檔案
    - 預設讀取本程式所在資料夾中修改日期最新的 xlsx。
    - 亦可在執行CLI命令時，用 --file 指定要讀取的檔案 xlsx。
    - 略過 ~$ 暫存檔及異常報告。
2. 檢查資料列
    - 第一階段 (check_stage1)：重複列、月結年月、日期、數量/折價/金額、中分類、收支、申報個帳、認列金額範圍
    - 第二階段 (check_stage2)：認列碼與月結年月/申報個帳是否一致、認列碼跳號、認列金額與金額、收支與正負號
3. 輸出
3-1. 文字介面：顯示讀取的檔名與資料筆數、各項問題的等級/原因/列數，以及異常報告檔名
    家庭共帳_年結2026.xlsx: (119, 25)
    [warning] 認列金額超過 +8000 或低於 -2000：6 列
    [error] 認列金額的絕對值不等於金額：3 列
    檢查結果：_異常報告_202609_20261009_061317.xlsx
3-2. 檔案：產生一份 _異常報告_<範圍>_<時間戳記>.xlsx，所有儲存格皆為文字格式，
    - 範圍為 全部、YYYY、YYYYMM 或 YYYYMM-YYYYMM；沒有問題列時不產生檔案。
"""
import argparse
import re
from datetime import datetime, time
from pathlib import Path

import pandas as pd

SRC_DIR = Path(__file__).resolve().parent

COL_YYYYMM = "月結年月"
REPORT_TAG = "_異常報告_"

WARN_AMOUNT_MAX = 8000    # 認列金額超過此值發出 warning
WARN_AMOUNT_MIN = -2000   # 認列金額低於此值發出 warning


def find_latest_file():
    """回傳 SRC_DIR 中修改日期最新的 xlsx（略過 Excel 開檔時產生的 ~$ 暫存檔及異常報告）。"""
    files = [p for p in SRC_DIR.glob("*.xlsx") if not p.name.startswith("~$") and REPORT_TAG not in p.name]
    if not files:
        raise FileNotFoundError(f"{SRC_DIR} 找不到 xlsx 檔案")
    return max(files, key=lambda p: p.stat().st_mtime)


def to_text(v):
    """
    轉成輸出用的文字
    空值為空字串、整數值的浮點數去掉 .0、午夜的日期只留日期。
    """
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d" if v.time() == time(0) else "%Y-%m-%d %H:%M:%S")
    return str(v)


def write_issues(df, issues, path):
    """
    把問題列寫成 xlsx，欄位為：列數、等級、原因、原始欄位...

    issues 為 [(等級, 原因, 問題列遮罩), ...]，遮罩的 index 須與 df 相同，
    列數換算成 Excel 列號（index + 2，第 1 列是標題）。
    所有儲存格都寫成文字格式，避免 Excel 自動轉型（例如去掉開頭的 0）。
    """
    frames = []
    for level, reason, mask in issues:
        rows = df[mask]
        frames.append(pd.concat([
            pd.DataFrame({"列數": rows.index + 2, "等級": level, "原因": reason}, index=rows.index),
            rows,
        ], axis=1))
    out = pd.concat(frames).apply(lambda col: col.map(to_text))

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="異常報告", freeze_panes=(1, 0))
        for row in writer.sheets["異常報告"].iter_rows():
            for cell in row:
                cell.number_format = "@"


def to_number(s):
    """轉成數字；容許千分位逗號（例如 "82,042"），無法轉換的變成 NaN。"""
    return pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(), errors="coerce")


def to_date(v):
    """轉成日期；容許 Excel 日期、"YYYY-MM-DD"、"M/D/YYYY"，其餘變成 NaT。"""
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
            try:
                return pd.to_datetime(v.strip(), format=fmt)
            except ValueError:
                pass
        return pd.NaT
    return pd.to_datetime(v, errors="coerce")


def check_stage1(df):
    """
    第一階段：檢查單一欄位的內容，回傳 [(等級, 原因, 問題列遮罩), ...]。
    整列空白的列只回報一次，不再列入其他檢查。
    """
    blank = df.isna().all(axis=1)
    issues = [("error", "整列空白", blank)]

    def add(level, reason, mask):
        issues.append((level, reason, mask & ~blank))

    #--1. 重複列：申報號碼_單頭、店名、日期、品項或說明、數量、金額、中分類、收支都相同，每一筆都標出
    dup_cols = ["申報號碼_單頭", "店名", "日期", "品項或說明", "數量", "金額", "中分類", "收支"]
    add("error", f"重複列（{'、'.join(dup_cols)}相同）", df.duplicated(subset=dup_cols, keep=False))

    #--2. 月結年月必須是 YYYYMM
    yyyymm = to_number(df[COL_YYYYMM])
    valid_ym = (yyyymm % 1 == 0) & yyyymm.between(190001, 999912) & (yyyymm % 100).between(1, 12)
    add("error", f"{COL_YYYYMM}不是 YYYYMM", ~valid_ym)

    #--3. 日期、核對處理日須為有效日期、不晚於今天
    today = pd.Timestamp.today().normalize()
    for col in ("日期", "核對處理日"):
        date = pd.to_datetime(df[col].map(to_date))
        add("error", f"{col}空白或不是有效日期", date.isna())
        add("error", f"{col}晚於今天", date > today)

    #--4. 數量、折價、金額須為正整數或 0
    for col in ("數量", "折價", "金額"):
        n = to_number(df[col])
        add("error", f"{col}不是正整數或0", ~((n >= 0) & (n % 1 == 0)))

    #--5. 中分類：第 1-3 碼英數字，第 4-7 碼中文字(一-鿿)
    add("error", "中分類格式不是 3 碼英數字 + 4 個中文字",
        ~df["中分類"].fillna("").astype(str).str.fullmatch(r"[A-Za-z0-9]{3}[\u4e00-\u9fff]{4}"))

    #--6. 收支只有收、支、期初
    add("error", "收支不是收、支、期初", ~df["收支"].isin(["收", "支", "期初"]))

    #--7. 申報個帳只有 A-E
    add("error", "申報個帳不是 A/B/C/D/E", ~df["申報個帳"].isin(list("ABCDE")))

    #--8. 認列金額超過 WARN_AMOUNT_MAX 或低於 WARN_AMOUNT_MIN
    n = to_number(df["認列金額"])
    add("warning", f"認列金額超過 {WARN_AMOUNT_MAX:+} 或低於 {WARN_AMOUNT_MIN}",
        (n > WARN_AMOUNT_MAX) | (n < WARN_AMOUNT_MIN))

    return issues


def check_stage2(df):
    """
    第二階段：檢查多個欄位之間的關係，回傳 [(等級, 原因, 問題列遮罩), ...]。
    整列空白的列已在第一階段回報，這裡略過。
    """
    blank = df.isna().all(axis=1)
    issues = []

    def add(level, reason, mask):
        issues.append((level, reason, mask.reindex(df.index, fill_value=False) & ~blank))

    code = df["認列碼"].fillna("").astype(str).str.strip()
    yyyymm = to_number(df[COL_YYYYMM])
    amount = to_number(df["金額"])
    recognized = to_number(df["認列金額"])

    #--1. 認列碼前兩碼 = 月結年月末兩碼
    ym_tail = yyyymm.astype("Int64").astype(str).str[-2:]
    add("error", f"認列碼前兩碼不等於{COL_YYYYMM}末兩碼", code.str[:2] != ym_tail)

    #--2. 認列碼第 3 碼 = 申報個帳
    add("error", "認列碼第3碼不等於申報個帳", code.str[2] != df["申報個帳"])

    #--3. 依月結年月 + 申報個帳分組，認列碼第 4-5 碼須從 00 起連號；缺號時標出缺號後的第一個號碼
    seq = to_number(code.str[3:5])
    for (ym, account), group in seq.groupby([yyyymm, df["申報個帳"]]):
        used = set(group.dropna().astype(int))
        missing = sorted(set(range(max(used, default=-1) + 1)) - used)
        if missing:
            after_gap = [n for n in used if n - 1 in missing]
            add("error", f"認列碼跳號：{int(ym)} {account} 缺 {', '.join(f'{n:02d}' for n in missing)}",
                group.isin(after_gap))

    #--4. 認列金額的絕對值 = 金額
    add("error", "認列金額的絕對值不等於金額", recognized.abs() != amount)

    #--5. 收支為「收」或「期初」時認列金額 >= 0，收支為「支」時 <= 0
    io = df["收支"]
    add("error", "收支為收或期初，認列金額卻是負數", io.isin(["收", "期初"]) & (recognized < 0))
    add("error", "收支為支，認列金額卻是正數", (io == "支") & (recognized > 0))

    return issues


def parse_month(text):
    """檢查 YYYYMM 字串格式，回傳整數 YYYYMM。"""
    if not re.fullmatch(r"\d{4}(0[1-9]|1[0-2])", text):
        raise argparse.ArgumentTypeError(f"格式應為 YYYYMM，收到：{text}")
    return int(text)


def build_parser():
    """建立命令列參數；--year、--month 二者擇一，--file 可搭配使用。"""
    parser = argparse.ArgumentParser(description="讀取家庭共帳年結檔")
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--year",
        type=int,
        metavar="YYYY",
        help=f"只取{COL_YYYYMM}屬於該年度的資料（例如 2026）",
    )
    group.add_argument(
        "--month",
        nargs="+",
        type=parse_month,
        metavar="YYYYMM",
        help=f"依{COL_YYYYMM}篩選；一個月別只取該月（例如 202609），兩個月別取起訖之間（例如 202607 202609）",
    )
    parser.add_argument(
        "--file",
        type=Path,
        metavar="FILE",
        help="指定要讀的檔案；只給檔名時會到 SRC_DIR 找。未指定時讀 SRC_DIR 中修改日期最新的 xlsx",
    )
    return parser


def main():
    #--輸入並解析月結年月
    parser = build_parser()
    args = parser.parse_args()
    #--檢查月結年月
    if args.month and len(args.month) > 2:
        parser.error("--month 最多只能給兩個月別（起月 迄月）")
    if args.month and len(args.month) == 2 and args.month[0] > args.month[1]:
        parser.error(f"--month 起月 {args.month[0]} 不可晚於迄月 {args.month[1]}")
    #--檢查檔案是否存在
    if args.file:
        path = args.file if args.file.exists() else SRC_DIR / args.file
        if not path.exists():
            raise FileNotFoundError(f"找不到檔案：{args.file}")
    else:
        #--若未輸入檔名，找最大修改日期的excel
        path = find_latest_file()

    #--讀取檔案
    df = pd.read_excel(path, sheet_name=0)
    yyyymm = to_number(df[COL_YYYYMM])
    #--指定的範圍超出資料集的月結年月範圍時，即時停止
    data_min, data_max = int(yyyymm.min()), int(yyyymm.max())
    out_of_range = (
        (args.year and args.year not in set((yyyymm // 100).dropna().astype(int)))
        or (args.month and (args.month[0] < data_min or args.month[-1] > data_max))
    )
    if out_of_range:
        wanted = args.year if args.year else "-".join(map(str, dict.fromkeys(args.month)))
        raise SystemExit(f"指定範圍 {wanted} 超出資料集的{COL_YYYYMM}範圍 {data_min}-{data_max}，停止執行")
    #--取出指定月結年/月結年月的資料集
    if args.year:
        df = df[yyyymm // 100 == args.year]
        scope = str(args.year)
    elif args.month:
        start, end = args.month[0], args.month[-1]
        df = df[yyyymm.between(start, end)]
        scope = str(start) if start == end else f"{start}-{end}"
    else:
        scope = "全部"

    print(f"{path.name}: {df.shape}")

    #--建立時戳並寫檔
    stamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    out_path = SRC_DIR / f"{REPORT_TAG}{scope}_{stamp}.xlsx"
    #--執行資料前處理(找出有問題/有疑慮的資料列)
    issues = [issue for issue in check_stage1(df) + check_stage2(df) if issue[2].any()]
    #--把問題列資料，輸出到畫面
    for level, reason, mask in issues:
        print(f"[{level}] {reason}：{mask.sum()} 列")
    #--把問題列資料，輸出成寫檔
    if issues:
        write_issues(df, issues, out_path)
        print(f"檢查結果：{out_path.name}")
    else:
        print("檢查通過，沒有問題列")


if __name__ == "__main__":
    main()
