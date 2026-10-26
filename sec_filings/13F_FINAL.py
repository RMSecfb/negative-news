from __future__ import annotations

from collections import Counter
from pathlib import Path
import os
import time

import pandas as pd
from edgar import Company, set_identity


# 0. 設定
# ============================================================
SEC_IDENTITY = os.environ.get("SEC_IDENTITY", "").strip()
if not SEC_IDENTITY:
    raise RuntimeError(
        "未設定 SEC_IDENTITY。SEC 要求自動化存取聲明名稱與聯絡信箱，"
        "請在 GitHub Secrets 或環境變數設定，例如：公司名稱 contact@公司網域"
    )

BASE_DIR = Path(__file__).resolve().parent

# Manager List：機構清單
MANAGER_FILE = BASE_DIR / "13F_Manager_List.xlsx"

# Ticker Universe：股票代碼清單
TICKER_UNIVERSE_FILE = BASE_DIR / "13F_Ticker_Universe.xlsx"

OUTPUT_DIR = BASE_DIR / "institutional_13f"
SLEEP_SECONDS = 0.15


# 1. 初始化
# ============================================================
set_identity(SEC_IDENTITY)

def get_manager_file() -> Path:
    if not MANAGER_FILE.exists():
        raise FileNotFoundError(
            f"找不到 Manager List：{MANAGER_FILE}"
        )
    return MANAGER_FILE

def normalize_period(value) -> str:
    if value is None or pd.isna(value):
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return str(value).strip()
    return parsed.strftime("%Y-%m-%d")


def period_to_quarter(period: str) -> str:
    parsed = pd.to_datetime(period, errors="coerce")
    if pd.isna(parsed):
        return period.replace("-", "")
    quarter = ((parsed.month - 1) // 3) + 1
    return f"{parsed.year}Q{quarter}"


def to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


# 2. 讀取 Ticker Universe
# ============================================================

def load_ticker_universe(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"找不到 Ticker Universe：{path}")

    frame = pd.read_excel(path, sheet_name=0, dtype=str)
    frame.columns = [
        str(c).replace("\u3000", " ").strip()
        for c in frame.columns
    ]

    aliases = {
        "Ticker": ("Ticker", "Symbol", "股票代號", "標的代碼"),
        "Company": ("Company", "Name", "股票名稱", "標的名稱"),
        "Active": ("Active", "啟用", "是否啟用"),
    }

    columns_lower = {
        str(c).strip().lower(): c
        for c in frame.columns
    }

    def find_col(names):
        for name in names:
            if name in frame.columns:
                return name
            if name.lower() in columns_lower:
                return columns_lower[name.lower()]
        return None

    ticker_col = find_col(aliases["Ticker"])
    company_col = find_col(aliases["Company"])
    active_col = find_col(aliases["Active"])

    if ticker_col is None:
        raise ValueError(
            "Ticker Universe 至少需要 Ticker / Symbol / 股票代號 欄位"
        )

    result = pd.DataFrame()
    result["Ticker"] = (
        frame[ticker_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )

    if company_col:
        result["Company"] = (
            frame[company_col]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    else:
        result["Company"] = ""

    if active_col:
        result["Active"] = (
            frame[active_col]
            .fillna("Y")
            .astype(str)
            .str.strip()
            .str.upper()
        )
    else:
        result["Active"] = "Y"

    result = result[
        (result["Ticker"] != "")
        & (result["Active"] == "Y")
    ].copy()

    result = result.drop_duplicates("Ticker").reset_index(drop=True)

    if result.empty:
        raise ValueError("Ticker Universe 沒有任何 Active=Y 的股票")

    return result


# 3. 清理單一期 holdings：
# 讀 report.holdings + 排除 PUT CALL + 只保留Universe
# 同 Institution × Ticker 最後只留一列
# ============================================================

def clean_holdings(report, target_tickers: set[str]) -> pd.DataFrame:

    holdings = report.holdings

    if holdings is None or len(holdings) == 0:
        return pd.DataFrame(columns=["Ticker", "Shares"])

    df = holdings.copy()

    required = {"Ticker", "SharesPrnAmount"}
    missing = required - set(df.columns)

    if missing:
        raise ValueError(
            f"holdings 缺少必要欄位：{sorted(missing)}；"
            f"實際欄位={df.columns.tolist()}"
        )

    df["Ticker"] = (
        df["Ticker"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 明確排除 option holdings
    if "PutCall" in df.columns:
        put_call = (
            df["PutCall"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        df = df[~put_call.isin({"PUT", "CALL"})].copy()

    df = df[df["Ticker"].isin(target_tickers)].copy()

    if df.empty:
        return pd.DataFrame(columns=["Ticker", "Shares"])

    df["SharesPrnAmount"] = to_numeric(df["SharesPrnAmount"])

    return (
        df.groupby("Ticker", as_index=False)
        .agg(Shares=("SharesPrnAmount", "sum"))
    )


def status_from_shares(current: float, previous: float) -> str:
    if previous == 0 and current > 0:
        return "NEW"
    if previous > 0 and current == 0:
        return "CLOSED"
    if current > previous:
        return "INCREASED"
    if current < previous:
        return "DECREASED"
    return "UNCHANGED"


def change_pct(current: float, previous: float):
    if previous == 0:
        return None if current > 0 else 0.0
    return (current - previous) / previous


# 4. 13F 訊號定義
# 持股量方向：看 Aggregate Shares QoQ
# 機構持股方向：看增持家數 vs 減持家數
# 13F訊號：同時保留「量」與「家數」兩種資訊
# ============================================================

def holding_amount_direction(qoq):
    if qoq is None:
        return "無前期比較資料"
    if qoq > 0:
        return "持股增加"
    if qoq < 0:
        return "持股下降"
    return "持股不變"


def institution_direction(increased: int, decreased: int):
    if increased > decreased:
        return "增持家數較多"
    if increased < decreased:
        return "減持家數較多"
    if increased == 0 and decreased == 0:
        return "無增減持"
    return "增減持家數相同"


def build_13f_signal(
    amount_direction: str,
    institution_holding_direction: str,
):
    if amount_direction == "無前期比較資料":
        return "無前期比較資料"

    return (
        f"{amount_direction}／"
        f"機構持有家數{institution_holding_direction}"
    )


# 5. 讀取 Manager List
# ============================================================

manager_file = get_manager_file()
managers = pd.read_excel(manager_file)

required_manager_columns = {
    "Institution",
    "CIK",
    "Active",
    "Validation Status",
    "Report Period",
}

missing_manager_columns = (
    required_manager_columns - set(managers.columns)
)

if missing_manager_columns:
    raise ValueError(
        f"Manager List 缺少欄位："
        f"{sorted(missing_manager_columns)}"
    )

for col in ["Institution", "Active", "Validation Status"]:
    managers[col] = (
        managers[col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

managers["Active"] = managers["Active"].str.upper()
managers["Validation Status"] = (
    managers["Validation Status"].str.upper()
)

if "Include in Aggregate" not in managers.columns:
    managers["Include in Aggregate"] = "Y"

managers["Include in Aggregate"] = (
    managers["Include in Aggregate"]
    .fillna("Y")
    .astype(str)
    .str.strip()
    .str.upper()
)

eligible = managers[
    (managers["Active"] == "Y")
    & (managers["Validation Status"] == "OK")
].copy()

if eligible.empty:
    raise RuntimeError(
        "沒有 Active=Y 且 Validation Status=OK 的機構"
    )


# 6. 載入所有股票清單
# ============================================================
universe_df = load_ticker_universe(TICKER_UNIVERSE_FILE)
TARGET_TICKERS = set(universe_df["Ticker"])

ticker_company_map = (
    universe_df
    .set_index("Ticker")["Company"]
    .to_dict()
)

print("=" * 100)
print("13F FINAL - FULL UNIVERSE FETCHER")
print("=" * 100)

print(f"Manager file : {manager_file}")
print(f"Ticker file  : {TICKER_UNIVERSE_FILE}")
print(f"Managers     : {len(eligible):,}")
print(f"Tickers      : {len(TARGET_TICKERS):,}")


# 7. 自動判斷共同最新季度
# ============================================================
periods = [
    normalize_period(v)
    for v in eligible["Report Period"]
    if normalize_period(v)
]

if not periods:
    raise RuntimeError("Manager List 沒有可用的 Report Period")

TARGET_REPORT_PERIOD = (
    Counter(periods)
    .most_common(1)[0][0]
)

QUARTER = period_to_quarter(
    TARGET_REPORT_PERIOD
)

print(f"Target period: {TARGET_REPORT_PERIOD}")
print(f"Quarter      : {QUARTER}")


# 8. Output
# ============================================================

QUARTER_DIR = OUTPUT_DIR / QUARTER
QUARTER_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

FINAL_FILE = (
    QUARTER_DIR
    / f"13F_{QUARTER}_Full_Universe.xlsx"
)


# 9. 選取各機構
# ============================================================

detail_frames = []
run_logs = []

for seq, (idx, original_row) in enumerate(
    eligible.iterrows(),
    start=1,
):

    institution = str(
        original_row["Institution"]
    ).strip()

    cik_raw = original_row["CIK"]

    include_setting = str(
        original_row.get(
            "Include in Aggregate",
            "Y",
        )
    ).strip().upper()

    try:
        cik = int(float(cik_raw))
    except Exception:
        run_logs.append({
            "Institution": institution,
            "CIK": cik_raw,
            "Status": "ERROR",
            "Aggregate": "N",
            "Target Period": TARGET_REPORT_PERIOD,
            "Actual Period": "",
            "Detail Rows": 0,
            "Message": "CIK 無法轉為整數",
        })
        continue

    print("\n" + "-" * 100)
    print(
        f"[{seq}/{len(eligible)}] "
        f"{institution} | CIK={cik}"
    )

    started = time.monotonic()

    try:
        manager = Company(cik)
        filings = manager.get_filings(
            form="13F-HR"
        )

        if filings is None or len(filings) == 0:
            raise ValueError(
                "找不到 13F-HR"
            )

        latest_filing = filings.latest()
        current_report = latest_filing.obj()

        actual_period = normalize_period(
            getattr(
                current_report,
                "report_period",
                None,
            )
        )

        # 最新資料非共同季度：STALE
        if actual_period != TARGET_REPORT_PERIOD:
            elapsed = time.monotonic() - started

            run_logs.append({
                "Institution": institution,
                "CIK": cik,
                "Status": "STALE",
                "Aggregate": "N",
                "Target Period": TARGET_REPORT_PERIOD,
                "Actual Period": actual_period,
                "Detail Rows": 0,
                "Seconds": round(elapsed, 2),
                "Message": "最新13F非目標季度",
            })

            print(
                f"STALE | actual={actual_period}"
            )

            time.sleep(SLEEP_SECONDS)
            continue

        # Manager governance
        if include_setting != "Y":
            elapsed = time.monotonic() - started

            run_logs.append({
                "Institution": institution,
                "CIK": cik,
                "Status": "REVIEW",
                "Aggregate": "N",
                "Target Period": TARGET_REPORT_PERIOD,
                "Actual Period": actual_period,
                "Detail Rows": 0,
                "Seconds": round(elapsed, 2),
                "Message": (
                    "Include in Aggregate="
                    f"{include_setting}"
                ),
            })

            print(
                f"REVIEW | "
                f"Include={include_setting}"
            )

            time.sleep(SLEEP_SECONDS)
            continue

        previous_report = (
            current_report
            .previous_holding_report()
        )

        if previous_report is None:
            raise ValueError(
                "找不到前一季 holding report"
            )

        previous_period = normalize_period(
            getattr(
                previous_report,
                "report_period",
                None,
            )
        )

        current_df = clean_holdings(
            current_report,
            TARGET_TICKERS,
        )

        previous_df = clean_holdings(
            previous_report,
            TARGET_TICKERS,
        )

        merged = current_df.merge(
            previous_df,
            on="Ticker",
            how="outer",
            suffixes=("", "_Previous"),
        )

        if "Shares" not in merged.columns:
            merged["Shares"] = 0.0

        if "Shares_Previous" not in merged.columns:
            merged["Shares_Previous"] = 0.0

        merged = merged.rename(
            columns={
                "Shares_Previous":
                    "Previous Shares",
            }
        )

        merged["Shares"] = (
            to_numeric(merged["Shares"])
        )

        merged["Previous Shares"] = (
            to_numeric(
                merged["Previous Shares"]
            )
        )

        merged["Change Shares"] = (
            merged["Shares"]
            - merged["Previous Shares"]
        )

        merged["Change %"] = merged.apply(
            lambda r: change_pct(
                float(r["Shares"]),
                float(r["Previous Shares"]),
            ),
            axis=1,
        )

        merged["Status"] = merged.apply(
            lambda r: status_from_shares(
                float(r["Shares"]),
                float(r["Previous Shares"]),
            ),
            axis=1,
        )

        merged["Quarter"] = QUARTER
        merged["Quarter End"] = (
            TARGET_REPORT_PERIOD
        )
        merged["Previous Quarter End"] = (
            previous_period
        )
        merged["Company"] = (
            merged["Ticker"]
            .map(ticker_company_map)
            .fillna("")
        )
        merged["Institution"] = institution
        merged["CIK"] = cik
        merged["Filing Date"] = str(
            getattr(
                latest_filing,
                "filing_date",
                "",
            )
        )

        if not merged.empty:
            merged = merged[
                [
                    "Quarter",
                    "Quarter End",
                    "Previous Quarter End",
                    "Ticker",
                    "Company",
                    "Institution",
                    "CIK",
                    "Filing Date",
                    "Shares",
                    "Previous Shares",
                    "Change Shares",
                    "Change %",
                    "Status",
                ]
            ]

            detail_frames.append(merged)

        elapsed = time.monotonic() - started

        run_logs.append({
            "Institution": institution,
            "CIK": cik,
            "Status": "OK",
            "Aggregate": "Y",
            "Target Period": TARGET_REPORT_PERIOD,
            "Actual Period": actual_period,
            "Previous Period": previous_period,
            "Detail Rows": len(merged),
            "Seconds": round(elapsed, 2),
            "Message": "",
        })

        print(
            f"OK | rows={len(merged):,} "
            f"| {elapsed:.2f}s"
        )

    except Exception as exc:
        elapsed = time.monotonic() - started

        run_logs.append({
            "Institution": institution,
            "CIK": cik,
            "Status": "ERROR",
            "Aggregate": "N",
            "Target Period": TARGET_REPORT_PERIOD,
            "Actual Period": "",
            "Detail Rows": 0,
            "Seconds": round(elapsed, 2),
            "Message": str(exc),
        })

        print(
            f"ERROR | {exc}"
        )

    time.sleep(SLEEP_SECONDS)

# 10. Detail
# ============================================================
if detail_frames:
    detail_df = pd.concat(
        detail_frames,
        ignore_index=True,
    )
else:
    detail_df = pd.DataFrame(
        columns=[
            "Quarter",
            "Quarter End",
            "Previous Quarter End",
            "Ticker",
            "Company",
            "Institution",
            "CIK",
            "Filing Date",
            "Shares",
            "Previous Shares",
            "Change Shares",
            "Change %",
            "Status",
        ]
    )

detail_df = (
    detail_df
    .sort_values(
        ["Ticker", "Institution"]
    )
    .reset_index(drop=True)
)

run_log_df = pd.DataFrame(
    run_logs
)

aggregate_manager_count = (
    int(
        (
            (run_log_df["Status"] == "OK")
            & (
                run_log_df["Aggregate"]
                == "Y"
            )
        ).sum()
    )
    if not run_log_df.empty
    else 0
)


# 11. Full Universe Summary
# ============================================================
summary_rows = []

for _, ticker_row in universe_df.iterrows():

    ticker = ticker_row["Ticker"]
    company = ticker_row["Company"]

    group = detail_df[
        detail_df["Ticker"] == ticker
    ].copy()

    current_shares = (
        float(group["Shares"].sum())
        if not group.empty
        else 0.0
    )

    previous_shares = (
        float(
            group["Previous Shares"].sum()
        )
        if not group.empty
        else 0.0
    )

    if previous_shares > 0:
        qoq = (
            current_shares
            - previous_shares
        ) / previous_shares
    else:
        qoq = (
            None
            if current_shares > 0
            else 0.0
        )

    status_counts = (
        group["Status"].value_counts()
        if not group.empty
        else pd.Series(dtype=int)
    )

    increased = int(
        status_counts.get(
            "INCREASED",
            0,
        )
    )

    decreased = int(
        status_counts.get(
            "DECREASED",
            0,
        )
    )

    new_positions = int(
        status_counts.get(
            "NEW",
            0,
        )
    )

    closed_positions = int(
        status_counts.get(
            "CLOSED",
            0,
        )
    )

    unchanged = int(
        status_counts.get(
            "UNCHANGED",
            0,
        )
    )

    net_institution_count = (
        increased - decreased
    )

    amount_dir = (
        holding_amount_direction(qoq)
    )

    institution_dir = (
        institution_direction(
            increased,
            decreased,
        )
    )

    signal = build_13f_signal(
        amount_dir,
        institution_dir,
    )

    summary_rows.append({
        "Quarter": QUARTER,
        "Quarter End": TARGET_REPORT_PERIOD,
        "Ticker": ticker,
        "Company": company,
        "追蹤Universe": len(eligible),
        "本季納入機構數":
            aggregate_manager_count,
        "本季持有機構數": (
            group.loc[
                group["Shares"] > 0,
                "Institution",
            ].nunique()
            if not group.empty
            else 0
        ),
        "前季持有機構數": (
            group.loc[
                group[
                    "Previous Shares"
                ] > 0,
                "Institution",
            ].nunique()
            if not group.empty
            else 0
        ),
        "Current Shares":
            current_shares,
        "Previous Shares":
            previous_shares,
        "機構持股QoQ": qoq,
        "增持家數": increased,
        "減持家數": decreased,
        "新建倉": new_positions,
        "清倉": closed_positions,
        "持股不變家數": unchanged,
        "淨增持家數":
            net_institution_count,
        "持股量方向":
            amount_dir,
        "機構持股方向":
            institution_dir,
        "13F訊號":
            signal,
    })

summary_df = pd.DataFrame(
    summary_rows
)


# 12. Dashboard 精簡版
# ============================================================
dashboard_df = summary_df[
    [
        "Quarter",
        "Ticker",
        "Company",
        "機構持股QoQ",
        "增持家數",
        "減持家數",
        "本季持有機構數",
        "持股量方向",
        "機構持股方向",
        "13F訊號",
    ]
].copy()

dashboard_df["13F顯示"] = (
    dashboard_df.apply(
        lambda r:
            (
                f"{'↑' if r['機構持股QoQ'] > 0 else '↓' if r['機構持股QoQ'] < 0 else '→'} "
                f"{r['機構持股QoQ']:.1%}"
                f"｜{int(r['增持家數'])}增/"
                f"{int(r['減持家數'])}減"
            )
            if pd.notna(r["機構持股QoQ"])
            else "無前期比較資料",
        axis=1,
    )
)


# 13. Excel 輸出，一次完成
# ============================================================
with pd.ExcelWriter(
    FINAL_FILE,
    engine="openpyxl",
) as writer:

    summary_df.to_excel(
        writer,
        sheet_name="Ticker Summary",
        index=False,
    )

    dashboard_df.to_excel(
        writer,
        sheet_name="Dashboard Summary",
        index=False,
    )

    detail_df.to_excel(
        writer,
        sheet_name="13F Detail",
        index=False,
    )

    run_log_df.to_excel(
        writer,
        sheet_name="Run Log",
        index=False,
    )

    eligible.to_excel(
        writer,
        sheet_name="Manager Universe",
        index=False,
    )

    universe_df.to_excel(
        writer,
        sheet_name="Ticker Universe",
        index=False,
    )


# 14. Console
# ============================================================

print("\n" + "=" * 100)
print("完成")
print("=" * 100)

if not run_log_df.empty:
    print("\nRun Status:")
    print(
        run_log_df["Status"]
        .value_counts(
            dropna=False
        )
    )

print(
    f"\nTicker Summary 筆數："
    f"{len(summary_df):,}"
)

print(
    f"13F Detail 筆數："
    f"{len(detail_df):,}"
)

print(
    f"\n輸出檔案："
    f"{FINAL_FILE}"
)

print(
    "\n13F 訊號說明："
    "\n- 持股量方向：看追蹤機構合計 Shares QoQ"
    "\n- 機構持股方向：比較增持家數與減持家數"
)
