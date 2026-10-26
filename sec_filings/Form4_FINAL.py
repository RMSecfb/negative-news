from __future__ import annotations

from datetime import datetime, timedelta, date
from pathlib import Path
import os
import math
import time

import pandas as pd
from edgar import Company, set_identity


# 0. 使用者設定
# ============================================================
SEC_IDENTITY = os.environ.get("SEC_IDENTITY", "").strip()
if not SEC_IDENTITY:
    raise RuntimeError(
        "未設定 SEC_IDENTITY。SEC 要求自動化存取聲明名稱與聯絡信箱，"
        "請在 GitHub Secrets 或環境變數設定，例如：公司名稱 contact@公司網域"
    )

BASE_DIR = Path(__file__).resolve().parent

# 與13F共用
TICKER_UNIVERSE_FILE = BASE_DIR / "13F_Ticker_Universe.xlsx"

OUTPUT_DIR = BASE_DIR / "insider_form4"

LOOKBACK_DAYS = 30

# 為避免filing date 比trade date 晚，抓取區間多往前15天
FILING_BUFFER_DAYS = 15

SLEEP_SECONDS = 0.12

# 方向性訊號：SEC Code的P / S
DIRECTIONAL_CODES = {"P", "S"}


# 1. 初始化
# ============================================================
set_identity(SEC_IDENTITY)

def safe_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def safe_float(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        x = float(value)
        if math.isnan(x):
            return None
        return x
    except Exception:
        return None


def normalize_date(value) -> str:
    if value is None:
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return safe_text(value)
    return parsed.strftime("%Y-%m-%d")


def date_in_window(value, start_date: date, end_date: date) -> bool:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return False
    d = parsed.date()
    return start_date <= d <= end_date

# 2. Ticker Universe
# ============================================================

def load_ticker_universe(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"找不到 Ticker Universe：{path}")

    frame = pd.read_excel(path, sheet_name=0, dtype=str)
    frame.columns = [
        str(c).replace("\u3000", " ").strip()
        for c in frame.columns
    ]

    lower_map = {
        str(c).strip().lower(): c
        for c in frame.columns
    }

    def find_col(names):
        for name in names:
            if name in frame.columns:
                return name
            if name.lower() in lower_map:
                return lower_map[name.lower()]
        return None

    ticker_col = find_col(
        ["Ticker", "Symbol", "股票代號", "標的代碼"]
    )
    company_col = find_col(
        ["Company", "Name", "股票名稱", "標的名稱"]
    )
    active_col = find_col(
        ["Active", "啟用", "是否啟用"]
    )

    if ticker_col is None:
        raise ValueError("Ticker Universe 缺少 Ticker / Symbol 欄位")

    result = pd.DataFrame()

    result["Ticker"] = (
        frame[ticker_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )

    result["Company"] = (
        frame[company_col].fillna("").astype(str).str.strip()
        if company_col
        else ""
    )

    result["Active"] = (
        frame[active_col].fillna("Y").astype(str).str.strip().str.upper()
        if active_col
        else "Y"
    )

    result = result[
        (result["Ticker"] != "")
        & (result["Active"] == "Y")
    ].copy()

    result = result.drop_duplicates("Ticker").reset_index(drop=True)

    if result.empty:
        raise ValueError("Ticker Universe 沒有任何 Active=Y 股票")

    return result


# 3. Form 4 metadata
# ============================================================

def get_insider_name(form4) -> str:
    # EdgarTools 官方 Form4 metadata
    value = getattr(form4, "insider_name", None)
    if value:
        return safe_text(value)

    try:
        summary = form4.get_ownership_summary()
        value = getattr(summary, "insider_name", None)
        if value:
            return safe_text(value)
    except Exception:
        pass

    return ""


def get_position(form4) -> str:
    value = getattr(form4, "position", None)
    if value:
        return safe_text(value)

    try:
        summary = form4.get_ownership_summary()
        value = getattr(summary, "position", None)
        if value:
            return safe_text(value)
    except Exception:
        pass

    return ""


def get_10b5_1(form4) -> str:
    try:
        summary = form4.get_ownership_summary()
        value = getattr(summary, "has_10b5_1_plan", None)

        if value is True:
            return "Y"
        if value is False:
            return "N"

    except Exception:
        pass

    return "Unknown"



def get_issuer_ticker(form4) -> str:
    """取得 Form 4真正的issuer ticker。"""
    try:
        issuer = getattr(form4, "issuer", None)
        if issuer is not None:
            ticker = getattr(issuer, "ticker", None)
            if ticker:
                return safe_text(ticker).upper()
    except Exception:
        pass

    for attr in ("issuer_ticker",):
        try:
            ticker = getattr(form4, attr, None)
            if ticker:
                return safe_text(ticker).upper()
        except Exception:
            pass

    return ""

# 4. DataFrame 欄位工具：回傳 {lowercase欄名: 原欄名}
# ============================================================

def normalize_columns(df: pd.DataFrame) -> dict[str, str]:
    return {
        str(c).strip().lower(): c
        for c in df.columns
    }


def pick_column(df: pd.DataFrame, candidates: list[str]):
    lower_map = normalize_columns(df)

    for candidate in candidates:
        if candidate in df.columns:
            return candidate

        key = candidate.strip().lower()
        if key in lower_map:
            return lower_map[key]

    return None


def row_value(row, column):
    if column is None:
        return None

    try:
        return row[column]
    except Exception:
        return None


# 5. Signal
# ============================================================

def calculate_npr(buy_count: int, sell_count: int):
    denominator = buy_count + sell_count

    if denominator == 0:
        return None

    return (buy_count - sell_count) / denominator


def signal_en(npr):
    if npr is None:
        return "No P/S Activity"
    if npr > 0:
        return "Net Buying"
    if npr < 0:
        return "Net Selling"
    return "Neutral"


def signal_zh(npr):
    if npr is None:
        return "無P/S交易"
    if npr > 0:
        return "近期淨買入"
    if npr < 0:
        return "近期淨賣出"
    return "買賣中性"


# 6. 日期設定
# ============================================================

TODAY = datetime.now().date()

ANALYSIS_END = TODAY

ANALYSIS_START = (
    ANALYSIS_END
    - timedelta(days=LOOKBACK_DAYS - 1)
)

FILING_FETCH_START = (
    ANALYSIS_START
    - timedelta(days=FILING_BUFFER_DAYS)
)

FILING_FETCH_END = ANALYSIS_END

FILING_DATE_RANGE = (
    f"{FILING_FETCH_START:%Y-%m-%d}:"
    f"{FILING_FETCH_END:%Y-%m-%d}"
)

RUN_STAMP = TODAY.strftime("%Y-%m-%d")

RUN_DIR = OUTPUT_DIR / RUN_STAMP
RUN_DIR.mkdir(parents=True, exist_ok=True)

FINAL_FILE = (
    RUN_DIR
    / f"Form4_{RUN_STAMP}_{LOOKBACK_DAYS}D_Full_Universe.xlsx"
)


# 7. 股票Universe
# ============================================================

universe_df = load_ticker_universe(TICKER_UNIVERSE_FILE)

print("=" * 100)
print("FORM 4 FINAL v2 - FULL UNIVERSE")
print("=" * 100)

print(f"Ticker file     : {TICKER_UNIVERSE_FILE}")
print(f"Tickers         : {len(universe_df):,}")
print(
    f"Analysis period : "
    f"{ANALYSIS_START:%Y-%m-%d} ~ {ANALYSIS_END:%Y-%m-%d}"
)
print(f"Filing fetch    : {FILING_DATE_RANGE}")
print("Directional     : P / S only")


# 8. 確認 Form 4 資料
# P/S 使用 form4.market_trades DataFrame
# 全交易使用 form4.to_dataframe()
# ============================================================
directional_rows = []
excluded_rows = []
run_logs = []


for seq, ticker_row in enumerate(
    universe_df.itertuples(index=False),
    start=1,
):

    ticker = safe_text(getattr(ticker_row, "Ticker")).upper()
    company_name = safe_text(getattr(ticker_row, "Company"))

    print("\n" + "-" * 100)
    print(
        f"[{seq}/{len(universe_df)}] "
        f"{ticker} | {company_name}"
    )

    started = time.monotonic()

    filing_count = 0
    amendment_count = 0
    parsed_transactions = 0
    directional_count = 0
    excluded_count = 0
    issuer_mismatch_count = 0
    issuer_unknown_count = 0
    parse_errors = 0

    try:
        company = Company(ticker)

        filings = company.get_filings(
            form="4",
            filing_date=FILING_DATE_RANGE,
        )

        try:
            amendments = company.get_filings(
                form="4/A",
                filing_date=FILING_DATE_RANGE,
            )
            amendment_count = (
                len(amendments)
                if amendments is not None
                else 0
            )
        except Exception:
            amendment_count = 0

        filings_iterable = (
            filings
            if filings is not None
            else []
        )

        filing_count = (
            len(filings)
            if filings is not None
            else 0
        )

        for filing in filings_iterable:

            try:
                form4 = filing.obj()

                if form4 is None:
                    continue

                # Issuer validation：
                # Company(ticker) 也可能因其為別家公司 reporting owner 而出現 Form 4。
                # 只有真正 issuer ticker == 目前分析 ticker 才納入。
                issuer_ticker = get_issuer_ticker(form4)

                if issuer_ticker:
                    if issuer_ticker != ticker:
                        issuer_mismatch_count += 1
                        continue
                else:
                    # 無法辨識 issuer 時不直接誤刪，但會在 Run Log。
                    issuer_unknown_count += 1

                insider = get_insider_name(form4)
                position = get_position(form4)
                plan_10b5 = get_10b5_1(form4)

                filing_date = normalize_date(
                    getattr(filing, "filing_date", "")
                )

                accession = safe_text(
                    getattr(
                        filing,
                        "accession_number",
                        getattr(filing, "accession_no", ""),
                    )
                )

                # ------------------------------------------------
                # A. 正式方向性交易：market_trades
                # 官方欄位：
                # Date / Security / Shares / Price / Remaining /
                # AcquiredDisposed / Code
                # ------------------------------------------------

                market_df = getattr(
                    form4,
                    "market_trades",
                    None,
                )

                if market_df is not None and not market_df.empty:

                    cols = normalize_columns(market_df)

                    date_col = pick_column(
                        market_df,
                        ["Date", "Transaction Date"]
                    )
                    security_col = pick_column(
                        market_df,
                        ["Security", "Security Title"]
                    )
                    shares_col = pick_column(
                        market_df,
                        ["Shares"]
                    )
                    price_col = pick_column(
                        market_df,
                        ["Price", "Price Per Share"]
                    )
                    remaining_col = pick_column(
                        market_df,
                        ["Remaining", "Shares After"]
                    )
                    acquired_col = pick_column(
                        market_df,
                        ["AcquiredDisposed", "Acquired/Disposed"]
                    )
                    code_col = pick_column(
                        market_df,
                        ["Code", "Transaction Code"]
                    )

                    for _, trade in market_df.iterrows():

                        trade_date = normalize_date(
                            row_value(trade, date_col)
                        )

                        # 真正分析以交易日為準
                        if not date_in_window(
                            trade_date,
                            ANALYSIS_START,
                            ANALYSIS_END,
                        ):
                            continue

                        code = safe_text(
                            row_value(trade, code_col)
                        ).upper()

                        # market_trades 原則上就是 P/S
                        if code not in DIRECTIONAL_CODES:
                            continue

                        shares = safe_float(
                            row_value(trade, shares_col)
                        )

                        price = safe_float(
                            row_value(trade, price_col)
                        )

                        amount = (
                            shares * price
                            if (
                                shares is not None
                                and price is not None
                            )
                            else None
                        )

                        directional_rows.append({
                            "Trade Date": trade_date,
                            "Filing Date": filing_date,
                            "Ticker": ticker,
                            "Issuer Ticker": issuer_ticker,
                            "Company": company_name,
                            "Insider": insider,
                            "Position": position,
                            "Transaction Code": code,
                            "Type": (
                                "Buy"
                                if code == "P"
                                else "Sell"
                            ),
                            "Security": safe_text(
                                row_value(
                                    trade,
                                    security_col,
                                )
                            ),
                            "Shares": shares,
                            "Price": price,
                            "Amount": amount,
                            "Acquired/Disposed": safe_text(
                                row_value(
                                    trade,
                                    acquired_col,
                                )
                            ),
                            "Shares After": safe_float(
                                row_value(
                                    trade,
                                    remaining_col,
                                )
                            ),
                            "10b5-1 Plan": plan_10b5,
                            "Accession Number": accession,
                        })

                        directional_count += 1

                # B. 全部交易：to_dataframe()
                # 用來保留 A/M/F/G/C... 等非 P/S
                # ------------------------------------------------

                all_df = form4.to_dataframe()

                if all_df is None or all_df.empty:
                    continue

                parsed_transactions += len(all_df)

                date_col = pick_column(
                    all_df,
                    ["Transaction Date", "Date"]
                )
                code_col = pick_column(
                    all_df,
                    ["Code", "Transaction Code"]
                )
                security_col = pick_column(
                    all_df,
                    ["Security", "Security Title"]
                )
                shares_col = pick_column(
                    all_df,
                    ["Shares"]
                )
                price_col = pick_column(
                    all_df,
                    ["Price", "Price Per Share"]
                )
                remaining_col = pick_column(
                    all_df,
                    ["Remaining", "Shares After"]
                )
                acquired_col = pick_column(
                    all_df,
                    ["AcquiredDisposed", "Acquired/Disposed"]
                )
                insider_col = pick_column(
                    all_df,
                    ["Insider"]
                )
                position_col = pick_column(
                    all_df,
                    ["Position"]
                )

                for _, trade in all_df.iterrows():

                    trade_date = normalize_date(
                        row_value(trade, date_col)
                    )

                    if not date_in_window(
                        trade_date,
                        ANALYSIS_START,
                        ANALYSIS_END,
                    ):
                        continue

                    code = safe_text(
                        row_value(trade, code_col)
                    ).upper()

                    # P/S 已由 market_trades 處理，這邊只保留其他 code。
                    if code in DIRECTIONAL_CODES:
                        continue

                    shares = safe_float(
                        row_value(trade, shares_col)
                    )

                    price = safe_float(
                        row_value(trade, price_col)
                    )

                    amount = (
                        shares * price
                        if (
                            shares is not None
                            and price is not None
                        )
                        else None
                    )

                    excluded_rows.append({
                        "Trade Date": trade_date,
                        "Filing Date": filing_date,
                        "Ticker": ticker,
                        "Issuer Ticker": issuer_ticker,
                        "Company": company_name,
                        "Insider": (
                            safe_text(
                                row_value(
                                    trade,
                                    insider_col,
                                )
                            )
                            or insider
                        ),
                        "Position": (
                            safe_text(
                                row_value(
                                    trade,
                                    position_col,
                                )
                            )
                            or position
                        ),
                        "Transaction Code": code,
                        "Type": "Excluded",
                        "Security": safe_text(
                            row_value(
                                trade,
                                security_col,
                            )
                        ),
                        "Shares": shares,
                        "Price": price,
                        "Amount": amount,
                        "Acquired/Disposed": safe_text(
                            row_value(
                                trade,
                                acquired_col,
                            )
                        ),
                        "Shares After": safe_float(
                            row_value(
                                trade,
                                remaining_col,
                            )
                        ),
                        "10b5-1 Plan": plan_10b5,
                        "Accession Number": accession,
                    })

                    excluded_count += 1

            except Exception as exc:

                parse_errors += 1

                print(
                    f"  Filing parse error: "
                    f"{getattr(filing, 'accession_number', '')} | {exc}"
                )

        elapsed = time.monotonic() - started

        status = (
            "OK"
            if parse_errors == 0
            else "PARTIAL"
        )

        run_logs.append({
            "Ticker": ticker,
            "Company": company_name,
            "Status": status,
            "Form 4 Filings": filing_count,
            "Form 4/A Amendments": amendment_count,
            "Parsed Transactions": parsed_transactions,
            "P/S Transactions": directional_count,
            "Excluded Transactions": excluded_count,
            "Issuer Mismatch Filings": issuer_mismatch_count,
            "Issuer Unknown Filings": issuer_unknown_count,
            "Parse Errors": parse_errors,
            "Analysis Start": ANALYSIS_START.isoformat(),
            "Analysis End": ANALYSIS_END.isoformat(),
            "Seconds": round(elapsed, 2),
            "Message": (
                "存在 Form 4/A，請於 Run Log 留意"
                if amendment_count > 0
                else ""
            ),
        })

        print(
            f"{status} | filings={filing_count} "
            f"| parsed={parsed_transactions} "
            f"| P/S={directional_count} "
            f"| excluded={excluded_count} "
            f"| issuer_mismatch={issuer_mismatch_count} "
            f"| issuer_unknown={issuer_unknown_count} "
            f"| amendments={amendment_count} "
            f"| {elapsed:.2f}s"
        )

    except Exception as exc:

        elapsed = time.monotonic() - started

        print(f"ERROR | {exc}")

        run_logs.append({
            "Ticker": ticker,
            "Company": company_name,
            "Status": "ERROR",
            "Form 4 Filings": filing_count,
            "Form 4/A Amendments": amendment_count,
            "Parsed Transactions": parsed_transactions,
            "P/S Transactions": directional_count,
            "Excluded Transactions": excluded_count,
            "Issuer Mismatch Filings": issuer_mismatch_count,
            "Issuer Unknown Filings": issuer_unknown_count,
            "Parse Errors": parse_errors,
            "Analysis Start": ANALYSIS_START.isoformat(),
            "Analysis End": ANALYSIS_END.isoformat(),
            "Seconds": round(elapsed, 2),
            "Message": str(exc),
        })

    time.sleep(SLEEP_SECONDS)


# 9. Detail
# ============================================================

detail_columns = [
    "Trade Date",
    "Filing Date",
    "Ticker",
    "Issuer Ticker",
    "Company",
    "Insider",
    "Position",
    "Transaction Code",
    "Type",
    "Security",
    "Shares",
    "Price",
    "Amount",
    "Acquired/Disposed",
    "Shares After",
    "10b5-1 Plan",
    "Accession Number",
]

detail_df = pd.DataFrame(
    directional_rows,
    columns=detail_columns,
)

excluded_df = pd.DataFrame(
    excluded_rows,
    columns=detail_columns,
)

if not detail_df.empty:
    detail_df = (
        detail_df
        .sort_values(
            ["Trade Date", "Ticker", "Insider"],
            ascending=[False, True, True],
        )
        .reset_index(drop=True)
    )

if not excluded_df.empty:
    excluded_df = (
        excluded_df
        .sort_values(
            ["Trade Date", "Ticker", "Insider"],
            ascending=[False, True, True],
        )
        .reset_index(drop=True)
    )



# 9.5 Transaction-level Dedup
# ============================================================

DEDUP_KEYS = [
    "Accession Number",
    "Trade Date",
    "Ticker",
    "Issuer Ticker",
    "Insider",
    "Transaction Code",
    "Security",
    "Shares",
    "Price",
]

def deduplicate_transactions(frame: pd.DataFrame):
    if frame is None or frame.empty:
        return frame, 0

    keys = [c for c in DEDUP_KEYS if c in frame.columns]
    before = len(frame)

    result = (
        frame
        .drop_duplicates(subset=keys, keep="last")
        .reset_index(drop=True)
    )

    return result, before - len(result)


detail_df, ps_duplicates_removed = deduplicate_transactions(detail_df)
excluded_df, excluded_duplicates_removed = deduplicate_transactions(excluded_df)


# 10. Summary
# ============================================================

summary_rows = []

for ticker_row in universe_df.itertuples(index=False):

    ticker = safe_text(
        getattr(ticker_row, "Ticker")
    )
    company_name = safe_text(
        getattr(ticker_row, "Company")
    )

    group = detail_df[
        detail_df["Ticker"] == ticker
    ].copy()

    buys = group[
        group["Transaction Code"] == "P"
    ]

    sells = group[
        group["Transaction Code"] == "S"
    ]

    buy_count = len(buys)
    sell_count = len(sells)

    buy_insider_count = (
        buys["Insider"]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    sell_insider_count = (
        sells["Insider"]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    buy_shares = (
        pd.to_numeric(
            buys["Shares"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    sell_shares = (
        pd.to_numeric(
            sells["Shares"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    buy_amount = (
        pd.to_numeric(
            buys["Amount"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    sell_amount = (
        pd.to_numeric(
            sells["Amount"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    net_amount = buy_amount - sell_amount

    npr = calculate_npr(
        buy_count,
        sell_count,
    )

    plan_ps_count = (
        int((group["10b5-1 Plan"] == "Y").sum())
        if not group.empty
        else 0
    )

    summary_rows.append({
        "As of Date": ANALYSIS_END.isoformat(),
        "Ticker": ticker,
        "Company": company_name,
        "Lookback": f"{LOOKBACK_DAYS}D",
        "Analysis Start": ANALYSIS_START.isoformat(),
        "Analysis End": ANALYSIS_END.isoformat(),
        "Buy Count": buy_count,
        "Sell Count": sell_count,
        "Buy Insider Count": int(buy_insider_count),
        "Sell Insider Count": int(sell_insider_count),
        "Buy Shares": float(buy_shares),
        "Sell Shares": float(sell_shares),
        "Buy Amount": float(buy_amount),
        "Sell Amount": float(sell_amount),
        "Net Amount": float(net_amount),
        "NPR": npr,
        "Signal": signal_en(npr),
        "內部人交易動向": signal_zh(npr),
        "10b5-1 P/S Count": plan_ps_count,
    })

summary_df = pd.DataFrame(summary_rows)


# 11. Dashboard Summary
# ============================================================

dashboard_df = summary_df[
    [
        "As of Date",
        "Ticker",
        "Company",
        "Lookback",
        "Buy Count",
        "Sell Count",
        "Buy Insider Count",
        "Sell Insider Count",
        "Net Amount",
        "NPR",
        "Signal",
        "內部人交易動向",
        "10b5-1 P/S Count",
    ]
].copy()


def build_dashboard_display(row):

    npr = row["NPR"]

    if pd.isna(npr):
        return "— 無P/S交易"

    arrow = (
        "↑"
        if npr > 0
        else "↓"
        if npr < 0
        else "→"
    )

    return (
        f"{arrow} {row['內部人交易動向']}"
        f"｜{int(row['Buy Count'])}買/"
        f"{int(row['Sell Count'])}賣"
    )


dashboard_df["Form4顯示"] = dashboard_df.apply(
    build_dashboard_display,
    axis=1,
)


# 12. Run Log + Methodology
# ============================================================

run_log_df = pd.DataFrame(run_logs)

methodology_df = pd.DataFrame(
    [
        {
            "項目": "資料來源",
            "定義": "SEC Form 4 via EdgarTools",
        },
        {
            "項目": "分析期間",
            "定義": (
                f"以 Trade Date 為準，最近 {LOOKBACK_DAYS} 日"
            ),
        },
        {
            "項目": "P/S 交易取得方式",
            "定義": (
                "使用 EdgarTools Form4.market_trades DataFrame"
            ),
        },
        {
            "項目": "方向性交易",
            "定義": (
                "P = Purchase；S = Sale；只有 P/S 納入 NPR"
            ),
        },
        {
            "項目": "Excluded Transactions",
            "定義": (
                "使用 Form4.to_dataframe() 保留非 P/S 交易，"
                "例如 A/M/F/G/C 等，不納入 NPR"
            ),
        },
        {
            "項目": "NPR",
            "定義": (
                "(Buy Count - Sell Count) / "
                "(Buy Count + Sell Count)，"
                "Buy/Sell Count 為 P/S 交易筆數"
            ),
        },
        {
            "項目": "Signal",
            "定義": (
                "NPR>0 Net Buying；NPR<0 Net Selling；"
                "NPR=0 Neutral；無P/S則 No P/S Activity"
            ),
        },
        {
            "項目": "Net Amount",
            "定義": (
                "Buy Amount - Sell Amount；"
                "獨立揭露，不取代 NPR"
            ),
        },
        {
            "項目": "10b5-1",
            "定義": (
                "保留 EdgarTools has_10b5_1_plan；"
                "第一版不排除、不降權"
            ),
        },
        {
            "項目": "Form 4/A",
            "定義": (
                "第一版不直接加入方向性交易；"
                "Run Log 顯示 amendment 數"
            ),
        },
        {
            "項目": "Issuer Validation",
            "定義": (
                "Form 4 issuer ticker 與目前分析 Ticker 不同者排除；"
                "避免把公司作為 reporting owner 的他社交易誤認為本公司 insider trade"
            ),
        },
        {
            "項目": "Transaction Dedup",
            "定義": (
                "P/S 與 Excluded 均依 accession、日期、ticker、insider、"
                "code、security、shares、price 等核心欄位移除重複"
            ),
        },
    ]
)


# 13. Excel
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
        sheet_name="Form4 Detail",
        index=False,
    )

    excluded_df.to_excel(
        writer,
        sheet_name="Excluded Transactions",
        index=False,
    )

    run_log_df.to_excel(
        writer,
        sheet_name="Run Log",
        index=False,
    )

    universe_df.to_excel(
        writer,
        sheet_name="Ticker Universe",
        index=False,
    )

    methodology_df.to_excel(
        writer,
        sheet_name="Methodology",
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
        .value_counts(dropna=False)
    )

print(
    f"\nTicker Summary 筆數：{len(summary_df):,}"
)

print(
    f"P/S Detail 筆數：{len(detail_df):,}"
)

print(
    f"Excluded Detail 筆數：{len(excluded_df):,}"
)

print(
    f"P/S 移除重複：{ps_duplicates_removed:,} 筆"
)

print(
    f"Excluded 移除重複：{excluded_duplicates_removed:,} 筆"
)

print(
    f"\n輸出檔案：{FINAL_FILE}"
)

