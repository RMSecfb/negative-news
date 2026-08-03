from __future__ import annotations

import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import unicodedata
import urllib.parse
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from difflib import SequenceMatcher
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
import streamlit as st
import altair as alt
import streamlit.components.v1 as components
from streamlit.runtime.scriptrunner import RerunData, get_script_run_ctx
from streamlit.runtime import Runtime
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def launch_streamlit_when_run_directly() -> None:
    """直接執行本檔案時，確認網站就緒後才開啟瀏覽器。"""
    if get_script_run_ctx(suppress_warning=True) is not None:
        return
    if os.environ.get("FUBON_STANDALONE_STREAMLIT_CHILD") == "1":
        return
    environment = os.environ.copy()
    environment["FUBON_STANDALONE_STREAMLIT_CHILD"] = "1"
    command = [
        sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()),
        "--server.address", "127.0.0.1", "--server.port", "8504",
        "--server.headless", "true", "--browser.gatherUsageStats", "false",
        "--client.toolbarMode", "viewer",
    ]
    script_dir = Path(__file__).resolve().parent
    log_path = script_dir / "啟動錯誤.log"
    url = "http://127.0.0.1:8504/"
    try:
        with log_path.open("w", encoding="utf-8") as log_file:
            process = subprocess.Popen(
                command, env=environment, cwd=str(script_dir),
                stdout=log_file, stderr=subprocess.STDOUT,
            )
            ready = False
            for _ in range(120):
                if process.poll() is not None:
                    break
                try:
                    with urllib.request.urlopen(f"{url}_stcore/health", timeout=1) as response:
                        ready = response.status == 200
                except Exception:
                    ready = False
                if ready:
                    break
                time.sleep(0.5)
            if not ready:
                if process.poll() is None:
                    process.terminate()
                print(f"網站啟動失敗，請查看：{log_path}")
                try:
                    input("按 Enter 關閉視窗……")
                except EOFError:
                    pass
                raise SystemExit(1)
            print(f"網站已啟動：{url}")
            print("請保留此視窗；關閉視窗後網站也會停止。")
            webbrowser.open(url)
            try:
                exit_code = process.wait()
            except KeyboardInterrupt:
                process.terminate()
                exit_code = process.wait(timeout=10)
            if exit_code not in (0, -15):
                print(f"網站已停止，詳細訊息請查看：{log_path}")
                try:
                    input("按 Enter 關閉視窗……")
                except EOFError:
                    pass
            raise SystemExit(exit_code)
    except OSError as exc:
        print(f"無法建立網站服務：{exc}")
        raise SystemExit(1) from exc


launch_streamlit_when_run_directly()


STANDALONE_DIR = Path(__file__).resolve().parent
SIMPLE_SITE = True
APP_TITLE = "美股負面新聞整合中心" if SIMPLE_SITE else "美股新聞風險指揮中心"
APP_VERSION = "雙方法合併精簡版 1.0" if SIMPLE_SITE else "全新重建版 1.0"
TAIPEI = timezone(timedelta(hours=8))
DATA_ROOT = Path(os.environ.get(
    "FUBON_REBUILT_DATA_DIR",
    str(STANDALONE_DIR / "負面新聞整合資料"),
))
OUTPUT_DIR = DATA_ROOT / "output"
UPLOAD_DIR = DATA_ROOT / "uploads"
CONFIG_DIR = DATA_ROOT / "config"
DB_PATH = DATA_ROOT / "system.db"
RULE_PATH = CONFIG_DIR / "Parameter_Event.xlsx"
COMPANY_PATH = CONFIG_DIR / "Company_List.xlsx"
BACKUP_PATH = DATA_ROOT / "backups" / "美股負面新聞系統_版面流程改進前_20260728.py"


def find_external_workbook(exact_name: str, patterns: tuple[str, ...]) -> Path | None:
    """從程式旁、既有 data/config 與常見個人資料夾尋找設定檔。"""
    search_roots = [
        STANDALONE_DIR,
        STANDALONE_DIR / "data" / "config",
        Path.home() / "Downloads",
        Path.home() / "Desktop",
        Path.home() / "Documents",
    ]
    candidates: list[Path] = []
    for root in search_roots:
        exact = root / exact_name
        if exact.is_file():
            candidates.append(exact)
        if root.is_dir():
            for pattern in patterns:
                candidates.extend(path for path in root.glob(pattern) if path.is_file())
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None

NEWS_COLUMNS = ["Ticker", "Company", "Published Time", "Title", "Source", "URL", "FinBERT"]
NEGATIVE_OUTPUT_COLUMNS = [
    "Published Time", "Ticker", "Company", "Title", "Title_ZH", "FinBERT",
    "Event_type", "Event_Code", "事件中文", "Level", "Action", "Keyword", "URL", "Source",
]

# 這份資料只用來「首次建立」可編輯的 Parameter_Event.xlsx。
# 實際分類永遠重新讀取 Excel，使用者修改 Excel 後不需改程式。
DEFAULT_RULE_ROWS = [
    ["E001", "Legal", "BANKRUPTCY", "破產／重整", 5, 100, "公司", "bankruptcy;chapter 11;insolvency;破產;重整", "立即檢視；持續追蹤", 30, "是", "是", "是", "", "公司進入破產、重整或無力償債程序"],
    ["E002", "Regulatory", "FRAUD", "詐欺／財報造假", 5, 98, "公司", "fraud;accounting irregularities;misstatement;財報造假;詐欺", "立即檢視；持續追蹤", 30, "是", "是", "是", "alleged fraud prevention", "重大詐欺或會計問題"],
    ["E003", "Regulatory", "INVESTIGATION", "監管調查", 4, 90, "公司", "investigation;probe;subpoena;監管調查;遭調查", "持續追蹤", 14, "是", "是", "是", "internal investigation", "政府或主管機關調查"],
    ["E004", "Legal", "LAWSUIT", "重大訴訟", 4, 88, "公司", "lawsuit;sued;litigation;class action;訴訟;集體訴訟", "持續追蹤", 14, "否", "是", "否", "settles lawsuit", "可能造成重大財務或營運影響的訴訟"],
    ["E005", "Regulatory", "ANTITRUST", "反壟斷調查", 4, 92, "公司", "antitrust;competition probe;反壟斷;反托拉斯", "持續追蹤", 21, "是", "是", "是", "", "競爭法或反壟斷事件"],
    ["E006", "Credit", "CREDIT_DOWNGRADE", "信用降評", 4, 85, "公司", "downgrade;rating cut;rating lowered;信用調降;信評調降", "持續追蹤", 14, "是", "是", "是", "upgrade", "信用評等遭調降"],
    ["E007", "Operations", "RECALL", "產品召回", 4, 82, "產品", "recall;product defect;產品召回;瑕疵產品", "持續追蹤", 14, "是", "是", "否", "", "產品安全或召回事件"],
    ["E008", "Operations", "CYBER", "資安事件", 4, 90, "公司", "cyberattack;data breach;ransomware;資料外洩;資安事件", "立即檢視；持續追蹤", 21, "是", "是", "是", "", "重大資安或資料外洩"],
    ["E009", "Operations", "SUPPLY", "供應鏈中斷", 3, 75, "營運", "supply disruption;production halt;factory shutdown;供應鏈中斷;停產", "持續追蹤", 10, "否", "是", "否", "", "生產或供應鏈受到重大影響"],
    ["E010", "Finance", "GUIDANCE_CUT", "下修財測", 4, 86, "財務", "cuts guidance;lowers outlook;profit warning;下修財測;獲利預警", "持續追蹤", 14, "是", "是", "是", "raises guidance", "公司下修營收或獲利展望"],
    ["E011", "Finance", "EARNINGS_MISS", "財報不如預期", 3, 70, "財務", "misses estimates;earnings miss;revenue miss;不如預期;低於預期", "檢視", 7, "否", "否", "否", "beats estimates", "財報低於市場預期"],
    ["E012", "Management", "EXEC_EXIT", "高階主管異動", 3, 68, "管理", "ceo resigns;cfo resigns;steps down abruptly;執行長辭職;財務長辭職", "檢視", 10, "否", "否", "否", "planned succession", "關鍵主管突然離任"],
    ["E013", "Workforce", "LAYOFF", "大規模裁員", 3, 72, "營運", "layoffs;job cuts;workforce reduction;裁員;縮減人力", "持續追蹤", 10, "否", "是", "否", "hiring", "大規模人力縮減"],
    ["E014", "Regulatory", "FINE", "監管裁罰", 4, 88, "公司", "fine;penalty;sanction;罰款;裁罰", "持續追蹤", 14, "是", "是", "否", "", "主管機關重大裁罰"],
    ["E015", "Market", "TARIFF", "關稅／制裁衝擊", 3, 74, "外部", "tariff;export ban;sanctions;關稅;出口禁令;制裁", "檢視；持續追蹤", 10, "否", "是", "否", "tariff exemption", "關稅、出口限制或制裁造成風險"],
    ["SYS001", "System", "UNKNOWN", "待人工覆核", 1, 3, "系統", "", "人工覆核", 3, "否", "否", "否", "", "未命中事件規則且情緒非正面的新聞"],
    ["SYS002", "System", "NOT_RELEVANT", "無關新聞", 0, 2, "系統", "", "不處理", 0, "否", "否", "否", "", "未命中事件規則且情緒為正面的新聞"],
    ["SYS003", "System", "NON_CORE_SOURCE", "非主要來源", 0, 1, "系統", "", "不處理", 0, "否", "否", "否", "", "未命中事件規則且不屬於主要新聞來源"],
]

DEFAULT_CORE_SOURCES = [
    "CNBC", "MoneyDJ", "經濟日報", "鉅亨網", "TradingView", "Nasdaq",
    "Reuters", "Associated Press", "AP News", "Bloomberg", "The Wall Street Journal",
    "Financial Times", "MarketWatch", "Yahoo Finance", "Business Insider", "Forbes",
    "Barron's", "The New York Times", "The Washington Post", "BBC", "CNN",
    "Fox Business", "Fortune", "The Guardian", "Seeking Alpha", "Investing.com",
    "The Motley Fool", "TechCrunch", "The Verge",
]

DOW_30 = [
    ("MMM", "3M"), ("AXP", "American Express"), ("AMGN", "Amgen"),
    ("AMZN", "Amazon"), ("AAPL", "Apple"), ("BA", "Boeing"),
    ("CAT", "Caterpillar"), ("CVX", "Chevron"), ("CSCO", "Cisco"),
    ("KO", "Coca-Cola"), ("DIS", "Disney"), ("GS", "Goldman Sachs"),
    ("HD", "Home Depot"), ("HON", "Honeywell"), ("IBM", "IBM"),
    ("JNJ", "Johnson & Johnson"), ("JPM", "JPMorgan Chase"),
    ("MCD", "McDonald's"), ("MRK", "Merck"), ("MSFT", "Microsoft"),
    ("NKE", "Nike"), ("NVDA", "NVIDIA"), ("PG", "Procter & Gamble"),
    ("CRM", "Salesforce"), ("SHW", "Sherwin-Williams"),
    ("TRV", "Travelers"), ("UNH", "UnitedHealth"), ("VZ", "Verizon"),
    ("V", "Visa"), ("WMT", "Walmart"),
]


@st.cache_data(ttl=86400)
def load_sp500() -> pd.DataFrame:
    """每次快取到期後由公開成分表重新取得，不依賴舊專案名單。"""
    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    response = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    table = pd.read_html(io.StringIO(response.text), match="Symbol")[0]
    result = table[["Symbol", "Security"]].rename(columns={"Symbol": "Ticker", "Security": "Company"})
    result["Ticker"] = result["Ticker"].astype(str).str.replace(".", "-", regex=False).str.upper()
    result["Company"] = result["Company"].astype(str).str.strip()
    return result.drop_duplicates("Ticker").reset_index(drop=True)


def initialize() -> None:
    for folder in (DATA_ROOT, OUTPUT_DIR, UPLOAD_DIR, CONFIG_DIR):
        folder.mkdir(parents=True, exist_ok=True)
    rule_is_incomplete = not RULE_PATH.exists()
    if RULE_PATH.exists():
        try:
            workbook = load_workbook(RULE_PATH, read_only=True, data_only=True)
            rule_is_incomplete = "主要網站" not in workbook.sheetnames
            workbook.close()
        except Exception:
            rule_is_incomplete = True
    if rule_is_incomplete:
        external_rule = find_external_workbook("Parameter_Event.xlsx", ("Parameter_Event*.xlsx",))
        if external_rule and external_rule.resolve() != RULE_PATH.resolve():
            shutil.copy2(external_rule, RULE_PATH)
    if not COMPANY_PATH.exists():
        external_company = find_external_workbook(
            "Company_List.xlsx",
            ("*Company*List*.xlsx", "*公司*名單*.xlsx", "*成分股*.xlsx"),
        )
        if external_company:
            shutil.copy2(external_company, COMPANY_PATH)
    with sqlite3.connect(DB_PATH) as connection:
        connection.executescript("""
            CREATE TABLE IF NOT EXISTS runs (
                id INTEGER PRIMARY KEY, started_at TEXT, finished_at TEXT, status TEXT,
                start_time TEXT, end_time TEXT, universe TEXT, rows INTEGER, output_path TEXT, message TEXT
            );
            CREATE TABLE IF NOT EXISTS manual_status (
                company_key TEXT PRIMARY KEY, ticker TEXT, company TEXT, status TEXT,
                owner TEXT, next_review TEXT, note TEXT, updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
        """)
    if not RULE_PATH.exists():
        create_rule_workbook(RULE_PATH)


def last_success_end() -> datetime | None:
    with sqlite3.connect(DB_PATH) as connection:
        row = connection.execute(
            "SELECT end_time FROM runs WHERE status='success' ORDER BY id DESC LIMIT 1"
        ).fetchone()
    if not row or not row[0]:
        return None
    try:
        return datetime.fromisoformat(row[0]).astimezone(TAIPEI)
    except ValueError:
        return None


def latest_crawl_result(today_only: bool = False) -> dict | None:
    """從正式執行紀錄找回結果，不依賴目前瀏覽器的 session。"""
    with sqlite3.connect(DB_PATH) as connection:
        rows = connection.execute(
            """SELECT started_at,finished_at,start_time,end_time,universe,rows,output_path,message
               FROM runs WHERE status='success' AND (universe LIKE '方法一｜%' OR universe LIKE '方法二｜%' OR universe LIKE '方法一＋方法二｜%')
               ORDER BY id DESC"""
        ).fetchall()
    today = datetime.now(TAIPEI).date()
    for row in rows:
        try:
            end_at = datetime.fromisoformat(row[3]).astimezone(TAIPEI)
        except (TypeError, ValueError):
            continue
        if today_only and end_at.date() != today:
            continue
        output_path = Path(str(row[6]))
        if not output_path.is_file():
            continue
        method = str(row[4]).split("｜", 1)[0]
        stamp = end_at.strftime("%Y%m%d")
        event_candidate = OUTPUT_DIR / f"美股_{stamp}_負面新聞爬蟲.xlsx"
        event_path = event_candidate if event_candidate.is_file() else None
        finbert_path = None
        if method in ("方法二", "方法一＋方法二"):
            finbert_name = output_path.name.replace("MergeNews_", "MergeNews_FinBERT_").replace("Combined_News_", "Combined_FinBERT_News_").replace("All_News_", "FinBERT_News_")
            candidate = output_path.with_name(finbert_name)
            finbert_path = candidate if candidate.is_file() else None
        return {
            "started_at": row[0], "finished_at": row[1], "start_time": row[2], "end_time": row[3],
            "universe": row[4], "rows": int(row[5] or 0), "path": output_path,
            "event_path": event_path, "finbert_path": finbert_path, "message": row[7] or "", "method": method,
            "is_today": end_at.date() == today,
        }
    return None


def create_run(start: datetime, end: datetime, universe: str) -> int:
    with sqlite3.connect(DB_PATH) as connection:
        cursor = connection.execute(
            "INSERT INTO runs(started_at,status,start_time,end_time,universe,rows,message) VALUES(?,?,?,?,?,?,?)",
            (datetime.now(TAIPEI).isoformat(timespec="seconds"), "running", start.isoformat(), end.isoformat(), universe, 0, ""),
        )
        return int(cursor.lastrowid)


def finish_run(run_id: int, status: str, rows: int = 0, output_path: str = "", message: str = "") -> None:
    with sqlite3.connect(DB_PATH) as connection:
        connection.execute(
            "UPDATE runs SET finished_at=?,status=?,rows=?,output_path=?,message=? WHERE id=?",
            (datetime.now(TAIPEI).isoformat(timespec="seconds"), status, rows, output_path, message, run_id),
        )


class CrawlCancelled(Exception):
    """使用者主動停止抓取。"""


class BackgroundProgress:
    def __init__(self, job: dict):
        self.job = job

    def progress(self, value=0, text=""):
        if self.job["stop_event"].is_set():
            raise CrawlCancelled("使用者已停止抓取")
        self.job["progress"] = max(0.0, min(float(value or 0), 1.0))
        if text:
            self.job["progress_text"] = str(text)
        return self


@st.cache_resource
def crawl_job_registry() -> dict:
    return {}


def format_elapsed(seconds: float) -> str:
    total = max(0, int(seconds))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def begin_job_stage(job: dict, label: str) -> None:
    """切換主要執行階段，並保留前一階段的花費時間。"""
    now = time.monotonic()
    previous = job.get("current_stage")
    previous_started = job.get("current_stage_started")
    if previous and previous_started is not None:
        job.setdefault("stage_times", []).append({
            "label": previous,
            "seconds": max(0.0, now - float(previous_started)),
        })
    job["current_stage"] = label
    job["current_stage_started"] = now


def finish_job_stage(job: dict) -> None:
    now = time.monotonic()
    current = job.pop("current_stage", None)
    started = job.pop("current_stage_started", None)
    if current and started is not None:
        job.setdefault("stage_times", []).append({
            "label": current,
            "seconds": max(0.0, now - float(started)),
        })


def run_crawl_job(job: dict, method_name: str, universe: str, companies: pd.DataFrame,
                  start_dt: datetime, end_dt: datetime) -> None:
    run_id = job["run_id"]
    progress = BackgroundProgress(job)
    active_method_key = ""
    try:
        stamp = end_dt.strftime("%Y%m%d")
        if method_name == "方法一":
            begin_job_stage(job, "方法一｜多來源新聞擷取")
            active_method_key = "method1_monitor"
            job[active_method_key].update(status="running", stage="建立多來源新聞池")
            progress.progress(0, text="方法一正在建立新聞池")
            frame, all_errors, source_counts = fetch_method1_news(companies, start_dt, end_dt, progress)
            job[active_method_key].update(
                status="success", stage="完成", rows=len(frame), errors=list(all_errors),
                source_counts=dict(source_counts),
            )
            path = OUTPUT_DIR / f"RawNews_{stamp}.xlsx"
            begin_job_stage(job, "產生今日新聞檔案")
            source_summary = pd.DataFrame([{"來源": source, "新聞筆數": count} for source, count in source_counts.items()])
            progress.progress(0.94, text="正在產生新聞檔案")
            write_excel(path, {"新聞": frame, "來源摘要": source_summary})
            job["summary"] = f"方法一已完成：取得 {len(frame)} 則新聞；來源異常 {len(all_errors)} 項"
        elif method_name == "方法二":
            begin_job_stage(job, "方法二｜Google News 擷取")
            active_method_key = "method2_monitor"
            job[active_method_key].update(status="running", stage="分批抓取 Google News")
            progress.progress(0, text="方法二正在分批爬取 Google News")
            raw_frame, all_errors, batch_count = fetch_method2_raw(companies, start_dt, end_dt, progress)
            job[active_method_key].update(stage="FinBERT 評分中", raw_rows=len(raw_frame), batches=batch_count, errors=list(all_errors))
            begin_job_stage(job, "FinBERT 新聞評分")
            frame, negative_frame = score_method2_finbert(raw_frame, progress)
            job[active_method_key].update(
                status="success", stage="完成", rows=len(frame), negative_rows=len(negative_frame),
                batches=batch_count, errors=list(all_errors),
            )
            path = OUTPUT_DIR / f"All_News_{stamp}.xlsx"
            finbert_path = OUTPUT_DIR / f"FinBERT_News_{stamp}.xlsx"
            begin_job_stage(job, "產生今日新聞檔案")
            progress.progress(0.94, text="正在產生新聞檔案")
            write_excel(path, {"All_News": frame})
            write_excel(finbert_path, {"FinBERT_News": negative_frame})
            job["summary"] = f"方法二已完成：{batch_count} 批、爬得 {len(frame)} 則；FinBERT ≤ 0 共 {len(negative_frame)} 則"
        else:
            begin_job_stage(job, "方法一｜多來源新聞擷取")
            active_method_key = "method1_monitor"
            job[active_method_key].update(status="running", stage="建立多來源新聞池")
            progress.progress(0, text="完整整合 1/4｜執行方法一")
            method1_frame, method1_errors, source_counts = fetch_method1_news(companies, start_dt, end_dt, progress)
            job[active_method_key].update(
                status="success", stage="完成", rows=len(method1_frame), errors=list(method1_errors),
                source_counts=dict(source_counts),
            )
            active_method_key = "method2_monitor"
            begin_job_stage(job, "方法二｜Google News 擷取")
            job[active_method_key].update(status="running", stage="分批抓取 Google News")
            progress.progress(0, text="完整整合 2/4｜執行方法二 Google News")
            method2_frame, method2_errors, batch_count = fetch_method2_raw(companies, start_dt, end_dt, progress)
            job[active_method_key].update(
                stage="FinBERT 評分中", raw_rows=len(method2_frame), batches=batch_count,
                errors=list(method2_errors),
            )
            begin_job_stage(job, "合併並刪除重複")
            progress.progress(0.55, text="完整整合 3/4｜合併並去除重複新聞")
            merged_frame, duplicate_log = merge_frames(method1_frame, method2_frame)
            begin_job_stage(job, "FinBERT 新聞評分")
            progress.progress(0.58, text="完整整合 4/4｜對全部整合新聞執行 FinBERT")
            frame, negative_frame = score_method2_finbert(merged_frame, progress)
            job[active_method_key].update(
                status="success", stage="完成", rows=len(method2_frame), merged_rows=len(frame),
                negative_rows=len(negative_frame), batches=batch_count, errors=list(method2_errors),
            )
            all_errors = method1_errors + method2_errors
            path = OUTPUT_DIR / f"MergeNews_{stamp}.xlsx"
            finbert_path = OUTPUT_DIR / f"MergeNews_FinBERT_{stamp}.xlsx"
            source_summary = pd.DataFrame(
                [{"來源": source, "新聞筆數": count} for source, count in source_counts.items()]
                + [{"來源": "Google News（方法二）", "新聞筆數": len(method2_frame)}]
            )
            begin_job_stage(job, "產生完整整合檔案")
            progress.progress(0.94, text="正在產生完整整合檔案")
            write_excel(path, {"完整整合新聞": frame, "重複紀錄": duplicate_log, "來源摘要": source_summary})
            write_excel(finbert_path, {"FinBERT小於等於0": negative_frame})
            job["summary"] = f"完整整合已完成：去重後保留 {len(frame)} 則；FinBERT ≤ 0 共 {len(negative_frame)} 則"
        begin_job_stage(job, "負面事件分類與中文翻譯")
        progress.progress(0.97, text="正在執行負面事件分類")
        rule_modified_ns = RULE_PATH.stat().st_mtime_ns
        event_frame, unknown_frame, irrelevant_frame = classify_news_sets(
            frame, load_rules(rule_modified_ns), load_core_sources(rule_modified_ns), progress
        )
        negative_translation_failures = int(event_frame.attrs.get("translation_failures", 0))
        unknown_translation_failures = int(unknown_frame.attrs.get("translation_failures", 0))
        irrelevant_translation_failures = int(irrelevant_frame.attrs.get("translation_failures", 0))
        translation_failures = negative_translation_failures + unknown_translation_failures + irrelevant_translation_failures
        event_path = OUTPUT_DIR / f"美股_{stamp}_負面新聞爬蟲.xlsx"
        begin_job_stage(job, "產生負面新聞檔案")
        write_excel(event_path, {
            "負面新聞": event_frame,
            "待人工覆核": unknown_frame,
            "無關新聞": irrelevant_frame,
        })
        finish_run(run_id, "success", len(frame), str(path), "；".join(all_errors))
        job.update(status="success", progress=1.0, progress_text="抓取與分類全部完成",
                   rows=len(frame), event_rows=len(event_frame), unknown_rows=len(unknown_frame),
                   irrelevant_rows=len(irrelevant_frame), translation_failures=translation_failures,
                   negative_translation_failures=negative_translation_failures,
                   unknown_translation_failures=unknown_translation_failures,
                   irrelevant_translation_failures=irrelevant_translation_failures,
                   path=str(path))
    except CrawlCancelled:
        if active_method_key and job.get(active_method_key, {}).get("status") == "running":
            job[active_method_key].update(status="stopped", stage="使用者停止")
        finish_run(run_id, "stopped", message="使用者主動停止")
        job.update(status="stopped", progress_text="已停止；未完成的檔案不會列為正式結果")
    except Exception as exc:
        if active_method_key:
            job[active_method_key].update(status="failed", stage="執行失敗", fatal_error=str(exc))
        finish_run(run_id, "failed", message=str(exc))
        job.update(status="failed", error=str(exc), progress_text="執行失敗")
    finally:
        finish_job_stage(job)
        job["finished_monotonic"] = time.monotonic()
        job["finished_at"] = datetime.now(TAIPEI).isoformat(timespec="seconds")
        # 背景工作只在結束瞬間通知原頁面完整更新一次；執行期間不輪詢、不刷新。
        script_context = job.get("script_context")
        try:
            runtime = Runtime.instance()
            session_info = runtime._session_mgr.get_session_info(job.get("session_id", ""))
            if session_info is not None:
                session = session_info.session
                session._event_loop.call_soon_threadsafe(session.request_rerun, None)
            elif script_context is not None and script_context.script_requests is not None:
                script_context.script_requests.request_rerun(RerunData())
        except Exception:
            pass


def create_rule_workbook(path: Path) -> None:
    headers = [
        "Event_ID", "Event_type", "Event_Code", "事件中文", "Level", "Primary_Priority",
        "Event_Role", "Keyword", "Action", "Review_Days", "Notify", "Watchlist",
        "Position_Review", "Exclude_Keyword", "Definition",
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Event"
    sheet.append(headers)
    for row in DEFAULT_RULE_ROWS:
        sheet.append(row)
    action = workbook.create_sheet("Action")
    action.append(["狀態", "用途"])
    for row in [("待確認", "尚未人工判讀"), ("追蹤中", "持續觀察"), ("已通知", "已通知相關人員"), ("已完成", "處理完成"), ("解除追蹤", "不再列入待辦")]:
        action.append(row)
    sources = workbook.create_sheet("主要網站")
    sources.append(["Source", "說明"])
    for source in DEFAULT_CORE_SOURCES:
        sources.append([source, "內建主要新聞來源"])
    for current in workbook.worksheets:
        current.freeze_panes = "A2"
        current.auto_filter.ref = current.dimensions
        for cell in current[1]:
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.font = Font(color="FFFFFF", bold=True)
        for column in range(1, current.max_column + 1):
            current.column_dimensions[get_column_letter(column)].width = 22
    workbook.save(path)
    workbook.close()


def split_terms(value: object) -> list[str]:
    return [item.strip() for item in re.split(r"[;；\n]+", str(value or "")) if item.strip()]


@st.cache_data(ttl=60)
def load_rules(modified_ns: int) -> list[dict]:
    del modified_ns
    workbook = load_workbook(RULE_PATH, read_only=True, data_only=True)
    sheet = workbook["Event"]
    headers = [str(cell.value or "").strip() for cell in sheet[1]]
    rules = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        item = dict(zip(headers, values))
        if not str(item.get("Event_Code") or "").strip():
            continue
        item["_keywords"] = split_terms(item.get("Keyword"))
        item["_exclusions"] = split_terms(item.get("Exclude_Keyword"))
        rules.append(item)
    workbook.close()
    return rules


@st.cache_data(ttl=60)
def load_core_sources(modified_ns: int) -> set[str]:
    del modified_ns
    workbook = load_workbook(RULE_PATH, read_only=True, data_only=True)
    if "主要網站" not in workbook.sheetnames:
        workbook.close()
        raise ValueError("目前關鍵字表缺少「主要網站」工作表")
    sheet = workbook["主要網站"]
    headers = [str(cell.value or "").strip() for cell in sheet[1]]
    if "Source" not in headers:
        workbook.close()
        raise ValueError("「主要網站」工作表缺少 Source 欄位")
    source_index = headers.index("Source")
    sources = {
        str(values[source_index]).strip().casefold()
        for values in sheet.iter_rows(min_row=2, values_only=True)
        if source_index < len(values) and values[source_index] not in (None, "")
    }
    workbook.close()
    return sources


def normalize_news(frame: pd.DataFrame) -> pd.DataFrame:
    # 外部工具常在欄名尾端留下不可見空白；先統一清理再做欄位對應。
    frame = frame.copy()
    frame.columns = [str(column).replace("\u3000", " ").strip() for column in frame.columns]
    aliases = {
        "Ticker": ("Ticker", "股票代號", "標的代碼"),
        "Company": ("Company", "股票名稱", "標的名稱"),
        "Published Time": ("Published Time", "發布時間", "發布日期"),
        "Title": ("Title", "新聞主旨", "主旨"),
        "Source": ("Source", "新聞來源"),
        "URL": ("URL", "新聞網址", "新聞連結"),
        "FinBERT": ("FinBERT",),
    }
    result = pd.DataFrame(index=frame.index)
    for target, names in aliases.items():
        source = next((name for name in names if name in frame.columns), None)
        result[target] = frame[source] if source else ""
    for column in NEWS_COLUMNS:
        result[column] = result[column].fillna("").astype(str).str.strip()
    result["Ticker"] = result["Ticker"].str.replace(r"\.0$", "", regex=True).str.upper()
    result = result[(result["Title"] != "") & (result["URL"] != "")]
    return result.reset_index(drop=True)


TRACKING_QUERY_NAMES = {
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "from", "source", "ref", "gclid", "fbclid", "mc_cid", "mc_eid",
}


def normalize_url(value: str) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        parts = urllib.parse.urlsplit(raw)
        if not parts.scheme or not parts.netloc:
            return raw.rstrip("/").casefold()
        query = [
            (key, val) for key, val in urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
            if key.casefold() not in TRACKING_QUERY_NAMES and not key.casefold().startswith("utm_")
        ]
        return urllib.parse.urlunsplit((
            parts.scheme.casefold(), parts.netloc.casefold(), parts.path.rstrip("/"),
            urllib.parse.urlencode(query), "",
        ))
    except ValueError:
        return raw.rstrip("/").casefold()


def url_signatures(value: str) -> set[str]:
    """EXE 同款：解開 Google 轉址、URL encoding 與常見內嵌目標網址。"""
    raw = str(value or "").strip()
    if not raw:
        return set()
    decoded_versions = [raw]
    for _ in range(3):
        decoded = urllib.parse.unquote(decoded_versions[-1])
        if decoded == decoded_versions[-1]:
            break
        decoded_versions.append(decoded)
    candidates = set(decoded_versions)
    for text in decoded_versions:
        candidates.update(re.findall(r"https?://[^\s<>'\"]+", text, flags=re.I))
        try:
            for key, target in urllib.parse.parse_qsl(urllib.parse.urlsplit(text).query, keep_blank_values=False):
                if key.casefold() in {"q", "url", "target", "dest", "destination", "continue"}:
                    candidates.add(urllib.parse.unquote(target))
        except ValueError:
            pass
    return {signature for candidate in candidates if (signature := normalize_url(candidate))}


def exact_title(value: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value or ""))).strip().casefold()


def title_for_similarity(value: str) -> str:
    text = exact_title(value)
    text = re.sub(r"\b(?:19|20)\d{2}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b", " ", text)
    text = re.sub(r"\b(?:mon|tue|wed|thu|fri|sat|sun),?\s+\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(?:19|20)\d{2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?(?:\s+gmt)?\b", " ", text, flags=re.I)
    text = re.sub(r"\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2},?\s+(?:19|20)\d{2}\b", " ", text, flags=re.I)
    text = re.sub(r"\s+(?:by|via)\s+[\w.&'’-]+(?:\s+[\w.&'’-]+){0,5}\s*$", "", text, flags=re.I)
    text = re.sub(r"\s+(?:[-–—|:]\s*)?(?:www\.)?[a-z0-9][a-z0-9.-]*\.(?:com|net|org|co|io|tw|ca|uk|news)(?:\.[a-z]{2})?\s*$", "", text, flags=re.I)
    return re.sub(r"[^0-9a-z\u3400-\u9fff]+", " ", text).strip()


def normalized_title(value: str) -> str:
    """給單一方法的快速去重使用，與 EXE 的標題清理一致。"""
    return title_for_similarity(value).replace(" ", "")


def numeric_facts(value: str) -> tuple[str, ...]:
    return tuple(sorted(re.findall(r"\d+(?:\.\d+)?%?", title_for_similarity(value))))


def title_similarity(left: str, right: str) -> float:
    left_key, right_key = title_for_similarity(left), title_for_similarity(right)
    if min(len(left_key), len(right_key)) < 12 or numeric_facts(left) != numeric_facts(right):
        return 0.0
    compact_score = SequenceMatcher(None, left_key.replace(" ", ""), right_key.replace(" ", "")).ratio()
    token_score = SequenceMatcher(None, " ".join(sorted(left_key.split())), " ".join(sorted(right_key.split()))).ratio()
    return max(compact_score, token_score)


def merge_company_key(record: dict) -> tuple[str, str] | None:
    ticker = re.sub(r"[\s_\-]+", "", unicodedata.normalize("NFKC", str(record.get("Ticker", ""))).strip().casefold())
    company = re.sub(r"[\s_\-]+", "", unicodedata.normalize("NFKC", str(record.get("Company", ""))).strip().casefold())
    return (ticker, company) if ticker and company else None


def merge_frames(first: pd.DataFrame, second: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """EXE 同款去重；但保留資料較完整的一筆，並固定輸出網站七欄。"""
    input_rows = (
        first.assign(_file="檔案一").to_dict("records")
        + second.assign(_file="檔案二").to_dict("records")
    )
    kept: list[dict] = []
    duplicates: list[dict] = []
    seen_urls: dict[tuple[tuple[str, str], str], int] = {}
    seen_exact_titles: dict[tuple[tuple[str, str], str], int] = {}
    retained_by_company: dict[tuple[str, str], list[int]] = {}

    def remember(record: dict, index: int) -> None:
        company_key = merge_company_key(record)
        if company_key is None:
            return
        for signature in url_signatures(record.get("URL", "")):
            seen_urls.setdefault((company_key, signature), index)
        title_key = exact_title(record.get("Title", ""))
        if title_key:
            seen_exact_titles.setdefault((company_key, title_key), index)
        indexes = retained_by_company.setdefault(company_key, [])
        if index not in indexes:
            indexes.append(index)

    for record in input_rows:
        company_key = merge_company_key(record)
        match_index: int | None = None
        reason = ""
        if company_key is not None:
            for signature in url_signatures(record.get("URL", "")):
                match_index = seen_urls.get((company_key, signature))
                if match_index is not None:
                    reason = "AB相同且網址相同（已忽略Google轉址前綴與追蹤參數）"
                    break
            if match_index is None:
                title_key = exact_title(record.get("Title", ""))
                if title_key:
                    match_index = seen_exact_titles.get((company_key, title_key))
                    if match_index is not None:
                        reason = "AB相同且主旨完全相同"
            if match_index is None:
                best_score, best_index = 0.0, None
                for candidate_index in retained_by_company.get(company_key, []):
                    score = title_similarity(kept[candidate_index].get("Title", ""), record.get("Title", ""))
                    if score > best_score:
                        best_score, best_index = score, candidate_index
                if best_index is not None and best_score >= 0.98:
                    match_index = best_index
                    reason = f"AB相同且主旨高度相似（{best_score:.0%}，數字資訊一致）"

        if match_index is None:
            kept.append(record)
            remember(record, len(kept) - 1)
            continue

        old = kept[match_index]
        if sum(bool(str(record.get(column, "")).strip()) for column in NEWS_COLUMNS) > sum(bool(str(old.get(column, "")).strip()) for column in NEWS_COLUMNS):
            kept[match_index] = record
        remember(record, match_index)
        duplicates.append({
            "原因": reason, "Ticker": record.get("Ticker", ""), "Company": record.get("Company", ""),
            "保留標題": kept[match_index].get("Title", ""), "重複標題": record.get("Title", ""),
            "保留網址": kept[match_index].get("URL", ""), "刪除網址": record.get("URL", ""),
            "來源檔": record.get("_file", ""),
        })

    result = pd.DataFrame(kept).drop(columns=["_file"], errors="ignore")
    if result.empty:
        result = pd.DataFrame(columns=NEWS_COLUMNS)
    return result.reindex(columns=NEWS_COLUMNS), pd.DataFrame(duplicates)


def translate_text_zh(text: str) -> str:
    text = str(text or "").strip()
    if not text:
        return ""
    for attempt in range(3):
        try:
            response = requests.get(
                "https://translate.googleapis.com/translate_a/single",
                params={"client": "gtx", "sl": "auto", "tl": "zh-TW", "dt": "t", "q": text},
                headers={"User-Agent": "Mozilla/5.0"}, timeout=15,
            )
            response.raise_for_status()
            payload = response.json()
            translated = "".join(str(part[0]) for part in payload[0] if part and part[0]).strip()
            if translated:
                return translated
        except Exception:
            pass
        if attempt < 2:
            time.sleep(1.5 * (attempt + 1))
    return ""


def translate_title_zh(title: str) -> str:
    return translate_text_zh(title)


def translate_output_rows(rows: list[dict], label: str, progress=None,
                          progress_start: float = 0.97, progress_end: float = 0.995) -> int:
    """批次翻譯輸出列；批次無法正確分割時再逐篇重試。"""
    if not rows:
        return 0
    separator = "<<<FUBON_NEWS_SPLIT>>>"
    batch_size = 15
    failures = 0
    batches = [rows[start:start + batch_size] for start in range(0, len(rows), batch_size)]

    def translate_batch(batch: list[dict]) -> list[str]:
        titles = [str(row.get("Title", "") or "").strip() for row in batch]
        joined = f"\n{separator}\n".join(titles)
        translated = translate_text_zh(joined)
        parts = [part.strip() for part in translated.split(separator)] if translated else []
        if len(parts) != len(batch):
            parts = [translate_title_zh(title) for title in titles]
        return parts

    done = 0
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(translate_batch, batch): batch for batch in batches}
        for future in as_completed(futures):
            batch = futures[future]
            try:
                parts = future.result()
            except Exception:
                parts = [""] * len(batch)
            for row, title_zh in zip(batch, parts):
                row["Title_ZH"] = title_zh
                if not title_zh:
                    failures += 1
            done += len(batch)
            if progress is not None:
                value = progress_start + (progress_end - progress_start) * done / len(rows)
                progress.progress(value, text=f"翻譯{label}標題｜{done}/{len(rows)}")
    return failures


def event_output(record: dict, rule: dict, keyword: str = "", title_zh: str = "") -> dict:
    return {
        "Published Time": record.get("Published Time", ""),
        "Ticker": record.get("Ticker", ""),
        "Company": record.get("Company", ""),
        "Title": record.get("Title", ""),
        "Title_ZH": title_zh,
        "FinBERT": record.get("FinBERT", ""),
        "Event_type": rule.get("Event_type", ""),
        "Event_Code": rule.get("Event_Code", ""),
        "事件中文": rule.get("事件中文", ""),
        "Level": rule.get("Level", ""),
        "Action": rule.get("Action", ""),
        "Keyword": keyword,
        "URL": record.get("URL", ""),
        "Source": record.get("Source", ""),
    }


def classify_news_sets(frame: pd.DataFrame, rules: list[dict], core_sources: set[str], progress=None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    negative_rows: list[dict] = []
    unknown_rows, irrelevant_rows = [], []
    unknown_rule = next((rule for rule in rules if str(rule.get("Event_Code")) == "UNKNOWN"), None)
    irrelevant_rule = next((rule for rule in rules if str(rule.get("Event_Code")) == "NOT_RELEVANT"), None)
    non_core_rule = next((rule for rule in rules if str(rule.get("Event_Code")) == "NON_CORE_SOURCE"), None)
    if unknown_rule is None or irrelevant_rule is None or non_core_rule is None:
        raise ValueError("目前關鍵字表必須包含 UNKNOWN、NOT_RELEVANT 與 NON_CORE_SOURCE 事件")
    for record in frame.to_dict("records"):
        title = record["Title"].lower()
        matches = []
        for rule in rules:
            if not rule["_keywords"]:
                continue
            if any(term.lower() in title for term in rule["_exclusions"]):
                continue
            hits = [term for term in rule["_keywords"] if term.lower() in title]
            if hits:
                matches.append((rule, hits))
        if not matches:
            source = str(record.get("Source") or "").strip().casefold()
            if source not in core_sources:
                irrelevant_rows.append(event_output(record, non_core_rule))
            else:
                try:
                    finbert = float(record.get("FinBERT"))
                except (TypeError, ValueError):
                    finbert = None
                fallback = irrelevant_rule if finbert is not None and finbert > 0 else unknown_rule
                target = irrelevant_rows if fallback is irrelevant_rule else unknown_rows
                target.append(event_output(record, fallback))
            continue
        # 所有判定值皆取自目前 Parameter_Event.xlsx；不再另建新聞分數。
        matches.sort(key=lambda pair: (-int(pair[0].get("Primary_Priority") or 0), str(pair[0].get("Event_ID") or "")))
        primary, hits = matches[0]
        output = event_output(record, primary, "；".join(hits))
        code = str(primary.get("Event_Code") or "")
        event_type = str(primary.get("Event_type") or "").upper()
        if code == "UNKNOWN":
            unknown_rows.append(output)
        elif event_type == "SYSTEM":
            irrelevant_rows.append(output)
        else:
            negative_rows.append(output)
    negative_translation_failures = translate_output_rows(negative_rows, "負面新聞", progress, 0.970, 0.980)
    unknown_translation_failures = translate_output_rows(unknown_rows, "待人工覆核", progress, 0.980, 0.990)
    irrelevant_translation_failures = translate_output_rows(irrelevant_rows, "無關新聞", progress, 0.990, 0.998)

    def make_frame(rows: list[dict]) -> pd.DataFrame:
        result = pd.DataFrame(rows, columns=NEGATIVE_OUTPUT_COLUMNS)
        if not result.empty:
            result["Level"] = pd.to_numeric(result["Level"], errors="coerce")
            result = result.sort_values(["Level", "Published Time"], ascending=[False, False]).reset_index(drop=True)
        return result

    negative_frame = make_frame(negative_rows)
    unknown_frame = make_frame(unknown_rows)
    irrelevant_frame = make_frame(irrelevant_rows)
    negative_frame.attrs["translation_failures"] = negative_translation_failures
    unknown_frame.attrs["translation_failures"] = unknown_translation_failures
    irrelevant_frame.attrs["translation_failures"] = irrelevant_translation_failures
    return negative_frame, unknown_frame, irrelevant_frame


def classify(frame: pd.DataFrame, rules: list[dict]) -> pd.DataFrame:
    """相容既有呼叫：只回傳正式負面新聞。"""
    modified_ns = RULE_PATH.stat().st_mtime_ns
    negative, _, _ = classify_news_sets(frame, rules, load_core_sources(modified_ns))
    return negative


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
            sheet = writer.sheets[name[:31]]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            is_negative_output = list(frame.columns) == NEGATIVE_OUTPUT_COLUMNS
            for index, cell in enumerate(sheet[1], 1):
                if is_negative_output:
                    cell.fill = PatternFill("solid", fgColor="FFF200" if 6 <= index <= 12 else "E2F0D9")
                    cell.font = Font(color="000000", bold=True)
                else:
                    cell.fill = PatternFill("solid", fgColor="1F4E78")
                    cell.font = Font(color="FFFFFF", bold=True)
            negative_widths = {
                "Published Time": 21, "Ticker": 12, "Company": 24, "Title": 60,
                "Title_ZH": 45, "FinBERT": 12, "Event_type": 18, "Event_Code": 27,
                "事件中文": 22, "Level": 10, "Action": 45, "Keyword": 30, "URL": 65, "Source": 18,
            }
            for index, column in enumerate(frame.columns, 1):
                width = negative_widths.get(column, 22) if is_negative_output else (80 if column in {"URL", "Definition"} else 60 if column == "Title" else 22)
                sheet.column_dimensions[get_column_letter(index)].width = width
            sheet.row_dimensions[1].height = 24
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)


METHOD1_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,zh-TW;q=0.8",
}
METHOD1_POOL_LIMIT = 300
METHOD1_EXCHANGES = {"AAPL", "AMZN", "CSCO", "GOOGL", "MSFT", "NVDA", "AMGN", "HON"}
METHOD1_THREAD_LOCAL = threading.local()


def method1_http_session() -> requests.Session:
    session = getattr(METHOD1_THREAD_LOCAL, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update(METHOD1_HEADERS)
        METHOD1_THREAD_LOCAL.session = session
    return session


def parse_method1_time(value: object, reference: datetime | None = None) -> datetime | None:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return None
    now = reference or datetime.now(TAIPEI)
    relative = re.search(r"(\d+)\s+(minute|minutes|hour|hours|day|days)\s+ago", text, re.I)
    if relative:
        amount = int(relative.group(1))
        unit = relative.group(2).lower()
        delta = timedelta(minutes=amount) if unit.startswith("minute") else timedelta(hours=amount) if unit.startswith("hour") else timedelta(days=amount)
        return now - delta
    try:
        return parsedate_to_datetime(text).astimezone(TAIPEI)
    except Exception:
        pass
    cleaned = re.sub(r"\b(?:ET|EST|EDT|CT|CST|CDT)\b", "", text, flags=re.I).strip()
    parsed = pd.to_datetime(cleaned, errors="coerce")
    if pd.isna(parsed):
        return None
    result = parsed.to_pydatetime()
    if result.tzinfo is None:
        result = result.replace(tzinfo=TAIPEI)
    return result.astimezone(TAIPEI)


def method1_in_window(published: datetime | None, start: datetime, end: datetime) -> bool:
    return published is not None and start <= published.astimezone(TAIPEI) <= end


def method1_company_terms(ticker: str, company: str) -> list[str]:
    name = re.sub(r"\s+", " ", company).strip()
    simplified = re.sub(
        r"\s+(?:incorporated|inc\.?|corporation|corp\.?|company|co\.?|limited|ltd\.?|plc|group|holdings?)$",
        "", name, flags=re.I,
    ).strip()
    terms = [name]
    if simplified and simplified.lower() != name.lower() and len(simplified) >= 4:
        terms.append(simplified)
    if ticker == "JNJ":
        terms.extend(["Johnson & Johnson", "J&J", "嬌生"])
    return list(dict.fromkeys(term for term in terms if len(term) >= 4))


def method1_match_company(title: str, summary: str, companies: pd.DataFrame) -> tuple[str, str] | None:
    text = re.sub(r"\s+", " ", f"{title} {summary}").strip()
    candidates: list[tuple[int, str, str]] = []
    for row in companies.itertuples(index=False):
        ticker, company = str(row.Ticker).upper(), str(row.Company)
        structured = re.search(rf"(?:\${re.escape(ticker)}\b|\({re.escape(ticker)}\)|(?:NASDAQ|NYSE|AMEX)\s*:\s*{re.escape(ticker)}\b)", text, re.I)
        matched_terms = [term for term in method1_company_terms(ticker, company) if re.search(rf"(?<![A-Za-z0-9]){re.escape(term)}(?![A-Za-z0-9])", text, re.I)]
        if structured or matched_terms:
            strength = 1000 if structured else max(len(term) for term in matched_terms)
            candidates.append((strength, ticker, company))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], -len(item[2]), item[1]))
    return candidates[0][1], candidates[0][2]


def method1_rss(source: str, url: str, start: datetime, end: datetime) -> list[dict]:
    response = requests.get(url, headers=METHOD1_HEADERS, timeout=25)
    response.raise_for_status()
    root = ET.fromstring(response.content)
    rows = []
    for node in root.findall(".//item") + root.findall(".//{*}entry"):
        title = (node.findtext("title") or node.findtext("{*}title") or "").strip()
        link = (node.findtext("link") or "").strip()
        if not link:
            link_node = node.find("{*}link")
            link = (link_node.get("href", "") if link_node is not None else "").strip()
        description = node.findtext("description") or node.findtext("{*}summary") or ""
        published_text = node.findtext("pubDate") or node.findtext("{*}published") or node.findtext("{*}updated") or ""
        published = parse_method1_time(published_text)
        if title and link and method1_in_window(published, start, end):
            rows.append({"Title": title, "URL": link, "Source": source, "Published": published, "Summary": BeautifulSoup(description, "html.parser").get_text(" ", strip=True)})
    return rows


def method1_moneydj(start: datetime, end: datetime) -> list[dict]:
    rows = []
    for url in ("https://www.moneydj.com/kmdj/news/newsreal_list.aspx?a=MB010000", "https://www.moneydj.com/kmdj/news/newsreal_list.aspx?a=MB020000"):
        response = requests.get(url, headers=METHOD1_HEADERS, timeout=25)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        for link in soup.select('a[href*="newsviewer.aspx"]'):
            title = link.get_text(" ", strip=True)
            container = link.find_parent(["li", "tr", "div"]) or link
            published = parse_method1_time(container.get_text(" ", strip=True))
            if title and method1_in_window(published, start, end):
                rows.append({"Title": title, "URL": urljoin("https://www.moneydj.com", link.get("href", "")), "Source": "MoneyDJ", "Published": published, "Summary": container.get_text(" ", strip=True)})
    return rows[:METHOD1_POOL_LIMIT]


def method1_udn(start: datetime, end: datetime) -> list[dict]:
    rows = []
    for url in ("https://money.udn.com/rssfeed/news/1001/5591", "https://money.udn.com/rssfeed/news/1001/5590"):
        rows.extend(method1_rss("經濟日報", url, start, end))
    return rows[:METHOD1_POOL_LIMIT]


def method1_cnbc(start: datetime, end: datetime) -> list[dict]:
    rows = []
    for url in ("https://www.cnbc.com/id/100003114/device/rss/rss.html", "https://www.cnbc.com/id/10001147/device/rss/rss.html"):
        rows.extend(method1_rss("CNBC", url, start, end))
    return rows[:METHOD1_POOL_LIMIT]


def method1_cnyes(start: datetime, end: datetime) -> list[dict]:
    rows = []
    for category in ("us_stock", "wd_stock", "headline"):
        url = f"https://api.cnyes.com/media/api/v1/newslist/category/{category}"
        for page in range(1, 7):
            response = requests.get(url, params={"limit": 50, "page": page}, headers=METHOD1_HEADERS, timeout=25)
            if response.status_code == 422:
                break
            response.raise_for_status()
            payload = response.json().get("items") or response.json().get("data", {}).get("items") or []
            records = payload.get("data", []) if isinstance(payload, dict) else payload
            if not records:
                break
            stop_category = False
            for record in records:
                timestamp = record.get("publishAt") or record.get("createdAt") or record.get("created_at")
                try:
                    stamp = int(timestamp)
                    if stamp > 10_000_000_000:
                        stamp //= 1000
                    published = datetime.fromtimestamp(stamp, TAIPEI)
                except Exception:
                    published = None
                if published and published < start:
                    stop_category = True
                    continue
                if not method1_in_window(published, start, end):
                    continue
                raw_title = str(record.get("title") or record.get("name") or "")
                title = BeautifulSoup(raw_title, "html.parser").get_text(" ", strip=True) if "<" in raw_title and ">" in raw_title else raw_title.strip()
                news_id = record.get("newsId") or record.get("id")
                link = record.get("url") or (f"https://news.cnyes.com/news/id/{news_id}" if news_id else "")
                if title and link:
                    rows.append({"Title": title, "URL": link, "Source": "鉅亨網", "Published": published, "Summary": ""})
            if stop_category or len(rows) >= METHOD1_POOL_LIMIT:
                break
    return rows[:METHOD1_POOL_LIMIT]


def method1_tradingview(ticker: str, company: str, start: datetime, end: datetime,
                        session: requests.Session | None = None) -> list[dict]:
    primary = "NASDAQ" if ticker in METHOD1_EXCHANGES else "NYSE"
    response = None
    used_symbol = ""
    for exchange in (primary, "NYSE" if primary == "NASDAQ" else "NASDAQ"):
        used_symbol = f"{exchange}:{ticker}"
        response = (session or requests).get(
            "https://news-mediator.tradingview.com/public/news-flow/v2/news",
            params=[("filter", "lang:zh-Hant"), ("filter", f"symbol:{used_symbol}"), ("client", "landing"), ("streaming", "false"), ("user_prostatus", "non_pro")],
            headers={**METHOD1_HEADERS, "Referer": f"https://tw.tradingview.com/symbols/{used_symbol.replace(':', '-')}/news/"},
            timeout=25,
        )
        if response.status_code != 422:
            response.raise_for_status()
            break
    if response is None or response.status_code == 422:
        return []
    rows = []
    for record in response.json().get("items", []):
        try:
            published = datetime.fromtimestamp(int(record.get("published")), TAIPEI)
        except Exception:
            continue
        title = str(record.get("title") or "").strip()
        link = str(record.get("link") or "").strip()
        if link.startswith("/"):
            link = urljoin("https://tw.tradingview.com", link)
        if not link:
            link = f"https://tw.tradingview.com/symbols/{used_symbol.replace(':', '-')}/news/"
        if title and method1_in_window(published, start, end):
            rows.append({"Ticker": ticker, "Company": company, "Published Time": published.strftime("%Y-%m-%d %H:%M:%S"), "Title": title, "Source": "TradingView", "URL": link, "FinBERT": ""})
    return rows


def fetch_method1_news(companies: pd.DataFrame, start: datetime, end: datetime, progress=None) -> tuple[pd.DataFrame, list[str], dict[str, int]]:
    """方法一：四個新聞池，加 TradingView 逐公司補抓；不使用 Nasdaq。"""
    errors: list[str] = []
    source_counts: dict[str, int] = {name: 0 for name in ("MoneyDJ", "經濟日報", "鉅亨網", "CNBC", "TradingView")}
    rows: list[dict] = []
    pool_fetchers = (("MoneyDJ", method1_moneydj), ("經濟日報", method1_udn), ("鉅亨網", method1_cnyes), ("CNBC", method1_cnbc))
    pool_items: list[dict] = []
    for source, fetcher in pool_fetchers:
        try:
            fetched = fetcher(start, end)
            pool_items.extend(fetched)
        except Exception as exc:
            errors.append(f"{source} 新聞池：{exc}")
        if isinstance(progress, BackgroundProgress):
            progress.job["method1_monitor"].update(
                stage=f"新聞池處理中｜{source}", errors=list(errors), source_counts=dict(source_counts)
            )
    for item in pool_items:
        matched = method1_match_company(item["Title"], item.get("Summary", ""), companies)
        if not matched:
            continue
        ticker, company = matched
        rows.append({"Ticker": ticker, "Company": company, "Published Time": item["Published"].strftime("%Y-%m-%d %H:%M:%S"), "Title": item["Title"], "Source": item["Source"], "URL": item["URL"], "FinBERT": ""})
        source_counts[item["Source"]] += 1

    def fetch_company(record: dict) -> tuple[list[dict], list[str]]:
        ticker, company = record["Ticker"], record["Company"]
        company_rows, company_errors = [], []
        session = method1_http_session()
        for source, fetcher in (("TradingView", method1_tradingview),):
            try:
                company_rows.extend(fetcher(ticker, company, start, end, session=session))
            except Exception as exc:
                company_errors.append(f"{source} {ticker}：{exc}")
        return company_rows, company_errors

    records = companies[["Ticker", "Company"]].to_dict("records")
    # TradingView 是方法一最主要的耗時階段；連線重用加上 12 個工作執行緒，
    # 可降低 S&P 500 逐家請求的總等待時間，同時避免過高並行導致限流。
    executor = ThreadPoolExecutor(max_workers=12)
    futures = {executor.submit(fetch_company, record): record for record in records}
    try:
        for completed, future in enumerate(as_completed(futures), 1):
            if isinstance(progress, BackgroundProgress) and progress.job["stop_event"].is_set():
                raise CrawlCancelled("使用者已停止抓取")
            company_rows, company_errors = future.result()
            rows.extend(company_rows)
            errors.extend(company_errors)
            for row in company_rows:
                source_counts[row["Source"]] += 1
            if progress is not None:
                progress.progress(completed / len(records), text=f"方法一逐公司補抓｜{completed}/{len(records)}")
            if isinstance(progress, BackgroundProgress):
                progress.job["method1_monitor"].update(
                    stage=f"逐公司補抓｜{completed}/{len(records)}", rows=len(rows),
                    errors=list(errors), source_counts=dict(source_counts),
                )
    finally:
        stopping = isinstance(progress, BackgroundProgress) and progress.job["stop_event"].is_set()
        if stopping:
            for future in futures:
                future.cancel()
        executor.shutdown(wait=not stopping, cancel_futures=stopping)
    frame = normalize_news(pd.DataFrame(rows, columns=NEWS_COLUMNS))
    if frame.empty:
        return frame, errors, source_counts
    frame["_url_key"] = frame["URL"].map(normalize_url)
    frame["_title_key"] = frame["Title"].map(normalized_title)
    frame = frame.drop_duplicates(["Company", "_url_key"]).drop_duplicates(["Company", "_title_key"]).drop(columns=["_url_key", "_title_key"])
    frame = frame.sort_values("Published Time", ascending=False).reset_index(drop=True)
    return frame[NEWS_COLUMNS], errors, source_counts


def method2_matches(title: str, companies: pd.DataFrame) -> list[tuple[str, str]]:
    """比對 Google News 標題；同一則新聞可對應多家公司。"""
    text = str(title).strip()
    lowered = text.lower()
    matches: list[tuple[str, str]] = []
    ambiguous = {"A", "AI", "ALL", "AN", "ARE", "AT", "BE", "BY", "FOR", "IT", "ON", "OR", "SO"}
    for row in companies.itertuples(index=False):
        ticker, company = str(row.Ticker).strip().upper(), str(row.Company).strip()
        ticker_hit = bool(ticker and ticker not in ambiguous and re.search(rf"(?<![A-Za-z0-9]){re.escape(ticker)}(?![A-Za-z0-9])", text, re.I))
        company_hit = bool(company and len(company) >= 3 and company.lower() in lowered)
        if ticker_hit or company_hit:
            matches.append((ticker, company))
    return matches


def fetch_method2_raw(companies: pd.DataFrame, start: datetime, end: datetime, progress=None) -> tuple[pd.DataFrame, list[str], int]:
    """方法二第一階段：每五家公司一批查詢 Google News RSS。"""
    records = companies[["Ticker", "Company"]].drop_duplicates("Ticker").to_dict("records")
    batches = [records[index:index + 5] for index in range(0, len(records), 5)]
    rows: list[dict] = []
    errors: list[str] = []
    before_day = (end.astimezone(TAIPEI).date() + timedelta(days=1)).isoformat()
    after_day = start.astimezone(TAIPEI).date().isoformat()
    for batch_index, batch in enumerate(batches, 1):
        terms: list[str] = []
        batch_frame = pd.DataFrame(batch)
        for item in batch:
            terms.extend([f'"{item["Ticker"]}"', f'"{item["Company"]}"'])
        query = f'intitle:({" OR ".join(terms)}) after:{after_day} before:{before_day}'
        rss_url = "https://news.google.com/rss/search?" + urllib.parse.urlencode({"q": query, "hl": "en-US", "gl": "US", "ceid": "US:en"})
        try:
            response = requests.get(rss_url, headers=METHOD1_HEADERS, timeout=12)
            response.raise_for_status()
            root = ET.fromstring(response.content)
            for item in root.findall(".//item"):
                full_title = (item.findtext("title") or "").strip()
                title, source = full_title, "未知媒體"
                if " - " in full_title:
                    title, source = full_title.rsplit(" - ", 1)
                source = (item.findtext("source") or source).strip()
                published = parse_method1_time(item.findtext("pubDate"))
                if not method1_in_window(published, start, end):
                    continue
                for ticker, company in method2_matches(title, batch_frame):
                    rows.append({"Ticker": ticker, "Company": company, "Published Time": published.astimezone(TAIPEI).strftime("%Y-%m-%d %H:%M:%S"), "Title": title, "Source": source, "URL": (item.findtext("link") or "").strip(), "FinBERT": ""})
        except Exception as exc:
            tickers = ", ".join(str(item["Ticker"]) for item in batch)
            errors.append(f"Google News 批次 {tickers}：{exc}")
        if progress is not None:
            progress.progress(batch_index / max(len(batches), 1) * 0.55, text=f"方法二爬取新聞｜{batch_index}/{len(batches)} 批")
        if isinstance(progress, BackgroundProgress):
            progress.job["method2_monitor"].update(
                stage=f"Google News 批次｜{batch_index}/{len(batches)}", raw_rows=len(rows),
                batches=batch_index, errors=list(errors),
            )
        time.sleep(1)
    result = pd.DataFrame(rows, columns=NEWS_COLUMNS)
    if not result.empty:
        result["_url"] = result["URL"].map(normalize_url)
        result["_title"] = result["Title"].map(normalized_title)
        result = result.drop_duplicates(["Ticker", "_url"]).drop_duplicates(["Ticker", "_title"]).drop(columns=["_url", "_title"])
        result = result.sort_values("Published Time", ascending=False).reset_index(drop=True)
    return result, errors, len(batches)


@st.cache_resource(show_spinner=False)
def load_finbert_model():
    try:
        import torch
        from transformers import AutoModelForSequenceClassification, AutoTokenizer
    except ImportError as exc:
        raise RuntimeError("方法二需要 transformers 與 torch，請先執行 pip install -r requirements.txt") from exc
    model_name = "ProsusAI/finbert"
    device = "cuda" if torch.cuda.is_available() else "cpu"
    if device == "cpu":
        # 本機為 8 個邏輯核心；保留少數核心給網站與作業系統，實測 6 執行緒最快。
        cpu_threads = min(6, max(1, int(os.cpu_count() or 4) - 2))
        torch.set_num_threads(cpu_threads)
    try:
        tokenizer = AutoTokenizer.from_pretrained(model_name, local_files_only=True)
        model = AutoModelForSequenceClassification.from_pretrained(model_name, local_files_only=True).to(device)
    except OSError:
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        model = AutoModelForSequenceClassification.from_pretrained(model_name).to(device)
    model.eval()
    return torch, tokenizer, model, device


def score_method2_finbert(frame: pd.DataFrame, progress=None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """方法二第二階段：FinBERT = positive probability - negative probability。"""
    scored = frame.copy()
    if scored.empty:
        scored["FinBERT"] = pd.Series(dtype=float)
        return scored, scored.copy()
    torch, tokenizer, model, device = load_finbert_model()
    titles = scored["Title"].fillna("").astype(str).tolist()
    # 本機 CPU 用較大批次降低模型呼叫開銷；依標題長度分組可減少 padding 的無效運算。
    batch_size = 128
    ordered_titles = sorted(enumerate(titles), key=lambda item: len(item[1].split()))
    scores: list[float] = [0.0] * len(titles)
    labels = {int(key): str(value).lower() for key, value in model.config.id2label.items()}
    with torch.inference_mode():
        for index in range(0, len(ordered_titles), batch_size):
            batch_items = ordered_titles[index:index + batch_size]
            original_indexes = [item[0] for item in batch_items]
            batch = [item[1] for item in batch_items]
            inputs = tokenizer(batch, padding=True, truncation=True, max_length=64, return_tensors="pt").to(device)
            probabilities = torch.nn.functional.softmax(model(**inputs).logits, dim=-1).cpu().numpy()
            for original_index, values in zip(original_indexes, probabilities):
                mapped = {labels[position]: float(values[position]) for position in range(len(values))}
                scores[original_index] = round(mapped.get("positive", 0.0) - mapped.get("negative", 0.0), 3)
            if progress is not None:
                done = min(index + batch_size, len(titles))
                progress.progress(0.55 + 0.45 * done / len(titles), text=f"方法二 FinBERT 評分｜{done}/{len(titles)}")
    scored["FinBERT"] = scores
    scored = scored.sort_values("FinBERT").reset_index(drop=True)
    filtered = scored[(scored["FinBERT"] >= -1) & (scored["FinBERT"] <= 0)].reset_index(drop=True)
    return scored, filtered


def load_company_upload(upload) -> pd.DataFrame:
    frame = pd.read_excel(io.BytesIO(upload.getvalue()), dtype=str)
    columns = {str(column).strip().lower(): column for column in frame.columns}
    ticker_col = next((columns[key] for key in ("ticker", "symbol", "股票代號") if key in columns), None)
    company_col = next((columns[key] for key in ("company", "name", "股票名稱") if key in columns), None)
    if not ticker_col or not company_col:
        raise ValueError("名單需要 Ticker/Symbol 與 Company/Name 欄位")
    result = frame[[ticker_col, company_col]].rename(columns={ticker_col: "Ticker", company_col: "Company"}).dropna()
    result["Ticker"] = result["Ticker"].astype(str).str.strip().str.upper()
    result["Company"] = result["Company"].astype(str).str.strip()
    return result[(result["Ticker"] != "") & (result["Company"] != "")].drop_duplicates("Ticker")


def load_saved_companies() -> pd.DataFrame:
    if not COMPANY_PATH.is_file():
        raise ValueError("尚未建立 S&P 500 公司名單版本，請先上傳新版 Excel")
    payload = io.BytesIO(COMPANY_PATH.read_bytes())
    return load_company_upload(payload)


@st.cache_data(ttl=3600)
def stock_prices(ticker: str, start_text: str, end_text: str) -> pd.DataFrame:
    start = int(datetime.fromisoformat(start_text).replace(tzinfo=timezone.utc).timestamp())
    end = int((datetime.fromisoformat(end_text) + timedelta(days=1)).replace(tzinfo=timezone.utc).timestamp())
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(ticker)}?period1={start}&period2={end}&interval=1d&events=history"
    response = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    result = response.json()["chart"]["result"][0]
    timestamps = result.get("timestamp", [])
    quotes = result["indicators"]["quote"][0]
    rows = []
    for index, stamp in enumerate(timestamps):
        close = quotes["close"][index]
        if close is not None:
            rows.append({"Date": datetime.fromtimestamp(stamp, timezone.utc).date(), "Close": float(close)})
    return pd.DataFrame(rows)


def output_files() -> list[Path]:
    return sorted(OUTPUT_DIR.glob("美股_*_負面新聞爬蟲.xlsx"), key=lambda path: (re.search(r"20\d{6}", path.name).group(0), path.stat().st_mtime), reverse=True)


def load_negative_files() -> pd.DataFrame:
    frames = []
    for path in output_files():
        frame = pd.read_excel(path, dtype=str)
        match = re.search(r"20\d{6}", path.name)
        frame["資料日期"] = match.group(0) if match else ""
        frames.append(frame)
    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()


def continuous_events(frame: pd.DataFrame) -> pd.DataFrame:
    """同公司、同 Event_Code 跨工作簿日期彙整；至少兩個日期才列入。"""
    if frame.empty:
        return pd.DataFrame()
    data = frame.copy()
    data["公司鍵"] = data["Ticker"].fillna("").where(data["Ticker"].fillna("") != "", data["Company"].fillna("").str.lower())
    data["Level"] = pd.to_numeric(data["Level"], errors="coerce").fillna(0)
    rows = []
    for (_, event_code), group in data.groupby(["公司鍵", "Event_Code"], dropna=False):
        days = sorted(set(group["資料日期"].dropna().astype(str)))
        if len(days) < 2:
            continue
        newest = group.sort_values(["資料日期", "Published Time"]).iloc[-1]
        rows.append({
            "Ticker": newest.get("Ticker", ""), "Company": newest.get("Company", ""), "Event_Code": event_code,
            "Event_type": newest.get("Event_type", ""), "首次出現": days[0], "最近出現": days[-1],
            "出現天數": len(days), "新聞篇數": len(group), "最高事件等級": int(group["Level"].max()),
            "最新標題": newest.get("Title", ""), "URL": newest.get("URL", ""),
        })
    return pd.DataFrame(rows).sort_values(["最高事件等級", "最近出現"], ascending=[False, False]) if rows else pd.DataFrame()


def save_status(ticker: str, company: str, status: str, owner: str, next_review: str, note: str) -> None:
    key = ticker or company.lower()
    with sqlite3.connect(DB_PATH) as connection:
        connection.execute(
            "INSERT INTO manual_status VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(company_key) DO UPDATE SET status=excluded.status,owner=excluded.owner,next_review=excluded.next_review,note=excluded.note,updated_at=excluded.updated_at",
            (key, ticker, company, status, owner, next_review, note, datetime.now(TAIPEI).isoformat(timespec="seconds")),
        )


PAGE_COPY = {
    "今日任務": ("今日新聞擷取", "選擇公司範圍與時間，一次取得可交付的新聞 Excel。"),
    "風險儀表板": ("負面新聞風險儀表板", "先掌握高風險事件，再查看公司與新聞明細。"),
    "持續追蹤": ("持續追蹤工作台", "集中查看跨日事件、人工處理進度與事件後股價。"),
    "設定與說明": ("設定與系統說明", "管理事件規則、資料位置與版本恢復。"),
}


def page_heading(page_name: str) -> None:
    title, description = PAGE_COPY[page_name]
    st.markdown(f"<div class='page-kicker'>{page_name}</div><div class='page-title'>{title}</div><div class='page-description'>{description}</div>", unsafe_allow_html=True)


initialize()
st.set_page_config(page_title=APP_TITLE, page_icon="🛡️", layout="wide")
st.markdown("""
<style>
  .stApp { background:#F7F9FC; }
  .block-container { max-width:1500px; padding-top:1.4rem; padding-bottom:4rem; }
  .app-hero { background:linear-gradient(120deg,#102A43,#1D4ED8); color:white; border-radius:16px; padding:20px 26px; margin-bottom:14px; box-shadow:0 10px 30px rgba(15,42,67,.12); }
  .app-hero-title { font-size:25px; font-weight:750; letter-spacing:.02em; }
  .app-hero-sub { opacity:.78; font-size:13px; margin-top:4px; }
  .page-kicker { color:#2563EB; font-size:13px; font-weight:700; letter-spacing:.08em; margin-top:24px; }
  .page-title { color:#172B4D; font-size:31px; line-height:1.25; font-weight:780; margin-top:4px; }
  .page-description { color:#64748B; font-size:16px; margin:7px 0 20px; }
  div[data-testid='stRadio'] > div { background:white; border:1px solid #E2E8F0; border-radius:13px; padding:10px; box-shadow:0 3px 12px rgba(15,23,42,.04); }
  div[data-testid='stRadio'] div[role='radiogroup'] { display:flex; gap:8px; width:100%; }
  div[data-testid='stRadio'] label[data-baseweb='radio'] { flex:1; min-height:50px; display:flex; align-items:center; justify-content:center; border-radius:9px; padding:8px 10px; color:#64748B; font-weight:700; cursor:pointer; transition:all .16s ease; }
  div[data-testid='stRadio'] label[data-baseweb='radio']:hover { background:#F1F5F9; color:#1D4ED8; }
  div[data-testid='stRadio'] label[data-baseweb='radio']:has(input:checked) { background:#EAF2FF; color:#1D4ED8; box-shadow:inset 0 0 0 1px #BFDBFE; }
  div[data-testid='stRadio'] label[data-baseweb='radio'] > div:first-child { display:none; }
  .method-card { background:white; border:1px solid #DBEAFE; border-left:6px solid #2563EB; border-radius:12px; padding:16px 18px; margin:4px 0 20px; }
  .method-title { color:#1E3A8A; font-weight:750; font-size:17px; }
  .method-desc { color:#64748B; font-size:14px; margin-top:5px; }
  .status-title { color:#172B4D; font-size:27px; line-height:1.15; font-weight:750; margin:0 0 4px; }
  .previous-result-grid { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:16px; margin:4px 0 16px; }
  .previous-result-card { background:white; border:1px solid #E2E8F0; border-radius:12px; padding:16px 18px; min-height:108px; box-shadow:0 2px 8px rgba(15,23,42,.035); }
  .previous-result-label { color:#334155; font-size:15px; margin-bottom:10px; }
  .previous-result-value { color:#1E293B; font-size:24px; line-height:1.3; font-weight:520; white-space:normal; overflow-wrap:anywhere; }
  div[data-testid='stMetric'] { background:white; border:1px solid #E2E8F0; border-radius:12px; padding:14px 16px; box-shadow:0 2px 8px rgba(15,23,42,.035); }
  div[data-testid='stFileUploader'] { background:white; border-radius:12px; }
  div[data-testid='stDataFrame'] { border:1px solid #E2E8F0; border-radius:12px; overflow:hidden; }
  div.stButton > button[kind='primary'] { min-height:46px; border-radius:10px; font-weight:700; }
  div.stDownloadButton > button { min-height:43px; border-radius:10px; font-weight:650; }
  h3 { color:#172B4D !important; margin-top:1.4rem !important; }
  @media (max-width:1000px) { .previous-result-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
  @media (max-width:800px) { div[data-testid='stRadio'] div[role='radiogroup'] { display:block; } div[data-testid='stRadio'] label[data-baseweb='radio'] { justify-content:flex-start; } .page-title { font-size:25px; } .previous-result-grid { grid-template-columns:1fr; } }
</style>
""", unsafe_allow_html=True)
hero_subtitle = (
    f"{APP_VERSION}｜抓取兩種來源、合併去重並整理負面新聞"
    if SIMPLE_SITE else f"{APP_VERSION}｜資料獨立保存，不使用舊系統執行資料"
)
st.markdown(f"<div class='app-hero'><div class='app-hero-title'>🛡️ {APP_TITLE}</div><div class='app-hero-sub'>{hero_subtitle}</div></div>", unsafe_allow_html=True)

if not SIMPLE_SITE and st.session_state.pop("_go_dashboard", False):
    st.session_state["page"] = "風險儀表板"
if SIMPLE_SITE:
    page = "今日任務"
    page_heading(page)
    st.markdown("<div class='method-card'><div class='method-title'>今日任務｜雙方法合併與負面新聞整理</div><div class='method-desc'>系統會依序抓取多來源新聞與 Google News，合併兩份結果、刪除重複新聞、執行 FinBERT 評分，最後依事件規則整理負面新聞。</div></div>", unsafe_allow_html=True)
else:
    nav_icons = {"今日任務": "① 今日任務", "風險儀表板": "② 風險儀表板", "持續追蹤": "🔔 持續追蹤", "設定與說明": "🧰 設定與說明"}
    page = st.radio("頁面", list(nav_icons), horizontal=True, label_visibility="collapsed", key="page", format_func=lambda value: nav_icons[value])

if page == "今日任務":
    st.markdown("### 1. 選擇公司範圍")
    company_universe_options = ["Dow Jones 30", "S&P 500", "自行上傳最新版"]
    universe = st.radio("本次要使用的公司名單", company_universe_options, index=1, horizontal=True)
    company_upload = None
    if universe == "自行上傳最新版":
        company_upload = st.file_uploader("上傳最新公司名單 Excel", type=["xlsx"], key="company_list_upload")
        if st.button("檢查並套用新版", use_container_width=True, disabled=company_upload is None):
            try:
                checked_companies = load_company_upload(company_upload)
                if checked_companies.empty:
                    raise ValueError("公司名單沒有可用資料")
                backup_dir = DATA_ROOT / "backups"
                backup_dir.mkdir(parents=True, exist_ok=True)
                if COMPANY_PATH.is_file():
                    backup_stamp = datetime.now(TAIPEI).strftime("%Y%m%d_%H%M%S")
                    (backup_dir / f"Company_List_{backup_stamp}.xlsx").write_bytes(COMPANY_PATH.read_bytes())
                COMPANY_PATH.write_bytes(company_upload.getvalue())
                applied_at = datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
                with sqlite3.connect(DB_PATH) as connection:
                    connection.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('company_original_name',?)", (company_upload.name,))
                    connection.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('company_updated_at',?)", (applied_at,))
                st.success(f"新版公司名單已套用：{company_upload.name}，共 {len(checked_companies):,} 家。之後會持續沿用到下次更新。")
            except Exception as exc:
                st.error(f"新版公司名單無法套用：{exc}")
    elif universe == "S&P 500" and COMPANY_PATH.is_file():
        with sqlite3.connect(DB_PATH) as connection:
            saved_company_name = connection.execute("SELECT value FROM settings WHERE key='company_original_name'").fetchone()
        current_company_name = saved_company_name[0] if saved_company_name else COMPANY_PATH.name
        st.success(f"目前使用：{current_company_name}（下次上傳新版後會自動更換）")
    now = datetime.now(TAIPEI).replace(hour=9, minute=0, second=0, microsecond=0)
    previous_end = last_success_end()
    default_start = (previous_end or (now - timedelta(days=3))).replace(hour=9, minute=0, second=0, microsecond=0)
    if previous_end:
        st.caption(f"前次成功結束日期：{previous_end:%Y-%m-%d}；本次時間預設為 09:00。")
    st.markdown("### 2. 設定新聞期間")
    start_col, end_col = st.columns(2)
    start_date = start_col.date_input("開始日期", value=default_start.date())
    start_hour = start_col.selectbox("開始時間", list(range(24)), index=default_start.hour, format_func=lambda hour: f"{hour:02d}:00")
    end_date = end_col.date_input("結束日期", value=now.date())
    end_hour = end_col.selectbox("結束時間", list(range(24)), index=now.hour, format_func=lambda hour: f"{hour:02d}:00")
    start_dt = datetime.combine(start_date, datetime.min.time(), TAIPEI) + timedelta(hours=start_hour)
    end_dt = datetime.combine(end_date, datetime.min.time(), TAIPEI) + timedelta(hours=end_hour)
    if start_dt > end_dt:
        st.error("開始時間不可晚於結束時間")
    st.caption(f"預計擷取：{start_dt:%Y-%m-%d %H:%M} ～ {end_dt:%Y-%m-%d %H:%M}（台北時間）")
    st.markdown("### 3. 今日新聞擷取")
    crawl_method = st.radio("選擇擷取方法", ["方法一｜多來源快速擷取", "方法二｜Google News＋FinBERT", "方法一＋方法二｜完整整合"], index=2, horizontal=True)
    if crawl_method == "方法一｜多來源快速擷取":
        st.markdown("<div class='method-card'><div class='method-title'>方法一｜多來源快速擷取</div><div class='method-desc'>彙整 MoneyDJ、經濟日報、鉅亨網與 CNBC，再由 TradingView 逐公司補抓；不使用 Nasdaq。</div></div>", unsafe_allow_html=True)
    elif crawl_method == "方法二｜Google News＋FinBERT":
        st.markdown("<div class='method-card'><div class='method-title'>方法二｜Google News＋FinBERT</div><div class='method-desc'>分批查詢英文 Google News RSS，完成後以 ProsusAI/finbert 分析標題。</div></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='method-card'><div class='method-title'>方法一＋方法二｜完整整合</div><div class='method-desc'>依序完成兩種爬蟲、合併去重，再統一執行 FinBERT 評分。其中方法一不使用 Nasdaq。</div></div>", unsafe_allow_html=True)
    method_name = {"方法一｜多來源快速擷取": "方法一", "方法二｜Google News＋FinBERT": "方法二", "方法一＋方法二｜完整整合": "方法一＋方法二"}[crawl_method]
    registry = crawl_job_registry()
    current_job = registry.get("current")
    is_running = bool(current_job and current_job.get("status") in ("running", "stopping"))
    if st.button(
        f"開始執行{method_name}", type="primary",
        disabled=start_dt > end_dt or is_running or universe == "自行上傳最新版",
        use_container_width=True,
    ):
        try:
            if universe.startswith("Dow"):
                companies = pd.DataFrame(DOW_30, columns=["Ticker", "Company"])
            elif universe.startswith("S&P"):
                companies = load_saved_companies() if COMPANY_PATH.is_file() else load_sp500()
            else:
                raise ValueError("請先套用新版名單，再選擇 S&P 500 開始執行")
            run_id = create_run(start_dt, end_dt, f"{method_name}｜{universe}")
            script_context = get_script_run_ctx()
            use_method1 = method_name in ("方法一", "方法一＋方法二")
            use_method2 = method_name in ("方法二", "方法一＋方法二")
            job = {
                "run_id": run_id, "status": "running", "method": method_name,
                "started_monotonic": time.monotonic(), "started_at": datetime.now(TAIPEI).isoformat(timespec="seconds"),
                "started_epoch_ms": int(time.time() * 1000),
                "progress": 0.0, "progress_text": "準備開始", "stop_event": threading.Event(),
                "stage_times": [], "current_stage": None, "current_stage_started": None,
                "script_context": script_context,
                "session_id": script_context.session_id if script_context is not None else "",
                "method1_monitor": {
                    "status": "pending" if use_method1 else "skipped", "stage": "等待執行" if use_method1 else "本次未執行",
                    "rows": 0, "errors": [], "source_counts": {},
                },
                "method2_monitor": {
                    "status": "pending" if use_method2 else "skipped", "stage": "等待執行" if use_method2 else "本次未執行",
                    "rows": 0, "raw_rows": 0, "negative_rows": 0, "batches": 0, "errors": [],
                },
            }
            registry["current"] = job
            worker = threading.Thread(
                target=run_crawl_job,
                args=(job, method_name, universe, companies.copy(), start_dt, end_dt),
                name=f"news-crawl-{run_id}", daemon=True,
            )
            job["thread"] = worker
            worker.start()
            st.rerun()
        except Exception as exc:
            st.error(f"無法開始抓取：{exc}")

    def render_crawl_status():
        job = crawl_job_registry().get("current")
        if not job:
            return False
        end_tick = job.get("finished_monotonic", time.monotonic())
        elapsed_text = format_elapsed(end_tick - job["started_monotonic"])
        status = job.get("status", "running")
        with st.container(border=True):
            left, right = st.columns([3, 1])
            with left:
                st.markdown("<div class='status-title'>執行狀態</div>", unsafe_allow_html=True)
                st.progress(job.get("progress", 0), text=job.get("progress_text", "執行中"))
            if status in ("running", "stopping"):
                started_epoch_ms = int(job.get("started_epoch_ms", time.time() * 1000))
                with right:
                    components.html(
                        f"""
                        <div style="font-family:Arial,sans-serif;border:1px solid #e2e8f0;border-radius:12px;
                                    padding:10px 16px;background:white;height:78px;box-sizing:border-box;">
                          <div style="font-size:14px;color:#334155;margin-bottom:3px;">抓新聞所用時間</div>
                          <div id="crawl-elapsed" style="font-size:26px;color:#1e293b;line-height:1.08;">{elapsed_text}</div>
                        </div>
                        <script>
                          const started = {started_epoch_ms};
                          const pad = n => String(n).padStart(2, '0');
                          const update = () => {{
                            const total = Math.max(0, Math.floor((Date.now() - started) / 1000));
                            const h = Math.floor(total / 3600);
                            const m = Math.floor((total % 3600) / 60);
                            const s = total % 60;
                            document.getElementById('crawl-elapsed').textContent = `${{pad(h)}}:${{pad(m)}}:${{pad(s)}}`;
                          }};
                          update(); setInterval(update, 1000);
                        </script>
                        """,
                        height=82,
                    )
            else:
                right.metric("抓新聞所用時間", elapsed_text)

            stage_rows = [
                {"執行階段": item["label"], "所用時間": format_elapsed(item["seconds"]), "狀態": "已完成"}
                for item in job.get("stage_times", [])
            ]
            current_stage = job.get("current_stage")
            current_stage_started = job.get("current_stage_started")
            if current_stage and current_stage_started is not None:
                stage_rows.append({
                    "執行階段": current_stage,
                    "所用時間": format_elapsed(time.monotonic() - float(current_stage_started)),
                    "狀態": "執行中" if status == "running" else "停止中",
                })
            if stage_rows and not SIMPLE_SITE:
                st.markdown("#### 各階段所用時間")
                st.dataframe(pd.DataFrame(stage_rows), hide_index=True, use_container_width=True)

            if job.get("method") == "方法一＋方法二" and not SIMPLE_SITE:
                st.markdown("#### 方法別監控")
                monitor_columns = st.columns(2)
                status_labels = {
                    "pending": "⏳ 等待中", "running": "🔄 執行中", "success": "✅ 已完成",
                    "failed": "❌ 執行異常", "stopped": "⏹️ 已停止", "skipped": "— 本次未執行",
                }

                def render_method_monitor(container, title: str, monitor: dict, is_method1: bool) -> None:
                    errors = list(monitor.get("errors") or [])
                    shown_status = monitor.get("status", "pending")
                    status_text = status_labels.get(shown_status, shown_status)
                    if shown_status == "success" and errors:
                        status_text = "⚠️ 完成但有異常"
                    with container:
                        with st.container(border=True):
                            st.markdown(f"**{title}｜{status_text}**")
                            st.caption(str(monitor.get("stage") or "等待更新"))
                            metric_columns = st.columns(2)
                            metric_columns[0].metric("新聞筆數", f"{int(monitor.get('rows') or monitor.get('raw_rows') or 0):,}")
                            metric_columns[1].metric("異常數", f"{len(errors):,}")
                            if is_method1:
                                counts = monitor.get("source_counts") or {}
                                if counts:
                                    source_text = "｜".join(f"{source} {int(count):,}" for source, count in sorted(counts.items()))
                                    st.caption(f"來源筆數：{source_text}")
                                else:
                                    st.caption("來源筆數：尚未產生")
                            else:
                                st.caption(
                                    f"Google News 批次：{int(monitor.get('batches') or 0):,}｜"
                                    f"FinBERT ≤ 0：{int(monitor.get('negative_rows') or 0):,}"
                                )
                            if errors:
                                with st.expander(f"查看 {len(errors)} 項異常"):
                                    for error in errors:
                                        st.write(f"• {error}")
                            if monitor.get("fatal_error"):
                                st.error(str(monitor["fatal_error"]))

                render_method_monitor(monitor_columns[0], "方法一｜多來源", job.get("method1_monitor", {}), True)
                render_method_monitor(monitor_columns[1], "方法二｜Google News＋FinBERT", job.get("method2_monitor", {}), False)

            button_labels = {
                "running": "停止抓取", "stopping": "停止中…", "success": "✅ 已完成",
                "stopped": "已停止", "failed": "執行失敗",
            }
            stop_col, refresh_col = st.columns(2)
            stop_clicked = stop_col.button(
                button_labels.get(status, "目前不可用"),
                type="secondary", use_container_width=True, key="stop_crawl",
                disabled=status != "running",
            )
            if status == "success" and not SIMPLE_SITE:
                if refresh_col.button("前往風險儀表板 →", type="primary", use_container_width=True, key="completed_go_dashboard"):
                    st.session_state["_go_dashboard"] = True
                    st.rerun()
            else:
                refresh_col.button("更新目前進度", use_container_width=True, key="refresh_crawl_status")
            if stop_clicked:
                job["status"] = "stopping"
                job["progress_text"] = "正在安全停止，請稍候…"
                job["stop_event"].set()
            status_messages = {
                "running": "抓取正在背景執行，可繼續停留在本頁查看進度。",
                "stopping": "已收到停止要求，正在結束目前的網路請求並清理未完成工作。",
                "success": (
                    f"{job.get('summary', '抓取完成')}；負面新聞 {job.get('event_rows', 0)} 則；"
                    f"待人工覆核 {job.get('unknown_rows', 0)} 則；無關新聞 {job.get('irrelevant_rows', 0)} 則。"
                    f"中文翻譯失敗 {job.get('translation_failures', 0)} 則。"
                    f"總耗時 {elapsed_text}"
                ),
                "stopped": f"抓取已停止。總耗時 {elapsed_text}；未完成的結果不會覆蓋前次成功紀錄。",
                "failed": f"抓取失敗：{job.get('error', '未知錯誤')}（耗時 {elapsed_text}）",
            }
            status_message = status_messages.get(status, "正在更新執行狀態…")
            if status == "success":
                st.success(f"✅ 全部完成｜{status_message}")
            elif status == "failed":
                st.error(status_message)
            elif status in ("stopping", "stopped"):
                st.warning(status_message)
            else:
                st.info(status_message)
        return False

    auto_refresh_running = render_crawl_status()
    saved_crawl = latest_crawl_result()
    if saved_crawl:
        current_job = crawl_job_registry().get("current")
        is_current_result = bool(
            current_job
            and current_job.get("status") == "success"
            and current_job.get("path")
            and Path(current_job["path"]) == saved_crawl["path"]
        )
        st.markdown("### 本次抓取結果" if is_current_result else "### 最近一次抓取結果")
        try:
            saved_start = datetime.fromisoformat(saved_crawl["start_time"]).astimezone(TAIPEI)
            saved_end = datetime.fromisoformat(saved_crawl["end_time"]).astimezone(TAIPEI)
            started = datetime.fromisoformat(saved_crawl["started_at"]).astimezone(TAIPEI)
            finished = datetime.fromisoformat(saved_crawl["finished_at"]).astimezone(TAIPEI)
            period_text = f"{saved_start:%m/%d %H:%M} ～ {saved_end:%m/%d %H:%M}"
            finished_text = f"{finished:%Y-%m-%d %H:%M}"
            elapsed_text = format_elapsed((finished - started).total_seconds())
        except (TypeError, ValueError):
            period_text, finished_text, elapsed_text = "時間紀錄無法解析", str(saved_crawl["finished_at"] or ""), "—"
        st.markdown(
            f"""
            <div class="previous-result-grid">
              <div class="previous-result-card"><div class="previous-result-label">擷取方法</div><div class="previous-result-value">{saved_crawl['method']}</div></div>
              <div class="previous-result-card"><div class="previous-result-label">新聞筆數</div><div class="previous-result-value">{saved_crawl['rows']:,} 筆</div></div>
              <div class="previous-result-card"><div class="previous-result-label">新聞期間</div><div class="previous-result-value">{period_text}</div></div>
              <div class="previous-result-card"><div class="previous-result-label">抓新聞所用時間</div><div class="previous-result-value">{elapsed_text}</div></div>
              <div class="previous-result-card"><div class="previous-result-label">完成時間</div><div class="previous-result-value">{finished_text}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        file_prefix = "今日" if saved_crawl["is_today"] else "前次"
        download_count = 1 + int(saved_crawl["event_path"] is not None) + int(saved_crawl["finbert_path"] is not None)
        download_columns = st.columns(download_count)
        main_label = "完整整合新聞" if saved_crawl["method"] == "方法一＋方法二" else "新聞"
        if saved_crawl["method"] == "方法一＋方法二":
            main_download_name = saved_crawl["path"].name
        else:
            try:
                download_stamp = datetime.fromisoformat(saved_crawl["end_time"]).astimezone(TAIPEI).strftime("%Y%m%d")
            except (TypeError, ValueError):
                download_stamp = datetime.now(TAIPEI).strftime("%Y%m%d")
            main_download_name = f"新聞爬蟲_{download_stamp}.xlsx"
        download_columns[0].download_button(f"下載{file_prefix}{main_label}", saved_crawl["path"].read_bytes(), main_download_name, use_container_width=True, key="saved_crawl_all")
        column_index = 1
        if saved_crawl["event_path"]:
            download_columns[column_index].download_button(f"下載{file_prefix}負面新聞", saved_crawl["event_path"].read_bytes(), saved_crawl["event_path"].name, use_container_width=True, key="saved_crawl_event")
            column_index += 1
        if saved_crawl["finbert_path"]:
            download_columns[column_index].download_button(f"下載{file_prefix} FinBERT ≤ 0", saved_crawl["finbert_path"].read_bytes(), saved_crawl["finbert_path"].name, use_container_width=True, key="saved_crawl_finbert")
        if not SIMPLE_SITE and saved_crawl["event_path"] and st.button("下一步：查看風險儀表板 →", type="primary", use_container_width=True):
            st.session_state["_go_dashboard"] = True
            st.rerun()
        if not saved_crawl["is_today"]:
            st.caption("今天尚未完成新的抓取；目前提供的是最近一次成功結果。")
    else:
        st.info("尚無成功抓取紀錄。完成第一次抓取後，結果會保留在這裡，重新整理或重新開啟網站也能下載。")
if page == "風險儀表板":
    page_heading(page)
    files = output_files()
    if not files:
        st.info("目前還沒有可顯示的負面新聞結果。請先到「今日任務」執行任一擷取方法，系統會自動去重並分類。")
    else:
        selected = st.selectbox("查看哪一次結果", files, format_func=lambda path: path.name)
        data = pd.read_excel(selected)
        if data.empty:
            st.info("這次分類結果為 0 筆：沒有新聞命中負面事件規則。可返回今日任務下載完整新聞，或改看其他日期結果。")
            st.stop()
        data["Level"] = pd.to_numeric(data["Level"], errors="coerce").fillna(0).astype(int)
        match = re.search(r"20\d{6}", selected.name)
        data_day = datetime.strptime(match.group(0), "%Y%m%d").strftime("%Y-%m-%d") if match else "未知"
        st.caption(f"資料日期：{data_day}｜檔案更新：{datetime.fromtimestamp(selected.stat().st_mtime):%Y-%m-%d %H:%M}")
        metrics = st.columns(4)
        metrics[0].metric("負面新聞", len(data))
        metrics[1].metric("涉及公司", data["Ticker"].replace("", pd.NA).nunique())
        metrics[2].metric("重大事件", int((data["Level"] >= 4).sum()))
        metrics[3].metric("持續追蹤", int(data["Action"].fillna("").str.contains("持續追蹤").sum()))
        chart_left, chart_right = st.columns([1.45, 1], gap="large")
        active_rules = load_rules(RULE_PATH.stat().st_mtime_ns)
        event_name_map = {
            str(rule.get("Event_Code") or ""): str(rule.get("事件中文") or rule.get("Event_Code") or "")
            for rule in active_rules
        }
        event_names = data["Event_Code"].fillna("").astype(str).map(event_name_map)
        event_names = event_names.where(event_names.fillna("").str.strip() != "", data["Event_Code"].fillna("未命名事件").astype(str))
        event_chart = (
            event_names.value_counts().head(10)
            .rename_axis("事件類型").reset_index(name="新聞筆數")
        )
        chart_left.markdown("#### 最常出現的負面事件")
        chart_left.caption("依新聞筆數由高到低排列，滑鼠移到長條可查看數字。")
        event_axis_max = max(int(event_chart["新聞筆數"].max() * 1.15), 1)
        event_plot = (
            alt.Chart(event_chart)
            .mark_bar(color="#2563EB", cornerRadiusEnd=5, size=20)
            .encode(
                x=alt.X("新聞筆數:Q", title="新聞筆數", scale=alt.Scale(domain=[0, event_axis_max]), axis=alt.Axis(tickMinStep=1, gridColor="#E8EDF5")),
                y=alt.Y("事件類型:N", title=None, sort="-x", axis=alt.Axis(labelLimit=150)),
                tooltip=[alt.Tooltip("事件類型:N"), alt.Tooltip("新聞筆數:Q", format=",d")],
            )
        )
        event_labels = event_plot.mark_text(
            align="left", baseline="middle", dx=5, color="#334155", fontSize=12
        ).encode(text=alt.Text("新聞筆數:Q", format=",d"))
        chart_left.altair_chart(
            (event_plot + event_labels).properties(height=330).configure_view(strokeWidth=0),
            use_container_width=True,
        )

        chart_right.markdown("#### 風險等級分布")
        chart_right.caption("等級越高代表越需要優先處理。")
        total_events = max(len(data), 1)
        level_cards = [
            (5, "立即關注", "#B91C1C", "#FEF2F2"),
            (4, "高度關注", "#C2410C", "#FFF7ED"),
            (3, "一般關注", "#A16207", "#FEFCE8"),
        ]
        for level, label, color, background in level_cards:
            count = int((data["Level"] == level).sum()) if level > 3 else int((data["Level"] <= 3).sum())
            percentage = count / total_events * 100
            shown_level = f"Level {level}" if level > 3 else "Level 3 以下"
            chart_right.markdown(
                f"""
                <div style="background:{background};border-left:6px solid {color};border-radius:10px;
                            padding:14px 16px;margin:0 0 12px 0;display:flex;align-items:center;justify-content:space-between;">
                  <div><div style="font-weight:700;color:{color};font-size:17px;">{shown_level}｜{label}</div>
                       <div style="color:#64748B;font-size:13px;margin-top:3px;">占全部新聞 {percentage:.1f}%</div></div>
                  <div style="font-size:30px;font-weight:750;color:#1E293B;">{count}<span style="font-size:14px;font-weight:500;color:#64748B;"> 筆</span></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown("### 新聞明細與篩選")
        filter_columns = st.columns([2, 1, 1])
        search = filter_columns[0].text_input("搜尋公司、Ticker、標題或事件")
        level_filter = filter_columns[1].selectbox("事件等級", ["全部", "Level 5", "Level 4", "Level 3 以下"])
        tracking_only = filter_columns[2].checkbox("只看持續追蹤")
        filtered = data.copy()
        if search:
            mask = filtered[["Ticker", "Company", "Title", "Event_type", "Event_Code"]].fillna("").astype(str).apply(lambda column: column.str.contains(search, case=False, regex=False)).any(axis=1)
            filtered = filtered[mask]
        if level_filter == "Level 5":
            filtered = filtered[filtered["Level"] == 5]
        elif level_filter == "Level 4":
            filtered = filtered[filtered["Level"] == 4]
        elif level_filter == "Level 3 以下":
            filtered = filtered[filtered["Level"] <= 3]
        if tracking_only:
            filtered = filtered[filtered["Action"].fillna("").str.contains("持續追蹤")]
        st.caption(f"篩選結果：顯示 {len(filtered):,}／{len(data):,} 筆；可點最右側「開啟」查看原始新聞。")
        st.dataframe(filtered, use_container_width=True, hide_index=True, column_config={"URL": st.column_config.LinkColumn("新聞", display_text="開啟")})
        st.markdown("### 跨日連續新聞")
        cross_day = continuous_events(load_negative_files())
        if cross_day.empty:
            st.caption("目前沒有同公司、同事件出現在兩個不同資料日期。")
        else:
            st.dataframe(cross_day, use_container_width=True, hide_index=True, column_config={"URL": st.column_config.LinkColumn("新聞", display_text="開啟")})

if page == "持續追蹤":
    page_heading(page)
    st.markdown("<div class='method-card'><div class='method-title'>這裡只顯示需要持續追蹤的公司</div><div class='method-desc'>同一家公司跨日出現會合併成一列；先用期間或公司名稱篩選，再點選公司查看事件、處理狀態與股價。</div></div>", unsafe_allow_html=True)
    all_data = load_negative_files()
    if all_data.empty:
        st.info("尚無負面新聞資料。")
    else:
        tracked = all_data[all_data["Action"].fillna("").str.contains("持續追蹤")].copy()
        period_column, search_column = st.columns([1, 2])
        period = period_column.selectbox("查看期間", ["最近 7 天", "最近 14 天", "最近 30 天", "全部日期"])
        search = search_column.text_input("搜尋公司或 Ticker", placeholder="例如：NVIDIA、NVDA")
        tracked["日期值"] = pd.to_datetime(tracked["資料日期"], format="%Y%m%d", errors="coerce")
        days = {"最近 7 天": 7, "最近 14 天": 14, "最近 30 天": 30}.get(period)
        if days:
            latest_day = tracked["日期值"].max()
            tracked = tracked[tracked["日期值"] >= latest_day - pd.Timedelta(days=days - 1)]
        if search:
            mask = tracked[["Ticker", "Company"]].fillna("").astype(str).apply(lambda column: column.str.contains(search, case=False, regex=False)).any(axis=1)
            tracked = tracked[mask]
        if tracked.empty:
            st.warning("此期間或搜尋條件沒有持續追蹤公司。")
            st.stop()
        else:
            tracked = tracked.sort_values(["日期值", "Published Time"])
            grouped = tracked.groupby(["Ticker", "Company"], dropna=False)
            summary = grouped.agg(首次列入=("資料日期", "min"), 最近列入=("資料日期", "max"), 出現天數=("資料日期", "nunique"), 新聞篇數=("Title", "count")).reset_index()
            event_types = grouped["Event_Code"].agg(lambda values: "、".join(dict.fromkeys(str(value) for value in values if pd.notna(value) and str(value).strip()))).reset_index(name="事件類型")
            latest_news = tracked.groupby(["Ticker", "Company"], as_index=False).tail(1)[["Ticker", "Company", "Title", "URL"]].rename(columns={"Title": "最新標題", "URL": "最新新聞"})
            summary = summary.merge(event_types, on=["Ticker", "Company"], how="left").merge(latest_news, on=["Ticker", "Company"], how="left")
            summary.insert(0, "市場", "美股")
            summary = summary.sort_values(["出現天數", "新聞篇數", "最近列入"], ascending=[False, False, False]).reset_index(drop=True)
            summary["首次列入"] = pd.to_datetime(summary["首次列入"], format="%Y%m%d", errors="coerce").dt.strftime("%Y-%m-%d")
            summary["最近列入"] = pd.to_datetime(summary["最近列入"], format="%Y%m%d", errors="coerce").dt.strftime("%Y-%m-%d")
            with sqlite3.connect(DB_PATH) as connection:
                saved_states = connection.execute("SELECT company_key,status,owner,next_review FROM manual_status").fetchall()
            state_map = {row[0]: row[1:] for row in saved_states}
            company_keys = [str(row.Ticker) if str(row.Ticker).strip() else str(row.Company).lower() for row in summary.itertuples()]
            summary["處理狀態"] = [state_map.get(key, ("待確認", "", ""))[0] for key in company_keys]
            summary["負責人"] = [state_map.get(key, ("待確認", "", ""))[1] for key in company_keys]
            summary["下次檢查"] = [state_map.get(key, ("待確認", "", ""))[2] for key in company_keys]
            metric_columns = st.columns(4)
            metric_columns[0].metric("持續追蹤公司", len(summary))
            metric_columns[1].metric("跨日出現公司", int((summary["出現天數"] >= 2).sum()))
            metric_columns[2].metric("累計新聞", f"{len(tracked)} 篇")
            first_day, last_day = tracked["日期值"].min(), tracked["日期值"].max()
            date_range = f"{first_day:%Y-%m-%d}" if first_day == last_day else f"{first_day:%Y-%m-%d} ～ {last_day:%m-%d}"
            metric_columns[3].metric("資料日期", date_range)
            st.markdown("### 追蹤公司清單")
            st.caption("點選任一公司列，下方會顯示事件明細、人工處理狀態與事件後股價。")
            table_event = st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                key="tracking_company_table",
                column_config={"最新新聞": st.column_config.LinkColumn("最新新聞", display_text="開啟新聞")},
            )
            selected_rows = table_event.selection.rows
            action_columns = st.columns([2, 1])
            company_options = [f"{row.Ticker}｜{row.Company}" for row in summary.itertuples()]
            quick_company = action_columns[0].selectbox("快速查看公司詳情", ["請選擇"] + company_options)
            action_columns[1].download_button(
                "下載目前追蹤名單 CSV",
                summary.to_csv(index=False).encode("utf-8-sig"),
                f"美股持續追蹤名單_{last_day:%Y%m%d}.csv",
                mime="text/csv",
                use_container_width=True,
            )
            selected_company = None
            if selected_rows:
                selected_company = summary.iloc[selected_rows[0]]
            elif quick_company != "請選擇":
                quick_ticker, quick_name = quick_company.split("｜", 1)
                selected_company = summary[(summary["Ticker"].astype(str) == quick_ticker) & (summary["Company"].astype(str) == quick_name)].iloc[0]
            if selected_company is None:
                st.info("請點選表格中的公司列，或使用「快速查看公司詳情」。")
                st.stop()
            ticker, company = str(selected_company["Ticker"]), str(selected_company["Company"])
            company_news = tracked[(tracked["Ticker"].astype(str) == ticker) & (tracked["Company"].astype(str) == company)]
            st.markdown(f"### {ticker}｜{company}")
            detail_columns = st.columns(4)
            detail_columns[0].metric("首次列入", selected_company["首次列入"])
            detail_columns[1].metric("最近列入", selected_company["最近列入"])
            detail_columns[2].metric("出現天數", int(selected_company["出現天數"]))
            detail_columns[3].metric("相關新聞", int(selected_company["新聞篇數"]))
            st.dataframe(company_news[["資料日期", "Event_type", "Event_Code", "Level", "Action", "Keyword", "Title", "URL"]], use_container_width=True, hide_index=True, column_config={"URL": st.column_config.LinkColumn("新聞", display_text="開啟")})
        st.markdown("### 更新處理狀態")
        with sqlite3.connect(DB_PATH) as connection:
            current = connection.execute("SELECT status,owner,next_review,note FROM manual_status WHERE company_key=?", (ticker,)).fetchone() or ("待確認", "", "", "")
        status_options = ["待確認", "追蹤中", "已通知", "已完成", "解除追蹤"]
        manual_columns = st.columns(3)
        status = manual_columns[0].selectbox("狀態", status_options, index=status_options.index(current[0]) if current[0] in status_options else 0, key=f"status_{ticker}")
        owner = manual_columns[1].text_input("負責人", value=current[1], key=f"owner_{ticker}")
        next_review = manual_columns[2].text_input("下次檢查日期", value=current[2], placeholder="YYYY-MM-DD", key=f"review_{ticker}")
        note = st.text_area("備註", value=current[3], key=f"note_{ticker}")
        if st.button("儲存處理狀態", type="primary", use_container_width=True):
            save_status(ticker, company, status, owner, next_review, note)
            st.success("已儲存；重新分類不會覆蓋人工紀錄。")
        st.markdown("### 事件後股價變化")
        event_day = datetime.strptime(company_news["資料日期"].min(), "%Y%m%d").date()
        try:
            prices = stock_prices(ticker, (event_day - timedelta(days=10)).isoformat(), (date.today() + timedelta(days=15)).isoformat())
            st.line_chart(prices.set_index("Date")["Close"])
            later = prices[prices["Date"] >= event_day].reset_index(drop=True)
            if not later.empty:
                base = later.iloc[0]["Close"]
                columns = st.columns(4)
                for column, horizon in zip(columns, (1, 3, 5, 10)):
                    value = None if len(later) <= horizon else (later.iloc[horizon]["Close"] / base - 1) * 100
                    column.metric(f"{horizon} 交易日", "尚無資料" if value is None else f"{value:+.2f}%")
        except Exception as exc:
            st.warning(f"股價暫時無法取得：{exc}")

if page == "設定與說明":
    page_heading(page)
    st.markdown("<div class='method-card'><div class='method-title'>全新獨立系統</div><div class='method-desc'>依交接規格從空白建立，不讀取舊 V3 程式或舊執行資料。下方可管理事件規則；一般每日操作不需要進入此頁。</div></div>", unsafe_allow_html=True)
    st.markdown("### 負面新聞評分關鍵字版本")
    try:
        current_book = load_workbook(RULE_PATH, read_only=True, data_only=True)
        current_rule_count = (
            sum(1 for row in current_book["Event"].iter_rows(min_row=2, values_only=True) if any(value not in (None, "") for value in row))
            if "Event" in current_book.sheetnames else 0
        )
        current_book.close()
    except Exception:
        current_rule_count = 0
    with sqlite3.connect(DB_PATH) as connection:
        saved_rule_name = connection.execute("SELECT value FROM settings WHERE key='rule_original_name'").fetchone()
        saved_rule_time = connection.execute("SELECT value FROM settings WHERE key='rule_updated_at'").fetchone()
    current_rule_name = saved_rule_name[0] if saved_rule_name else RULE_PATH.name
    current_rule_time = saved_rule_time[0] if saved_rule_time else datetime.fromtimestamp(RULE_PATH.stat().st_mtime, TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
    info_columns = st.columns(3)
    info_columns[0].metric("目前版本", current_rule_name)
    info_columns[1].metric("規則筆數", f"{current_rule_count:,} 筆")
    info_columns[2].metric("套用時間", current_rule_time)
    st.info("目前版本會用於之後所有新聞分類；只有再次上傳並套用新版時才會更換。")
    rule_choice = st.radio(
        "本次要使用的關鍵字版本",
        ["使用目前版本", "自行上傳最新版"],
        index=0,
        horizontal=True,
    )
    st.download_button("下載目前關鍵字 Excel", RULE_PATH.read_bytes(), current_rule_name, use_container_width=True)
    replacement = st.file_uploader(
        "上傳最新版關鍵字 Excel",
        type=["xlsx"],
        key="rule_upload",
        disabled=rule_choice != "自行上傳最新版",
    )
    if rule_choice == "使用目前版本":
        st.success(f"目前持續使用：{current_rule_name}")
    if st.button(
        "檢查並永久套用新版",
        type="primary",
        disabled=rule_choice != "自行上傳最新版" or not replacement,
        use_container_width=True,
    ):
        try:
            test_path = UPLOAD_DIR / "Parameter_Event_pending.xlsx"
            test_path.write_bytes(replacement.getvalue())
            workbook = load_workbook(test_path, read_only=True, data_only=True)
            if "Event" not in workbook.sheetnames:
                raise ValueError("缺少 Event 工作表")
            if "主要網站" not in workbook.sheetnames:
                raise ValueError("缺少「主要網站」工作表")
            headers = [str(cell.value or "").strip() for cell in workbook["Event"][1]]
            missing_headers = [name for name in ("Event_Code", "Keyword", "Level", "Action") if name not in headers]
            if missing_headers:
                raise ValueError(f"Event 工作表缺少必要欄位：{'、'.join(missing_headers)}")
            new_rule_count = sum(
                1 for row in workbook["Event"].iter_rows(min_row=2, values_only=True)
                if any(value not in (None, "") for value in row)
            )
            if new_rule_count == 0:
                raise ValueError("Event 工作表沒有規則資料")
            source_headers = [str(cell.value or "").strip() for cell in workbook["主要網站"][1]]
            if "Source" not in source_headers:
                raise ValueError("「主要網站」工作表缺少 Source 欄位")
            event_code_index = headers.index("Event_Code")
            event_codes = {
                str(row[event_code_index] or "").strip()
                for row in workbook["Event"].iter_rows(min_row=2, values_only=True)
            }
            required_system_codes = {"UNKNOWN", "NOT_RELEVANT", "NON_CORE_SOURCE"}
            missing_system_codes = sorted(required_system_codes - event_codes)
            if missing_system_codes:
                raise ValueError(f"Event 工作表缺少系統事件：{'、'.join(missing_system_codes)}")
            workbook.close()
            backup_dir = DATA_ROOT / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_stamp = datetime.now(TAIPEI).strftime("%Y%m%d_%H%M%S")
            (backup_dir / f"Parameter_Event_{backup_stamp}.xlsx").write_bytes(RULE_PATH.read_bytes())
            RULE_PATH.write_bytes(test_path.read_bytes())
            applied_at = datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M:%S")
            with sqlite3.connect(DB_PATH) as connection:
                connection.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('rule_original_name',?)", (replacement.name,))
                connection.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('rule_updated_at',?)", (applied_at,))
            load_rules.clear()
            load_core_sources.clear()
            st.success(f"新版已套用：{replacement.name}，共 {new_rule_count:,} 筆規則。之後會持續使用此版本，直到下次更新。")
            st.rerun()
        except Exception as exc:
            st.error(f"新版關鍵字檔無法套用：{exc}")
    st.divider()
    st.markdown("### 版本恢復")
    if BACKUP_PATH.is_file():
        st.caption("已保存這次版面與流程改進前的完整主程式；若不喜歡新版，可恢復到改版前狀態。新聞資料與人工紀錄不會刪除。")
        confirm_restore = st.checkbox("我確認要恢復成版面流程改進前版本")
        if st.button("恢復改版前版本", disabled=not confirm_restore):
            try:
                current_path = Path(__file__).resolve()
                safety_copy = BACKUP_PATH.parent / "美股負面新聞系統_恢復前自動備份.py"
                safety_copy.write_bytes(current_path.read_bytes())
                current_path.write_bytes(BACKUP_PATH.read_bytes())
                st.success("已恢復改版前版本，網站將重新載入。")
                st.rerun()
            except Exception as exc:
                st.error(f"版本恢復失敗：{exc}")
    else:
        st.warning("找不到改版前備份，為安全起見不提供恢復按鈕。")
