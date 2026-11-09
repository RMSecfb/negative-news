"""Pre-filter before spending LLM allowance (set PREFILTER=0 to disable).

Skip only records with no risk-event keyword/price-drop pattern AND either clearly
positive FinBERT (>0.2) or routine phrasing (price target, dividend, holdings...).
Backtest 10/23-11/08 on LLM-analysed rows: skipped 10%, lost 5% of LLM-found events;
on queued rows it removes about 35% of the candidates. Skipped rows become
NOT_RELEVANT with 最終判斷來源='規則前置過濾' and can still be corrected by review.
"""
import re
I=re.I
R=[ # (code, regex on title+zh title+summary) ordered by priority
('BANKRUPTCY',r'\b(files? for|filed for|enters?|seeks?) (chapter 1[15]|bankruptcy)|bankruptcy (filing|protection)|申請破產|破產保護'),
('DELISTING',r'\bdelist(ed|ing)\b|下市'),
('ACCOUNTING_FRAUD',r'accounting (fraud|irregularit)|fraud(ulent)? (charge|scheme)|會計舞弊'),
('RESTATEMENT',r'\brestate(s|d|ment)\b.{0,30}(financial|earnings|results)|重編財報'),
('GOING_CONCERN',r'going concern|繼續經營'),
('GUIDANCE_WITHDRAWN',r'(withdraws?|suspends?|pulls?) (its |full-year |annual )?(guidance|outlook|forecast)|撤回財測'),
('EARNINGS_WARNING',r'(cuts?|lowers?|slashes?|reduces?|trims?) (its |full-year |annual |fy\d* )?(guidance|outlook|forecast|profit forecast|revenue forecast)|profit warning|warns? (of|on) (weak|lower|slower|profit|earnings|sales|revenue)|下修財測|財測下修|財測.{0,6}下調'),
('AUDITOR_RESIGNATION',r'auditor (resign|quit|step)|會計師辭'),
('ANTITRUST_INVESTIGATION',r'antitrust (probe|investigation|lawsuit|suit|case)|反壟斷調查|反壟斷.{0,4}(訴|調查)'),
('REGULATORY_PENALTY',r'\b(fined|fines|to pay|agrees? to pay|settles?)\b.{0,40}(sec|ftc|doj|cftc|finra|regulator|penalt|million|billion)|(sec|ftc|doj|finra|cfpb|occ)\b.{0,30}\b(fines?|penalt|charges?|settle)|civil penalt|被罰款|罰款'),
('REGULATORY_INVESTIGATION',r'\b(sec|ftc|doj|cfpb|fcc|eu commission|regulators?)\b.{0,30}\b(probe|investigat|subpoena|inquiry)|\b(probe|investigation|subpoena)\b.{0,25}\b(sec|ftc|doj|regulator)|(under|faces?) (federal |government |regulatory )?investigation|監管調查|被調查'),
('FDA_WARNING',r'fda (warning letter|issues? warning|rejects?|refuses|complete response)|complete response letter|FDA警告'),
('PRODUCT_RECALL',r'\brecalls?\b|\brecalled\b|產品召回|召回'),
('DATA_BREACH',r'data breach|breach(ed)? (of )?(customer|user|personal)|customer data (leak|exposed|stolen)|資料外洩|個資外洩'),
('CYBER_ATTACK',r'cyber ?attack|ransomware|hackers?\b|hacked|cyber ?incident|資安攻擊|駭客|勒索軟體'),
('WORKFORCE_REDUCTION',r'\blay ?offs?\b|laying off|lays off|laid off|job cuts?|cuts? (\d[\d,]* |\d+% of )?(jobs|workforce|staff|positions)|workforce reduction|redundanc(y|ies)|裁員'),
('FACTORY_SHUTDOWN',r'(plant|factory|facility|mill|refinery)\b.{0,30}\b(shut ?down|shutter|closure|close[sd]?|idle)|(shut(s|ting)? down|closes?|idles?) (a |its |the )?(plant|factory|facility)|工廠停工|停產'),
('CREDIT_DOWNGRADE',r'(moody\'?s|s&p global ratings|s&p|fitch)\b.{0,40}\b(downgrad|cuts? (rating|credit)|lowers?|negative outlook)|downgrade[sd]? (credit|debt|rating)|credit rating (cut|downgrad)|信用調降|調降.{0,6}信用評等'),
('CEO_DEPARTURE',r'\bceo\b.{0,30}\b(resign|steps? down|to step down|departs?|ousted|fired|exit|retire)|(resign|ousts?|fires?|replaces?)\b.{0,15}\bceo\b|執行長.{0,6}(辭|離職|請辭)'),
('CFO_DEPARTURE',r'\bcfo\b.{0,30}\b(resign|steps? down|to step down|departs?|ousted|exit)|(resign|ousts?)\b.{0,15}\bcfo\b|財務長.{0,6}(辭|離職)'),
('SHORT_SELLER_REPORT',r'short[- ]seller|short report|放空機構|hindenburg|muddy waters'),
('MAJOR_LITIGATION',r'class[- ]action|securities (fraud|litigation|lawsuit)|investors? (who )?(lost|suffered|have lost)|shareholder (lawsuit|suit)|\b(sued|sues|lawsuit|suit)\b.{0,40}(billion|million|\$\d)|(billion|million|\$\d)\b.{0,40}\b(lawsuit|verdict|judgment)|集體訴訟|集團訴訟|證券訴訟'),
('PATENT_LITIGATION',r'patent (infringement|lawsuit|suit|litigation|dispute)|專利訴訟|專利侵權'),
('TARIFF',r'\btariffs?\b.{0,40}(hit|impact|cost|hurt|weigh|pressure)|關稅.{0,10}(衝擊|影響|成本)'),
('EXPORT_RESTRICTION',r'export (control|restriction|ban|curb)|出口管制'),
('ANALYST_DOWNGRADE',r'\bdowngrad(e|es|ed|ing)\b|cuts? (price )?target|lowers? (price )?target|reduces? (price )?target|降評|調降.{0,6}(評級|目標價)'),
('EARNINGS_MISS',r'(miss(es|ed)?|falls? short of|below|lags?|trails?)\b.{0,25}\b(estimates?|expectations?|forecasts?|consensus|analysts?)|(earnings|revenue|sales|eps)\b.{0,15}\b(miss|disappoint|shortfall)|財報.{0,6}(不如預期|低於預期)|(營收|獲利).{0,8}(不如|低於)預期'),
('OPERATING_DETERIORATION',r'(profit|earnings|net income|operating income|margin)s?\b.{0,15}\b(plunge|plummet|tumble|collapse|sink|slump|crater)|net loss widens|loss(es)? widen|營運惡化'),
('SUPPLY_CHAIN_DISRUPTION',r'supply[- ]chain (disruption|issue|problem|shortage|crisis)|parts? shortage|供應鏈中斷'),
('PRODUCTION_DELAY',r'production (delay|halt|pause|cut)|delays? (production|launch|shipment)|生產延遲'),
('PRODUCT_DEFECT',r'product defect|safety (defect|issue)|重大瑕疵'),
('INSIDER_TRADING',r'insider trading|內線交易'),
('SANCTION',r'\bsanctions?\b.{0,30}(on|against|imposed|hit)|制裁'),
('EXECUTIVE_MISCONDUCT',r'misconduct|executive (charged|arrested|indicted)|indicted|高管不當'),
('MAJOR_CUSTOMER_LOSS',r'loses? (major |key |largest )?(customer|contract)|customer loss|重大客戶流失'),
]
RX=[(c,re.compile(p,I)) for c,p in R]
# routine / non-risk phrasing to veto weak codes
VETO_ANALYST=re.compile(r'\bupgrad|raises? (price )?target|buy rating|outperform|overweight|strong buy|reiterat|maintains? (buy|outperform)|initiat',I)
MOVE=re.compile(r'(moved|moves|down|drops?|dropped|falls?|fell|slides?|slid|sinks?|sank|plunges?|plunged|tumbles?|tumbled|slumps?|sinks?|declines?|declined|sheds?|lost|loses)\b[^%]{0,40}?(\d+(\.\d+)?)\s?%|(\d+(\.\d+)?)\s?%[^%]{0,25}\b(drop|fall|decline|slide|plunge|slump|tumble)',I)
PCT=re.compile(r'(\d+(?:\.\d+)?)\s?%')
ROUTINE=re.compile(r'stock (price )?(moved|up|down)|shares? (rise|rose|gain|jump|climb|surge)|dividend|price target|valuation|fair value|undervalued|overvalued|holdings|stake|13f|position in|shares? (bought|sold|acquired)|should you buy|buy or sell|is it a buy|what to know|here\'s why|why .* (stock|shares)|earnings preview|ahead of earnings|to report earnings|earnings (call|date|calendar)|upgrade',I)


REASON = '規則前置過濾：無風險事件字詞，且情緒偏正面或屬例行資訊'


def _text(record):
    return ' || '.join(str(record.get(k, '') or '') for k in ('Title', 'Title_ZH', 'Summary'))


# Wording that can signal trouble even when the legacy patterns and FinBERT do not.
GUARD = re.compile(r"probe|investigat|inquir|subpoena|\bspy\b|espionage|alleg|fraud|misconduct|scandal|sues?\b|sued|lawsuit|settle|"
                   r"resign|steps? down|fired|ousted|halt|suspend|outage|breach|hack|recall|\bcuts?\b|downgrad|slash|warn|"
                   r"\bcds\b|spreads?\b|\brout\b|credit risk|concern|weakness|plunge|tumble|slump|crash|layoff|delist|default|bankrupt|"
                   r"信用風險|調查|起訴|訴訟|罰款|和解|辭職|下架|召回|裁員|下調|警告|暴跌|違約|破產|停產|停牌", re.I)


def skip_llm(record):
    """Return a reason when the LLM call can be saved, else ''."""
    text = _text(record)
    if any(rx.search(text) for _, rx in RX) or MOVE.search(text) or GUARD.search(text):
        return ''
    try:
        score = float(record.get('FinBERT'))
    except (TypeError, ValueError):
        return ''
    routine = bool(ROUTINE.search(str(record.get('Title', '')) + ' ' + str(record.get('Title_ZH', ''))))
    return REASON if score > 0.2 or (routine and score > -0.3) else ''
