#!/usr/bin/env python3
"""구글 플레이 주문 내역 JSON -> 결제내역 엑셀 변환기.

사용법:  python convert.py Order_History.json [-o 결과.xlsx]
경로를 생략하면 파일 선택창(없으면 터미널 입력)으로 고릅니다.

필요한 항목(날짜/항목/금액/결제수단)만 골라 담는 방식이라,
구글이 개인정보 칸을 새로 추가해도 결과물에 섞이지 않습니다.
"""
import argparse
import json
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

KST = ZoneInfo("Asia/Seoul")

# 통화 기호 -> 통화 코드 (긴 기호가 먼저 와야 "US$"가 "$"보다 먼저 매칭됨)
CURRENCY_SYMBOLS = {
    "JP¥": "JPY", "US$": "USD", "CA$": "CAD", "A$": "AUD", "HK$": "HKD", "NT$": "TWD",
    "₩": "KRW", "$": "USD", "¥": "JPY", "￥": "JPY", "€": "EUR", "£": "GBP",
}
CURRENCY_CODES = {"KRW", "USD", "JPY", "EUR", "GBP", "CAD", "AUD", "HKD", "TWD"}
# 소수점 없는 통화
ZERO_DECIMAL = {"KRW", "JPY", "TWD"}

# 결제수단: 원문(카드 끝자리·이름·전화번호·이메일이 섞여 있음)은 저장하지도 출력하지도 않고,
# 아래 규칙에 맞는 "종류 이름"만 꺼냅니다. 규칙에 없으면 "기타"로 두고 경고합니다.
METHOD_RULES = [
    (re.compile(r"^\s*toss", re.I), "토스"),
    (re.compile(r"^\s*신한"), "신한"),
    (re.compile(r"^\s*비씨"), "비씨"),
    (re.compile(r"^\s*visa", re.I), "Visa"),
    (re.compile(r"^\s*unionpay", re.I), "UnionPay"),
    (re.compile(r"^\s*(kt\b|korea telecom)", re.I), "KT 휴대폰"),
    (re.compile(r"^\s*naver\s*pay", re.I), "NAVER Pay"),
    (re.compile(r"^\s*google play 잔액", re.I), "Google Play 잔액"),
]

# 앱/게임 이름 통일 (표기가 둘 이상인 것만). 형식: "항목명 속 표기": "통일할 이름"
APP_ALIASES = {
    "소녀전선 Girls' Frontline": "소녀전선",
    "Crusaders Quest": "크루세이더 퀘스트",
    "명조:워더링 웨이브 × 사이버펑크 콜라보": "명조:워더링 웨이브",
    "KakaoTalk: Free Calls & Text": "카카오톡",
    "KakaoTalk : Messenger": "카카오톡",
    "원신-1주년": "원신",
    "캐치잇 잉글리시-Catch It English": "캐치잇 잉글리시",
}
# 항목명 모양으로 묶는 규칙: (패턴, 이름)
APP_PATTERNS = [(re.compile(r"상당 쿠폰$"), "Google Play 쿠폰")]

# 결과물에 남아 있으면 안 되는 개인정보 패턴
PII_PATTERNS = {
    "이메일": re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),
    "긴 숫자(7자리 이상)": re.compile(r"\d{7,}"),
    "전화번호": re.compile(r"\+?\d{2,3}[-\s]\d{3,4}[-\s]\d{4}"),
}


class ConvertError(Exception):
    """사용자에게 이유를 그대로 보여줄 오류."""


# ---------------------------------------------------------------- 읽기
def load_orders(path: Path) -> list:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ConvertError(f"파일을 찾을 수 없습니다: {path}")
    except UnicodeDecodeError:
        raise ConvertError("UTF-8 텍스트 파일이 아닙니다. 구글에서 받은 JSON이 맞는지 확인하세요.")
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ConvertError(f"JSON 형식이 아닙니다 ({e.lineno}번째 줄 {e.colno}번째 칸: {e.msg}).")
    if isinstance(data, dict):  # 주문이 하나뿐이거나 한 겹 감싼 경우
        data = data.get("orders") or [data]
    if not isinstance(data, list) or not data:
        raise ConvertError("주문 목록을 찾지 못했습니다. 비어 있거나 예상과 다른 구조입니다.")
    return data


# ---------------------------------------------------------------- 변환
def parse_time(value: str) -> datetime:
    """UTC 문자열 -> 한국시간(시간대 정보 없는 datetime, 엑셀 저장용)."""
    s = value.strip().replace("Z", "+00:00")
    s = re.sub(r"(\.\d{6})\d+", r"\1", s)  # 나노초 등 6자리 초과 소수부 자르기
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(KST).replace(tzinfo=None)


def parse_price(text, warnings: list, ctx: str):
    """'₩29,500' -> ('KRW', 29500). 해석 못 하면 (None, None)."""
    if text is None or str(text).strip() == "":
        return None, None
    s = str(text).strip()
    neg = "-" in s or "−" in s or s.startswith("(")
    cur = None
    for sym, code in CURRENCY_SYMBOLS.items():
        if sym in s:
            cur = code
            s = s.replace(sym, "")
            break
    else:
        m = re.search(r"\b([A-Z]{3})\b", s)
        if m:
            cur = m.group(1)
            s = s.replace(cur, "")
    num = re.sub(r"[^\d.,]", "", s)  # 부호는 위에서 이미 판정
    if not num or cur is None:
        warnings.append(f"금액을 해석하지 못했습니다 ({ctx}): {text!r}")
        return None, None
    if cur not in CURRENCY_CODES:
        warnings.append(f"처음 보는 통화입니다 ({ctx}): {cur}")
    num = num.replace(",", "")
    try:
        value = float(num)
    except ValueError:
        warnings.append(f"금액을 해석하지 못했습니다 ({ctx}): {text!r}")
        return None, None
    if cur in ZERO_DECIMAL:
        value = int(round(value))
    return cur, -value if neg else value


def classify_method(raw, idx: int, warnings: list) -> str:
    """결제수단 원문에서 종류만 판별한다. 원문은 어디에도 남기지 않는다."""
    if not isinstance(raw, str) or not raw.strip():
        warnings.append(f"결제수단을 찾지 못했습니다 ({idx}번째 주문).")
        return "기타"
    for pat, label in METHOD_RULES:
        if pat.search(raw):
            return label
    warnings.append(f"처음 보는 결제수단이라 '기타'로 분류했습니다 ({idx}번째 주문). "
                    "개인정보 보호를 위해 원문은 표시하지 않습니다. 필요하면 METHOD_RULES에 규칙을 추가하세요.")
    return "기타"


def app_name(title: str, doc_type) -> str:
    """'상품명 (앱·게임 이름 - 부가설명)' 에서 앱/게임 이름을 꺼낸다."""
    t = title.strip()
    if t.endswith(")"):
        depth = 0
        for i in range(len(t) - 1, -1, -1):          # 끝의 괄호와 짝이 맞는 여는 괄호 찾기
            depth += (t[i] == ")") - (t[i] == "(")
            if depth == 0:
                name = t[i + 1:-1].split(" - ")[0].strip()
                if name:
                    return APP_ALIASES.get(name, name)
                break
    if t.startswith("Google Play 잔액"):
        return "Google Play 잔액 충전"
    for pat, label in APP_PATTERNS:
        if pat.search(t):
            return label
    return APP_ALIASES.get(t, t)                    # 괄호가 없으면 항목명 자체가 앱/구독/도서 이름


def first(d: dict, *keys):
    for k in keys:
        if d.get(k) not in (None, ""):
            return d[k]
    return None


def extract(entry: dict, idx: int, warnings: list):
    """주문 하나에서 필요한 칸만 꺼낸다. 못 만들면 None."""
    o = entry.get("orderHistory", entry) if isinstance(entry, dict) else None
    if not isinstance(o, dict):
        warnings.append(f"{idx}번째 항목이 예상한 구조가 아닙니다.")
        return None
    ctx = f"{idx}번째 주문"

    t = first(o, "creationTime", "orderTime", "time")
    try:
        when = parse_time(t)
    except (TypeError, ValueError, AttributeError):
        warnings.append(f"날짜를 해석하지 못해 건너뜁니다 ({ctx}): {t!r}")
        return None

    cur, amount = parse_price(first(o, "totalPrice", "price"), warnings, ctx)
    if amount is None:
        warnings.append(f"금액이 없어 건너뜁니다 ({ctx}).")
        return None

    titles, apps = [], []
    for li in o.get("lineItem") or []:
        doc = li.get("doc") if isinstance(li, dict) else None
        title = doc.get("title") if isinstance(doc, dict) else None
        if title:
            if title not in titles:
                titles.append(title)
            a = app_name(title, doc.get("documentType"))
            if a not in apps:
                apps.append(a)
    if not titles:
        warnings.append(f"항목명이 없습니다 ({ctx}).")

    bi = o.get("billingInstrument")
    method = classify_method(bi.get("displayName") if isinstance(bi, dict) else None, idx, warnings)

    rcur, refund = parse_price(o.get("refundAmount"), warnings, ctx)
    if refund and rcur != cur:
        warnings.append(f"환불 통화가 결제 통화와 다릅니다 ({ctx}).")
    return {"date": when, "item": ", ".join(titles), "app": ", ".join(apps) or "(이름 없음)",
            "currency": cur, "amount": amount, "refund": refund or 0, "method": method}


# ---------------------------------------------------------------- 엑셀
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
BOLD = Font(bold=True)


def money_fmt(cur):
    return "#,##0" if cur in ZERO_DECIMAL else "#,##0.00"


def write_excel(rows: list, out: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "결제내역"

    currencies = sorted({r["currency"] for r in rows})
    header_row = 3 + len(currencies) + 1          # 합계 블록 아래 한 줄 띄우고 표 시작
    first_data, last_data = header_row + 1, header_row + len(rows)
    # 열: A 날짜 / B 항목 / C 앱·게임 / D 통화 / E 결제금액 / F 환불금액 / G 결제수단
    rng = lambda col: f"${col}${first_data}:${col}${last_data}"

    ws["A1"] = "통화별 합계"
    ws["A1"].font = Font(bold=True, size=13)
    for col, name in ((2, "결제"), (3, "환불"), (4, "순액(결제-환불)")):
        ws.cell(1, col, name).font = BOLD
    for i, cur in enumerate(currencies):
        r = 2 + i
        ws.cell(r, 1, cur).font = BOLD
        ws.cell(r, 2, f"=SUMIF({rng('D')},A{r},{rng('E')})")
        ws.cell(r, 3, f"=SUMIF({rng('D')},A{r},{rng('F')})")
        ws.cell(r, 4, f"=B{r}-C{r}")
        for col in (2, 3, 4):
            ws.cell(r, col).number_format = money_fmt(cur)
            ws.cell(r, col).font = BOLD

    for col, name in enumerate(["날짜(KST)", "항목", "앱/게임", "통화", "결제금액", "환불금액", "결제수단"], 1):
        c = ws.cell(header_row, col, name)
        c.font, c.fill = BOLD, HEAD_FILL
    for r, row in enumerate(rows, first_data):
        ws.cell(r, 1, row["date"]).number_format = "yyyy-mm-dd hh:mm"
        ws.cell(r, 2, row["item"])
        ws.cell(r, 3, row["app"])
        ws.cell(r, 4, row["currency"])
        ws.cell(r, 5, row["amount"]).number_format = money_fmt(row["currency"])
        ws.cell(r, 6, row["refund"]).number_format = money_fmt(row["currency"])
        ws.cell(r, 7, row["method"])
    ws.auto_filter.ref = f"A{header_row}:G{last_data}"
    ws.freeze_panes = ws.cell(first_data, 1)
    for col, width in enumerate([18, 50, 26, 8, 14, 14, 16], 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # 앱/게임 x 통화별 합계 (표의 SUMIFS 수식이라 원본 표를 고치면 같이 바뀜)
    pairs = {}
    for r in rows:
        k = (r["app"], r["currency"])
        pairs[k] = pairs.get(k, 0) + r["amount"] - r["refund"]
    order = sorted(pairs, key=lambda k: (k[1], -pairs[k], k[0]))
    ap = wb.create_sheet("앱별")
    for col, name in enumerate(["앱/게임", "통화", "건수", "결제", "환불", "순액"], 1):
        c = ap.cell(1, col, name)
        c.font, c.fill = BOLD, HEAD_FILL
    for r, (app, cur) in enumerate(order, 2):
        crit = f"'결제내역'!{rng('C')},$A{r},'결제내역'!{rng('D')},$B{r}"
        ap.cell(r, 1, app)
        ap.cell(r, 2, cur)
        ap.cell(r, 3, f"=COUNTIFS({crit})")
        ap.cell(r, 4, f"=SUMIFS('결제내역'!{rng('E')},{crit})")
        ap.cell(r, 5, f"=SUMIFS('결제내역'!{rng('F')},{crit})")
        ap.cell(r, 6, f"=D{r}-E{r}")
        for col in (4, 5, 6):
            ap.cell(r, col).number_format = money_fmt(cur)
    ap.auto_filter.ref = f"A1:F{len(order) + 1}"
    ap.freeze_panes = "A2"
    for col, width in enumerate([40, 8, 8, 14, 14, 14], 1):
        ap.column_dimensions[get_column_letter(col)].width = width

    info = wb.create_sheet("안내")
    lines = [
        "이 파일은 구글 플레이 주문 내역에서 결제 정보만 골라 만든 것입니다.",
        "",
        "가져온 항목: 주문 시각(한국시간), 항목명, 앱/게임 이름, 통화, 결제금액, 환불금액, 결제수단 종류",
        "제거한 정보: 이름, 주소, 전화번호, 이메일, IP, 주문번호, 카드 끝자리 등 나머지 전부",
        "결제수단은 원문(카드번호 끝자리·이름·전화번호 등이 섞여 있음)을 읽어 저장하지 않고 종류만 분류합니다.",
        "",
        "주의사항",
        "- 앱/게임 이름은 항목명 끝의 괄호 안 이름에서 뽑았습니다. 같은 앱의 다른 표기는 APP_ALIASES로 합칩니다.",
        "- 금액은 숫자로 저장되어 합계·필터를 바로 쓸 수 있습니다.",
        "- 합계와 앱별 시트는 수식이라 표를 고치면 함께 바뀝니다.",
        "- 환불된 주문은 결제금액에 그대로 남고 환불금액 열에 따로 표시됩니다. 순액 = 결제 - 환불.",
        "- 원본 JSON에는 개인정보가 들어 있으니 따로 보관하고 공유하지 마세요.",
        "- 구글이 내보내기 형식을 바꾸면 변환이 맞지 않을 수 있습니다. 경고 메시지를 확인하세요.",
    ]
    for i, line in enumerate(lines, 1):
        info.cell(i, 1, line)
    info["A1"].font = BOLD
    info["A7"].font = BOLD
    info.column_dimensions["A"].width = 90
    wb.save(out)


# ---------------------------------------------------------------- 검증
def verify(out: Path, rows: list, raw_orders: list, warnings: list) -> bool:
    ok = True
    wb = load_workbook(out)

    # 1) 개인정보 패턴 검사 (날짜 셀은 숫자 시각이라 문자열 셀만 검사)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    for name, pat in PII_PATTERNS.items():
                        if pat.search(c.value):
                            print(f"[검증 실패] {ws.title}!{c.coordinate}에 {name} 패턴이 있습니다.")
                            ok = False

    # 2) 원본에서 따로 계산한 합계 vs 엑셀 표 합계
    raw = Counter()
    for e in raw_orders:
        o = e.get("orderHistory", e) if isinstance(e, dict) else {}
        txt = str(first(o, "totalPrice", "price") or "")
        m = re.search(r"[\d,]+(?:\.\d+)?", txt)
        if not m:
            continue
        cur = next((c for s, c in CURRENCY_SYMBOLS.items() if s in txt), None) \
            or (re.search(r"\b[A-Z]{3}\b", txt) or [None])[0]
        raw[cur] += float(m.group().replace(",", "")) * (-1 if ("-" in txt or "−" in txt) else 1)

    sheet = Counter()
    ws = wb["결제내역"]
    for r in ws.iter_rows(values_only=True):
        if isinstance(r[4], (int, float)) and isinstance(r[3], str) and len(r[3]) == 3:
            sheet[r[3]] += r[4]
    for cur in sorted(set(raw) | set(sheet)):
        if abs(raw[cur] - sheet[cur]) > 0.005:
            print(f"[검증 실패] {cur} 합계 불일치: 원본 {raw[cur]:,.2f} / 엑셀 {sheet[cur]:,.2f}")
            ok = False
    if ok:
        print("[검증 통과] 개인정보 패턴 없음, 원본 합계와 엑셀 합계 일치")
    return ok


# ---------------------------------------------------------------- 실행
def pick_file() -> Path:
    try:
        import tkinter
        from tkinter import filedialog
        root = tkinter.Tk()
        root.withdraw()
        p = filedialog.askopenfilename(title="주문 내역 JSON 선택", filetypes=[("JSON", "*.json")])
        root.destroy()
        if p:
            return Path(p)
        raise ConvertError("파일을 선택하지 않았습니다.")
    except ImportError:
        p = input("JSON 파일 경로: ").strip().strip('"')
        if not p:
            raise ConvertError("경로를 입력하지 않았습니다.")
        return Path(p)
    except tkinter.TclError:  # 화면이 없는 환경
        p = input("JSON 파일 경로: ").strip().strip('"')
        if not p:
            raise ConvertError("경로를 입력하지 않았습니다.")
        return Path(p)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="구글 플레이 주문 내역 JSON -> 결제내역 엑셀")
    ap.add_argument("json_path", nargs="?", help="주문 내역 JSON 경로 (생략하면 선택창)")
    ap.add_argument("-o", "--output", help="저장할 엑셀 경로 (기본: 입력파일명_결제내역.xlsx)")
    args = ap.parse_args(argv)

    try:
        src = Path(args.json_path) if args.json_path else pick_file()
        orders = load_orders(src)
        warnings: list = []
        rows = [r for i, e in enumerate(orders, 1) if (r := extract(e, i, warnings))]
        if not rows:
            raise ConvertError("변환할 수 있는 주문이 하나도 없습니다. 아래 경고를 확인하세요.\n  " + "\n  ".join(warnings[:10]))
        rows.sort(key=lambda r: r["date"])
        out = Path(args.output) if args.output else src.with_name(src.stem + "_결제내역.xlsx")
        write_excel(rows, out)
    except ConvertError as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1
    except PermissionError as e:
        print(f"오류: 파일을 쓸 수 없습니다. 엑셀에서 열려 있지 않은지 확인하세요 ({e.filename}).", file=sys.stderr)
        return 1

    print(f"{len(rows)}/{len(orders)}건 변환 -> {out}")
    for cur, total in sorted(Counter(r["currency"] for r in rows).items()):
        s = sum(r["amount"] for r in rows if r["currency"] == cur)
        print(f"  {cur}: {s:,.2f} ({total}건)")
    seen = list(dict.fromkeys(warnings))  # 중복 제거, 순서 유지
    for w in seen:
        print(f"[경고] {w}")
    ok = verify(out, rows, orders, warnings)
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
