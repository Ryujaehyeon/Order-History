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

# 이미 알고 있는 결제수단. 여기 없는 종류가 나오면 경고합니다. 필요하면 추가하세요.
KNOWN_METHODS = {"토스", "신한", "비씨", "Visa", "UnionPay", "KT 휴대폰", "NAVER Pay",
                 "Google Play 잔액", "Google Play 기프트 카드"}
# 표기 통일 (정리한 뒤의 이름 -> 최종 이름)
METHOD_ALIASES = {"Toss": "토스", "KT": "KT 휴대폰", "Korea Telecom KR": "KT 휴대폰",
                  "Korea Telecom": "KT 휴대폰"}

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


def clean_method(raw) -> str:
    """카드 끝자리 등을 버리고 결제수단 종류만 남긴다."""
    if not raw:
        return ""
    s = str(raw).split(":")[0]                          # "Toss: 유재*" -> "Toss" (콜론 뒤는 이름/이메일/잔액)
    s = re.sub(r"\+?\d[\d\- ]{6,}", "", s)             # 전화번호
    s = re.sub(r"[\(\[（].*?[\)\]）]", "", s)           # (1234), [끝자리 1234]
    s = re.sub(r"[-–·•*\s]*(끝자리|ending in|ending)?[\s:]*[*•x]*\d{2,}\s*$", "", s, flags=re.I)
    s = re.sub(r"[*•x]{2,}", "", s, flags=re.I)          # ••••, ****
    s = re.sub(r"\s+", " ", s).strip(" -–·")
    return METHOD_ALIASES.get(s, s)


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

    titles = []
    for li in o.get("lineItem") or []:
        doc = li.get("doc") if isinstance(li, dict) else None
        title = (doc or {}).get("title") if isinstance(doc, dict) else None
        if title and title not in titles:
            titles.append(title)
    if not titles:
        warnings.append(f"항목명이 없습니다 ({ctx}).")

    bi = o.get("billingInstrument")
    method = clean_method(first(o, "paymentMethodTitle", "paymentMethod")
                          or (bi.get("displayName") if isinstance(bi, dict) else None))
    if not method:
        warnings.append(f"결제수단을 찾지 못했습니다 ({ctx}).")
    elif method not in KNOWN_METHODS:
        warnings.append(f"처음 보는 결제수단입니다: {method!r} (KNOWN_METHODS에 추가하면 경고가 사라집니다)")

    rcur, refund = parse_price(o.get("refundAmount"), warnings, ctx)
    if refund and rcur != cur:
        warnings.append(f"환불 통화가 결제 통화와 다릅니다 ({ctx}).")
    return {"date": when, "item": ", ".join(titles), "currency": cur,
            "amount": amount, "refund": refund or 0, "method": method}


# ---------------------------------------------------------------- 엑셀
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
BOLD = Font(bold=True)


def write_excel(rows: list, out: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "결제내역"

    currencies = sorted({r["currency"] for r in rows})
    header_row = 3 + len(currencies) + 1          # 합계 블록 아래 한 줄 띄우고 표 시작
    first_data, last_data = header_row + 1, header_row + len(rows)

    ws["A1"] = "통화별 합계"
    ws["A1"].font = Font(bold=True, size=13)
    rng = lambda col: f"${col}${first_data}:${col}${last_data}"
    for i, cur in enumerate(currencies):
        r = 2 + i
        fmt = "#,##0" if cur in ZERO_DECIMAL else "#,##0.00"
        ws.cell(r, 1, cur).font = BOLD
        ws.cell(r, 2, f"=SUMIF({rng('C')},A{r},{rng('D')})")   # 결제
        ws.cell(r, 3, f"=SUMIF({rng('C')},A{r},{rng('E')})")   # 환불
        ws.cell(r, 4, f"=B{r}-C{r}")                           # 순액
        for col in (2, 3, 4):
            ws.cell(r, col).number_format = fmt
            ws.cell(r, col).font = BOLD
    ws.cell(1, 2, "결제").font = BOLD
    ws.cell(1, 3, "환불").font = BOLD
    ws.cell(1, 4, "순액(결제-환불)").font = BOLD

    for col, name in enumerate(["날짜(KST)", "항목", "통화", "결제금액", "환불금액", "결제수단"], 1):
        c = ws.cell(header_row, col, name)
        c.font, c.fill = BOLD, HEAD_FILL
    for r, row in enumerate(rows, first_data):
        fmt = "#,##0" if row["currency"] in ZERO_DECIMAL else "#,##0.00"
        ws.cell(r, 1, row["date"]).number_format = "yyyy-mm-dd hh:mm"
        ws.cell(r, 2, row["item"])
        ws.cell(r, 3, row["currency"])
        ws.cell(r, 4, row["amount"]).number_format = fmt
        ws.cell(r, 5, row["refund"]).number_format = fmt
        ws.cell(r, 6, row["method"])

    ws.auto_filter.ref = f"A{header_row}:F{last_data}"
    ws.freeze_panes = ws.cell(first_data, 1)
    for col, width in enumerate([18, 50, 8, 14, 14, 22], 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    info = wb.create_sheet("안내")
    lines = [
        "이 파일은 구글 플레이 주문 내역에서 결제 정보만 골라 만든 것입니다.",
        "",
        "가져온 항목: 주문 시각(한국시간), 항목명, 통화, 결제금액, 환불금액, 결제수단 종류",
        "제거한 정보: 이름, 주소, 전화번호, 이메일, IP, 주문번호, 카드 끝자리 등 나머지 전부",
        "",
        "주의사항",
        "- 금액은 숫자로 저장되어 합계·필터를 바로 쓸 수 있습니다.",
        "- 맨 위 합계는 수식(SUMIF)이라 필터를 걸어도 전체 합계가 유지됩니다.",
        "- 환불된 주문은 결제금액에 그대로 남고 환불금액 열에 따로 표시됩니다. 순액 = 결제 - 환불.",
        "- 원본 JSON에는 개인정보가 들어 있으니 따로 보관하고 공유하지 마세요.",
        "- 구글이 내보내기 형식을 바꾸면 변환이 맞지 않을 수 있습니다. 경고 메시지를 확인하세요.",
    ]
    for i, line in enumerate(lines, 1):
        info.cell(i, 1, line)
    info["A1"].font = BOLD
    info["A6"].font = BOLD
    info.column_dimensions["A"].width = 80
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
        if isinstance(r[3], (int, float)) and isinstance(r[2], str) and len(r[2]) == 3:
            sheet[r[2]] += r[3]
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
