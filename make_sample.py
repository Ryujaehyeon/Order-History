#!/usr/bin/env python3
"""연습용 가짜 주문 내역 JSON 생성. 실제 개인정보는 들어 있지 않습니다.
사용법: python make_sample.py [sample_order_history.json]
"""
import json
import sys

PERSON = {"name": "홍길동", "email": "test.user@example.com",
          "phoneNumber": "010-1234-5678", "address": "서울시 가짜구 가짜로 1"}


def order(n, time, total, method, titles, **extra):
    return {"orderHistory": {
        "orderId": f"GPA.0000-0000-0000-{n:05d}", "creationTime": time,
        "associatedContact": [PERSON], "emailAddress": [PERSON["email"]], "ipAddress": "203.0.113.7",
        "billingInstrument": {"displayName": method}, "totalPrice": total, "tax": "₩0", "refundAmount": "₩0",
        "lineItem": [{"doc": {"documentType": "Android Apps", "title": t}, "quantity": 1} for t in titles],
        **extra}}


orders = [
    order(1, "2023-01-01T14:59:59.123456789Z", "₩29,500", "신한-1234", ["가짜 RPG 패키지"]),
    order(2, "2023-02-14T01:00:00Z", "₩1,100", "Toss: 홍길*", ["가짜 퍼즐 코인"]),
    order(3, "2023-03-03T12:30:00Z", "$4.99", "Google Play 잔액: ₩12,000", ["Fake Game Pass"], discount=""),
    order(4, "2023-05-20T09:00:00Z", "₩12,000", "Visa-5678", ["월정액", "추가 아이템"]),
    order(5, "2023-06-01T00:00:00Z", "₩1,100", "Toss: 홍길*", ["가짜 퍼즐 코인"], refundAmount="₩1,100"),
    order(6, "2023-07-07T07:07:07Z", "€2.49", "신한-1234", ["유로 앱"]),
    order(8, "2023-09-09T09:09:09Z", "JP¥500", "KT +821012345678", ["엔화 앱"]),
    order(7, "2023-08-08T08:08:08Z", "₩5,500", "새로운결제사-9999", ["새 앱"]),
]
path = sys.argv[1] if len(sys.argv) > 1 else "sample_order_history.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(orders, f, ensure_ascii=False, indent=2)
print(f"가짜 데이터 {len(orders)}건 -> {path}")
