# 구글 플레이 주문 내역 → 결제내역 엑셀

구글 Takeout의 주문 내역 JSON에서 **날짜·항목·금액·결제수단**만 골라 엑셀로 만듭니다.
이름, 주소, 전화번호, 이메일, 주문번호, 카드 끝자리는 가져오지 않습니다.

## 설치
```
pip install -r requirements.txt
```
Python 3.9 이상이 필요합니다.

## 사용
```
python convert.py Order_History.json            # → Order_History_결제내역.xlsx
python convert.py Order_History.json -o out.xlsx
python convert.py                               # 파일 선택창
```

연습용 가짜 데이터: `python make_sample.py` → `sample_order_history.json`

## 결과물
- **결제내역** 시트: 맨 위에 통화별 합계(SUMIF 수식), 아래에 날짜(KST)·항목·통화·결제금액·환불금액·결제수단 표. 합계는 통화별로 결제/환불/순액을 보여줍니다. 필터와 합계를 바로 쓸 수 있도록 금액은 숫자로 저장됩니다.
- **앱별** 시트: 앱/게임 × 통화별 건수·결제·환불·순액(수식). 앱 이름은 항목명 끝 괄호(`상품명 (앱 이름)`)에서 뽑고, 표기가 다른 같은 앱은 `convert.py`의 `APP_ALIASES`로 합칩니다.
- **안내** 시트: 제거한 정보와 주의사항.

## 자동 검증과 경고
실행이 끝나면 스스로 확인합니다.
- 결과물에 이메일·7자리 이상 숫자·전화번호 패턴이 남았는지 검사
- 원본에서 따로 계산한 통화별 합계와 엑셀 표 합계 비교
- 모르는 통화, 처음 보는 결제수단, 해석 못 한 날짜·금액은 `[경고]`로 출력

결제수단은 원문(카드 끝자리·이름·전화번호·이메일이 섞여 있음)을 저장하거나 출력하지 않고, `METHOD_RULES`에 맞는 종류만 꺼냅니다. 규칙에 없으면 `기타`로 두고 경고합니다(원문은 경고에도 표시하지 않음).
검증이 실패하면 종료 코드 2, 입력 오류는 1을 반환합니다.

## 주의
- 원본 JSON과 변환 결과는 개인정보가 섞일 수 있으니 공유하지 말고 따로 보관하세요. `.gitignore`가 `*.json`, `*.xlsx`를 막아 둡니다.
- 구글이 내보내기 형식을 바꾸면 맞지 않을 수 있습니다. 경고를 확인하세요.
- 읽는 키는 `creationTime`, `totalPrice`, `refundAmount`, `billingInstrument.displayName`(종류 판별에만 사용), `lineItem[].doc.title`입니다. 구조가 바뀌면 `extract()`의 키 이름만 고치면 됩니다.
