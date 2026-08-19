import openpyxl
import sys
import os
import datetime
from app import parse_kakao_text, generate_order_excel, append_to_erp_upload_list

sys.stdout.reconfigure(encoding='utf-8')

print("="*70)
print("1. KAKAO PARSE TEST WITH GREETINGS & CASUAL TEXT")
print("="*70)

kakao_with_greetings = """안녕하세요 발주하려고 합니다 ^^
사장님 오늘도 수고 많으십니다!

1. 납품처 
   매장명 : 스트릿캔 구리인창점
   주소 : 경기 구리시 인창동 672-2
   현장담당 이름, 전화번호 : 송재호 010-9388-9989
   납품일 및 희망 시간 : 8월 14일

2. 모델명과 수량

스타리온 1200 반찬냉장고 1/3 앞작업대 / SRV12EIEVF
+ 1200 2단선반
스타리온 25BOX 냉장,냉동 / 직냉식 / SR-E25B1F
스타리온 1500 냉동냉장 / 직냉식 / SR-T15B1F
스타리온 1500 올냉동 / 직냉식 / SR-T15BAFC
음료아날로그 / SR-SC44RW - 2대

확인 부탁드립니다. 감사합니다!
"""

parsed_aug = parse_kakao_text(kakao_with_greetings)
print("매장명:", parsed_aug["store_name"])
print("배송요청일:", parsed_aug["delivery_date"])
print("추출된 품목 개수:", len(parsed_aug["items"]))

# 품명에 인사말이 들어갔는지 엄격 검증
for it in parsed_aug["items"]:
    print(f"  NO {it['no']}: [품명] {it['item_name']} | [모델명] {it['model_name']} | [수량] {it['qty']} | [매입] {it['buy_price']:,}원 | [수주] {it['sell_price']:,}원")
    assert "안녕하세요" not in it["item_name"], "인사말이 품명에 포함되었습니다!"
    assert "수고" not in it["item_name"], "잡담이 품명에 포함되었습니다!"
    assert "감사합니다" not in it["item_name"], "감사 문구가 품명에 포함되었습니다!"

assert len(parsed_aug["items"]) == 5, f"품목 개수 오류: {len(parsed_aug['items'])} (5개여야 함)"
print("\n[PASS] 인사말 필터링 완벽 통과! (순수 5개 장비 품목만 정확히 추출)")

print("\n" + "="*70)
print("2. AUGUST ORDER -> ERP TAB '26.8' GENERATION TEST")
print("="*70)

# 기존 파일 초기화 후 8월 데이터 기록
erp_path = append_to_erp_upload_list(parsed_aug, replace_all=True)
wb_erp = openpyxl.load_workbook(erp_path, data_only=True)
print("현재 엑셀 탭 목록:", wb_erp.sheetnames)
assert "26.8" in wb_erp.sheetnames, "26.8 탭이 생성되지 않았습니다."

ws_8 = wb_erp["26.8"]
print(f"26.8 탭 행 수: {ws_8.max_row}")
for r in range(1, ws_8.max_row + 1):
    vals = [ws_8.cell(r, c).value for c in range(1, 17)]
    print(f"  R{r:2d}: A={repr(vals[0])}, B={repr(vals[1])}, C(거래처)={repr(vals[2])}, D(거래처명)={repr(vals[3])}, E={repr(vals[4])}, F(품명)={repr(vals[5])}, G(모델)={repr(vals[6])}, H(수량)={vals[7]}, I(매입)={vals[8]}, K(수주)={vals[10]}, M(마진)={vals[12]}")
    if r >= 2:
        # C열(거래처), D열(거래처명)이 None(빈칸)인지 검증
        assert vals[2] is None, f"C열(거래처)이 빈칸이 아닙니다: {vals[2]}"
        assert vals[3] is None, f"D열(거래처명)이 빈칸이 아닙니다: {vals[3]}"

print("\n" + "="*70)
print("3. SEPTEMBER ORDER -> NEW ERP TAB '26.9' AUTOMATIC CREATION TEST")
print("="*70)

kakao_sep = """안녕하세요 사장님 주문서 보냅니다~

1. 납품처
   매장명 : 미쉐 서면점
   주소 : 부산광역시 부산진구 중앙대로 686
   현장담당 이름, 전화번호 : 김현수 010-5555-4321
   납품일 및 희망 시간 : 09월 15일 오후

2. 모델명과 수량
스타리온 1500 냉동냉장 / 직냉식 / SR-T15B1F - 1대
스타리온 25BOX 올냉동 / 직냉식 / SR-E25BAFC - 1대
음료아날로그 / SR-SC44RW - 2대
"""

parsed_sep = parse_kakao_text(kakao_sep)
erp_path_sep = append_to_erp_upload_list(parsed_sep, replace_all=False)

wb_erp_reloaded = openpyxl.load_workbook(erp_path_sep, data_only=True)
print("9월 추가 후 엑셀 탭 목록:", wb_erp_reloaded.sheetnames)
assert "26.8" in wb_erp_reloaded.sheetnames, "26.8 탭이 보존되어야 합니다."
assert "26.9" in wb_erp_reloaded.sheetnames, "26.9 탭이 새로 생성되어야 합니다."

ws_9 = wb_erp_reloaded["26.9"]
print(f"\n26.9 탭 행 수: {ws_9.max_row}")
for r in range(1, ws_9.max_row + 1):
    vals = [ws_9.cell(r, c).value for c in range(1, 17)]
    print(f"  R{r:2d}: A={repr(vals[0])}, B={repr(vals[1])}, C(거래처)={repr(vals[2])}, D(거래처명)={repr(vals[3])}, E={repr(vals[4])}, F(품명)={repr(vals[5])}, G(모델)={repr(vals[6])}, H(수량)={vals[7]}, I(매입)={vals[8]}, K(수주)={vals[10]}, M(마진)={vals[12]}")
    if r >= 2:
        assert vals[2] is None, "9월 탭 C열 거래처가 빈칸이어야 합니다."
        assert vals[3] is None, "9월 탭 D열 거래처명이 빈칸이어야 합니다."

print("\n>>> ALL TESTS PASSED SUCCESSFULLY! 100% VERIFIED <<<")
