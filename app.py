import os
import json
import re
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Optional, Dict, Any

from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
ORDERS_DIR = os.path.join(OUTPUT_DIR, "orders")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(ORDERS_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="카카오톡 주문 원버튼 발주서 및 ERP 연동 시스템")

# ----------------- Helper Functions ----------------- #

def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "naverworks": {
            "smtp_server": "smtp.worksmobile.com",
            "smtp_port": 465,
            "use_ssl": True,
            "sender_email": "",
            "sender_password": "",
            "sender_name": "발주담당자"
        },
        "company_info": {
            "company_name": "(주)바이브컴퍼니",
            "sender_name": "발주팀 담당자",
            "phone": "010-1234-5678",
            "email": "order@vibecompany.com"
        },
        "default_mail_template": {
            "subject": "[발주서] {vendor_name} 귀하 - {date} 발주 건 ({item_summary})",
            "body": "안녕하세요, {vendor_name} 담당자님.\n\n발주서 첨부하여 전달드립니다.\n확인 후 납기에 맞춰 진행 부탁드립니다.\n\n감사합니다."
        },
        "vendors": []
    }

def save_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def get_price_master() -> Dict[str, dict]:
    """price_master.xlsx에서 품번별/품명별 단가 정보를 읽어옵니다."""
    master_path = os.path.join(TEMPLATES_DIR, "price_master.xlsx")
    master = {}
    if not os.path.exists(master_path):
        return master
        
    wb = openpyxl.load_workbook(master_path, data_only=True)
    ws = wb.active
    
    # Header: 품번(1), 품명(2), 규격(3), 매입가(4), 수주가(5), 비고(6)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        item_code = str(row[0]).strip()
        item_name = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        spec = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        
        try:
            buy_price = int(float(str(row[3]).replace(",", ""))) if len(row) > 3 and row[3] is not None else 0
        except (ValueError, TypeError):
            buy_price = 0
            
        try:
            sell_price = int(float(str(row[4]).replace(",", ""))) if len(row) > 4 and row[4] is not None else 0
        except (ValueError, TypeError):
            sell_price = 0
            
        remark = str(row[5]).strip() if len(row) > 5 and row[5] else ""
        
        item_data = {
            "item_code": item_code,
            "item_name": item_name,
            "spec": spec,
            "buy_price": buy_price,
            "sell_price": sell_price,
            "remark": remark
        }
        master[item_code.upper()] = item_data
        if item_name:
            master[item_name.upper()] = item_data
            
    return master

# ----------------- Parser Logic ----------------- #

def parse_kakao_text(text: str) -> dict:
    """카카오톡 주문 메시지를 파싱하여 구조화된 데이터를 반환합니다."""
    lines = [line.strip() for line in text.strip().split("\n") if line.strip()]
    price_master = get_price_master()
    config = load_config()
    
    result = {
        "vendor_name": "",
        "recipient_name": "",
        "recipient_phone": "",
        "recipient_address": "",
        "order_date": datetime.date.today().strftime("%Y-%m-%d"),
        "memo": "",
        "items": []
    }
    
    # 1. 거래처 찾기 (config의 vendors 이름과 대조 또는 키워드 추출)
    vendor_matched = False
    for v in config.get("vendors", []):
        if v["name"] in text:
            result["vendor_name"] = v["name"]
            vendor_matched = True
            break
            
    if not vendor_matched:
        vendor_match = re.search(r'(?:업체|거래처|공급처|수신|상호)\s*[:：\-]?\s*([가-힣a-zA-Z0-9_㈜()（）\s]+)', text)
        if vendor_match:
            v_name = vendor_match.group(1).strip()
            # 줄바꿈이나 다른 태그 전까지만
            v_name = v_name.split()[0] if v_name else ""
            result["vendor_name"] = v_name
            
    # 2. 연락처(전화번호) 추출
    phone_match = re.search(r'(01[016789]-?\d{3,4}-?\d{4}|02-?\d{3,4}-?\d{4}|0\d{2}-?\d{3,4}-?\d{4})', text)
    if phone_match:
        result["recipient_phone"] = phone_match.group(1)
        
    # 3. 받는사람/수령인 추출
    recipient_match = re.search(r'(?:받는\s*분|수령인|받는사람|담당자|성함|이름)\s*[:：\-]?\s*([가-힣a-zA-Z]{2,10})', text)
    if recipient_match:
        result["recipient_name"] = recipient_match.group(1).strip()
        
    # 4. 주소 추출 (시/도/구/동/로/길/번지 등)
    addr_match = re.search(r'(?:주소|배송지|납품처|배송처)\s*[:：\-]?\s*([가-힣0-9a-zA-Z\s\(\)\-\,\.\~]+)', text)
    if addr_match:
        result["recipient_address"] = addr_match.group(1).strip()
    else:
        # 키워드 없어도 일반 한국 주소 패턴 탐지 (서울/경기/인천/부산... 구/시/군/동/로/길)
        addr_fallback = re.search(r'((?:서울|경기|인천|강원|충북|충남|대전|경북|경남|대구|전북|전남|광주|울산|부산|제주|세종)[^\n,]{8,60})', text)
        if addr_fallback:
            result["recipient_address"] = addr_fallback.group(1).strip()
            
    # 5. 메모/요청사항 추출
    memo_match = re.search(r'(?:메모|요청사항|특이사항|배송메모|비고)\s*[:：\-]?\s*([^\n]+)', text)
    if memo_match:
        result["memo"] = memo_match.group(1).strip()
        
    # 6. 품목 및 수량 추출
    extracted_items = []
    
    # 6-A. 먼저 단가 마스터에 등록된 품번/품명이 텍스트에 포함되어 있는지 탐색
    for line in lines:
        # 배송지나 연락처 헤더 라인은 제외
        if any(keyword in line for keyword in ["주소:", "배송지:", "받는분:", "연락처:", "전화:"]):
            continue
            
        found_code = None
        for code, info in price_master.items():
            if code in line.upper():
                found_code = info["item_code"]
                break
                
        if found_code:
            info = price_master.get(found_code.upper(), {})
            # 수량 탐색: '100개', '500권', '10박스', '50ea', '300'
            qty_match = re.search(r'(\d+)\s*(?:개|권|박스|box|ea|EA|매|장|set|세트)?', line.replace(found_code, ""))
            qty = int(qty_match.group(1)) if qty_match else 1
            
            extracted_items.append({
                "item_code": info.get("item_code", found_code),
                "item_name": info.get("item_name", found_code),
                "spec": info.get("spec", ""),
                "qty": qty,
                "buy_price": info.get("buy_price", 0),
                "sell_price": info.get("sell_price", 0),
                "remark": info.get("remark", "")
            })
            
    # 6-B. 만약 위에서 못 찾았다면, 품목 라인 패턴 (예: "BK-101 50개", "소프트커버 / 100", "품목: ABC-123, 50개") 정규식 탐지
    if not extracted_items:
        for line in lines:
            if any(k in line for k in ["주소", "받는", "수령", "연락처", "전화", "업체", "거래처", "배송", "메모"]):
                continue
            
            # 패턴 1: 영문/숫자 품번 + 수량 (예: BK-101 500, PROD-01 20개)
            m1 = re.search(r'([A-Za-z0-9\-_]{3,20})\s*[:\-\/,]?\s*(\d+)\s*(?:개|권|박스|box|ea|EA|매|장|set)?', line)
            if m1:
                code = m1.group(1).upper()
                qty = int(m1.group(2))
                info = price_master.get(code, {})
                extracted_items.append({
                    "item_code": code,
                    "item_name": info.get("item_name", code),
                    "spec": info.get("spec", ""),
                    "qty": qty,
                    "buy_price": info.get("buy_price", 0),
                    "sell_price": info.get("sell_price", 0),
                    "remark": ""
                })
                continue
                
            # 패턴 2: 한글 품명 + 수량 (예: 복사용지 10박스, 코팅지 500개)
            m2 = re.search(r'([가-힣a-zA-Z0-9\s]{2,20})\s*[:\-\/,]\s*(\d+)\s*(?:개|권|박스|box|ea|EA|매|장|set)?', line)
            if m2:
                name = m2.group(1).strip()
                qty = int(m2.group(2))
                info = price_master.get(name.upper(), {})
                extracted_items.append({
                    "item_code": info.get("item_code", name),
                    "item_name": name,
                    "spec": info.get("spec", ""),
                    "qty": qty,
                    "buy_price": info.get("buy_price", 0),
                    "sell_price": info.get("sell_price", 0),
                    "remark": ""
                })
                
    # 중복 제거 (품번 기준 합산)
    combined_items = {}
    for item in extracted_items:
        code = item["item_code"]
        if code in combined_items:
            combined_items[code]["qty"] += item["qty"]
        else:
            combined_items[code] = item
            
    result["items"] = list(combined_items.values())
    
    # 기본 발주일자 생성
    now = datetime.datetime.now()
    result["order_no"] = f"PO-{now.strftime('%Y%m%d')}-{now.strftime('%H%M%S')[-3:]}"
    
    return result

# ----------------- Excel Generators ----------------- #

def generate_order_excel(order_data: dict) -> str:
    """발주서 엑셀 파일을 생성하여 저장하고 파일명을 반환합니다."""
    template_path = os.path.join(TEMPLATES_DIR, "order_template.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError("발주서 템플릿(order_template.xlsx) 파일이 없습니다.")
        
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    config = load_config()
    company_info = config.get("company_info", {})
    
    # 발주일자 / 발주번호 / 발주처 / 담당자
    ws["B4"] = order_data.get("order_date", datetime.date.today().strftime("%Y-%m-%d"))
    ws["B5"] = order_data.get("order_no", f"PO-{datetime.datetime.now().strftime('%Y%m%d%H%M')}")
    ws["B6"] = company_info.get("company_name", "(주)바이브컴퍼니")
    ws["B7"] = f"{company_info.get('sender_name', '발주팀')} / {company_info.get('phone', '')}"
    
    # 수신 공급업체 / 담당자 / 수령인 / 배송지
    vendor_name = order_data.get("vendor_name", "거래처")
    ws["F4"] = vendor_name
    ws["F5"] = order_data.get("vendor_contact", "")
    
    rec_name = order_data.get("recipient_name", "")
    rec_phone = order_data.get("recipient_phone", "")
    ws["F6"] = f"{rec_name} / {rec_phone}".strip(" /")
    ws["F7"] = order_data.get("recipient_address", "")
    
    # 품목 데이터 기입 (Row 10부터 19까지)
    items = order_data.get("items", [])
    for idx in range(10): # 최대 10행 기본 처리
        row_num = 10 + idx
        if idx < len(items):
            item = items[idx]
            ws.cell(row=row_num, column=1, value=idx + 1)
            ws.cell(row=row_num, column=2, value=item.get("item_code", ""))
            ws.cell(row=row_num, column=3, value=item.get("item_name", ""))
            ws.cell(row=row_num, column=4, value=item.get("spec", ""))
            
            qty = item.get("qty", 0)
            ws.cell(row=row_num, column=5, value=qty)
            
            buy_price = item.get("buy_price", 0)
            ws.cell(row=row_num, column=6, value=buy_price)
            
            # 공급가액 (수식 = E{row}*F{row})
            ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}")
            ws.cell(row=row_num, column=8, value=item.get("remark", ""))
        else:
            # 빈 행
            ws.cell(row=row_num, column=1, value=idx + 1)
            ws.cell(row=row_num, column=2, value="")
            ws.cell(row=row_num, column=3, value="")
            ws.cell(row=row_num, column=4, value="")
            ws.cell(row=row_num, column=5, value=0)
            ws.cell(row=row_num, column=6, value=0)
            ws.cell(row=row_num, column=7, value=f"=E{row_num}*F{row_num}")
            ws.cell(row=row_num, column=8, value="")
            
    # 특이사항 메모
    if order_data.get("memo"):
        ws["A22"] = f"【특이사항 및 전달 메모】\n* {order_data['memo']}\n* 납기일정 준수 요망"
        
    safe_vendor = re.sub(r'[\/:*?"<>|]', '_', vendor_name) if vendor_name else "발주서"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"발주서_{safe_vendor}_{timestamp}.xlsx"
    out_file_path = os.path.join(ORDERS_DIR, file_name)
    
    wb.save(out_file_path)
    return file_name

def append_to_erp_upload_list(order_data: dict) -> str:
    """사내 ERP 업로드용 엑셀 파일(erp_upload_list.xlsx)에 거래처/품번/수량/매입가/수주가 행을 누적 추가합니다."""
    erp_file_path = os.path.join(OUTPUT_DIR, "erp_upload_list.xlsx")
    
    headers = [
        "발주일자", "발주번호", "거래처명", "품번", "품명", "규격",
        "수량", "매입단가", "매입합계", "수주단가", "수주합계", "예상마진",
        "수령인", "연락처", "배송지주소", "비고"
    ]
    
    if os.path.exists(erp_file_path):
        wb = openpyxl.load_workbook(erp_file_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ERP_업로드목록"
        ws.append(headers)
        
        # 헤더 스타일
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
        
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    order_date = order_data.get("order_date", datetime.date.today().strftime("%Y-%m-%d"))
    order_no = order_data.get("order_no", "")
    vendor_name = order_data.get("vendor_name", "")
    rec_name = order_data.get("recipient_name", "")
    rec_phone = order_data.get("recipient_phone", "")
    rec_addr = order_data.get("recipient_address", "")
    memo = order_data.get("memo", "")
    
    for item in order_data.get("items", []):
        qty = int(item.get("qty", 0))
        buy_p = int(item.get("buy_price", 0))
        sell_p = int(item.get("sell_price", 0))
        buy_total = qty * buy_p
        sell_total = qty * sell_p
        margin = sell_total - buy_total
        
        row_data = [
            order_date,
            order_no,
            vendor_name,
            item.get("item_code", ""),
            item.get("item_name", ""),
            item.get("spec", ""),
            qty,
            buy_p,
            buy_total,
            sell_p,
            sell_total,
            margin,
            rec_name,
            rec_phone,
            rec_addr,
            item.get("remark", memo)
        ]
        
        ws.append(row_data)
        current_row = ws.max_row
        ws.row_dimensions[current_row].height = 20
        
        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=current_row, column=c_idx)
            c.font = Font(name="맑은 고딕", size=10)
            c.border = thin_border
            # 숫자 서식 적용
            if c_idx in [7, 8, 9, 10, 11, 12]: # 수량, 매입단가, 매입합계, 수주단가, 수주합계, 예상마진
                c.number_format = '#,##0'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx in [1, 2, 4]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    # 컬럼 너비 자동 조정
    col_widths = {
        'A': 13, 'B': 18, 'C': 16, 'D': 16, 'E': 22, 'F': 16,
        'G': 10, 'H': 13, 'I': 14, 'J': 13, 'K': 14, 'L': 14,
        'M': 12, 'N': 16, 'O': 32, 'P': 20
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
        
    wb.save(erp_file_path)
    return erp_file_path

# ----------------- Naver Works Mail Sender ----------------- #

def send_naverworks_email(to_email: str, subject: str, body_text: str, attachment_filename: str) -> dict:
    """네이버웍스 SMTP(smtp.worksmobile.com:465)를 사용하여 발주서 첨부 메일을 발송합니다."""
    config = load_config()
    nw = config.get("naverworks", {})
    
    smtp_server = nw.get("smtp_server", "smtp.worksmobile.com")
    smtp_port = nw.get("smtp_port", 465)
    sender_email = nw.get("sender_email", "").strip()
    sender_password = nw.get("sender_password", "").strip()
    sender_name = nw.get("sender_name", "발주담당자")
    
    if not sender_email or not sender_password:
        raise ValueError("네이버웍스 이메일 계정 정보(이메일, 비밀번호)가 설정되어 있지 않습니다. 설정 화면에서 입력해주세요.")
        
    if not to_email:
        raise ValueError("수신자 이메일 주소가 비어 있습니다.")
        
    # MIME 메시지 생성
    msg = MIMEMultipart()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg["Date"] = datetime.datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0900")
    
    msg.attach(MIMEText(body_text, "plain", "utf-8"))
    
    # 첨부파일 처리
    attach_path = os.path.join(ORDERS_DIR, attachment_filename)
    if os.path.exists(attach_path):
        with open(attach_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            
            # 한글 파일명 헤더 인코딩
            part.add_header(
                "Content-Disposition",
                f"attachment; filename*=UTF-8''{attachment_filename}"
            )
            msg.attach(part)
    else:
        raise FileNotFoundError(f"첨부할 발주서 파일이 존재하지 않습니다: {attachment_filename}")
        
    # SMTP_SSL 전송
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=15)
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [to_email], msg.as_string())
        server.quit()
        return {"success": True, "message": f"'{to_email}'(으)로 발주서 메일이 성공적으로 전송되었습니다."}
    except Exception as e:
        raise RuntimeError(f"네이버웍스 메일 발송 실패: {str(e)}")

# ----------------- API Endpoints ----------------- #

class ParseRequest(BaseModel):
    text: str

class OrderItem(BaseModel):
    item_code: str
    item_name: str = ""
    spec: str = ""
    qty: int = 1
    buy_price: int = 0
    sell_price: int = 0
    remark: str = ""

class CreateOrderRequest(BaseModel):
    vendor_name: str
    order_no: Optional[str] = None
    order_date: Optional[str] = None
    vendor_contact: Optional[str] = ""
    recipient_name: Optional[str] = ""
    recipient_phone: Optional[str] = ""
    recipient_address: Optional[str] = ""
    memo: Optional[str] = ""
    items: List[OrderItem]

class SendMailRequest(BaseModel):
    to_email: str
    subject: str
    body: str
    attachment_filename: str

class TestMailRequest(BaseModel):
    to_email: str

@app.post("/api/parse-order")
async def api_parse_order(req: ParseRequest):
    """카카오톡 주문 텍스트를 파싱하여 기본 정보 및 품번 단가 매칭 결과를 반환합니다."""
    if not req.text.strip():
        raise HTTPException(status_code=400, detail="카톡 주문 내용을 입력해주세요.")
    data = parse_kakao_text(req.text)
    
    # 해당 거래처의 등록된 이메일이 있는지 확인
    config = load_config()
    vendor_email = ""
    for v in config.get("vendors", []):
        if v["name"].lower() == data["vendor_name"].lower():
            vendor_email = v.get("email", "")
            break
    data["vendor_email"] = vendor_email
    
    return {"success": True, "data": data}

@app.post("/api/create-order-and-erp")
async def api_create_order_and_erp(order: CreateOrderRequest):
    """[1차 버튼] 발주서 엑셀 파일을 생성하고, 사내 ERP 누적 엑셀에 행을 추가합니다."""
    order_dict = order.model_dump() if hasattr(order, 'model_dump') else order.dict()

    
    # 1. 거래처 발송용 발주서 엑셀 생성
    excel_filename = generate_order_excel(order_dict)
    
    # 2. 사내 ERP 업로드용 엑셀 행 누적
    append_to_erp_upload_list(order_dict)
    
    # 3. 메일 템플릿 기본 내용 생성
    config = load_config()
    mail_tpl = config.get("default_mail_template", {})
    comp_info = config.get("company_info", {})
    
    item_summary = f"{order_dict['items'][0]['item_name'] or order_dict['items'][0]['item_code']} 외 {len(order_dict['items'])-1}건" if len(order_dict['items']) > 1 else (order_dict['items'][0]['item_name'] or order_dict['items'][0]['item_code']) if order_dict['items'] else "발주"
    
    # 거래처 이메일 자동 매칭
    to_email = ""
    for v in config.get("vendors", []):
        if v["name"].lower() == order_dict["vendor_name"].lower():
            to_email = v.get("email", "")
            break
            
    subject = mail_tpl.get("subject", "[발주서] {vendor_name} 귀하").format(
        vendor_name=order_dict["vendor_name"],
        date=order_dict.get("order_date", datetime.date.today().strftime("%Y-%m-%d")),
        item_summary=item_summary,
        company_name=comp_info.get("company_name", ""),
        sender_name=comp_info.get("sender_name", "")
    )
    
    body = mail_tpl.get("body", "").format(
        vendor_name=order_dict["vendor_name"],
        date=order_dict.get("order_date", datetime.date.today().strftime("%Y-%m-%d")),
        item_summary=item_summary,
        recipient_name=order_dict.get("recipient_name", ""),
        recipient_phone=order_dict.get("recipient_phone", ""),
        recipient_address=order_dict.get("recipient_address", ""),
        company_name=comp_info.get("company_name", ""),
        sender_name=comp_info.get("sender_name", "")
    )
    
    return {
        "success": True,
        "message": "발주서 엑셀 및 사내 ERP 리스트 생성이 완료되었습니다.",
        "order_file": excel_filename,
        "mail_draft": {
            "to_email": to_email,
            "subject": subject,
            "body": body,
            "attachment_filename": excel_filename
        }
    }

@app.post("/api/send-mail")
async def api_send_mail(req: SendMailRequest):
    """[2차 버튼] 네이버웍스를 통해 발주서 첨부 메일을 발송합니다."""
    try:
        result = send_naverworks_email(
            to_email=req.to_email,
            subject=req.subject,
            body_text=req.body,
            attachment_filename=req.attachment_filename
        )
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/test-naverworks")
async def api_test_naverworks(req: TestMailRequest):
    """네이버웍스 SMTP 연동 테스트 메일을 발송합니다."""
    try:
        config = load_config()
        nw = config.get("naverworks", {})
        sender_email = nw.get("sender_email", "")
        if not sender_email or not nw.get("sender_password"):
            raise ValueError("네이버웍스 계정 정보(이메일, 비밀번호)가 설정되어 있지 않습니다.")
            
        msg = MIMEMultipart()
        msg["From"] = f"{nw.get('sender_name', '발주시스템')} <{sender_email}>"
        msg["To"] = req.to_email
        msg["Subject"] = "[테스트] 네이버웍스 연동 테스트 메일입니다."
        msg.attach(MIMEText("네이버웍스 SMTP 연동이 정상적으로 작동하고 있습니다.\n\n발주 자동화 시스템 준비 완료!", "plain", "utf-8"))
        
        server = smtplib.SMTP_SSL(nw.get("smtp_server", "smtp.worksmobile.com"), nw.get("smtp_port", 465), timeout=10)
        server.login(sender_email, nw.get("sender_password"))
        server.sendmail(sender_email, [req.to_email], msg.as_string())
        server.quit()
        return {"success": True, "message": f"'{req.to_email}'(으)로 테스트 메일이 성공적으로 전송되었습니다!"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"네이버웍스 연결 실패: {str(e)}")

# ----------------- Download Endpoints ----------------- #

@app.post("/api/open-order-folder")
async def api_open_order_folder():
    """윈도우 탐색기에서 발주서 폴더(output/orders)를 바로 엽니다."""
    try:
        import subprocess
        os.makedirs(ORDERS_DIR, exist_ok=True)
        subprocess.Popen(f'explorer "{ORDERS_DIR}"')
        return {"success": True, "message": "발주서 폴더를 열었습니다."}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.get("/api/download/order/{filename}")
async def download_order_file(filename: str):
    file_path = os.path.join(ORDERS_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    return FileResponse(file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/download/erp-list")
async def download_erp_list():
    file_path = os.path.join(OUTPUT_DIR, "erp_upload_list.xlsx")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="ERP 누적 엑셀 파일이 아직 없습니다. 발주서를 먼저 생성해주세요.")
    return FileResponse(file_path, filename="erp_upload_list.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/download/price-master")
async def download_price_master():
    file_path = os.path.join(TEMPLATES_DIR, "price_master.xlsx")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="단가표 파일이 없습니다.")
    return FileResponse(file_path, filename="price_master.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/download/order-template")
async def download_order_template():
    file_path = os.path.join(TEMPLATES_DIR, "order_template.xlsx")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="발주서 템플릿 파일이 없습니다.")
    return FileResponse(file_path, filename="order_template.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------- Master & Config Management ----------------- #

@app.get("/api/config")
async def api_get_config():
    return load_config()

@app.post("/api/config")
async def api_save_config(config_data: dict):
    save_config(config_data)
    return {"success": True, "message": "설정이 저장되었습니다."}

@app.get("/api/price-master")
async def api_get_price_master():
    master = get_price_master()
    # Unique item list
    unique_items = {}
    for k, v in master.items():
        code = v["item_code"]
        if code not in unique_items:
            unique_items[code] = v
    return {"success": True, "items": list(unique_items.values())}

@app.post("/api/upload/price-master")
async def api_upload_price_master(file: UploadFile = File(...)):
    """단가 마스터 엑셀 파일 업로드"""
    file_path = os.path.join(TEMPLATES_DIR, "price_master.xlsx")
    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)
    return {"success": True, "message": "단가표(price_master.xlsx)가 성공적으로 업데이트되었습니다."}

@app.post("/api/upload/order-template")
async def api_upload_order_template(file: UploadFile = File(...)):
    """발주서 양식 엑셀 파일 업로드"""
    file_path = os.path.join(TEMPLATES_DIR, "order_template.xlsx")
    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)
    return {"success": True, "message": "발주서 양식(order_template.xlsx)이 성공적으로 업데이트되었습니다."}

@app.get("/api/erp-summary")
async def api_get_erp_summary():
    """ERP 누적 목록의 요약 통계 및 최근 내역을 반환합니다."""
    erp_file_path = os.path.join(OUTPUT_DIR, "erp_upload_list.xlsx")
    if not os.path.exists(erp_file_path):
        return {"exists": False, "total_rows": 0, "recent_items": []}
        
    wb = openpyxl.load_workbook(erp_file_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    
    recent = []
    for r in reversed(rows[-10:]): # 최근 10건
        if r and len(r) >= 12:
            recent.append({
                "date": str(r[0]),
                "order_no": str(r[1]),
                "vendor": str(r[2]),
                "code": str(r[3]),
                "name": str(r[4]),
                "qty": r[6],
                "buy_price": r[7],
                "sell_price": r[9],
                "margin": r[11]
            })
            
    return {
        "exists": True,
        "total_rows": len(rows),
        "recent_items": recent
    }

# ----------------- Static Files ----------------- #
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def serve_index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))

if __name__ == "__main__":
    import uvicorn
    print("Server started at http://127.0.0.1:8000 (or http://localhost:8000)")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
