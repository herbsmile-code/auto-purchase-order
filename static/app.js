let currentOrderData = null;
let currentOrderFile = "";
let currentOrderHtml = "";

// Sample Kakao Messages
const SAMPLES = {
  1: `1. 납품처 
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
음료아날로그 / SR-SC44RW - 2대`,

  2: `1. 납품처
   매장명 : 미쉐 서면점
   주소 : 부산광역시 부산진구 중앙대로 686
   현장담당 이름, 전화번호 : 김현수 010-5555-4321
   납품일 및 희망 시간 : 08월 25일 오전

2. 모델명과 수량

스타리온 1500 냉동냉장 / 직냉식 / SR-T15B1F - 1대
스타리온 25BOX 올냉동 / 직냉식 / SR-E25BAFC - 1대
음료아날로그 / SR-SC44RW - 1대`
};

// Safe DOM Helpers
function safeSetVal(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = (val !== null && val !== undefined) ? val : "";
}

function safeGetVal(id, defaultVal = "") {
  const el = document.getElementById(id);
  return el ? el.value.trim() : defaultVal;
}

function safeSetText(id, txt) {
  const el = document.getElementById(id);
  if (el) el.innerText = (txt !== null && txt !== undefined) ? txt : "";
}

function safeAddEvent(id, event, handler) {
  const el = document.getElementById(id);
  if (el) el.addEventListener(event, handler);
}

// DOM Ready
document.addEventListener("DOMContentLoaded", () => {
  try {
    initEventListeners();
    loadErpSummary();
    loadConfig();
  } catch (e) {
    console.error("Initialization error:", e);
  }
});

function initEventListeners() {
  // Clear button
  safeAddEvent("btnClear", "click", () => {
    safeSetVal("kakaoInput", "");
    const s2 = document.getElementById("step2Card");
    if (s2) s2.style.display = "none";
  });

  // STEP 1 Button: Create Order & ERP List
  safeAddEvent("btnStep1Submit", "click", handleStep1Submit);

  // STEP 2: Items Table Add Row & Regenerate
  safeAddEvent("btnAddRow", "click", handleAddTableRow);
  safeAddEvent("btnRegenerateOrder", "click", handleRegenerateOrder);

  // STEP 2: Copy Mail Body (HTML Table + Text)
  safeAddEvent("btnCopyMailBody", "click", handleCopyMailBody);

  // STEP 2: Open Folder
  safeAddEvent("btnOpenFolder", "click", handleOpenFolder);

  // STEP 2: Open Edge Naver Works Webmail
  safeAddEvent("btnOpenWebmail", "click", handleOpenWebmail);

  // STEP 2: Send Mail (SMTP)
  safeAddEvent("btnSendMail", "click", handleSendMail);

  // Download buttons
  safeAddEvent("btnDownloadOrder", "click", () => {
    if (currentOrderFile) {
      window.open(`/api/download/order/${encodeURIComponent(currentOrderFile)}`, "_blank");
    }
  });

  safeAddEvent("btnDownloadErp", "click", () => {
    window.open("/api/download/erp-list", "_blank");
  });

  safeAddEvent("btnModalDownloadErp", "click", () => {
    window.open("/api/download/erp-list", "_blank");
  });

  // Header Modal Triggers
  safeAddEvent("btnErpSummary", "click", openErpModal);
  safeAddEvent("btnOpenPriceMaster", "click", openPriceMasterModal);
  safeAddEvent("btnOpenSettings", "click", openSettingsModal);

  // ERP Clear button in modal
  safeAddEvent("btnClearErpList", "click", handleClearErpList);

  // Settings Modal actions
  safeAddEvent("btnSaveSettings", "click", saveConfig);
  safeAddEvent("btnTestMail", "click", handleTestMail);
  safeAddEvent("btnAddVendor", "click", () => addVendorRow());

  // Template Download / Upload actions
  safeAddEvent("btnDownloadTemplate", "click", () => {
    window.open("/api/download/order-template", "_blank");
  });
  safeAddEvent("orderTemplateFileInput", "change", handleOrderTemplateUpload);
}

// ----------------- Safe Fetch Helper ----------------- //

async function fetchJson(url, options = {}) {
  const res = await fetch(url, options);
  const text = await res.text();
  
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    if (!res.ok) {
      throw new Error(`서버 오류 (${res.status}): 요청 처리 중 문제가 발생했습니다.`);
    }
    throw new Error("서버 응답을 해석할 수 없습니다.");
  }

  if (!res.ok) {
    throw new Error(data.detail || data.message || `오류 발생 (상태 코드: ${res.status})`);
  }

  return data;
}

// ----------------- STEP 1: Process Kakao Order ----------------- //

function loadSample(num) {
  safeSetVal("kakaoInput", SAMPLES[num] || "");
  showToast("샘플 출고요청서 내용이 입력되었습니다.", "info");
}

async function handleStep1Submit() {
  const kakaoText = safeGetVal("kakaoInput");
  if (!kakaoText) {
    showToast("카카오톡 주문 내용을 입력해주세요.", "error");
    return;
  }

  const chkReplace = document.getElementById("chkReplaceErp");
  const replaceErp = chkReplace ? chkReplace.checked : false;

  const btn = document.getElementById("btnStep1Submit");
  const originalHtml = btn ? btn.innerHTML : "";
  if (btn) {
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 분석 및 출고요청서 생성 중...`;
    btn.disabled = true;
  }

  try {
    // 1. 카톡 파싱 API 호출 (모델명 및 단가표 매칭)
    const parseData = await fetchJson("/api/parse-order", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: kakaoText })
    });

    const order = parseData.data;
    order.replace_erp = replaceErp;

    // 2. 발주서 및 ERP 리스트 생성 API 호출
    const createData = await fetchJson("/api/create-order-and-erp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(order)
    });

    currentOrderData = order;
    currentOrderFile = createData.order_file;

    // 3. UI 렌더링
    renderStep2(order, createData);
    loadErpSummary();

    showToast("🎉 출고요청서 엑셀 및 ERP 리스트 생성이 완료되었습니다!", "success");

  } catch (err) {
    console.error(err);
    showToast(`오류: ${err.message}`, "error");
  } finally {
    if (btn) {
      btn.innerHTML = originalHtml;
      btn.disabled = false;
    }
  }
}

// ----------------- STEP 2: Render & Review ----------------- //

function renderStep2(order, createData) {
  const step2Card = document.getElementById("step2Card");
  if (step2Card) step2Card.style.display = "block";

  // Basic Info
  safeSetVal("orderStore", order.store_name || order.recipient_name || "");
  safeSetVal("orderCustomerName", order.customer_name || "");
  safeSetVal("orderDeliveryDate", order.delivery_date || "");
  safeSetVal("orderContact", order.recipient_contact || `${order.recipient_name || ''} ${order.recipient_phone || ''}`.trim());
  safeSetVal("orderAddress", order.recipient_address || "");

  // File Name & HTML Body
  safeSetText("orderFileName", createData.order_file || "");
  const mailDraft = createData.mail_draft || {};
  currentOrderHtml = mailDraft.body_html || "";

  // Mail Draft
  safeSetVal("mailTo", mailDraft.to_email || order.vendor_email || "hj.seo@starion.co.kr, gscheon@starion.co.kr, cth@ohjin.co.kr");
  safeSetVal("mailSubject", mailDraft.subject || "");
  safeSetVal("mailBody", mailDraft.body || "");

  // Render Items Table
  renderItemsTable(order.items || []);

  // Smooth scroll to Step 2
  if (step2Card) {
    step2Card.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function renderItemsTable(items) {
  const tbody = document.getElementById("itemsTableBody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!items || items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted" style="padding: 20px;">등록된 품목이 없습니다. [품목 추가]를 눌러주세요.</td></tr>`;
    calculateTotals();
    return;
  }

  items.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index;

    const no = index + 1;
    const itemName = item.item_name || "";
    const modelName = item.model_name || item.item_code || "";
    const qty = parseInt(item.qty) || 1;
    const buyPrice = parseInt(item.buy_price) || 0;
    const sellPrice = parseInt(item.sell_price) || 0;
    const margin = sellPrice - buyPrice;
    const remark = item.remark || "스타리온 직배송 요청";

    tr.innerHTML = `
      <td class="text-center font-bold">${no}</td>
      <td><input type="text" class="table-input item-name" value="${escapeHtml(itemName)}" placeholder="품목명 (예: 스타리온 1200 반찬냉장고)" /></td>
      <td><input type="text" class="table-input item-model font-mono text-primary font-bold text-center" value="${escapeHtml(modelName)}" placeholder="모델명 (예: SRV12EIEVF)" /></td>
      <td><input type="number" class="table-input item-qty text-center font-bold" value="${qty}" min="1" /></td>
      <td><input type="number" class="table-input item-buy text-right font-bold" value="${buyPrice}" placeholder="매입단가" /></td>
      <td><input type="number" class="table-input item-sell text-right font-bold text-primary" value="${sellPrice}" placeholder="수주단가" /></td>
      <td class="text-right font-bold item-margin-cell" style="color: ${margin >= 0 ? '#4ade80' : '#f87171'}; padding-right: 10px;">${formatNumber(margin)}원</td>
      <td><input type="text" class="table-input item-remark text-center" value="${escapeHtml(remark)}" placeholder="비고" /></td>
      <td class="text-center">
        <button class="btn-danger-sm" onclick="handleDeleteRow(this)" title="행 삭제"><i class="fa-solid fa-trash"></i></button>
      </td>
    `;

    // Realtime listeners
    const qtyInp = tr.querySelector(".item-qty");
    const buyInp = tr.querySelector(".item-buy");
    const sellInp = tr.querySelector(".item-sell");

    if (qtyInp) qtyInp.addEventListener("input", () => updateRowMarginAndTotals(tr));
    if (buyInp) buyInp.addEventListener("input", () => updateRowMarginAndTotals(tr));
    if (sellInp) sellInp.addEventListener("input", () => updateRowMarginAndTotals(tr));

    tbody.appendChild(tr);
  });

  calculateTotals();
}

function updateRowMarginAndTotals(tr) {
  const buyInp = tr.querySelector(".item-buy");
  const sellInp = tr.querySelector(".item-sell");
  const marginCell = tr.querySelector(".item-margin-cell");

  const buy = parseInt(buyInp ? buyInp.value : 0) || 0;
  const sell = parseInt(sellInp ? sellInp.value : 0) || 0;
  const margin = sell - buy;

  if (marginCell) {
    marginCell.innerText = `${formatNumber(margin)}원`;
    marginCell.style.color = margin >= 0 ? '#4ade80' : '#f87171';
  }

  calculateTotals();
}

function handleAddTableRow() {
  const tbody = document.getElementById("itemsTableBody");
  if (!tbody) return;

  const rowCount = tbody.querySelectorAll("tr").length + 1;
  const tr = document.createElement("tr");
  
  tr.innerHTML = `
    <td class="text-center font-bold">${rowCount}</td>
    <td><input type="text" class="table-input item-name" placeholder="품목명" /></td>
    <td><input type="text" class="table-input item-model font-mono text-primary font-bold text-center" placeholder="모델명 (예: SRV12EIEVF)" /></td>
    <td><input type="number" class="table-input item-qty text-center font-bold" value="1" min="1" /></td>
    <td><input type="number" class="table-input item-buy text-right font-bold" value="0" placeholder="매입단가" /></td>
    <td><input type="number" class="table-input item-sell text-right font-bold text-primary" value="0" placeholder="수주단가" /></td>
    <td class="text-right font-bold item-margin-cell" style="color: #4ade80; padding-right: 10px;">0원</td>
    <td><input type="text" class="table-input item-remark text-center" value="스타리온 직배송 요청" placeholder="비고" /></td>
    <td class="text-center">
      <button class="btn-danger-sm" onclick="handleDeleteRow(this)" title="행 삭제"><i class="fa-solid fa-trash"></i></button>
    </td>
  `;

  const qtyInp = tr.querySelector(".item-qty");
  const buyInp = tr.querySelector(".item-buy");
  const sellInp = tr.querySelector(".item-sell");

  if (qtyInp) qtyInp.addEventListener("input", () => updateRowMarginAndTotals(tr));
  if (buyInp) buyInp.addEventListener("input", () => updateRowMarginAndTotals(tr));
  if (sellInp) sellInp.addEventListener("input", () => updateRowMarginAndTotals(tr));

  tbody.appendChild(tr);
  calculateTotals();
}

function handleDeleteRow(btn) {
  const tr = btn.closest("tr");
  if (tr) tr.remove();
  // Re-index row numbers
  document.querySelectorAll("#itemsTableBody tr").forEach((row, idx) => {
    const noCell = row.querySelector("td:first-child");
    if (noCell) noCell.innerText = idx + 1;
  });
  calculateTotals();
}

function calculateTotals() {
  const rows = document.querySelectorAll("#itemsTableBody tr");
  let totQty = 0;
  let totBuy = 0;
  let totSell = 0;

  rows.forEach(tr => {
    const qtyInp = tr.querySelector(".item-qty");
    const buyInp = tr.querySelector(".item-buy");
    const sellInp = tr.querySelector(".item-sell");

    if (!qtyInp) return;
    const qty = parseInt(qtyInp.value) || 0;
    const buy = parseInt(buyInp ? buyInp.value : 0) || 0;
    const sell = parseInt(sellInp ? sellInp.value : 0) || 0;

    totQty += qty;
    totBuy += (qty * buy);
    totSell += (qty * sell);
  });

  const totMargin = totSell - totBuy;

  safeSetText("totQty", `${formatNumber(totQty)} 대`);
  safeSetText("totBuyPrice", `${formatNumber(totBuy)} 원`);
  safeSetText("totSellPrice", `${formatNumber(totSell)} 원`);

  const elMargin = document.getElementById("totMargin");
  if (elMargin) {
    elMargin.innerText = `${formatNumber(totMargin)} 원`;
    elMargin.style.color = totMargin >= 0 ? '#4ade80' : '#f87171';
  }
}

function gatherCurrentOrderData() {
  const items = [];
  const rows = document.querySelectorAll("#itemsTableBody tr");
  
  rows.forEach((tr, idx) => {
    const nameInp = tr.querySelector(".item-name");
    const modelInp = tr.querySelector(".item-model");
    const qtyInp = tr.querySelector(".item-qty");
    const buyInp = tr.querySelector(".item-buy");
    const sellInp = tr.querySelector(".item-sell");
    const remarkInp = tr.querySelector(".item-remark");
    
    if (!nameInp && !modelInp) return;
    
    const itemName = nameInp ? nameInp.value.trim() : "";
    const modelName = modelInp ? modelInp.value.trim() : "";
    if (!itemName && !modelName) return;

    const qty = parseInt(qtyInp ? qtyInp.value : 1) || 1;
    const buyPrice = parseInt(buyInp ? buyInp.value : 0) || 0;
    const sellPrice = parseInt(sellInp ? sellInp.value : 0) || 0;

    items.push({
      no: idx + 1,
      item_code: modelName,
      model_name: modelName,
      item_name: itemName,
      qty: qty,
      buy_price: buyPrice,
      sell_price: sellPrice,
      margin: sellPrice - buyPrice,
      remark: remarkInp ? remarkInp.value.trim() : "스타리온 직배송 요청"
    });
  });

  const chkReplace = document.getElementById("chkReplaceErp");

  return {
    vendor_name: "스타리온",
    customer_name: safeGetVal("orderCustomerName"),
    store_name: safeGetVal("orderStore"),
    delivery_date: safeGetVal("orderDeliveryDate"),
    recipient_contact: safeGetVal("orderContact"),
    recipient_address: safeGetVal("orderAddress"),
    replace_erp: chkReplace ? chkReplace.checked : false,
    items: items
  };
}

async function handleRegenerateOrder() {
  const updatedOrder = gatherCurrentOrderData();
  if (updatedOrder.items.length === 0) {
    showToast("최소 1개 이상의 품목을 입력해주세요.", "error");
    return;
  }

  const btn = document.getElementById("btnRegenerateOrder");
  const originalHtml = btn ? btn.innerHTML : "";
  if (btn) {
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 재반영 중...`;
    btn.disabled = true;
  }

  try {
    const data = await fetchJson("/api/create-order-and-erp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(updatedOrder)
    });

    currentOrderFile = data.order_file;
    safeSetText("orderFileName", data.order_file || "");
    
    // 메일 내용 동기화
    if (data.mail_draft) {
      safeSetVal("mailSubject", data.mail_draft.subject || "");
      safeSetVal("mailBody", data.mail_draft.body || "");
    }
    
    loadErpSummary();
    showToast("수정된 단가 및 내용으로 출고요청서 및 ERP 리스트가 재반영되었습니다!", "success");
  } catch (err) {
    showToast(`오류: ${err.message}`, "error");
  } finally {
    if (btn) {
      btn.innerHTML = originalHtml;
      btn.disabled = false;
    }
  }
}

// ----------------- STEP 2: In-Memory Order Image Generator & Webmail ----------------- //

function createOrderImageBlob(orderData) {
  return new Promise((resolve) => {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");

    // 최적의 밸런스 사이즈 (너비 620px, 1x 정밀 렌더링)
    const width = 620;
    const items = orderData.items || [];
    const numItems = Math.max(items.length, 1);
    const rowHeight = 26;
    const tableHeight = (numItems + 1) * rowHeight;
    const height = 145 + tableHeight + 45;

    canvas.width = width;
    canvas.height = height;

    // 1. Background
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    // 2. Metadata (상단 정보)
    ctx.textBaseline = "middle";
    let y = 18;
    const xStart = 14;

    // 납품처명
    ctx.font = "bold 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillStyle = "#000000";
    ctx.textAlign = "left";
    ctx.fillText("납품처명:    ", xStart, y);
    ctx.font = "normal 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText(orderData.store_name || orderData.recipient_name || "", xStart + 80, y);

    y += 20;
    // 배송요청일 (파란색 볼드)
    ctx.font = "bold 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText("배송요청일 : ", xStart, y);
    ctx.fillStyle = "#0033cc";
    ctx.fillText(orderData.delivery_date || "", xStart + 80, y);
    ctx.fillStyle = "#000000";

    y += 20;
    // 배송장소
    ctx.font = "bold 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText("배송장소 :   ", xStart, y);
    ctx.font = "normal 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText(orderData.recipient_address || "", xStart + 80, y);

    y += 20;
    // 담당자
    ctx.font = "bold 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText("담당자 :     ", xStart, y);
    ctx.font = "normal 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText(orderData.recipient_contact || "", xStart + 80, y);

    y += 28;
    ctx.font = "bold 12.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.fillText("아래와 같이 기기류에 대해 견적하오니 검토 바랍니다.", xStart, y);

    // 3. Table (품목 표 - 넉넉한 열 너비 분배)
    y += 18;
    const tableX = 14;
    const tableW = width - 28;
    const colWidths = [34, 285, 125, 38, 110];
    const colsX = [tableX];
    for (let w of colWidths) {
      colsX.push(colsX[colsX.length - 1] + w);
    }

    // Header
    ctx.fillStyle = "#a6a6a6";
    ctx.fillRect(tableX, y, tableW, rowHeight);
    ctx.strokeStyle = "#777777";
    ctx.lineWidth = 1;
    ctx.strokeRect(tableX, y, tableW, rowHeight);

    const headers = ["NO", "품목", "모델명", "수량", "비고"];
    ctx.fillStyle = "#000000";
    ctx.font = "bold 12px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.textAlign = "center";

    for (let i = 0; i < headers.length; i++) {
      const cx = colsX[i] + colWidths[i] / 2;
      ctx.fillText(headers[i], cx, y + rowHeight / 2);
      ctx.beginPath();
      ctx.moveTo(colsX[i], y);
      ctx.lineTo(colsX[i], y + rowHeight);
      ctx.stroke();
    }

    // Rows
    let currY = y + rowHeight;
    ctx.font = "normal 11.5px '맑은 고딕', 'Malgun Gothic', sans-serif";

    for (let idx = 0; idx < items.length; idx++) {
      const it = items[idx];
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(tableX, currY, tableW, rowHeight);
      ctx.strokeRect(tableX, currY, tableW, rowHeight);

      // NO
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      ctx.fillText(String(idx + 1), colsX[0] + colWidths[0] / 2, currY + rowHeight / 2);

      // 품목 (maxWidth 지정 및 클리핑 마스크로 다음 열 침범/겹침 100% 방지)
      ctx.save();
      ctx.beginPath();
      ctx.rect(colsX[1] + 1, currY + 1, colWidths[1] - 2, rowHeight - 2);
      ctx.clip();
      ctx.fillStyle = "#000000";
      ctx.textAlign = "left";
      const itemMaxW = colWidths[1] - 12;
      ctx.fillText(it.item_name || "", colsX[1] + 6, currY + rowHeight / 2, itemMaxW);
      ctx.restore();

      // 모델명 (클리핑 및 maxWidth)
      ctx.save();
      ctx.beginPath();
      ctx.rect(colsX[2] + 1, currY + 1, colWidths[2] - 2, rowHeight - 2);
      ctx.clip();
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      const modelMaxW = colWidths[2] - 8;
      ctx.fillText(it.model_name || it.item_code || "", colsX[2] + colWidths[2] / 2, currY + rowHeight / 2, modelMaxW);
      ctx.restore();

      // 수량
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      ctx.fillText(String(it.qty || 1), colsX[3] + colWidths[3] / 2, currY + rowHeight / 2);

      // Grid lines
      for (let cx of colsX) {
        ctx.beginPath();
        ctx.moveTo(cx, currY);
        ctx.lineTo(cx, currY + rowHeight);
        ctx.stroke();
      }

      currY += rowHeight;
    }

    // 비고 Merged Area
    const remarkText = (items.length > 0 && items[0].remark) ? items[0].remark : "스타리온 직배송 요청";
    const remarkStartY = y + rowHeight;
    const remarkHeight = currY - remarkStartY;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(colsX[4], remarkStartY, colWidths[4], remarkHeight);
    ctx.strokeRect(colsX[4], remarkStartY, colWidths[4], remarkHeight);

    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.fillText(remarkText, colsX[4] + colWidths[4] / 2, remarkStartY + remarkHeight / 2, colWidths[4] - 8);

    // 4. 특이사항 Box
    currY += 9;
    const boxH = 28;
    ctx.strokeRect(tableX, currY, 130, boxH);
    ctx.strokeRect(tableX + 130, currY, tableW - 130, boxH);

    ctx.font = "bold 12px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("특이사항", tableX + 65, currY + boxH / 2);

    ctx.font = "normal 11.5px '맑은 고딕', 'Malgun Gothic', sans-serif";
    ctx.textAlign = "left";
    ctx.fillText(orderData.memo || "", tableX + 140, currY + boxH / 2, tableW - 150);

    canvas.toBlob((blob) => {
      resolve(blob);
    }, "image/png");
  });
}

async function handleCopyMailBody() {
  const currentOrder = gatherCurrentOrderData();
  const bodyText = safeGetVal("mailBody");

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      // 메모리에서 50% 축소 발주서 이미지를 실시간 렌더링하여 클립보드에 이미지로 복사
      const imgBlob = await createOrderImageBlob(currentOrder);
      if (imgBlob) {
        await navigator.clipboard.write([
          new ClipboardItem({
            "image/png": imgBlob
          })
        ]);
        showToast("🖼️ 발주서 이미지가 클립보드에 복사되었습니다! 메일 본문에서 [Ctrl + V]를 누르세요.", "success");
        return;
      }
    }
  } catch (err) {
    console.log("Image clipboard write fallback:", err);
  }

  // Fallback to text copy
  if (navigator.clipboard && bodyText) {
    await navigator.clipboard.writeText(bodyText);
    showToast("📋 메일 본문이 클립보드에 복사되었습니다! (Ctrl+V)", "info");
  }
}

async function handleOpenFolder() {
  try {
    const filename = currentOrderFile || "";
    await fetch(`/api/open-order-folder${filename ? `?filename=${encodeURIComponent(filename)}` : ""}`, { method: "POST" });
    showToast("📂 윈도우 탐색기에서 출고요청서 엑셀 파일이 선택되었습니다.", "info");
  } catch (e) {
    showToast("폴더 열기 실패", "error");
  }
}

async function handleOpenWebmail() {
  const to = safeGetVal("mailTo") || "hj.seo@starion.co.kr, gscheon@starion.co.kr, cth@ohjin.co.kr";
  const subject = safeGetVal("mailSubject") || "";
  const body = safeGetVal("mailBody") || "";

  // 1. 발주서 파일이 선택된 상태로 윈도우 탐색기 열기 (드래그&드롭 즉시 준비)
  await handleOpenFolder();

  // 2. 발주서 이미지(50% 크기)를 메모리에서 실시간 생성하여 클립보드에 자동 복사
  await handleCopyMailBody();

  // 3. 네이버웍스 웹메일 메일쓰기 팝업 URL 생성 (to, subject, body 기본 자동 주입)
  const encodedTo = encodeURIComponent(to);
  const encodedSubject = encodeURIComponent(subject);
  const encodedBody = encodeURIComponent(body);

  const popupUrl = `https://mail.worksmobile.com/write/popup?to=${encodedTo}&subject=${encodedSubject}&body=${encodedBody}`;
  
  const popupWidth = 1080;
  const popupHeight = 850;
  const left = Math.max(0, (window.screen.width - popupWidth) / 2);
  const top = Math.max(0, (window.screen.height - popupHeight) / 2);
  
  const popupWindow = window.open(
    popupUrl,
    "NaverWorksMailWrite",
    `width=${popupWidth},height=${popupHeight},top=${top},left=${left},resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no`
  );

  if (popupWindow) {
    popupWindow.focus();
    showToast("✉️ 네이버웍스 창이 열렸습니다! 기본 내용 확인 후 본문에서 [Ctrl + V]를 누르면 이미지가 들어갑니다!", "success");
  } else {
    window.open(popupUrl, "_blank");
    showToast("🌐 새 탭으로 네이버웍스 화면이 열렸습니다. 본문에서 [Ctrl + V]를 누르세요!", "info");
  }
}

// ----------------- STEP 2: Send Mail via Naver Works (SMTP) ----------------- //

async function handleSendMail() {
  const toEmail = safeGetVal("mailTo");
  const subject = safeGetVal("mailSubject");
  const body = safeGetVal("mailBody");

  if (!toEmail) {
    showToast("받는 사람(거래처) 이메일 주소를 입력해주세요.", "error");
    const mailToEl = document.getElementById("mailTo");
    if (mailToEl) mailToEl.focus();
    return;
  }
  if (!currentOrderFile) {
    showToast("첨부할 발주서 파일이 없습니다. 1차 생성을 먼저 진행해주세요.", "error");
    return;
  }

  const btn = document.getElementById("btnSendMail");
  const originalHtml = btn ? btn.innerHTML : "";
  if (btn) {
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 네이버웍스 전송 중...`;
    btn.disabled = true;
  }

  try {
    const data = await fetchJson("/api/send-mail", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to_email: toEmail,
        subject: subject,
        body: body,
        attachment_filename: currentOrderFile
      })
    });

    showToast(`✉️ ${data.message || '출고요청서 메일이 성공적으로 전송되었습니다!'}`, "success");
  } catch (err) {
    console.error(err);
    showToast(`메일 발송 실패: ${err.message}`, "error");
  } finally {
    if (btn) {
      btn.innerHTML = originalHtml;
      btn.disabled = false;
    }
  }
}

// ----------------- Modals & Settings ----------------- //

function closeModal(modalId) {
  const m = document.getElementById(modalId);
  if (m) m.classList.remove("active");
}

async function openSettingsModal() {
  await loadConfig();
  const m = document.getElementById("modalSettings");
  if (m) m.classList.add("active");
}

async function loadConfig() {
  try {
    const cfg = await fetchJson("/api/config");
    const nw = cfg.naverworks || {};
    safeSetVal("nwSenderEmail", nw.sender_email || "");
    safeSetVal("nwSenderPassword", nw.sender_password || "");
    safeSetVal("nwSenderName", nw.sender_name || "");

    renderVendorList(cfg.vendors || []);
  } catch (e) {
    console.error("Config load error:", e);
  }
}

function renderVendorList(vendors) {
  const container = document.getElementById("vendorList");
  if (!container) return;
  container.innerHTML = "";
  vendors.forEach((v) => {
    const row = document.createElement("div");
    row.className = "vendor-item-row";
    row.innerHTML = `
      <input type="text" class="form-input v-name" value="${escapeHtml(v.name || '')}" placeholder="거래처명" />
      <input type="email" class="form-input v-email" value="${escapeHtml(v.email || '')}" placeholder="이메일" />
      <input type="text" class="form-input v-phone" value="${escapeHtml(v.phone || '')}" placeholder="연락처" />
      <button class="btn-danger-sm" onclick="this.closest('.vendor-item-row').remove()"><i class="fa-solid fa-xmark"></i></button>
    `;
    container.appendChild(row);
  });
}

function addVendorRow() {
  const container = document.getElementById("vendorList");
  if (!container) return;
  const row = document.createElement("div");
  row.className = "vendor-item-row";
  row.innerHTML = `
    <input type="text" class="form-input v-name" placeholder="거래처명" />
    <input type="email" class="form-input v-email" placeholder="이메일" />
    <input type="text" class="form-input v-phone" placeholder="연락처" />
    <button class="btn-danger-sm" onclick="this.closest('.vendor-item-row').remove()"><i class="fa-solid fa-xmark"></i></button>
  `;
  container.appendChild(row);
}

async function saveConfig() {
  const vendors = [];
  document.querySelectorAll("#vendorList .vendor-item-row").forEach(row => {
    const nameEl = row.querySelector(".v-name");
    const emailEl = row.querySelector(".v-email");
    const phoneEl = row.querySelector(".v-phone");

    const name = nameEl ? nameEl.value.trim() : "";
    const email = emailEl ? emailEl.value.trim() : "";
    const phone = phoneEl ? phoneEl.value.trim() : "";
    if (name) {
      vendors.push({ name, email, phone });
    }
  });

  const senderName = safeGetVal("nwSenderName") || "발주담당자";
  const configData = {
    naverworks: {
      smtp_server: "smtp.worksmobile.com",
      smtp_port: 465,
      use_ssl: true,
      sender_email: safeGetVal("nwSenderEmail"),
      sender_password: safeGetVal("nwSenderPassword"),
      sender_name: senderName
    },
    company_info: {
      company_name: "(주)바이브컴퍼니",
      sender_name: senderName
    },
    vendors: vendors
  };

  try {
    const data = await fetchJson("/api/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(configData)
    });
    if (data.success) {
      showToast("설정이 성공적으로 저장되었습니다.", "success");
      closeModal("modalSettings");
    }
  } catch (err) {
    showToast(`저장 실패: ${err.message}`, "error");
  }
}

async function handleTestMail() {
  const testEmail = safeGetVal("testEmailInput");
  if (!testEmail) {
    showToast("테스트 메일을 수신할 이메일 주소를 입력해주세요.", "error");
    return;
  }

  await saveConfig();

  const btn = document.getElementById("btnTestMail");
  const originalHtml = btn ? btn.innerHTML : "";
  if (btn) {
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 전송 중...`;
    btn.disabled = true;
  }

  try {
    const data = await fetchJson("/api/test-naverworks", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to_email: testEmail })
    });

    showToast(data.message, "success");
  } catch (err) {
    showToast(`테스트 실패: ${err.message}`, "error");
  } finally {
    if (btn) {
      btn.innerHTML = `<i class="fa-solid fa-vial"></i> 테스트 전송`;
      btn.disabled = false;
    }
  }
}

// ----------------- Price Master & Template Modal ----------------- //

async function openPriceMasterModal() {
  const m = document.getElementById("modalPriceMaster");
  if (m) m.classList.add("active");
  const tbody = document.getElementById("priceMasterTableBody");
  if (!tbody) return;
  tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted" style="padding: 20px;"><i class="fa-solid fa-spinner fa-spin"></i> 단가표 불러오는 중...</td></tr>`;

  try {
    const data = await fetchJson("/api/price-master");
    if (data.success) {
      tbody.innerHTML = "";
      if (!data.items || data.items.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted" style="padding: 20px;">등록된 단가 데이터가 없습니다.</td></tr>`;
        return;
      }
      data.items.forEach(item => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td class="font-bold text-primary font-mono">${escapeHtml(item.model_name || item.item_code)}</td>
          <td class="text-muted font-mono text-sm">${escapeHtml(item.item_code || '')}</td>
          <td>${escapeHtml(item.item_name || '')}</td>
          <td class="text-sm">${escapeHtml(item.spec || '')}</td>
          <td class="text-right font-bold text-danger">${formatNumber(item.buy_price)}원</td>
          <td class="text-right font-bold text-primary">${formatNumber(item.sell_price)}원</td>
          <td class="text-muted text-sm">${escapeHtml(item.remark || '')}</td>
        `;
        tbody.appendChild(tr);
      });
    }
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-danger">단가표를 불러오지 못했습니다.</td></tr>`;
  }
}

async function handleOrderTemplateUpload(e) {
  const file = e.target.files ? e.target.files[0] : null;
  if (!file) return;

  const formData = new FormData();
  formData.append("file", file);

  try {
    showToast("발주서 양식 파일 업로드 중...", "info");
    const res = await fetch("/api/upload/order-template", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      showToast("발주서 템플릿(order_template.xlsx)이 성공적으로 교체되었습니다!", "success");
    } else {
      throw new Error(data.detail || "업로드 실패");
    }
  } catch (err) {
    showToast(`업로드 실패: ${err.message}`, "error");
  } finally {
    e.target.value = "";
  }
}

// ----------------- ERP Summary Modal & Clear ----------------- //

async function openErpModal() {
  const m = document.getElementById("modalErp");
  if (m) m.classList.add("active");
  const tbody = document.getElementById("erpRecentTableBody");
  if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted" style="padding: 20px;"><i class="fa-solid fa-spinner fa-spin"></i> 로딩 중...</td></tr>`;

  try {
    const data = await fetchJson("/api/erp-summary");

    const totalRows = data.total_rows || 0;
    safeSetText("erpRowCount", totalRows);
    safeSetText("modalErpTotalRows", totalRows);

    if (!tbody) return;
    tbody.innerHTML = "";
    if (!data.recent_items || data.recent_items.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted" style="padding: 20px;">저장된 ERP 발주 데이터가 없습니다.</td></tr>`;
      return;
    }

    data.recent_items.forEach(item => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="text-center text-sm">${escapeHtml(item.date)}</td>
        <td class="text-sm font-mono">${escapeHtml(item.order_no)}</td>
        <td class="font-bold">${escapeHtml(item.store || '')}</td>
        <td>${escapeHtml(item.item_name || '')}</td>
        <td class="text-primary font-mono font-bold">${escapeHtml(item.model_name || '')}</td>
        <td class="text-center font-bold">${formatNumber(item.qty)}</td>
        <td class="text-right">${formatNumber(item.buy_price)}원</td>
        <td class="text-right text-primary font-bold">${formatNumber(item.sell_price)}원</td>
        <td class="text-right text-success font-bold">${formatNumber(item.margin)}원</td>
      `;
      tbody.appendChild(tr);
    });
  } catch (err) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-danger">데이터를 불러오지 못했습니다.</td></tr>`;
  }
}

async function handleClearErpList() {
  if (!confirm("⚠️ 정말로 기존 ERP 누적 리스트의 모든 기록을 삭제하고 초기화하시겠습니까?")) {
    return;
  }

  try {
    const res = await fetch("/api/clear-erp-list", { method: "POST" });
    const data = await res.json();
    if (data.success) {
      showToast("ERP 누적 리스트가 성공적으로 초기화되었습니다.", "success");
      openErpModal();
      loadErpSummary();
    }
  } catch (err) {
    showToast("ERP 초기화 실패", "error");
  }
}

async function loadErpSummary() {
  try {
    const res = await fetch("/api/erp-summary");
    const data = await res.json();
    const totalRows = data.total_rows || 0;
    const countEl = document.getElementById("erpRowCount");
    const modalCountEl = document.getElementById("modalErpTotalRows");
    if (countEl) countEl.innerText = totalRows;
    if (modalCountEl) modalCountEl.innerText = totalRows;
  } catch (e) {
    console.error("loadErpSummary error:", e);
  }
}

// ----------------- Utilities ----------------- //

function showToast(msg, type = "info") {
  const toast = document.getElementById("toast");
  toast.innerText = msg;
  toast.className = `toast show toast-${type}`;
  setTimeout(() => {
    toast.className = "toast";
  }, 4000);
}

function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return "0";
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
