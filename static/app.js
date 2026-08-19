// Global State
let currentOrderData = null;
let currentOrderFile = "";

// Sample Kakao Messages
const SAMPLES = {
  1: `[발주요청]
업체: 인쇄나라
품목:
BK-101 500개
BK-202 50개
받는분: 김철수
연락처: 010-3333-7777
주소: 서울특별시 마포구 월드컵북로 120 인쇄빌딩 3층
메모: 파손주의 및 오후 2시 이전 출고 요망`,

  2: `한빛지류 대리님 안녕하세요~
BK-102 300권이랑 PAPER-A4 20박스 급하게 발주 넣습니다!
수령인: 이영희 과장
전화: 010-8888-9999
납품처: 경기도 파주시 문발로 456 출판단지 A동 201호
특이사항: 거래명세서 동봉 부탁드립니다`,

  3: `동서제본 발주건입니다.
- BK-201 / 200개
- BK-103 / 50개
수령인: 박민수 (010-5555-1234)
배송지: 부산광역시 해운대구 센텀중앙로 78 12층
배송메모: 도착 전 기사님 연락 필수`
};

// DOM Ready
document.addEventListener("DOMContentLoaded", () => {
  initEventListeners();
  loadErpSummary();
  loadConfig();
});

function initEventListeners() {
  // Clear button
  document.getElementById("btnClear").addEventListener("click", () => {
    document.getElementById("kakaoInput").value = "";
    document.getElementById("step2Card").style.display = "none";
  });

  // STEP 1 Button: Create Order & ERP List
  document.getElementById("btnStep1Submit").addEventListener("click", handleStep1Submit);

  // STEP 2: Items Table Add Row
  document.getElementById("btnAddRow").addEventListener("click", handleAddTableRow);

  // STEP 2: Open Folder
  document.getElementById("btnOpenFolder").addEventListener("click", handleOpenFolder);

  // STEP 2: Open Edge Naver Works Webmail
  document.getElementById("btnOpenWebmail").addEventListener("click", handleOpenWebmail);

  // STEP 2: Send Mail (SMTP)
  document.getElementById("btnSendMail").addEventListener("click", handleSendMail);

  // Download buttons
  document.getElementById("btnDownloadOrder").addEventListener("click", () => {
    if (currentOrderFile) {
      window.open(`/api/download/order/${encodeURIComponent(currentOrderFile)}`, "_blank");
    }
  });

  document.getElementById("btnDownloadErp").addEventListener("click", () => {
    window.open("/api/download/erp-list", "_blank");
  });

  document.getElementById("btnModalDownloadErp").addEventListener("click", () => {
    window.open("/api/download/erp-list", "_blank");
  });

  // Header Modal Triggers
  document.getElementById("btnErpSummary").addEventListener("click", openErpModal);
  document.getElementById("btnOpenPriceMaster").addEventListener("click", openPriceMasterModal);
  document.getElementById("btnOpenSettings").addEventListener("click", openSettingsModal);

  // Settings Modal actions
  document.getElementById("btnSaveSettings").addEventListener("click", saveConfig);
  document.getElementById("btnTestMail").addEventListener("click", handleTestMail);
  document.getElementById("btnAddVendor").addEventListener("click", () => addVendorRow());

  // Price Master Modal actions
  document.getElementById("btnDownloadPriceMaster").addEventListener("click", () => {
    window.open("/api/download/price-master", "_blank");
  });

  document.getElementById("priceMasterFileInput").addEventListener("change", handlePriceMasterUpload);
}

// ----------------- STEP 1: Process Kakao Order ----------------- //

function loadSample(num) {
  document.getElementById("kakaoInput").value = SAMPLES[num] || "";
  showToast("샘플 카톡 메시지가 입력되었습니다.", "info");
}

async function handleStep1Submit() {
  const kakaoText = document.getElementById("kakaoInput").value.trim();
  if (!kakaoText) {
    showToast("카카오톡 주문 내용을 입력해주세요.", "error");
    return;
  }

  const btn = document.getElementById("btnStep1Submit");
  const originalHtml = btn.innerHTML;
  btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 분석 및 생성 중...`;
  btn.disabled = true;

  try {
    // 1. 카톡 파싱 API 호출
    const parseRes = await fetch("/api/parse-order", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: kakaoText })
    });
    const parseData = await parseRes.json();
    if (!parseRes.ok || !parseData.success) {
      throw new Error(parseData.detail || "주문 분석에 실패했습니다.");
    }

    const order = parseData.data;

    // 2. 발주서 및 ERP 리스트 생성 API 호출
    const createRes = await fetch("/api/create-order-and-erp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(order)
    });
    const createData = await createRes.json();
    if (!createRes.ok || !createData.success) {
      throw new Error(createData.detail || "발주서 엑셀 생성에 실패했습니다.");
    }

    currentOrderData = order;
    currentOrderFile = createData.order_file;

    // 3. UI 렌더링
    renderStep2(order, createData);
    loadErpSummary();

    showToast("🎉 발주서 엑셀 및 ERP 누적 리스트 생성이 완료되었습니다!", "success");

  } catch (err) {
    console.error(err);
    showToast(`오류: ${err.message}`, "error");
  } finally {
    btn.innerHTML = originalHtml;
    btn.disabled = false;
  }
}

// ----------------- STEP 2: Render & Review ----------------- //

function renderStep2(order, createData) {
  const step2Card = document.getElementById("step2Card");
  step2Card.style.display = "block";

  // Basic Info
  document.getElementById("orderVendor").value = order.vendor_name || "";
  document.getElementById("orderNo").value = order.order_no || "";
  document.getElementById("recName").value = order.recipient_name || "";
  document.getElementById("recPhone").value = order.recipient_phone || "";
  document.getElementById("recAddress").value = order.recipient_address || "";

  // File Name
  document.getElementById("orderFileName").innerText = createData.order_file || "";

  // Mail Draft
  const mailDraft = createData.mail_draft || {};
  document.getElementById("mailTo").value = mailDraft.to_email || order.vendor_email || "";
  document.getElementById("mailSubject").value = mailDraft.subject || "";
  document.getElementById("mailBody").value = mailDraft.body || "";

  // Render Items Table
  renderItemsTable(order.items || []);

  // Smooth scroll to Step 2
  step2Card.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderItemsTable(items) {
  const tbody = document.getElementById("itemsTableBody");
  tbody.innerHTML = "";

  if (!items || items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" class="text-center text-muted" style="padding: 20px;">등록된 품목이 없습니다. [품목 추가]를 눌러주세요.</td></tr>`;
    calculateTotals();
    return;
  }

  items.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index;

    const qty = parseInt(item.qty) || 0;
    const buyPrice = parseInt(item.buy_price) || 0;
    const sellPrice = parseInt(item.sell_price) || 0;
    const margin = (sellPrice - buyPrice) * qty;

    tr.innerHTML = `
      <td><input type="text" class="table-input item-code" value="${escapeHtml(item.item_code || '')}" placeholder="품번" /></td>
      <td><input type="text" class="table-input item-name" value="${escapeHtml(item.item_name || '')}" placeholder="품명" /></td>
      <td><input type="text" class="table-input item-spec" value="${escapeHtml(item.spec || '')}" placeholder="규격" /></td>
      <td><input type="number" class="table-input item-qty text-right" value="${qty}" min="1" /></td>
      <td><input type="number" class="table-input item-buy-price text-right" value="${buyPrice}" /></td>
      <td><input type="number" class="table-input item-sell-price text-right" value="${sellPrice}" /></td>
      <td class="text-right font-bold item-margin text-success">${formatNumber(margin)}원</td>
      <td class="text-center">
        <button class="btn-danger-sm" onclick="handleDeleteRow(this)"><i class="fa-solid fa-trash"></i></button>
      </td>
    `;

    // Add input event listeners for auto-calc
    tr.querySelectorAll("input").forEach(inp => {
      inp.addEventListener("input", () => handleRowInputChange(tr));
    });

    tbody.appendChild(tr);
  });

  calculateTotals();
}

function handleRowInputChange(tr) {
  const qty = parseInt(tr.querySelector(".item-qty").value) || 0;
  const buyPrice = parseInt(tr.querySelector(".item-buy-price").value) || 0;
  const sellPrice = parseInt(tr.querySelector(".item-sell-price").value) || 0;
  const margin = (sellPrice - buyPrice) * qty;

  const marginEl = tr.querySelector(".item-margin");
  marginEl.innerText = `${formatNumber(margin)}원`;
  if (margin < 0) {
    marginEl.className = "text-right font-bold item-margin text-danger";
  } else {
    marginEl.className = "text-right font-bold item-margin text-success";
  }

  calculateTotals();
}

function handleAddTableRow() {
  const tbody = document.getElementById("itemsTableBody");
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input type="text" class="table-input item-code" placeholder="예: BK-101" /></td>
    <td><input type="text" class="table-input item-name" placeholder="품명" /></td>
    <td><input type="text" class="table-input item-spec" placeholder="규격" /></td>
    <td><input type="number" class="table-input item-qty text-right" value="1" min="1" /></td>
    <td><input type="number" class="table-input item-buy-price text-right" value="0" /></td>
    <td><input type="number" class="table-input item-sell-price text-right" value="0" /></td>
    <td class="text-right font-bold item-margin text-success">0원</td>
    <td class="text-center">
      <button class="btn-danger-sm" onclick="handleDeleteRow(this)"><i class="fa-solid fa-trash"></i></button>
    </td>
  `;
  tr.querySelectorAll("input").forEach(inp => {
    inp.addEventListener("input", () => handleRowInputChange(tr));
  });
  tbody.appendChild(tr);
  calculateTotals();
}

function handleDeleteRow(btn) {
  const tr = btn.closest("tr");
  tr.remove();
  calculateTotals();
}

function calculateTotals() {
  const rows = document.querySelectorAll("#itemsTableBody tr");
  let totQty = 0;
  let totBuy = 0;
  let totSell = 0;
  let totMargin = 0;

  rows.forEach(tr => {
    const qtyInp = tr.querySelector(".item-qty");
    if (!qtyInp) return;
    const qty = parseInt(qtyInp.value) || 0;
    const buy = parseInt(tr.querySelector(".item-buy-price").value) || 0;
    const sell = parseInt(tr.querySelector(".item-sell-price").value) || 0;

    totQty += qty;
    totBuy += (qty * buy);
    totSell += (qty * sell);
  });

  totMargin = totSell - totBuy;

  document.getElementById("totQty").innerText = formatNumber(totQty);
  document.getElementById("totBuyAmount").innerText = `${formatNumber(totBuy)}원`;
  document.getElementById("totSellAmount").innerText = `${formatNumber(totSell)}원`;
  document.getElementById("totMargin").innerText = `${formatNumber(totMargin)}원`;
}

function gatherCurrentOrderData() {
  const items = [];
  const rows = document.querySelectorAll("#itemsTableBody tr");
  rows.forEach(tr => {
    const codeInp = tr.querySelector(".item-code");
    if (!codeInp) return;
    const code = codeInp.value.trim();
    if (!code) return;

    items.push({
      item_code: code,
      item_name: tr.querySelector(".item-name").value.trim(),
      spec: tr.querySelector(".item-spec").value.trim(),
      qty: parseInt(tr.querySelector(".item-qty").value) || 1,
      buy_price: parseInt(tr.querySelector(".item-buy-price").value) || 0,
      sell_price: parseInt(tr.querySelector(".item-sell-price").value) || 0,
      remark: ""
    });
  });

  return {
    vendor_name: document.getElementById("orderVendor").value.trim(),
    order_no: document.getElementById("orderNo").value.trim(),
    recipient_name: document.getElementById("recName").value.trim(),
    recipient_phone: document.getElementById("recPhone").value.trim(),
    recipient_address: document.getElementById("recAddress").value.trim(),
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
  btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 반영 중...`;
  btn.disabled = true;

  try {
    const res = await fetch("/api/create-order-and-erp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(updatedOrder)
    });
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.detail || "재생성 실패");

    currentOrderFile = data.order_file;
    document.getElementById("orderFileName").innerText = data.order_file;
    loadErpSummary();
    showToast("수정된 내용으로 발주서 및 ERP 리스트가 재반영되었습니다.", "success");
  } catch (err) {
    showToast(`오류: ${err.message}`, "error");
  } finally {
    btn.innerHTML = `<i class="fa-solid fa-arrows-rotate"></i> 엑셀 내용 재반영`;
    btn.disabled = false;
  }
}

// ----------------- STEP 2: Open Folder & Webmail (Edge) ----------------- //

async function handleOpenFolder() {
  try {
    await fetch("/api/open-order-folder", { method: "POST" });
    showToast("📂 윈도우 탐색기에서 발주서 폴더가 열렸습니다.", "info");
  } catch (e) {
    showToast("폴더 열기 실패", "error");
  }
}

async function handleOpenWebmail() {
  const toEmail = document.getElementById("mailTo").value.trim();
  const subject = document.getElementById("mailSubject").value.trim();
  const body = document.getElementById("mailBody").value.trim();

  // 1. 발주서 폴더 탐색기로 열기
  await handleOpenFolder();

  // 2. 메일 본문 클립보드에 자동 복사
  if (navigator.clipboard && body) {
    try {
      await navigator.clipboard.writeText(body);
      showToast("📋 메일 본문이 클립보드에 복사되었습니다!", "info");
    } catch (err) {
      console.log("Clipboard write error:", err);
    }
  }

  // 3. 네이버웍스 웹메일 작성 URL 새 탭으로 열기
  // (Edge에 이미 로그인되어 있으므로 바로 작성 화면으로 이동)
  const naverWorksUrl = `https://mail.worksmobile.com/`;
  window.open(naverWorksUrl, "_blank");

  showToast("🌐 Edge 네이버웍스 새 탭이 열렸습니다. 발주서 파일을 끌어다 첨부 후 발송하세요!", "success");
}

// ----------------- STEP 2: Send Mail via Naver Works (SMTP) ----------------- //

async function handleSendMail() {
  const toEmail = document.getElementById("mailTo").value.trim();
  const subject = document.getElementById("mailSubject").value.trim();
  const body = document.getElementById("mailBody").value.trim();

  if (!toEmail) {
    showToast("받는 사람(거래처) 이메일 주소를 입력해주세요.", "error");
    document.getElementById("mailTo").focus();
    return;
  }
  if (!currentOrderFile) {
    showToast("첨부할 발주서 파일이 없습니다. 1차 생성을 먼저 진행해주세요.", "error");
    return;
  }

  const btn = document.getElementById("btnSendMail");
  const originalHtml = btn.innerHTML;
  btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 네이버웍스 전송 중...`;
  btn.disabled = true;

  try {
    const res = await fetch("/api/send-mail", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to_email: toEmail,
        subject: subject,
        body: body,
        attachment_filename: currentOrderFile
      })
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      throw new Error(data.detail || data.message || "메일 발송에 실패했습니다.");
    }

    showToast(`✉️ ${data.message || '발주서 메일이 성공적으로 전송되었습니다!'}`, "success");
  } catch (err) {
    console.error(err);
    showToast(`메일 발송 실패: ${err.message}`, "error");
  } finally {
    btn.innerHTML = originalHtml;
    btn.disabled = false;
  }
}

// ----------------- Modals & Settings ----------------- //

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove("active");
}

async function openSettingsModal() {
  await loadConfig();
  document.getElementById("modalSettings").classList.add("active");
}

async function loadConfig() {
  try {
    const res = await fetch("/api/config");
    const cfg = await res.json();

    const nw = cfg.naverworks || {};
    document.getElementById("nwSenderEmail").value = nw.sender_email || "";
    document.getElementById("nwSenderPassword").value = nw.sender_password || "";
    document.getElementById("nwSenderName").value = nw.sender_name || "";

    renderVendorList(cfg.vendors || []);
  } catch (e) {
    console.error("Config load error:", e);
  }
}

function renderVendorList(vendors) {
  const container = document.getElementById("vendorList");
  container.innerHTML = "";
  vendors.forEach((v, idx) => {
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
    const name = row.querySelector(".v-name").value.trim();
    const email = row.querySelector(".v-email").value.trim();
    const phone = row.querySelector(".v-phone").value.trim();
    if (name) {
      vendors.push({ name, email, phone });
    }
  });

  const configData = {
    naverworks: {
      smtp_server: "smtp.worksmobile.com",
      smtp_port: 465,
      use_ssl: true,
      sender_email: document.getElementById("nwSenderEmail").value.trim(),
      sender_password: document.getElementById("nwSenderPassword").value.trim(),
      sender_name: document.getElementById("nwSenderName").value.trim()
    },
    company_info: {
      company_name: "(주)바이브컴퍼니",
      sender_name: document.getElementById("nwSenderName").value.trim() || "발주담당자"
    },
    vendors: vendors
  };

  try {
    const res = await fetch("/api/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(configData)
    });
    const data = await res.json();
    if (data.success) {
      showToast("설정이 성공적으로 저장되었습니다.", "success");
      closeModal("modalSettings");
    }
  } catch (err) {
    showToast(`저장 실패: ${err.message}`, "error");
  }
}

async function handleTestMail() {
  const testEmail = document.getElementById("testEmailInput").value.trim();
  if (!testEmail) {
    showToast("테스트 메일을 수신할 이메일 주소를 입력해주세요.", "error");
    return;
  }

  // 먼저 현재 폼의 네이버웍스 계정 정보 저장
  await saveConfig();

  const btn = document.getElementById("btnTestMail");
  btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 전송 중...`;
  btn.disabled = true;

  try {
    const res = await fetch("/api/test-naverworks", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to_email: testEmail })
    });
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.detail || "테스트 실패");

    showToast(data.message, "success");
  } catch (err) {
    showToast(`테스트 실패: ${err.message}`, "error");
  } finally {
    btn.innerHTML = `<i class="fa-solid fa-vial"></i> 테스트 전송`;
    btn.disabled = false;
  }
}

// ----------------- Price Master Modal ----------------- //

async function openPriceMasterModal() {
  document.getElementById("modalPriceMaster").classList.add("active");
  const tbody = document.getElementById("priceMasterTableBody");
  tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted" style="padding: 20px;"><i class="fa-solid fa-spinner fa-spin"></i> 단가표 불러오는 중...</td></tr>`;

  try {
    const res = await fetch("/api/price-master");
    const data = await res.json();
    if (data.success) {
      tbody.innerHTML = "";
      data.items.forEach(item => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td class="font-bold text-primary">${escapeHtml(item.item_code)}</td>
          <td>${escapeHtml(item.item_name || '')}</td>
          <td>${escapeHtml(item.spec || '')}</td>
          <td class="text-right text-danger">${formatNumber(item.buy_price)}원</td>
          <td class="text-right text-primary">${formatNumber(item.sell_price)}원</td>
          <td class="text-muted text-sm">${escapeHtml(item.remark || '')}</td>
        `;
        tbody.appendChild(tr);
      });
    }
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center text-danger">단가표를 불러오지 못했습니다.</td></tr>`;
  }
}

async function handlePriceMasterUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append("file", file);

  try {
    showToast("단가표 파일 업로드 중...", "info");
    const res = await fetch("/api/upload/price-master", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      showToast("단가표가 성공적으로 업데이트되었습니다!", "success");
      openPriceMasterModal();
    } else {
      throw new Error(data.detail || "업로드 실패");
    }
  } catch (err) {
    showToast(`업로드 실패: ${err.message}`, "error");
  } finally {
    e.target.value = "";
  }
}

// ----------------- ERP Summary Modal ----------------- //

async function openErpModal() {
  document.getElementById("modalErp").classList.add("active");
  const tbody = document.getElementById("erpRecentTableBody");
  tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted" style="padding: 20px;"><i class="fa-solid fa-spinner fa-spin"></i> 로딩 중...</td></tr>`;

  try {
    const res = await fetch("/api/erp-summary");
    const data = await res.json();

    document.getElementById("modalErpTotalRows").innerText = data.total_rows || 0;

    tbody.innerHTML = "";
    if (!data.recent_items || data.recent_items.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted" style="padding: 20px;">누적된 ERP 데이터가 없습니다.</td></tr>`;
      return;
    }

    data.recent_items.forEach(item => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="text-sm">${escapeHtml(item.date)}</td>
        <td class="text-sm font-mono">${escapeHtml(item.order_no)}</td>
        <td class="font-bold">${escapeHtml(item.vendor)}</td>
        <td class="text-primary">${escapeHtml(item.code)}</td>
        <td>${escapeHtml(item.name)}</td>
        <td class="text-right">${formatNumber(item.qty)}</td>
        <td class="text-right">${formatNumber(item.buy_price)}원</td>
        <td class="text-right">${formatNumber(item.sell_price)}원</td>
        <td class="text-right font-bold text-success">${formatNumber(item.margin)}원</td>
      `;
      tbody.appendChild(tr);
    });
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-danger">데이터를 불러오지 못했습니다.</td></tr>`;
  }
}

async function loadErpSummary() {
  try {
    const res = await fetch("/api/erp-summary");
    const data = await res.json();
    document.getElementById("erpRowCount").innerText = data.total_rows || 0;
  } catch (e) {
    console.error(e);
  }
}

// ----------------- Utilities ----------------- //

function showToast(msg, type = "info") {
  const toast = document.getElementById("toast");
  toast.innerText = msg;
  toast.className = `toast show ${type}`;
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
