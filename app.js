/**
 * 공공임대주택 차량관리 시스템 - Core Logic
 * Nature Jeju Theme 적용
 */

// --- 기본 데이터 (최초 실행 시에만 사용) ---
const DEFAULT_DATA = [
  {
    id: "1",
    name: "홍길동",
    dong: "101",
    ho: "502",
    model: "그랜저",
    carNumber: "12가 3456",
  },
  {
    id: "2",
    name: "김철수",
    dong: "102",
    ho: "101",
    model: "테슬라",
    carNumber: "78나 9012",
  },
  {
    id: "3",
    name: "이영희",
    dong: "101",
    ho: "304",
    model: "아반떼",
    carNumber: "34다 5678",
  },
];

// --- 데이터 영속화 (localStorage) ---
const STORAGE_KEY = "jpdc_vehicle_data";

function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) return JSON.parse(saved);
  } catch (e) {
    console.warn("저장된 데이터를 불러오지 못했습니다.", e);
  }
  return JSON.parse(JSON.stringify(DEFAULT_DATA)); // 기본 데이터 복제
}

function saveData() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(vehicleData));
  } catch (e) {
    console.warn("데이터 저장에 실패했습니다.", e);
  }
}

let vehicleData = loadData();

// --- State Management ---
const state = {
  currentPage: "page-search",
};

// --- DOM Elements ---
const pages = document.querySelectorAll(".page");
const carTableBody = document.querySelector("#car-table tbody");

// --- QR에 담을 접속 URL 생성 ---
// QR을 스캔하면 이 URL로 접속되고, ?vid= 파라미터로 자동 조회된다.
function buildQrUrl(id) {
  const base = location.href.split("?")[0].split("#")[0];
  return `${base}?vid=${encodeURIComponent(id)}`;
}

// --- 스캔/입력 텍스트로 차량 찾기 ---
// URL(?vid=), 순수 id, 차량번호 모두 대응
function resolveVehicle(text) {
  if (!text) return null;
  let key = String(text).trim();

  // URL 형태이면 vid 파라미터 추출
  try {
    const url = new URL(key);
    const vid = url.searchParams.get("vid");
    if (vid) key = vid;
  } catch (e) {
    // URL이 아니면 그대로 사용
  }

  return (
    vehicleData.find((v) => v.id === key) ||
    vehicleData.find((v) => v.carNumber === key) ||
    vehicleData.find((v) => v.carNumber.includes(key)) ||
    null
  );
}

// --- Navigation Logic ---
function navigateTo(pageId) {
  pages.forEach((page) => page.classList.remove("active"));
  const targetPage = document.getElementById(pageId);
  if (targetPage) {
    targetPage.classList.add("active");
    state.currentPage = pageId;
  }
  // 페이지 이동 시 스캐너가 열려 있으면 닫기
  stopScanner();
}

// --- Search Logic ---
function searchVehicle(text) {
  const vehicle = resolveVehicle(text);
  if (vehicle) {
    showResult(vehicle);
  } else {
    alert("해당 차량 정보를 찾을 수 없습니다.");
  }
}

function showResult(vehicle) {
  document.getElementById("res-car-number").textContent = vehicle.carNumber;
  document.getElementById("res-name").textContent = vehicle.name;
  document.getElementById(
    "res-address"
  ).textContent = `${vehicle.dong}동 ${vehicle.ho}호`;
  document.getElementById("res-model").textContent = vehicle.model || "-";
  navigateTo("page-result");
}

// --- Admin Logic: Table Rendering ---
function renderAdminTable(filter = "") {
  carTableBody.innerHTML = "";
  const filteredData = vehicleData.filter(
    (v) => v.carNumber.includes(filter) || v.name.includes(filter)
  );

  if (filteredData.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6" class="empty-row">등록된 차량이 없습니다.</td>`;
    carTableBody.appendChild(tr);
    return;
  }

  filteredData.forEach((v) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
            <td>${v.name}</td>
            <td>${v.dong}-${v.ho}</td>
            <td>${v.model || "-"}</td>
            <td>${v.carNumber}</td>
            <td><div class="qr-cell" id="qr-${v.id}" title="클릭하여 저장"></div></td>
            <td><button class="btn-link btn-delete" onclick="deleteVehicle('${v.id}')">삭제</button></td>
        `;
    carTableBody.appendChild(tr);

    const qrContainer = document.getElementById(`qr-${v.id}`);
    qrContainer.innerHTML = "";
    // QR에는 접속 URL을 담아 스캔 시 자동 조회되게 한다.
    new QRCode(qrContainer, {
      text: buildQrUrl(v.id),
      width: 96,
      height: 96,
      correctLevel: QRCode.CorrectLevel.H,
    });

    // 캔버스가 생성될 때까지 반복 확인 후 로고 합성 및 이벤트 바인딩
    const checkCanvas = setInterval(() => {
      const canvas = qrContainer.querySelector("canvas");
      if (canvas) {
        clearInterval(checkCanvas);
        window.addLogoToCanvas(canvas, "JPDC");

        qrContainer.onclick = () => {
          window.saveCanvasAsImage(canvas, `QR_${v.carNumber}.png`);
        };
        qrContainer.style.cursor = "pointer";
      }
    }, 50);

    setTimeout(() => clearInterval(checkCanvas), 2000);
  });
}

function deleteVehicle(id) {
  if (confirm("정말 삭제하시겠습니까?")) {
    vehicleData = vehicleData.filter((v) => v.id !== id);
    saveData();
    renderAdminTable(document.getElementById("admin-search").value);
  }
}

// --- Excel Upload Logic (SheetJS) ---
function handleExcelUpload(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(worksheet);

      const newVehicles = json
        .map((row, index) => {
          const id = Date.now().toString() + index;
          return {
            id: id,
            name: row["입주민 성함"] || row["성함"] || "미상",
            dong: String(row["동"] || ""),
            ho: String(row["호"] || ""),
            model: row["차종"] || "",
            carNumber: String(row["차량번호"] || "").trim(),
          };
        })
        .filter((v) => v.carNumber);

      if (newVehicles.length > 0) {
        vehicleData = [...vehicleData, ...newVehicles];
        saveData();
        alert(`${newVehicles.length}건의 데이터가 성공적으로 업로드되었습니다.`);
        navigateTo("page-admin");
        renderAdminTable();
      } else {
        alert("업로드 가능한 유효한 데이터가 없습니다. 양식을 확인해주세요.");
      }
    } catch (err) {
      console.error(err);
      alert("엑셀 파일을 읽는 중 오류가 발생했습니다.");
    }
  };
  reader.readAsArrayBuffer(file);
}

// --- Excel Template Download ---
function downloadTemplate() {
  const sample = [
    {
      "입주민 성함": "홍길동",
      동: "101",
      호: "502",
      차종: "그랜저",
      차량번호: "12가 3456",
    },
  ];
  const ws = XLSX.utils.json_to_sheet(sample);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "차량목록");
  XLSX.writeFile(wb, "차량등록_양식.xlsx");
}

// --- QR 로고 합성 (전역) ---
window.addLogoToCanvas = function (canvas, text) {
  const ctx = canvas.getContext("2d");
  const size = canvas.width;

  const logoSize = size * 0.24;
  const x = (size - logoSize) / 2;
  const y = (size - logoSize) / 2;

  ctx.fillStyle = "#FFFFFF";
  if (ctx.roundRect) {
    ctx.beginPath();
    ctx.roundRect(x, y, logoSize, logoSize, 4);
    ctx.fill();
  } else {
    ctx.fillRect(x, y, logoSize, logoSize);
  }

  ctx.fillStyle = "#1A365D";
  ctx.font = `bold ${logoSize * 0.32}px 'Pretendard', Arial`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.shadowColor = "rgba(0,0,0,0.1)";
  ctx.shadowBlur = 2;
  ctx.fillText(text, size / 2, size / 2);
  ctx.shadowBlur = 0;
};

// --- QR PNG 저장 (전역) ---
window.saveCanvasAsImage = function (canvas, filename) {
  const link = document.createElement("a");
  link.download = filename;
  link.href = canvas.toDataURL("image/png");
  link.click();
};

// --- QR 카메라 스캐너 ---
let html5QrScanner = null;

function startScanner() {
  const overlay = document.getElementById("scanner-overlay");
  const msg = document.getElementById("scanner-msg");

  if (typeof Html5Qrcode === "undefined") {
    // 라이브러리 로드 실패 시(오프라인 등) 수동 입력으로 대체
    const manual = prompt(
      "카메라 스캐너를 사용할 수 없습니다.\n차량번호를 직접 입력하세요:"
    );
    if (manual) searchVehicle(manual);
    return;
  }

  overlay.hidden = false;
  msg.textContent = "카메라를 QR 코드에 비춰주세요";
  html5QrScanner = new Html5Qrcode("reader");

  html5QrScanner
    .start(
      { facingMode: "environment" },
      { fps: 10, qrbox: { width: 220, height: 220 } },
      (decodedText) => {
        stopScanner();
        searchVehicle(decodedText);
      },
      () => {} // 스캔 실패 프레임은 무시
    )
    .catch((err) => {
      console.error(err);
      msg.textContent =
        "카메라를 열 수 없습니다. (HTTPS 또는 카메라 권한이 필요합니다)";
    });
}

function stopScanner() {
  const overlay = document.getElementById("scanner-overlay");
  if (html5QrScanner && html5QrScanner.isScanning) {
    html5QrScanner
      .stop()
      .then(() => html5QrScanner.clear())
      .catch(() => {});
  }
  html5QrScanner = null;
  if (overlay) overlay.hidden = true;
}

// --- Event Listeners ---
document.addEventListener("DOMContentLoaded", () => {
  // Navigation
  document
    .getElementById("btn-go-admin")
    .addEventListener("click", () => navigateTo("page-admin"));
  document
    .getElementById("btn-back-home")
    .addEventListener("click", () => navigateTo("page-search"));
  document
    .getElementById("btn-back-home-admin")
    .addEventListener("click", () => navigateTo("page-search"));
  document
    .getElementById("btn-back-search")
    .addEventListener("click", () => navigateTo("page-search"));
  document
    .getElementById("btn-go-save")
    .addEventListener("click", () => navigateTo("page-save"));

  // Search
  document.getElementById("btn-search").addEventListener("click", () => {
    const val = document.getElementById("input-car-number").value;
    searchVehicle(val);
  });
  document
    .getElementById("input-car-number")
    .addEventListener("keydown", (e) => {
      if (e.key === "Enter") searchVehicle(e.target.value);
    });

  // QR Scan (실제 카메라)
  document
    .getElementById("btn-qr-scan")
    .addEventListener("click", startScanner);
  document
    .getElementById("btn-close-scanner")
    .addEventListener("click", stopScanner);

  // Manual Input + QR Preview
  const carForm = document.getElementById("form-car-input");
  const qrPreviewContainer = document.getElementById("qr-preview-container");
  const qrcodeElement = document.getElementById("qrcode");
  const previewQr = new QRCode(qrcodeElement, {
    width: 128,
    height: 128,
    correctLevel: QRCode.CorrectLevel.H,
  });

  carForm.addEventListener("input", () => {
    const carNumber = carForm.elements["carNumber"].value.trim();
    if (carNumber) {
      qrPreviewContainer.style.display = "block";
      // 미리보기 QR도 실제 접속 URL(차량번호 기준)로 생성
      previewQr.makeCode(buildQrUrl(carNumber));
      setTimeout(() => {
        const canvas = qrcodeElement.querySelector("canvas");
        if (canvas) window.addLogoToCanvas(canvas, "JPDC");
      }, 50);
    } else {
      qrPreviewContainer.style.display = "none";
    }
  });

  carForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const id = Date.now().toString();
    const newVehicle = {
      id: id,
      name: formData.get("name"),
      dong: formData.get("dong"),
      ho: formData.get("ho"),
      model: formData.get("model"),
      carNumber: formData.get("carNumber").trim(),
    };
    vehicleData.push(newVehicle);
    saveData();
    alert("차량 정보가 저장되었습니다.");
    e.target.reset();
    qrPreviewContainer.style.display = "none";
    navigateTo("page-admin");
    renderAdminTable();
  });

  // Excel Upload (클릭 + 드래그앤드롭)
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-upload");

  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) handleExcelUpload(e.target.files[0]);
  });
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", () =>
    dropZone.classList.remove("dragover")
  );
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0)
      handleExcelUpload(e.dataTransfer.files[0]);
  });

  // Template Download
  document
    .getElementById("btn-download-template")
    .addEventListener("click", downloadTemplate);

  // Admin Search
  document.getElementById("admin-search").addEventListener("input", (e) => {
    renderAdminTable(e.target.value);
  });

  // Initial Render
  renderAdminTable();

  // URL Parameter Handling (QR Direct Access)
  const urlParams = new URLSearchParams(window.location.search);
  const vid = urlParams.get("vid");
  if (vid) {
    const vehicle = resolveVehicle(vid);
    if (vehicle) showResult(vehicle);
  }
});
