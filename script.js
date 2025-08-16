const dbName = "ExcelDB";
const storeName = "sheetData";
const stateKey = "rowStates";
const rowsPerPage = 100;

let currentPage = 1;
let allData = [];

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(dbName, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(storeName)) {
        db.createObjectStore(storeName, { keyPath: "id" });
      }
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function saveData(data) {
  const db = await openDB();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).put({ id: 1, data });
  return tx.complete;
}

async function loadData() {
  const db = await openDB();
  return new Promise((resolve) => {
    const tx = db.transaction(storeName, "readonly");
    const store = tx.objectStore(storeName);
    const req = store.get(1);
    req.onsuccess = () => resolve(req.result ? req.result.data : null);
    req.onerror = () => resolve(null);
  });
}

async function clearData() {
  const db = await openDB();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).delete(1);
  return tx.complete;
}

function loadStates() {
  try { return JSON.parse(localStorage.getItem(stateKey) || "{}"); }
  catch { return {}; }
}

function saveStates(states) {
  localStorage.setItem(stateKey, JSON.stringify(states));
}

function clearStates() {
  localStorage.removeItem(stateKey);
}

function getRowState(idx) {
  const states = loadStates();
  return states[idx] || { messaged: false, bookmark: "" };
}

function setRowState(idx, partial) {
  const states = loadStates();
  const prev = states[idx] || { messaged: false, bookmark: "" };
  states[idx] = { ...prev, ...partial };
  saveStates(states);
}

function setRowsInfo() {
  const info = document.getElementById("rowsInfo");
  if (!allData.length) {
    info.textContent = "";
    info.style.opacity = 0.75;
    return;
  }
  const totalPages = Math.max(1, Math.ceil(allData.length / rowsPerPage));
  const start = (currentPage - 1) * rowsPerPage + 1;
  const end = Math.min(currentPage * rowsPerPage, allData.length);
  info.textContent = `Rows ${start}–${end} of ${allData.length} • Page ${currentPage}/${totalPages}`;
  info.style.opacity = 0.8;
}

function applyRowStyle(row, val) {
  row.classList.remove("start-row", "end-row", "none-row");

  if (val === "start") {
    row.classList.add("start-row");
  } else if (val === "end") {
    row.classList.add("end-row");
  } else {
    row.classList.add("none-row");
  }

  row.classList.add("changed");
  setTimeout(() => row.classList.remove("changed"), 300);
}


function renderTable(page = 1) {
  currentPage = page;
  localStorage.setItem("currentPage", currentPage);
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  if (!allData.length) {
    setRowsInfo();
    updateMessageCount();
    renderPagination();
    return;
  }

  const start = (currentPage - 1) * rowsPerPage;
  const end = Math.min(start + rowsPerPage, allData.length);

  for (let i = start; i < end; i++) {
    const row = allData[i];
    const state = getRowState(i);

    const tr = document.createElement("tr");

    const tdNum = document.createElement("td");
    tdNum.className = "number-col";
    tdNum.textContent = i + 1;
    tr.appendChild(tdNum);

    const tdBm = document.createElement("td");

    const dropdown = document.createElement("div");
    dropdown.className = "customDropdown";
    dropdown.dataset.rowIndex = i;

    const selected = document.createElement("div");
    selected.className = "selected";
    selected.textContent = state.bookmark
      ? (state.bookmark === "start" ? "Start" : "End")
      : "None";

    selected.classList.remove("start-bg", "end-bg", "none-bg");
    if (state.bookmark === "start") selected.classList.add("start-bg");
    else if (state.bookmark === "end") selected.classList.add("end-bg");
    else selected.classList.add("none-bg");

    dropdown.appendChild(selected);

    const list = document.createElement("ul");
    list.className = "options";
    list.innerHTML = `
    <li data-value="">None</li>
    <li data-value="start" class="bmstart">Start</li>
    <li data-value="end" class="bmend">End</li>
  `;
    dropdown.appendChild(list);

    selected.addEventListener("click", () => {
      dropdown.classList.toggle("open");
      const isLastTwoInPage = i >= end - 2;

      if (isLastTwoInPage) {
        dropdown.classList.add("upward");
      } else {
        const rect = list.getBoundingClientRect();
        if (rect.bottom > window.innerHeight) {
          dropdown.classList.add("upward");
        } else {
          dropdown.classList.remove("upward");
        }
      }
    });
    list.querySelectorAll("li").forEach(li => {
      li.addEventListener("click", (e) => {
        const val = e.target.dataset.value;
        selected.textContent = e.target.textContent;
        dropdown.classList.remove("open");

        setRowState(i, { bookmark: val });
        applyRowStyle(tr, val);
        updateMessageCount();

        selected.classList.remove("start-bg", "end-bg", "none-bg");
        if (val === "start") selected.classList.add("start-bg");
        else if (val === "end") selected.classList.add("end-bg");
        else selected.classList.add("none-bg");
      });
    });

    tdBm.appendChild(dropdown);
    tr.appendChild(tdBm);

    const tdU = document.createElement("td"); tdU.textContent = row[0]; tr.appendChild(tdU);
    const tdF = document.createElement("td"); tdF.textContent = row[1]; tr.appendChild(tdF);
    const tdFy = document.createElement("td"); tdFy.textContent = row[2]; tr.appendChild(tdFy);
    const tdV = document.createElement("td"); tdV.textContent = row[3]; tr.appendChild(tdV);

    const tdMsg = document.createElement("td");
    tdMsg.classList.add("checkboxcell");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "messageCheckbox";
    cb.dataset.rowIndex = i;
    cb.checked = !!state.messaged;
    cb.addEventListener("change", (e) => {
      setRowState(i, { messaged: e.target.checked });
      updateMessageCount();
    });
    tdMsg.appendChild(cb);
    tr.appendChild(tdMsg);

    const tdLink = document.createElement("td");
    const url = (row[4] || "").toString().trim();
    if (url) {
      const a = document.createElement("a");
      a.href = url;
      a.target = "_blank";
      a.className = "link";
      a.textContent = url;
      tdLink.appendChild(a);
    }
    tr.appendChild(tdLink);

    applyRowStyle(tr, state.bookmark);
    tbody.appendChild(tr);
  }


  setRowsInfo();
  renderPagination();
  updateMessageCount();
}

function renderPagination() {
  const pag = document.getElementById("pagination");
  pag.innerHTML = "";

  if (!allData.length) return;

  const totalPages = Math.ceil(allData.length / rowsPerPage);

  const prev = document.createElement("button");
  prev.textContent = "‹ Prev";
  prev.disabled = currentPage === 1;
  prev.addEventListener("click", () => {
    renderTable(currentPage - 1);
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
  pag.appendChild(prev);

  const maxButtons = 9;
  const half = Math.floor(maxButtons / 2);
  let start = Math.max(1, currentPage - half);
  let end = Math.min(totalPages, start + maxButtons - 1);
  if (end - start + 1 < maxButtons) start = Math.max(1, end - maxButtons + 1);

  for (let p = start; p <= end; p++) {
    const btn = document.createElement("button");
    btn.textContent = p;
    btn.disabled = p === currentPage;
    btn.addEventListener("click", () => {
      renderTable(p);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    pag.appendChild(btn);
  }

  const next = document.createElement("button");
  next.textContent = "Next ›";
  next.disabled = currentPage === totalPages;
  next.addEventListener("click", () => {
    renderTable(currentPage + 1);
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
  pag.appendChild(next);
}

function updateMessageCount() {
  const states = loadStates();

  let lastStart = -1, lastEnd = -1;
  const totalRows = allData.length;

  for (let i = 0; i < totalRows; i++) {
    const s = states[i];
    if (!s) continue;
    if (s.bookmark === "start") lastStart = i;
  }

  for (let i = 0; i < totalRows; i++) {
    const s = states[i];
    if (!s) continue;
    if (s.bookmark === "end") {
      if (lastStart === -1 || i >= lastStart) lastEnd = i;
    }
  }

  const startIdx = lastStart >= 0 ? lastStart : 0;
  const endIdx = lastEnd >= startIdx ? lastEnd : totalRows - 1;

  let total = 0;
  for (let i = startIdx; i <= endIdx; i++) {
    if (states[i]?.messaged) total++;
  }

  document.getElementById("updateMessageCount").textContent = total;

  const details = document.getElementById("countDetails");
  details.textContent = `Counting rows ${startIdx + 1} → ${endIdx + 1} • Checked: ${total}`;
}

document.getElementById("fileInput").addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    let json = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: true });

    json = json.map(row => [
      row[1] ?? "",
      row[2] ?? "",
      row[3] ?? "",
      row[4] ?? "",
      row[5] ?? ""
    ]);

    allData = json;
    saveData(json);
    currentPage = 1;
    renderTable(currentPage);
  };
  reader.readAsArrayBuffer(file);
});


window.addEventListener("load", async () => {
  const saved = await loadData();
  if (saved && Array.isArray(saved) && saved.length) {
    allData = saved;
  }

  const savedPage = parseInt(localStorage.getItem("currentPage")) || 1;
  currentPage = savedPage;

  renderTable(currentPage);
});


document.getElementById("clearStatesBtn").addEventListener("click", () => {
  if (confirm("Clear all checkmarks & bookmarks?")) {
    clearStates();
    renderTable(currentPage);
  }
});

document.getElementById("clearDataBtn").addEventListener("click", async () => {
  if (confirm("Clear uploaded data? (Table will be emptied)")) {
    await clearData();
    allData = [];
    renderTable(1);
  }
});

const icon = document.getElementById("infoIcon");
const hiddenText = document.getElementById("hiddenText");

icon.addEventListener("click", () => {
  hiddenText.style.display = hiddenText.style.display === "block" ? "none" : "block";
});

document.addEventListener("click", (e) => {
  if (!icon.contains(e.target) && !hiddenText.contains(e.target)) {
    hiddenText.style.display = "none";
  }
});

document.querySelectorAll("#hiddenText .tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.tab;

    document.querySelectorAll("#hiddenText .tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll("#hiddenText .tab-content").forEach(c => c.classList.remove("active"));

    btn.classList.add("active");
    document.getElementById(target).classList.add("active");
  });
});

