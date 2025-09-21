const dbName = "ExcelDB";
const storeName = "sheetData";
const stateKey = "rowStates";
const pagesKey = "filePages";
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

async function saveData(data, id) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readwrite");
    const store = tx.objectStore(storeName);
    const req = store.put({ id, data });
    req.onsuccess = () => { /* ok */ };
    req.onerror = (e) => reject(e.target.error);
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
    tx.onabort = (e) => reject(e.target.error);
  });
}

async function loadData(id) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readonly");
    const store = tx.objectStore(storeName);
    const req = store.get(id);
    req.onsuccess = () => resolve(req.result ? req.result.data : null);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function deleteData(id) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readwrite");
    const store = tx.objectStore(storeName);
    const req = store.delete(id);
    req.onsuccess = () => resolve();
    req.onerror = (e) => reject(e.target.error);
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

async function hashArrayBuffer(arrayBuffer) {
  if (!window.crypto || !crypto.subtle) {
    return "no-crypto";
  }
  const hashBuffer = await crypto.subtle.digest("SHA-256", arrayBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

function getCurrentFileKey() {
  return localStorage.getItem("currentFileKey") || null;
}

function loadAllStates() {
  try { return JSON.parse(localStorage.getItem(stateKey) || "{}"); } catch { return {}; }
}
function saveAllStates(obj) { localStorage.setItem(stateKey, JSON.stringify(obj)); }

function loadStatesForFile(fileKey) {
  const all = loadAllStates();
  return all[fileKey] || {};
}
function saveStatesForFile(fileKey, states) {
  const all = loadAllStates();
  all[fileKey] = states;
  saveAllStates(all);
}
function clearStatesForFile(fileKey) {
  const all = loadAllStates();
  delete all[fileKey];
  saveAllStates(all);
}

function loadPagesMap() {
  try { return JSON.parse(localStorage.getItem(pagesKey) || "{}"); } catch { return {}; }
}
function savePagesMap(map) { localStorage.setItem(pagesKey, JSON.stringify(map)); }

function getRowState(idx) {
  const fileKey = getCurrentFileKey() || "global";
  const states = loadStatesForFile(fileKey);
  return states[idx] || { messaged: false, bookmark: "" };
}

function setRowState(idx, partial) {
  const fileKey = getCurrentFileKey() || "global";
  const states = loadStatesForFile(fileKey);
  const prev = states[idx] || { messaged: false, bookmark: "" };
  states[idx] = { ...prev, ...partial };
  saveStatesForFile(fileKey, states);
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
  if (val === "start") row.classList.add("start-row");
  else if (val === "end") row.classList.add("end-row");
  else row.classList.add("none-row");

  row.classList.add("changed");
  setTimeout(() => row.classList.remove("changed"), 300);
}

function renderTable(page = 1) {
  currentPage = page;

  const currentFileKey = getCurrentFileKey();
  if (currentFileKey) {
    const pages = loadPagesMap();
    pages[currentFileKey] = currentPage;
    savePagesMap(pages);
  }

  localStorage.setItem("currentPage", currentPage);

  const tbody = document.querySelector("#dataTable tbody");
  if (!tbody) return;
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
    selected.textContent = state.bookmark ? (state.bookmark === "start" ? "Start" : "End") : "None";
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
      if (isLastTwoInPage) dropdown.classList.add("upward");
      else {
        const rect = list.getBoundingClientRect();
        if (rect.bottom > window.innerHeight) dropdown.classList.add("upward");
        else dropdown.classList.remove("upward");
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

      const visitedLinks = JSON.parse(localStorage.getItem("visitedLinks") || "[]");
      if (visitedLinks.includes(url)) {
        a.classList.add("visited");
      }

      a.addEventListener("click", () => {
        let links = JSON.parse(localStorage.getItem("visitedLinks") || "[]");
        if (!links.includes(url)) {
          links.push(url);
          localStorage.setItem("visitedLinks", JSON.stringify(links));
        }
        a.classList.add("visited");
      });

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
  if (!pag) return;
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
  const fileKey = getCurrentFileKey() || "global";
  const states = loadStatesForFile(fileKey);

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

  const counter = document.getElementById("updateMessageCount");
  if (counter) counter.textContent = total;

  const details = document.getElementById("countDetails");
  if (details) details.textContent = `Counting rows ${startIdx + 1} → ${endIdx + 1} • Checked: ${total}`;
}

document.getElementById("fileInput").addEventListener("change", (event) => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const arrayBuffer = e.target.result;
      const hash = await hashArrayBuffer(arrayBuffer);
      const fileKey = `${file.name}_${file.size}_${hash}`;

      localStorage.setItem("currentFileKey", fileKey);

      const data = new Uint8Array(arrayBuffer);
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
      await saveData(json, fileKey);

      const savedPages = loadPagesMap();
      currentPage = savedPages[fileKey] || 1;

      renderTable(currentPage);
    } catch (err) {
      console.error("Failed to load file:", err);
      alert("Failed to read file. See console for details.");
    }
  };
  reader.readAsArrayBuffer(file);
});

window.addEventListener("load", async () => {
  const currentFileKey = localStorage.getItem("currentFileKey");
  if (currentFileKey) {
    try {
      const saved = await loadData(currentFileKey);
      if (saved && Array.isArray(saved) && saved.length) {
        allData = saved;
      }
      const savedPages = loadPagesMap();
      currentPage = savedPages[currentFileKey] || 1;
    } catch (err) {
      console.warn("Could not load saved file from DB:", err);
      currentPage = 1;
    }
  } else {
    currentPage = 1;
  }
  renderTable(currentPage);
});

document.getElementById("clearStatesBtn").addEventListener("click", async () => {
  const fileKey = getCurrentFileKey() || "global";
  if (!fileKey) return;

  const confirmClear = await showConfirm("Clear all checkmarks & bookmarks for this file?");
  if (!confirmClear) return;

  clearStatesForFile(fileKey);
  renderTable(currentPage);

  await showInfo("All checkmarks and bookmarks were cleared for this file.");
});

document.getElementById("clearDataBtn").addEventListener("click", async () => {
  const fileKey = getCurrentFileKey();
  if (!fileKey) {
    const proceed = await showConfirm("No file loaded. Clear everything stored in DB/localStorage?");
    if (!proceed) return;
  }

  const confirmDelete = await showConfirm("Remove current file? (Table will be emptied, bookmarks/states will be kept)");
  if (!confirmDelete) return;

  try {
    if (fileKey) {
      await deleteData(fileKey);
      localStorage.removeItem("currentFileKey");
    }

    allData = [];
    renderTable(1);

    await showInfo("File removed. Bookmarks and checkmarks are saved.");
  } catch (err) {
    console.error("Failed to delete data:", err);
    await showInfo("Failed to delete data. See console.");
  }
});

function showConfirm(message) {
  return new Promise((resolve) => {
    const modal = document.getElementById("confirmModal");
    document.getElementById("confirmMessage").textContent = message;
    modal.classList.remove("hidden");

    const yesBtn = document.getElementById("confirmYes");
    const noBtn = document.getElementById("confirmNo");

    const cleanup = (result) => {
      modal.classList.add("hidden");
      yesBtn.onclick = null;
      noBtn.onclick = null;
      resolve(result);
    };

    yesBtn.onclick = () => cleanup(true);
    noBtn.onclick = () => cleanup(false);
  });
}

function showInfo(message) {
  return new Promise((resolve) => {
    const modal = document.getElementById("infoModal");
    document.getElementById("infoMessage").textContent = message;
    modal.classList.remove("hidden");

    const okBtn = document.getElementById("infoOk");
    okBtn.onclick = () => {
      modal.classList.add("hidden");
      resolve();
    };
  });
}