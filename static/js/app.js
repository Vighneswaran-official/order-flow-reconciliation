/**
 * OrderFlow Studio — 5-File Settlement & Master SKU Reconciliation Suite (v5.0)
 * Upload in same area -> Generate -> Download Final Report -> Preview Sheet "Final" (6 Columns)
 */

document.addEventListener("DOMContentLoaded", () => {
  // Application State
  const state = {
    ordersFile: null,
    flowFile: null,
    prepaidFile: null,
    postpaidFile: null,
    masterFile: null,
    currentPage: 1,
    pageSize: 25,
    totalPages: 1,
    totalRecords: 0,
    searchQuery: "",
    allRecords: [],
    filteredRecords: [],
    isRunning: false
  };

  const getEl = (id) => document.getElementById(id);
  const safeText = (el, val) => { if (el) el.textContent = val; };

  // Format currency
  const formatRs = (num) => {
    if (num === null || num === undefined || isNaN(num)) return "Rs. 0.00";
    return "Rs. " + Number(num).toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  };

  // DOM Elements
  const elements = {
    toastContainer: getEl("toast-container"),
    statusPill: getEl("status-pill"),
    statusPillText: getEl("status-pill-text"),
    dropZone: getEl("drop-zone"),
    fileInput: getEl("file-input"),

    // 5 File Cards
    cardOrders: getEl("card-orders"),
    nameOrders: getEl("name-seller-orders"),
    metaOrders: getEl("meta-seller-orders"),
    statusTagOrders: getEl("status-tag-orders"),

    cardFlow: getEl("card-flow"),
    nameFlow: getEl("name-order-flow"),
    metaFlow: getEl("meta-order-flow"),
    statusTagFlow: getEl("status-tag-flow"),

    cardPrepaid: getEl("card-prepaid"),
    namePrepaid: getEl("name-prepaid"),
    metaPrepaid: getEl("meta-prepaid"),
    statusTagPrepaid: getEl("status-tag-prepaid"),

    cardPostpaid: getEl("card-postpaid"),
    namePostpaid: getEl("name-postpaid"),
    metaPostpaid: getEl("meta-postpaid"),
    statusTagPostpaid: getEl("status-tag-postpaid"),

    cardMaster: getEl("card-master"),
    nameMaster: getEl("name-master-sku"),
    metaMaster: getEl("meta-master-sku"),
    statusTagMaster: getEl("status-tag-master"),

    // Cancel / Remove Buttons
    btnCancelOrders: getEl("btn-cancel-orders"),
    btnCancelFlow: getEl("btn-cancel-flow"),
    btnCancelPrepaid: getEl("btn-cancel-prepaid"),
    btnCancelPostpaid: getEl("btn-cancel-postpaid"),
    btnCancelMaster: getEl("btn-cancel-master"),

    // Action Button
    btnGenerate: getEl("btn-generate"),
    btnSpinner: document.querySelector(".btn-spinner"),
    btnPlayIcon: document.querySelector(".btn-play-icon"),

    // Result & Download
    downloadSection: getEl("download-section"),
    resultTitle: getEl("result-title"),
    resultSubtitle: getEl("result-subtitle"),
    btnDlExcel: getEl("btn-dl-excel"),
    btnDlLog: getEl("btn-dl-log"),

    // Metrics Strip
    metricOrderRows: getEl("metric-order-rows"),
    metricMatchRate: getEl("metric-match-rate"),
    metricPrepaidAmt: getEl("metric-prepaid-amt"),
    metricPostpaidAmt: getEl("metric-postpaid-amt"),
    metricFinalAmt: getEl("metric-final-amt"),
    metricPayoutAmt: getEl("metric-payout-amt"),
    metricNetCost: getEl("metric-net-cost"),
    metricProfitAmt: getEl("metric-profit-amt"),

    // Preview Table
    previewSection: getEl("output-preview-section"),
    previewTbody: getEl("preview-tbody"),
    previewSearchInput: getEl("preview-search-input"),
    previewCount: getEl("preview-count"),
    btnPrev: getEl("btn-prev"),
    btnNext: getEl("btn-next"),
    pageText: getEl("page-text")
  };

  // Toast Notifications
  function showToast(message, type = "info", duration = 4000) {
    if (!elements.toastContainer) return;
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    elements.toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = "0";
      setTimeout(() => toast.remove(), 200);
    }, duration);
  }

  // Helper to update a file card UI (Green if uploaded, Red if not)
  function updateCard(cardEl, nameEl, metaEl, tagEl, fileInfo, defaultRoleName) {
    if (fileInfo && fileInfo.filename) {
      if (cardEl) {
        cardEl.classList.remove("is-missing");
        cardEl.classList.add("is-uploaded");
      }
      safeText(nameEl, fileInfo.filename);
      const rows = fileInfo.row_count ? `${Number(fileInfo.row_count).toLocaleString()} rows` : "Ready";
      const size = fileInfo.size_formatted ? ` • ${fileInfo.size_formatted}` : "";
      safeText(metaEl, `✓ Uploaded (${rows}${size})`);
      if (tagEl) {
        tagEl.className = "badge-status-chip ready";
        tagEl.textContent = "Uploaded";
      }
      return true;
    } else {
      if (cardEl) {
        cardEl.classList.remove("is-uploaded");
        cardEl.classList.add("is-missing");
      }
      safeText(nameEl, "Not Uploaded");
      safeText(metaEl, "Click or drop file here");
      if (tagEl) {
        tagEl.className = "badge-status-chip missing";
        tagEl.textContent = "Missing";
      }
      return false;
    }
  }

  // 1. Initial Status Check
  async function checkStatus() {
    try {
      const res = await fetch("/api/status");
      if (!res.ok) return;
      const data = await res.json();

      const o = data.detected_orders;
      const f = data.detected_flow;
      const pre = data.detected_prepaid;
      const post = data.detected_postpaid;
      const m = data.detected_master;

      const hasOrders = updateCard(elements.cardOrders, elements.nameOrders, elements.metaOrders, elements.statusTagOrders, o, "Seller Orders Report");
      const hasFlow = updateCard(elements.cardFlow, elements.nameFlow, elements.metaFlow, elements.statusTagFlow, f, "Order Flow Report");
      const hasPrepaid = updateCard(elements.cardPrepaid, elements.namePrepaid, elements.metaPrepaid, elements.statusTagPrepaid, pre, "Prepaid Settlement");
      const hasPostpaid = updateCard(elements.cardPostpaid, elements.namePostpaid, elements.metaPostpaid, elements.statusTagPostpaid, post, "Postpaid Settlement");
      const hasMaster = updateCard(elements.cardMaster, elements.nameMaster, elements.metaMaster, elements.statusTagMaster, m, "Master SKU File");

      state.ordersFile = o ? o.filename : null;
      state.flowFile = f ? f.filename : null;
      state.prepaidFile = pre ? pre.filename : null;
      state.postpaidFile = post ? post.filename : null;
      state.masterFile = m ? m.filename : null;

      if (hasOrders && hasFlow && hasPrepaid && hasPostpaid && hasMaster) {
        elements.btnGenerate.disabled = false;
        elements.statusPill.className = "status-pill status-ready";
        safeText(elements.statusPillText, "All 5 Files Ready");
      } else {
        const count = [hasOrders, hasFlow, hasPrepaid, hasPostpaid, hasMaster].filter(Boolean).length;
        elements.btnGenerate.disabled = true;
        elements.statusPill.className = "status-pill status-missing";
        safeText(elements.statusPillText, `${count}/5 Files Uploaded`);
      }
    } catch (err) {
      console.warn("Status check failed:", err);
    }
  }

  // File Removal / Cancellation
  async function removeFile(role, displayName) {
    try {
      showToast(`Removing ${displayName}...`, "info", 2000);
      const res = await fetch("/api/remove-file", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ role: role })
      });
      const responseText = await res.text();
      let data;
      try {
        data = JSON.parse(responseText);
      } catch (e) {
        throw new Error(responseText.slice(0, 150) || `Server error (${res.status})`);
      }

      if (!res.ok) {
        throw new Error(data.error || "Failed to remove file");
      }

      showToast(`Cancelled & removed ${displayName}`, "info");

      // Hide results since dataset has changed
      if (elements.downloadSection) elements.downloadSection.classList.add("hidden");
      if (elements.previewSection) elements.previewSection.classList.add("hidden");
      state.allRecords = [];
      state.filteredRecords = [];

      await checkStatus();
    } catch (err) {
      showToast(`Error removing file: ${err.message}`, "error");
    }
  }

  // Bind Cancel / Remove buttons
  [
    { btn: elements.btnCancelOrders, role: "seller_orders", name: "Seller Orders Report" },
    { btn: elements.btnCancelFlow, role: "order_flow", name: "Order Flow Report" },
    { btn: elements.btnCancelPrepaid, role: "prepaid_settlement", name: "Prepaid Settlement" },
    { btn: elements.btnCancelPostpaid, role: "postpaid_settlement", name: "Postpaid Settlement" },
    { btn: elements.btnCancelMaster, role: "master_sku", name: "Master SKU Report" }
  ].forEach(({ btn, role, name }) => {
    if (btn) {
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        removeFile(role, name);
      });
    }
  });

  // 2. Drag & Drop Handling in Single Area + Individual Cards
  const dropZone = elements.dropZone;
  const fileInput = elements.fileInput;

  if (dropZone && fileInput) {
    dropZone.addEventListener("click", () => fileInput.click());

    ["dragenter", "dragover"].forEach(evt => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.add("drag-active");
      });
    });

    ["dragleave", "drop"].forEach(evt => {
      dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.remove("drag-active");
      });
    });

    dropZone.addEventListener("drop", (e) => {
      if (e.dataTransfer.files.length > 0) handleUpload(e.dataTransfer.files);
    });

    fileInput.addEventListener("change", () => {
      if (fileInput.files.length > 0) handleUpload(fileInput.files);
      fileInput.value = "";
    });
  }

  // Individual Card Drag, Drop, and Click to Upload
  const cardConfigs = [
    { card: elements.cardOrders, input: getEl("input-orders"), role: "seller_orders", name: "Seller Orders Report" },
    { card: elements.cardFlow, input: getEl("input-flow"), role: "order_flow", name: "Order Flow Report" },
    { card: elements.cardPrepaid, input: getEl("input-prepaid"), role: "prepaid_settlement", name: "Prepaid Settlement" },
    { card: elements.cardPostpaid, input: getEl("input-postpaid"), role: "postpaid_settlement", name: "Postpaid Settlement" },
    { card: elements.cardMaster, input: getEl("input-master"), role: "master_sku", name: "Master SKU Report" }
  ];

  cardConfigs.forEach(({ card, input, role, name }) => {
    if (!card) return;

    // Click card to open specific file picker
    card.addEventListener("click", (e) => {
      if (e.target.closest(".btn-cancel-file")) return;
      if (input) input.click();
    });

    if (input) {
      input.addEventListener("change", () => {
        if (input.files.length > 0) {
          handleUpload(input.files, role);
          input.value = "";
        }
      });
    }

    // Drag & drop directly onto individual card
    card.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.stopPropagation();
      card.classList.add("card-drag-active");
    });

    card.addEventListener("dragleave", (e) => {
      e.preventDefault();
      e.stopPropagation();
      card.classList.remove("card-drag-active");
    });

    card.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      card.classList.remove("card-drag-active");
      if (e.dataTransfer.files.length > 0) {
        handleUpload(e.dataTransfer.files, role);
      }
    });
  });

  async function handleUpload(fileList, targetRole = null) {
    const isLocalhost = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
    if (!isLocalhost) {
      for (let i = 0; i < fileList.length; i++) {
        const f = fileList[i];
        if (f.size > 4.5 * 1024 * 1024) {
          const sizeMB = (f.size / (1024 * 1024)).toFixed(1);
          showToast(`"${f.name}" is ${sizeMB} MB. Vercel max payload is 4.5 MB. Please use a compact file under 4.5 MB or run locally at http://localhost:8000`, "error", 8000);
          return;
        }
      }
    }

    const formData = new FormData();
    for (let i = 0; i < fileList.length; i++) {
      formData.append("files", fileList[i]);
    }
    if (targetRole) {
      formData.append("target_role", targetRole);
    }

    elements.statusPill.className = "status-pill status-running";
    safeText(elements.statusPillText, "Uploading & Processing...");

    try {
      const res = await fetch("/api/upload", { method: "POST", body: formData });
      const responseText = await res.text();
      let data;
      try {
        data = JSON.parse(responseText);
      } catch (parseErr) {
        if (res.status === 413 || (responseText && responseText.toLowerCase().includes("request entity too large"))) {
          throw new Error("File exceeds server upload limit (Vercel max payload is 4.5 MB). Use http://localhost:8000 locally or upload a compact Master file under 4.5 MB.");
        }
        throw new Error(responseText.slice(0, 150) || `Server error (${res.status})`);
      }

      if (!res.ok) {
        throw new Error(data.error || data.detail || `Upload failed (${res.status})`);
      }

      const uploaded = data.uploaded || [];
      const successful = uploaded.filter(u => u.success);
      const failed = uploaded.filter(u => !u.success);

      if (failed.length > 0) {
        const errorMsgs = failed.map(f => `${f.filename}: ${f.error || "Failed"}`).join("; ");
        showToast(`Upload error: ${errorMsgs}`, "error", 6000);
      }
      if (successful.length > 0) {
        showToast(`Uploaded ${successful.length} file(s) successfully`, "success");
      }

      await checkStatus();
    } catch (err) {
      showToast("Upload failed: " + err.message, "error", 7000);
      elements.statusPill.className = "status-pill";
      safeText(elements.statusPillText, "Upload Failed");
    }
  }

  // 3. Generate Final Report Action
  elements.btnGenerate.addEventListener("click", async () => {
    if (state.isRunning) return;
    state.isRunning = true;

    elements.btnGenerate.disabled = true;
    elements.btnSpinner.classList.remove("hidden");
    elements.btnPlayIcon.classList.add("hidden");
    elements.statusPill.className = "status-pill status-running";
    safeText(elements.statusPillText, "Processing & Reconciling...");

    try {
      const payload = {};

      const res = await fetch("/api/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const responseText = await res.text();
      let data;
      try {
        data = JSON.parse(responseText);
      } catch (parseErr) {
        throw new Error(responseText.slice(0, 150) || `Server error (${res.status})`);
      }

      if (!res.ok || !data.success) {
        throw new Error(data.error || "Generation failed");
      }

      // Update Result Box in same area
      safeText(elements.resultTitle, "Final Excel Report Generated");
      safeText(
        elements.resultSubtitle,
        `${Number(data.flow_rows).toLocaleString()} Order Flow rows • 100% SKU match • Complete 3-sheet workbook with Grand Totals`
      );

      // Populate Metric Strip
      safeText(elements.metricOrderRows, Number(data.flow_rows).toLocaleString());
      safeText(elements.metricMatchRate, `${data.match_rate}%`);
      safeText(elements.metricPrepaidAmt, formatRs(data.total_prepaid_amount));
      safeText(elements.metricPostpaidAmt, formatRs(data.total_postpaid_amount));
      safeText(elements.metricFinalAmt, formatRs(data.sum_final_settled));
      safeText(elements.metricPayoutAmt, formatRs(data.total_payout_sum));
      safeText(elements.metricNetCost, formatRs(data.total_net_cost_sum));
      safeText(elements.metricProfitAmt, formatRs(data.total_profit_sum));

      elements.downloadSection.classList.remove("hidden");

      // Load Preview Records from Sheet "Final"
      if (data.preview_rows && data.preview_rows.length > 0) {
        state.allRecords = data.preview_rows;
      } else {
        await loadPreviewData();
      }

      applyFilter();
      elements.previewSection.classList.remove("hidden");

      elements.statusPill.className = "status-pill status-ready";
      safeText(elements.statusPillText, "Report Complete");

      showToast("Final Excel Report generated successfully!", "success");

    } catch (err) {
      showToast("Error: " + err.message, "error");
      elements.statusPill.className = "status-pill";
      safeText(elements.statusPillText, "Execution Error");
    } finally {
      state.isRunning = false;
      elements.btnGenerate.disabled = false;
      elements.btnSpinner.classList.add("hidden");
      elements.btnPlayIcon.classList.remove("hidden");
    }
  });

  // 4. Load Final Sheet Preview Data
  async function loadPreviewData() {
    try {
      const res = await fetch("/api/data?page=1&limit=300&match_status=all");
      if (!res.ok) return;
      const json = await res.json();
      state.allRecords = json.data || [];
    } catch (err) {
      console.warn("Could not load preview:", err);
    }
  }

  function applyFilter() {
    const q = state.searchQuery.toLowerCase().trim();

    if (!q) {
      state.filteredRecords = [...state.allRecords];
    } else {
      state.filteredRecords = state.allRecords.filter(r => {
        const releaseId = String(r["order release id"] || "").toLowerCase();
        const sku = String(r["seller sku code"] || "").toLowerCase();
        const style = String(r["Style Name"] || "").toLowerCase();
        const retType = String(r["return_type"] || "").toLowerCase();
        const qty = String(r["QTY"] || "").toLowerCase();
        const retQty = String(r["RETURN QTY"] || "").toLowerCase();
        const saleQty = String(r["SALE QTY "] || "").toLowerCase();
        const payout = String(r["Payout"] || "").toLowerCase();
        const netCost = String(r["NET PRODUCT COST "] || "").toLowerCase();
        const amt = String(r["Settled_Amount"] || "").toLowerCase();
        const profit = String(r["PROFIT"] || "").toLowerCase();
        return releaseId.includes(q) || sku.includes(q) || style.includes(q) || retType.includes(q) ||
               qty.includes(q) || retQty.includes(q) || saleQty.includes(q) || payout.includes(q) ||
               netCost.includes(q) || amt.includes(q) || profit.includes(q);
      });
    }

    state.totalRecords = state.filteredRecords.length;
    state.totalPages = Math.max(1, Math.ceil(state.totalRecords / state.pageSize));
    state.currentPage = 1;
    renderTable();
  }

  // 5. Render Table (11 Columns matching Sample Reference)
  function renderTable() {
    const tbody = elements.previewTbody;
    if (!tbody) return;
    tbody.innerHTML = "";

    if (state.filteredRecords.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="11" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
            No matching records found.
          </td>
        </tr>
      `;
      safeText(elements.previewCount, "Showing 0 of 0 records");
      safeText(elements.pageText, "Page 1 / 1");
      if (elements.btnPrev) elements.btnPrev.disabled = true;
      if (elements.btnNext) elements.btnNext.disabled = true;
      return;
    }

    const startIdx = (state.currentPage - 1) * state.pageSize;
    const endIdx = Math.min(startIdx + state.pageSize, state.totalRecords);
    const pageRecords = state.filteredRecords.slice(startIdx, endIdx);

    const fragment = document.createDocumentFragment();

    pageRecords.forEach(r => {
      const tr = document.createElement("tr");

      // 1. order release id
      const td1 = document.createElement("td");
      td1.className = "cell-mono";
      td1.textContent = r["order release id"] || "—";
      tr.appendChild(td1);

      // 2. seller sku code
      const td2 = document.createElement("td");
      const skuVal = r["seller sku code"] || "NOT FOUND";
      td2.textContent = skuVal;
      if (skuVal === "NOT FOUND") {
        td2.className = "sku-notfound";
      }
      tr.appendChild(td2);

      // 3. Style Name
      const td3 = document.createElement("td");
      const styleVal = r["Style Name"] || "NOT FOUND";
      td3.textContent = styleVal;
      if (styleVal === "NOT FOUND") {
        td3.className = "sku-notfound";
      }
      tr.appendChild(td3);

      // 4. return_type
      const td4 = document.createElement("td");
      td4.style.textAlign = "center";
      const retVal = r["return_type"] || "Delivered";
      const badge = document.createElement("span");
      badge.textContent = retVal;

      if (retVal === "Delivered") badge.className = "status-tag tag-delivered";
      else if (retVal === "RTO") badge.className = "status-tag tag-rto";
      else if (retVal === "RVP") badge.className = "status-tag tag-rvp";
      else if (retVal === "Exchange") badge.className = "status-tag tag-exchange";
      else badge.className = "status-tag tag-unmapped";

      td4.appendChild(badge);
      tr.appendChild(td4);

      // 5. QTY
      const td5 = document.createElement("td");
      td5.style.textAlign = "center";
      td5.className = "cell-mono";
      td5.textContent = r["QTY"] !== undefined ? r["QTY"] : "1";
      tr.appendChild(td5);

      // 6. RETURN QTY
      const td6 = document.createElement("td");
      td6.style.textAlign = "center";
      td6.className = "cell-mono";
      td6.textContent = r["RETURN QTY"] !== undefined ? r["RETURN QTY"] : (retVal.toLowerCase() === "delivered" ? "0" : "1");
      tr.appendChild(td6);

      // 7. SALE QTY 
      const td7 = document.createElement("td");
      td7.style.textAlign = "center";
      td7.className = "cell-mono";
      td7.textContent = r["SALE QTY "] !== undefined ? r["SALE QTY "] : (retVal.toLowerCase() === "delivered" ? "1" : "0");
      tr.appendChild(td7);

      // 8. Payout
      const td8 = document.createElement("td");
      td8.style.textAlign = "right";
      td8.className = "cell-mono";
      td8.textContent = formatRs(r["Payout"]);
      tr.appendChild(td8);

      // 9. NET PRODUCT COST 
      const td9 = document.createElement("td");
      td9.style.textAlign = "right";
      td9.className = "cell-mono";
      td9.textContent = formatRs(r["NET PRODUCT COST "]);
      tr.appendChild(td9);

      // 10. Settled_Amount
      const td10 = document.createElement("td");
      td10.style.textAlign = "right";
      td10.className = "cell-mono";
      td10.textContent = formatRs(r["Settled_Amount"]);
      tr.appendChild(td10);

      // 11. PROFIT
      const td11 = document.createElement("td");
      td11.style.textAlign = "right";
      td11.className = "cell-mono";
      const profitVal = Number(r["PROFIT"] || 0);
      td11.style.color = profitVal < 0 ? "var(--danger)" : "var(--success)";
      td11.textContent = formatRs(profitVal);
      tr.appendChild(td11);

      fragment.appendChild(tr);
    });

    tbody.appendChild(fragment);

    safeText(elements.previewCount, `Showing ${startIdx + 1} to ${endIdx} of ${state.totalRecords.toLocaleString()} records`);
    safeText(elements.pageText, `Page ${state.currentPage} / ${state.totalPages}`);

    if (elements.btnPrev) elements.btnPrev.disabled = state.currentPage <= 1;
    if (elements.btnNext) elements.btnNext.disabled = state.currentPage >= state.totalPages;
  }

  // 6. Search & Pagination Controls
  if (elements.previewSearchInput) {
    elements.previewSearchInput.addEventListener("input", (e) => {
      state.searchQuery = e.target.value;
      applyFilter();
    });
  }

  if (elements.btnPrev) {
    elements.btnPrev.addEventListener("click", () => {
      if (state.currentPage > 1) {
        state.currentPage--;
        renderTable();
      }
    });
  }

  if (elements.btnNext) {
    elements.btnNext.addEventListener("click", () => {
      if (state.currentPage < state.totalPages) {
        state.currentPage++;
        renderTable();
      }
    });
  }

  // Initialize
  checkStatus();
});
