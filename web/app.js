(function () {
  const state = {
    files: [],
    result: null,
  };

  const elements = {};

  function normalizeText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value)
      .replace(/\n/g, "")
      .replace(/\s+/g, "")
      .replace(/（/g, "(")
      .replace(/）/g, ")")
      .trim();
  }

  function toNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : null;
    }

    if (typeof value === "string") {
      const cleaned = value.replace(/,/g, "").trim();
      if (!cleaned) {
        return null;
      }
      const parsed = Number(cleaned);
      return Number.isFinite(parsed) ? parsed : null;
    }

    return null;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("zh-CN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(value || 0);
  }

  function findHeaderRow(rows) {
    const searchEnd = Math.min(rows.length, 15);
    for (let index = 0; index < searchEnd; index += 1) {
      const values = (rows[index] || []).map(normalizeText);
      if (values.includes("销售日期") && values.includes("商品名称") && values.includes("数量")) {
        return index;
      }
    }
    return null;
  }

  function buildColumnMap(headerRow) {
    const aliases = {
      date: new Set(["销售日期"]),
      name: new Set(["商品名称"]),
      spec: new Set(["规格型号"]),
      unit: new Set(["单位"]),
      qty: new Set(["数量"]),
      price: new Set(["单价(元)"]),
      amount: new Set(["金额(元)"]),
      invoiceNo: new Set(["销售单据号"]),
    };

    const columnMap = {};
    (headerRow || []).forEach((cell, index) => {
      const normalized = normalizeText(cell);
      Object.entries(aliases).forEach(([field, validHeaders]) => {
        if (validHeaders.has(normalized)) {
          columnMap[field] = index;
        }
      });
    });

    const required = ["date", "name", "spec", "unit", "qty", "price", "amount"];
    return required.every((field) => field in columnMap) ? columnMap : null;
  }

  function findPrefixedValue(rows, prefixes, searchEndRow) {
    for (let index = 0; index < searchEndRow; index += 1) {
      const value = rows[index]?.[0];
      if (value === null || value === undefined) {
        continue;
      }

      const text = String(value).trim();
      if (prefixes.some((prefix) => text.startsWith(prefix))) {
        return text;
      }
    }
    return "";
  }

  function extractHeightInMeters(spec, name) {
    if (!spec) {
      return null;
    }

    const nameText = String(name || "");
    if (!nameText.includes("井座") && !nameText.includes("闸阀方井")) {
      return null;
    }

    const parts = String(spec).split("*");
    if (parts.length < 3) {
      return null;
    }

    const height = toNumber(parts[parts.length - 1]);
    return height === null ? null : height / 1000;
  }

  function parseWorkbook(fileName, arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const items = [];
    const customerInfos = [];
    const addressInfos = [];
    const processedSheets = [];
    const issues = [];
    let totalAmount = 0;

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: null,
      });

      const headerIndex = findHeaderRow(rows);
      if (headerIndex === null) {
        return;
      }

      const columnMap = buildColumnMap(rows[headerIndex]);
      if (!columnMap) {
        issues.push(`${fileName} / ${sheetName} 缺少必要列，已跳过。`);
        return;
      }

      processedSheets.push(`${fileName} / ${sheetName}`);

      const searchEnd = Math.min(headerIndex + 1, 10);
      const customerInfo = findPrefixedValue(rows, ["收货单位：", "客户："], searchEnd);
      const addressInfo = findPrefixedValue(rows, ["送货地址："], searchEnd);

      if (customerInfo && !customerInfos.includes(customerInfo)) {
        customerInfos.push(customerInfo);
      }
      if (addressInfo && !addressInfos.includes(addressInfo)) {
        addressInfos.push(addressInfo);
      }

      for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex] || [];
        const name = row[columnMap.name];
        const spec = row[columnMap.spec];
        const unit = row[columnMap.unit];
        const qty = row[columnMap.qty];
        const price = row[columnMap.price];
        const amount = row[columnMap.amount];
        const invoiceNo = columnMap.invoiceNo !== undefined ? row[columnMap.invoiceNo] : null;

        if ([name, spec, qty, amount].every((value) => value === null || value === undefined || value === "")) {
          continue;
        }

        if (normalizeText(name) === "") {
          continue;
        }

        if (String(name).includes("合计")) {
          continue;
        }

        const amountValue = toNumber(amount) ?? 0;
        const qtyValue = toNumber(qty) ?? 0;
        const priceValue = toNumber(price);
        const heightInMeters = extractHeightInMeters(spec, name);

        if (heightInMeters !== null) {
          items.push({
            name: String(name || ""),
            spec: spec ? String(spec) : "",
            unit: "米",
            qty: qtyValue * heightInMeters,
            price: priceValue ?? price ?? 0,
            amount: amountValue,
            invoiceNo: invoiceNo ?? null,
          });
        } else {
          items.push({
            name: String(name || ""),
            spec: spec ? String(spec) : "",
            unit: unit ? String(unit) : "",
            qty: qtyValue,
            price: priceValue ?? price ?? 0,
            amount: amountValue,
            invoiceNo: invoiceNo ?? null,
          });
        }

        totalAmount += amountValue;
      }
    });

    return {
      items,
      totalAmount,
      customerInfos,
      addressInfos,
      processedSheets,
      issues,
    };
  }

  function summarizeItems(items) {
    const summaryMap = new Map();

    items.forEach((item) => {
      const key = [item.name, item.spec, item.unit, item.price].join("__");
      const existing = summaryMap.get(key) || {
        name: item.name,
        spec: item.spec,
        unit: item.unit,
        qty: 0,
        price: item.price,
        amount: 0,
      };

      existing.qty += item.qty;
      existing.amount += item.amount;
      summaryMap.set(key, existing);
    });

    return Array.from(summaryMap.values()).sort((left, right) => {
      const nameCompare = left.name.localeCompare(right.name, "zh-CN");
      if (nameCompare !== 0) {
        return nameCompare;
      }

      const specCompare = left.spec.localeCompare(right.spec, "zh-CN");
      if (specCompare !== 0) {
        return specCompare;
      }

      return Number(left.price || 0) - Number(right.price || 0);
    });
  }

  function buildSummaryWorkbook(result) {
    const workbook = XLSX.utils.book_new();
    const customerLabel = result.customerInfos.length ? result.customerInfos.join("；") : "客户信息";
    const addressLabel = result.addressInfos.length ? result.addressInfos.join("；") : "送货地址信息";
    const sourceLabel = `来源文件：${result.sourceFiles.join("、")}`;
    const generatedLabel = `生成时间：${new Date().toLocaleString("zh-CN", { hour12: false })}`;

    const aoa = [
      ["销售发票"],
      [customerLabel],
      [addressLabel],
      [sourceLabel],
      [generatedLabel],
      ["商品名称", "规格型号", "数量", "单价(元)", "金额(元)"],
      ...result.items.map((item) => [item.name, item.spec, item.qty, item.price, item.amount]),
      ["合计", "", "", "", result.totalAmount],
    ];

    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    sheet["!merges"] = [
      XLSX.utils.decode_range("A1:E1"),
      XLSX.utils.decode_range("A2:E2"),
      XLSX.utils.decode_range("A3:E3"),
      XLSX.utils.decode_range("A4:E4"),
      XLSX.utils.decode_range("A5:E5"),
    ];
    sheet["!cols"] = [
      { wch: 20 },
      { wch: 22 },
      { wch: 14 },
      { wch: 14 },
      { wch: 16 },
    ];

    for (let rowIndex = 7; rowIndex <= result.items.length + 7; rowIndex += 1) {
      ["C", "D", "E"].forEach((col) => {
        const cell = sheet[`${col}${rowIndex}`];
        if (cell) {
          cell.z = "0.00";
        }
      });
    }

    const totalCell = sheet[`E${result.items.length + 7}`];
    if (totalCell) {
      totalCell.z = "0.00";
    }

    XLSX.utils.book_append_sheet(workbook, sheet, "开票汇总");
    return workbook;
  }

  function timestamp() {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, "0");
    return [
      now.getFullYear(),
      pad(now.getMonth() + 1),
      pad(now.getDate()),
      "_",
      pad(now.getHours()),
      pad(now.getMinutes()),
      pad(now.getSeconds()),
    ].join("");
  }

  async function buildResult() {
    if (!state.files.length) {
      throw new Error("请先选择至少一个 Excel 文件。")
    }

    const aggregate = {
      items: [],
      totalAmount: 0,
      customerInfos: [],
      addressInfos: [],
      processedSheets: [],
      sourceFiles: [],
      issues: [],
    };

    for (const file of state.files) {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(file.name, buffer);

      if (parsed.processedSheets.length) {
        aggregate.sourceFiles.push(file.name);
      }

      aggregate.items.push(...parsed.items);
      aggregate.totalAmount += parsed.totalAmount;
      aggregate.processedSheets.push(...parsed.processedSheets);
      aggregate.issues.push(...parsed.issues);

      parsed.customerInfos.forEach((entry) => {
        if (!aggregate.customerInfos.includes(entry)) {
          aggregate.customerInfos.push(entry);
        }
      });
      parsed.addressInfos.forEach((entry) => {
        if (!aggregate.addressInfos.includes(entry)) {
          aggregate.addressInfos.push(entry);
        }
      });
    }

    if (!aggregate.processedSheets.length) {
      throw new Error("没有在这些文件里找到可汇总的对账数据。")
    }

    aggregate.items = summarizeItems(aggregate.items);
    return aggregate;
  }

  function renderFiles() {
    elements.selectedFiles.innerHTML = "";

    state.files.forEach((file) => {
      const pill = document.createElement("div");
      pill.className = "file-pill";
      pill.textContent = `${file.name} · ${(file.size / 1024).toFixed(1)} KB`;
      elements.selectedFiles.appendChild(pill);
    });

    if (!state.files.length) {
      const hint = document.createElement("div");
      hint.className = "file-pill";
      hint.textContent = "还没有选择文件";
      elements.selectedFiles.appendChild(hint);
    }
  }

  function renderStats(result) {
    const cards = [
      ["来源文件", result.sourceFiles.length],
      ["来源工作表", result.processedSheets.length],
      ["商品行数", result.items.length],
      ["总金额", formatNumber(result.totalAmount)],
    ];

    elements.statsGrid.innerHTML = "";
    cards.forEach(([label, value], index) => {
      const card = document.createElement("article");
      card.className = `stat-card${index === 3 ? " accent" : ""}`;
      card.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
      elements.statsGrid.appendChild(card);
    });

    elements.metaStack.innerHTML = "";
    [
      ["客户", result.customerInfos.join("；") || "未识别"],
      ["地址", result.addressInfos.join("；") || "未识别"],
      ["来源", result.processedSheets.join("；")],
    ].forEach(([label, value]) => {
      const block = document.createElement("div");
      block.className = "meta-block";
      block.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
      elements.metaStack.appendChild(block);
    });
  }

  function renderTable(result) {
    elements.previewBody.innerHTML = "";

    result.items.forEach((item) => {
      const row = document.createElement("tr");
      row.innerHTML = [
        `<td>${item.name}</td>`,
        `<td>${item.spec || ""}</td>`,
        `<td class="num">${formatNumber(item.qty)}${item.unit}</td>`,
        `<td class="num">${formatNumber(Number(item.price) || 0)}</td>`,
        `<td class="num">${formatNumber(item.amount)}</td>`,
      ].join("");
      elements.previewBody.appendChild(row);
    });

    const totalRow = document.createElement("tr");
    totalRow.innerHTML = [
      "<td><strong>合计</strong></td>",
      "<td></td>",
      "<td></td>",
      "<td></td>",
      `<td class="num"><strong>${formatNumber(result.totalAmount)}</strong></td>`,
    ].join("");
    elements.previewBody.appendChild(totalRow);
  }

  function renderIssues(issues) {
    elements.issueList.innerHTML = "";
    if (!issues.length) {
      elements.issuesPanel.hidden = true;
      return;
    }

    elements.issuesPanel.hidden = false;
    issues.forEach((issue) => {
      const li = document.createElement("li");
      li.textContent = issue;
      elements.issueList.appendChild(li);
    });
  }

  function setStatus(text, isReady) {
    elements.statusBadge.textContent = text;
    elements.statusBadge.classList.toggle("ready", Boolean(isReady));
  }

  function syncButtons() {
    elements.generateBtn.disabled = !state.files.length;
    elements.resetBtn.disabled = !state.files.length;
    elements.downloadBtn.disabled = !state.result;
  }

  function resetResult() {
    state.result = null;
    elements.previewBody.innerHTML = '<tr><td colspan="5" class="empty-cell">选择 Excel 文件后生成预览。</td></tr>';
    elements.statsGrid.innerHTML = `
      <article class="stat-card muted"><span>来源文件</span><strong>0</strong></article>
      <article class="stat-card muted"><span>来源工作表</span><strong>0</strong></article>
      <article class="stat-card muted"><span>商品行数</span><strong>0</strong></article>
      <article class="stat-card accent muted"><span>总金额</span><strong>0.00</strong></article>
    `;
    elements.metaStack.innerHTML = `
      <div class="meta-block"><span>客户</span><strong>等待生成</strong></div>
      <div class="meta-block"><span>地址</span><strong>等待生成</strong></div>
    `;
    renderIssues([]);
    setStatus(state.files.length ? "等待生成" : "等待文件", false);
    syncButtons();
  }

  function setFiles(fileList) {
    state.files = Array.from(fileList || []).filter((file) => /\.(xlsx|xls)$/i.test(file.name));
    renderFiles();
    resetResult();
  }

  async function onGenerate() {
    try {
      setStatus("正在解析文件", false);
      const result = await buildResult();
      state.result = result;
      renderStats(result);
      renderTable(result);
      renderIssues(result.issues);
      setStatus("预览已生成", true);
      syncButtons();
    } catch (error) {
      state.result = null;
      renderIssues([error.message || "生成失败"]);
      setStatus("生成失败", false);
      syncButtons();
    }
  }

  function onDownload() {
    if (!state.result) {
      return;
    }

    const workbook = buildSummaryWorkbook(state.result);
    XLSX.writeFile(workbook, `开票汇总_${timestamp()}.xlsx`);
  }

  function initDom() {
    elements.fileInput = document.getElementById("file-input");
    elements.dropzone = document.getElementById("dropzone");
    elements.selectedFiles = document.getElementById("selected-files");
    elements.generateBtn = document.getElementById("generate-btn");
    elements.downloadBtn = document.getElementById("download-btn");
    elements.resetBtn = document.getElementById("reset-btn");
    elements.statsGrid = document.getElementById("stats-grid");
    elements.metaStack = document.getElementById("meta-stack");
    elements.previewBody = document.getElementById("preview-body");
    elements.statusBadge = document.getElementById("status-badge");
    elements.issuesPanel = document.getElementById("issues-panel");
    elements.issueList = document.getElementById("issue-list");
  }

  function bindEvents() {
    elements.fileInput.addEventListener("change", (event) => {
      setFiles(event.target.files);
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      elements.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropzone.classList.add("is-dragover");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      elements.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropzone.classList.remove("is-dragover");
      });
    });

    elements.dropzone.addEventListener("drop", (event) => {
      if (event.dataTransfer?.files?.length) {
        setFiles(event.dataTransfer.files);
      }
    });

    elements.generateBtn.addEventListener("click", onGenerate);
    elements.downloadBtn.addEventListener("click", onDownload);
    elements.resetBtn.addEventListener("click", () => {
      elements.fileInput.value = "";
      state.files = [];
      renderFiles();
      resetResult();
    });
  }

  function boot() {
    initDom();
    renderFiles();
    resetResult();

    if (!window.XLSX) {
      setStatus("组件加载失败", false);
      renderIssues(["未能加载 Excel 处理库，请检查网络后刷新页面。"]);
      return;
    }

    bindEvents();
  }

  window.addEventListener("DOMContentLoaded", boot);
})();