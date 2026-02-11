const { invoke } = window.__TAURI__.core;
const { open, save } = window.__TAURI__.dialog;

const btnSelect = document.getElementById("btn-select");
const btnExport = document.getElementById("btn-export");
const btnMerge = document.getElementById("btn-merge");
const filePath = document.getElementById("file-path");
const rowCount = document.getElementById("row-count");
const statusMsg = document.getElementById("status-msg");
const tableBody = document.getElementById("table-body");
const mappingPanel = document.getElementById("mapping-panel");
const mappingFields = document.getElementById("mapping-fields");
const btnApplyMapping = document.getElementById("btn-apply-mapping");
const btnCancelMapping = document.getElementById("btn-cancel-mapping");

let currentFilePath = null;
let availableColumns = [];

const fieldConfigs = [
  { key: "recipient_name", label: "收件人姓名", type: "string" },
  { key: "recipient_phone", label: "收件人手机号", type: "string" },
  { key: "delivery_address", label: "收货地址", type: "string" },
  { key: "product_name", label: "商品名称", type: "string" },
  { key: "product_spec", label: "商品规格", type: "string" },
  { key: "quantity", label: "商品数量", type: "number" },
  { key: "remarks", label: "备注", type: "string" },
];

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function showStatus(message, type) {
  statusMsg.textContent = message;
  statusMsg.className = "status-msg " + type;
}

function hideStatus() {
  statusMsg.className = "status-msg hidden";
}

function updateMergeButton(hasDuplicates, duplicateCount) {
  if (hasDuplicates) {
    btnMerge.textContent = `合并重复项 (${duplicateCount}条)`;
    btnMerge.style.display = "inline-block";
  } else {
    btnMerge.style.display = "none";
  }
}

function renderTable(rows) {
  if (!rows || rows.length === 0) {
    tableBody.innerHTML =
      '<tr><td colspan="8" class="empty-hint">没有数据</td></tr>';
    return;
  }

  tableBody.innerHTML = rows
    .map(
      (row, i) => {
        const groupClass = row.group_id > 0 ? ` dup-group-${(row.group_id - 1) % 6}` : "";
        return `<tr class="${groupClass}">
        <td>${i + 1}</td>
        <td>${escapeHtml(row.recipient_name)}</td>
        <td>${escapeHtml(row.recipient_phone)}</td>
        <td>${escapeHtml(row.delivery_address)}</td>
        <td>${escapeHtml(row.product_name)}</td>
        <td>${escapeHtml(row.product_spec)}</td>
        <td>${escapeHtml(row.quantity)}</td>
        <td>${escapeHtml(row.remarks)}</td>
      </tr>`;
      }
    )
    .join("");
}

function renderMappingPanel() {
  mappingFields.innerHTML = fieldConfigs.map(field => `
    <div class="field-mapping">
      <label>${field.label}:</label>
      <div class="mapping-controls">
        <select class="column-select" data-field="${field.key}" multiple size="3">
          ${availableColumns.map(col => `<option value="${col.index}">${col.code} - ${col.title}</option>`).join("")}
        </select>
        <select class="operation-select" data-field="${field.key}">
          ${field.type === "number" ? `
            <option value="add">加</option>
            <option value="subtract">减</option>
            <option value="multiply">乘</option>
            <option value="divide">除</option>
          ` : `
            <option value="concat">拼接</option>
          `}
        </select>
      </div>
    </div>
  `).join("");
}

function getMappings() {
  const mappings = {};
  fieldConfigs.forEach(field => {
    const select = mappingFields.querySelector(`select.column-select[data-field="${field.key}"]`);
    const opSelect = mappingFields.querySelector(`select.operation-select[data-field="${field.key}"]`);
    const selectedIndices = Array.from(select.selectedOptions).map(opt => parseInt(opt.value));
    if (selectedIndices.length > 0) {
      mappings[field.key] = {
        source_indices: selectedIndices,
        operation: opSelect.value,
      };
    }
  });
  return mappings;
}

btnSelect.addEventListener("click", async () => {
  try {
    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel 文件", extensions: ["xlsx", "xls"] }],
    });

    if (!selected) return;

    currentFilePath = selected;
    filePath.textContent = selected;
    filePath.title = selected;
    showStatus("正在读取列信息...", "loading");
    btnSelect.disabled = true;

    availableColumns = await invoke("read_columns", { path: selected });
    renderMappingPanel();
    mappingPanel.style.display = "block";
    showStatus("请配置列映射", "success");
  } catch (err) {
    showStatus("读取失败: " + err, "error");
  } finally {
    btnSelect.disabled = false;
  }
});

btnApplyMapping.addEventListener("click", async () => {
  try {
    const mappings = getMappings();
    showStatus("正在转换文件...", "loading");
    btnApplyMapping.disabled = true;

    const result = await invoke("convert_with_mapping", { path: currentFilePath, mappings });

    renderTable(result.rows);
    rowCount.textContent = `共 ${result.total_rows} 条数据`;
    btnExport.disabled = false;
    updateMergeButton(result.has_duplicates, result.duplicate_count);
    mappingPanel.style.display = "none";
    showStatus(`文件读取成功，共转换 ${result.total_rows} 条数据`, "success");
  } catch (err) {
    showStatus("转换失败: " + err, "error");
  } finally {
    btnApplyMapping.disabled = false;
  }
});

btnCancelMapping.addEventListener("click", () => {
  mappingPanel.style.display = "none";
  hideStatus();
});

btnExport.addEventListener("click", async () => {
  try {
    const savePath = await save({
      defaultPath: "output.xlsx",
      filters: [{ name: "Excel 文件", extensions: ["xlsx"] }],
    });

    if (!savePath) return;

    showStatus("正在导出...", "loading");
    btnExport.disabled = true;

    const message = await invoke("export_file", { outputPath: savePath });

    showStatus(message, "success");
  } catch (err) {
    showStatus("导出失败: " + err, "error");
  } finally {
    btnExport.disabled = false;
  }
});

btnMerge.addEventListener("click", async () => {
  try {
    showStatus("正在合并重复项...", "loading");
    btnMerge.disabled = true;

    const result = await invoke("merge_duplicates");

    renderTable(result.rows);
    rowCount.textContent = `共 ${result.total_rows} 条数据`;
    updateMergeButton(result.has_duplicates, result.duplicate_count);
    showStatus(`合并完成，当前共 ${result.total_rows} 条数据`, "success");
  } catch (err) {
    showStatus("合并失败: " + err, "error");
  } finally {
    btnMerge.disabled = false;
  }
});
