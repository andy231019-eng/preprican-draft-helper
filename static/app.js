const REPORTS = [
  { id: "monthly", label: "月報" },
  { id: "quarterly", label: "季報" },
  { id: "annual", label: "年報" },
];

const state = {
  recipients: [],
  recipientSummary: null,
  reportFiles: {
    monthly: null,
    quarterly: null,
    annual: null,
  },
  selectionMatrix: {},
  previews: [],
  templates: null,
  templateLang: "zh",
  languageMode: "recipient",
};

const $ = (selector) => document.querySelector(selector);

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function showAlert(message) {
  $("#alert").textContent = message;
  $("#alert").classList.toggle("hidden", !message);
}

function reportLabel(reportType) {
  return REPORTS.find((item) => item.id === reportType)?.label || reportType;
}

function isSendable(recipient) {
  return recipient.email && recipient.status === "active";
}

function selectedCountForRecipient(recipientId) {
  return REPORTS.filter((report) => state.selectionMatrix[recipientId]?.[report.id]).length;
}

function selectedCountForReport(reportType) {
  return state.recipients.filter((recipient) => state.selectionMatrix[recipient.recipient_id]?.[reportType]).length;
}

function draftCount() {
  return state.recipients.filter((recipient) => selectedCountForRecipient(recipient.recipient_id) > 0 && recipient.email).length;
}

function renderRecipientSummary() {
  const summary = state.recipientSummary;
  const el = $("#recipientSummary");
  if (summary) {
    el.innerHTML = `已載入 ${summary.total} 位收件人，可到下方勾選寄送對象`;
    el.classList.remove("hidden");
    el.classList.add("loaded");
  } else {
    el.classList.add("hidden");
    el.classList.remove("loaded");
  }
}

function renderPdfStatus() {
  $("#pdfStatus").innerHTML = REPORTS.map((report) => {
    const file = state.reportFiles[report.id];
    return `<span class="pill ${file ? "" : "warning"}">${report.label}：${file ? `已上傳 ${escapeHtml(file.fileName)}` : "尚未上傳"}</span>`;
  }).join(" ");
}

function filteredRecipients() {
  const keyword = $("#recipientSearch").value.trim().toLowerCase();
  const filter = $("#recipientFilter").value;
  return state.recipients.filter((recipient) => {
    const text = `${recipient.name} ${recipient.nickname} ${recipient.email}`.toLowerCase();
    const statusMatch =
      filter === "all" ||
      (filter === "sendable" && isSendable(recipient)) ||
      (filter === "missing" && !recipient.email) ||
      (filter === "inactive" && recipient.status !== "active");
    return (!keyword || text.includes(keyword)) && statusMatch;
  });
}

function renderMatrix() {
  const recipients = filteredRecipients();
  if (!state.recipients.length) {
    $("#matrixTable").innerHTML = `<div class="muted">請先上傳收件人 Excel。</div>`;
    return;
  }

  $("#matrixTable").innerHTML = `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>收件人</th>
            <th>暱稱</th>
            <th>Email</th>
            <th>CC</th>
            ${REPORTS.map((report) => `
              <th>
                ${report.label}
                <div class="th-actions">
                  <button type="button" class="tiny" data-select-report="${report.id}" data-value="true" ${state.reportFiles[report.id] ? "" : "disabled"}>全選</button>
                  <button type="button" class="tiny" data-select-report="${report.id}" data-value="false">全不選</button>
                </div>
              </th>
            `).join("")}
            <th>本人報告數</th>
            <th>狀態</th>
          </tr>
        </thead>
        <tbody>
          ${recipients.map((recipient) => {
            const disabled = !recipient.email || recipient.status !== "active";
            return `
              <tr class="${disabled ? "muted-row" : ""}">
                <td>${escapeHtml(recipient.name || "-")}</td>
                <td>${escapeHtml(recipient.nickname || "-")}</td>
                <td>${escapeHtml(recipient.email || "-")}</td>
                <td>${escapeHtml(recipient.cc || "-")}</td>
                ${REPORTS.map((report) => `
                  <td>
                    <input type="checkbox" data-recipient-id="${escapeHtml(recipient.recipient_id)}" data-report-type="${report.id}" ${state.selectionMatrix[recipient.recipient_id]?.[report.id] ? "checked" : ""} ${disabled || !state.reportFiles[report.id] ? "disabled" : ""} />
                  </td>
                `).join("")}
                <td>${selectedCountForRecipient(recipient.recipient_id)}</td>
                <td>${!recipient.email ? "缺少 Email" : recipient.status !== "active" ? "停用" : "可寄送"}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
        <tfoot>
          <tr>
            <td colspan="4">統計</td>
            ${REPORTS.map((report) => `<td>${report.label}收件人數：${selectedCountForReport(report.id)}</td>`).join("")}
            <td>總 Email 草稿數：${draftCount()}</td>
            <td></td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;

  $("#matrixTable").querySelectorAll("[data-recipient-id]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const recipientId = event.target.dataset.recipientId;
      const reportType = event.target.dataset.reportType;
      state.selectionMatrix[recipientId] ||= {};
      state.selectionMatrix[recipientId][reportType] = event.target.checked;
      state.previews = [];
      renderMatrix();
      renderPreviewRows();
    });
  });

  $("#matrixTable").querySelectorAll("[data-select-report]").forEach((button) => {
    button.addEventListener("click", () => {
      const reportType = button.dataset.selectReport;
      const value = button.dataset.value === "true";
      for (const recipient of filteredRecipients()) {
        if (isSendable(recipient) && (state.reportFiles[reportType] || !value)) {
          state.selectionMatrix[recipient.recipient_id] ||= {};
          state.selectionMatrix[recipient.recipient_id][reportType] = value;
        }
      }
      state.previews = [];
      renderMatrix();
      renderPreviewRows();
    });
  });
}

async function uploadRecipients() {
  const file = $("#recipientExcel").files?.[0];
  if (!file) return showAlert("請先選擇收件人 Excel。");
  if (!file.name.toLowerCase().endsWith(".xlsx")) return showAlert("只接受 .xlsx 格式。");

  const formData = new FormData();
  formData.append("recipients", file);
  const response = await fetch("/api/upload-recipients", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) return showAlert(data.error || "解析收件人資料失敗。");

  showAlert("");
  state.recipients = data.recipients;
  state.recipientSummary = data.summary;
  state.selectionMatrix = {};
  for (const recipient of state.recipients) state.selectionMatrix[recipient.recipient_id] = {};
  renderRecipientSummary();
  renderMatrix();
}

async function uploadReport(reportType, file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".pdf") || file.type !== "application/pdf") {
    showAlert("只接受 .pdf 格式的報告檔案。");
    return;
  }
  const formData = new FormData();
  formData.append("reportType", reportType);
  formData.append("report", file);
  const response = await fetch("/api/upload-report", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) return showAlert(data.error || "PDF 上傳失敗。");
  showAlert("");
  state.reportFiles[reportType] = data;
  renderPdfStatus();
  renderMatrix();
}

function previewPayload() {
  return {
    fundName: $("#fundName").value.trim(),
    reportPeriod: $("#reportPeriod").value.trim(),
    languageMode: state.languageMode,
    recipients: state.recipients,
    selectionMatrix: state.selectionMatrix,
  };
}

function validateBeforePreview() {
  const errors = [];
  if (!$("#fundName").value.trim()) errors.push("基金／專案名稱不可空白。");
  if (!$("#reportPeriod").value.trim()) errors.push("報告期間不可空白。");
  if (!state.recipients.length) errors.push("請先上傳收件人 Excel。");
  if (!draftCount()) errors.push("請至少勾選一位可寄送收件人的報告。");
  $("#validationList").innerHTML = errors.map((error) => `<li>${escapeHtml(error)}</li>`).join("");
  $("#validationList").classList.toggle("hidden", !errors.length);
  return errors;
}

async function generatePreview() {
  const errors = validateBeforePreview();
  if (errors.length) return showAlert(errors[0]);
  const response = await fetch("/api/preview-simple", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(previewPayload()),
  });
  const data = await response.json();
  if (!response.ok) return showAlert(data.error || "產生預覽失敗。");
  showAlert("");
  state.previews = data.previews.map((preview) => ({ ...preview, include: preview.status === "ready" }));
  renderPreviewRows();
}

function renderPreviewRows() {
  $("#previewCount").textContent = `${state.previews.filter((item) => item.include).length} 封預覽 Email`;
  const rows = $("#previewRows");
  if (!state.previews.length) {
    rows.innerHTML = `<tr><td colspan="9">請先勾選寄送對象並產生預覽。</td></tr>`;
    return;
  }
  rows.innerHTML = state.previews.map((preview) => `
    <tr class="${preview.include ? "" : "muted-row"}">
      <td><label class="inline-check"><input type="checkbox" data-preview-include="${escapeHtml(preview.recipientId)}" ${preview.include ? "checked" : ""} ${preview.status !== "ready" ? "disabled" : ""} /> ${preview.include ? "建立" : "不建立"}</label></td>
      <td>${escapeHtml(preview.recipientName)}</td>
      <td>${escapeHtml(preview.email || "-")}</td>
      <td>${escapeHtml(preview.cc || "-")}</td>
      <td>${escapeHtml(preview.subject)}</td>
      <td>${preview.attachmentCount}</td>
      <td>${preview.attachments.map(escapeHtml).join("<br>")}</td>
      <td><span class="pill ${preview.status === "ready" ? "" : "blocked"}">${preview.status === "ready" ? "正常" : "錯誤：" + escapeHtml(preview.errorMessage)}</span></td>
      <td><button type="button" class="icon-button text-icon" data-preview="${escapeHtml(preview.recipientId)}">預覽</button></td>
    </tr>
  `).join("");

  rows.querySelectorAll("[data-preview-include]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const preview = state.previews.find((item) => item.recipientId === event.target.dataset.previewInclude);
      preview.include = event.target.checked;
      renderPreviewRows();
    });
  });
  rows.querySelectorAll("[data-preview]").forEach((button) => {
    button.addEventListener("click", () => openPreview(state.previews.find((item) => item.recipientId === button.dataset.preview)));
  });
}

function openPreview(preview) {
  $("#emailMeta").innerHTML = `
    <dt>收件人</dt><dd>${escapeHtml(preview.email || "-")}</dd>
    <dt>副本</dt><dd>${escapeHtml(preview.cc || "-")}</dd>
    <dt>主旨</dt><dd>${escapeHtml(preview.subject)}</dd>
    <dt>附件</dt><dd>${preview.attachments.map(escapeHtml).join("<br>")}</dd>
  `;
  $("#emailBody").textContent = preview.body;
  $("#modal").classList.remove("hidden");
}

async function createDrafts() {
  if (!state.previews.length) return showAlert("請先產生預覽。");
  const included = state.previews.filter((preview) => preview.include && preview.status === "ready");
  if (!included.length) return showAlert("沒有可建立的 Gmail 草稿。");
  const button = $("#createDrafts");
  button.disabled = true;
  button.textContent = "草稿建立中...";
  const response = await fetch("/api/create-drafts-simple", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...previewPayload(), previews: included }),
  });
  const data = await response.json();
  button.disabled = false;
  button.textContent = "建立 Gmail 草稿";
  if (!response.ok) return showAlert(data.error || "建立 Gmail 草稿失敗。");
  renderDraftResults(data);
}

function renderDraftResults(data) {
  $("#draftResults").innerHTML = `
    <div class="result-summary">
      <strong>${data.successCount}</strong> 成功
      <strong>${data.failedCount}</strong> 失敗
      <span>批次 ID：${escapeHtml(data.batchId)}</span>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>收件人</th><th>狀態</th><th>草稿 ID</th><th>錯誤訊息</th></tr></thead>
        <tbody>${data.results.map((result) => `
          <tr>
            <td>${escapeHtml(result.recipientName)}</td>
            <td><span class="pill ${result.createStatus === "success" ? "" : "failed"}">${result.createStatus === "success" ? "成功" : "失敗"}</span></td>
            <td>${escapeHtml(result.gmailDraftId || "-")}</td>
            <td>${escapeHtml(result.errorMessage || "-")}</td>
          </tr>`).join("")}</tbody>
      </table>
    </div>
  `;
}

function renderTemplateForm() {
  if (!state.templates) return;
  const template = state.templates[state.templateLang];
  $("#templateName").value = template.name || "";
  $("#templateLanguage").value = template.language || state.templateLang;
  $("#templateSubject").value = template.subject || "";
  $("#templateBody").value = template.body || "";
  renderTemplatePreview();
}

function applyTemplate(text, vars) {
  let output = text;
  for (const [key, value] of Object.entries(vars)) output = output.replaceAll(`{{${key}}}`, value);
  return output;
}

function renderTemplatePreview() {
  const vars = {
    fund_name: $("#fundName").value || "TGVest Fund",
    report_period: $("#reportPeriod").value || "2026年4月",
    recipient_name: "林先生",
    nickname: "林先生",
    email: "lin.demo@example.com",
    cc: "",
    attachment_list: "- 月報\n- 季報",
    report_count: "2",
    report_names: "月報, 季報",
  };
  $("#templatePreviewMeta").innerHTML = `<dt>主旨</dt><dd>${escapeHtml(applyTemplate($("#templateSubject").value, vars))}</dd>`;
  $("#templatePreviewBody").textContent = applyTemplate($("#templateBody").value, vars);
}

async function loadTemplates() {
  const response = await fetch("/api/templates");
  const data = await response.json();
  state.templates = data.templates;
  renderTemplateForm();
}

async function saveTemplates() {
  state.templates[state.templateLang] = {
    name: $("#templateName").value,
    language: $("#templateLanguage").value,
    subject: $("#templateSubject").value,
    body: $("#templateBody").value,
  };
  const response = await fetch("/api/templates", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ templates: state.templates }),
  });
  const data = await response.json();
  if (!response.ok) return showAlert(data.error || "儲存模板失敗。");
  state.templates = data.templates;
  showAlert("");
  renderTemplateForm();
}

function switchMainTab(tab) {
  $("#draftPage").classList.toggle("hidden", tab !== "draft");
  $("#templatesPage").classList.toggle("hidden", tab !== "templates");
  document.querySelectorAll("[data-page-tab]").forEach((button) => button.classList.toggle("active", button.dataset.pageTab === tab));
  if (tab === "templates") renderTemplatePreview();
}

$("#uploadRecipients").addEventListener("click", uploadRecipients);
$("#recipientExcel").addEventListener("change", () => {
  const hasFile = !!$("#recipientExcel").files?.[0];
  const btn = $("#uploadRecipients");
  btn.disabled = !hasFile;
  btn.title = hasFile ? "" : "請先上傳收件人 Excel";
});
document.querySelectorAll("[data-report-upload]").forEach((input) => input.addEventListener("change", (event) => uploadReport(event.target.dataset.reportUpload, event.target.files?.[0])));
$("#recipientSearch").addEventListener("input", renderMatrix);
$("#recipientFilter").addEventListener("change", renderMatrix);
$("#generatePreview").addEventListener("click", generatePreview);
$("#createDrafts").addEventListener("click", createDrafts);
$("#saveTemplates").addEventListener("click", saveTemplates);
["templateName", "templateLanguage", "templateSubject", "templateBody", "fundName", "reportPeriod"].forEach((id) => $(`#${id}`).addEventListener("input", renderTemplatePreview));
$("#languageModeFields").addEventListener("click", (event) => {
  if (!event.target.dataset.language) return;
  state.languageMode = event.target.dataset.language;
  [...$("#languageModeFields").querySelectorAll("button")].forEach((button) => button.classList.toggle("active", button.dataset.language === state.languageMode));
});
document.querySelectorAll("[data-page-tab]").forEach((button) => button.addEventListener("click", () => switchMainTab(button.dataset.pageTab)));
document.querySelectorAll("[data-template-lang]").forEach((button) => button.addEventListener("click", () => {
  state.templateLang = button.dataset.templateLang;
  document.querySelectorAll("[data-template-lang]").forEach((item) => item.classList.toggle("active", item.dataset.templateLang === state.templateLang));
  renderTemplateForm();
}));
$("#closeModal").addEventListener("click", () => $("#modal").classList.add("hidden"));
$("#modal").addEventListener("click", (event) => {
  if (event.target.id === "modal") $("#modal").classList.add("hidden");
});

renderRecipientSummary();
renderPdfStatus();
renderMatrix();
renderPreviewRows();
loadTemplates();
