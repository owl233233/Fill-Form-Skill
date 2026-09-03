/* ============ 智填 FormFiller 前端逻辑 ============ */
"use strict";

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

/* ---------------- 工具 ---------------- */
async function api(url, options = {}) {
  const res = await fetch(url, options);
  let data = null;
  try { data = await res.json(); } catch (_) {}
  if (!res.ok) {
    const msg = data && (data.detail || data.message) || `请求失败（${res.status}）`;
    const err = new Error(typeof msg === "string" ? msg : "请求失败");
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

function toast(msg, type = "", ms = 3200) {
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${type === "ok" ? "✅" : type === "err" ? "⛔" : type === "warn" ? "⚠️" : "💡"}</span><span>${esc(msg)}</span>`;
  $("#toast-wrap").appendChild(el);
  setTimeout(() => { el.style.opacity = "0"; el.style.transition = "opacity .3s"; setTimeout(() => el.remove(), 320); }, ms);
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

let _confirmCb = null;
function confirmBox(title, body, onOk) {
  $("#modal-title").textContent = title;
  $("#modal-body").textContent = body;
  _confirmCb = onOk;
  $("#modal-mask").classList.remove("hidden");
}
$("#modal-cancel").onclick = () => { $("#modal-mask").classList.add("hidden"); _confirmCb = null; };
$("#modal-ok").onclick = () => { $("#modal-mask").classList.add("hidden"); if (_confirmCb) _confirmCb(); _confirmCb = null; };
$("#modal-mask").addEventListener("click", e => { if (e.target === $("#modal-mask")) $("#modal-cancel").click(); });

/* ---------------- Tab 切换 ---------------- */
$("#tabs").addEventListener("click", e => {
  const btn = e.target.closest(".tab");
  if (!btn) return;
  $$(".tab").forEach(t => t.classList.toggle("active", t === btn));
  $$(".page").forEach(p => p.classList.toggle("active", p.id === `page-${btn.dataset.tab}`));
  if (btn.dataset.tab === "library") loadLibrary();
  if (btn.dataset.tab === "templates") loadTemplates();
  if (btn.dataset.tab === "fill") initFill();
  if (btn.dataset.tab === "records") loadRecords();
});
function gotoTab(name) { $(`#tabs .tab[data-tab="${name}"]`).click(); }

/* ---------------- 文件类型图标 ---------------- */
function fileIconMeta(filename, isImage) {
  if (isImage) return { cls: "ft-img", label: "IMG" };
  const ext = (filename.split(".").pop() || "").toLowerCase();
  const map = { docx: ["ft-docx", "W"], xlsx: ["ft-xlsx", "X"], xls: ["ft-xlsx", "X"],
                pdf: ["ft-pdf", "P"], txt: ["ft-txt", "T"], md: ["ft-txt", "T"],
                csv: ["ft-txt", "T"], json: ["ft-txt", "T"] };
  const [cls, label] = map[ext] || ["ft-txt", "?"];
  return { cls, label };
}

/* ================= 资料库 ================= */
let libCache = { entries: [], files: [], images: [] };

async function loadLibrary() {
  try {
    libCache = await api("/api/library");
  } catch (e) { toast(e.message, "err"); return; }
  renderStats();
  renderEntries();
  renderFiles();
}

function renderStats() {
  const { entries, files, images } = libCache;
  $("#library-stats").innerHTML = `
    <div class="stat-card"><div class="stat-icon">🗂️</div><div><div class="stat-num">${entries.length}</div><div class="stat-label">信息条目</div></div></div>
    <div class="stat-card"><div class="stat-icon">📄</div><div><div class="stat-num">${files.length}</div><div class="stat-label">资料文件</div></div></div>
    <div class="stat-card"><div class="stat-icon">🖼️</div><div><div class="stat-num">${images.length}</div><div class="stat-label">图片素材</div></div></div>`;
  $("#entry-count").textContent = entries.length;
  $("#file-count").textContent = libCache.files.length;
}

function renderEntries() {
  const { entries } = libCache;
  $("#entry-empty").style.display = entries.length ? "none" : "block";
  $("#entry-tbody").innerHTML = entries.map(e => `
    <tr>
      <td class="entry-key">${esc(e.key)}</td>
      <td class="entry-value">${esc(e.value)}</td>
      <td class="entry-src">${esc(e.source)}</td>
      <td><button class="icon-btn" title="删除" onclick="delEntry('${esc(e.name)}')">🗑️</button></td>
    </tr>`).join("");
}

function renderFiles() {
  const { files, images } = libCache;
  const all = [
    ...files.map(f => ({ ...f, isImage: false })),
    ...images.map(i => ({ ...i, filename: i.filename, isImage: true, uploaded_at: i.uploaded_at })),
  ];
  $("#file-empty").style.display = all.length ? "none" : "block";
  $("#file-list").innerHTML = all.map(f => {
    const ic = fileIconMeta(f.filename, f.isImage);
    return `
    <div class="file-card">
      ${f.isImage
        ? `<img class="img-thumb" src="/data/library/${encodeURIComponent(f.filename)}" onerror="this.outerHTML='<div class=&quot;file-type-icon ${ic.cls}&quot;>${ic.label}</div>'">`
        : `<div class="file-type-icon ${ic.cls}">${ic.label}</div>`}
      <div class="file-meta">
        <div class="file-name">${esc(f.filename)}</div>
        <div class="file-sub">${esc(f.uploaded_at || "")}</div>
      </div>
      <button class="icon-btn" title="删除" onclick="delFile('${esc(f.id)}')">🗑️</button>
    </div>`;
  }).join("");
}

window.delEntry = async name => {
  confirmBox("删除信息", `确定删除这条信息吗？删除后填表时将无法自动匹配。`, async () => {
    try { await api(`/api/library/entry/${encodeURIComponent(name)}`, { method: "DELETE" }); toast("已删除", "ok"); loadLibrary(); }
    catch (e) { toast(e.message, "err"); }
  });
};

window.delFile = async id => {
  confirmBox("删除文件", `确定从资料库删除该文件吗？（已提取的信息会保留）`, async () => {
    try { await api(`/api/library/file/${id}`, { method: "DELETE" }); toast("已删除", "ok"); loadLibrary(); }
    catch (e) { toast(e.message, "err"); }
  });
};

$("#btn-add-entry").onclick = async () => {
  const key = $("#add-key").value.trim(), value = $("#add-value").value.trim();
  if (!key || !value) { toast("请填写字段名和内容", "warn"); return; }
  const fd = new FormData();
  fd.append("key", key); fd.append("value", value);
  try {
    const r = await api("/api/library/entry", { method: "POST", body: fd });
    toast(r.existed ? `已更新「${key}」` : `已添加「${key}」`, "ok");
    $("#add-key").value = ""; $("#add-value").value = "";
    loadLibrary();
  } catch (e) { toast(e.message, "err"); }
};

/* ---- 上传（拖拽 + 点击，进度提示）---- */
function bindDropzone(zoneId, inputId, handler) {
  const zone = $(zoneId), input = $(inputId);
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => { if (input.files.length) handler([...input.files]); input.value = ""; });
  ["dragover", "dragenter"].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add("dragover"); }));
  ["dragleave", "drop"].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove("dragover"); }));
  zone.addEventListener("drop", e => {
    const files = [...(e.dataTransfer?.files || [])];
    if (files.length) handler(files);
  });
}

bindDropzone("#lib-dropzone", "#lib-file-input", uploadLibraryFiles);

async function uploadLibraryFiles(files) {
  for (const file of files) {
    const fd = new FormData();
    fd.append("file", file);
    try {
      const r = await api("/api/library/upload", { method: "POST", body: fd });
      if (r.category === "image") {
        toast(`图片「${r.filename}」已加入素材库`, "ok");
      } else if (r.extracted_count > 0) {
        toast(`「${r.filename}」提取到 ${r.extracted_count} 条信息：${(r.added_keys || []).join("、")}`, "ok", 4200);
      } else {
        toast(`「${r.filename}」已上传（未识别到键值信息，可手动补充）`, "warn", 4200);
      }
    } catch (e) { toast(`「${file.name}」${e.message}`, "err", 4200); }
  }
  loadLibrary();
}

/* ================= 模板管理 ================= */
let tplCache = [];

async function loadTemplates() {
  try { tplCache = await api("/api/templates"); }
  catch (e) { toast(e.message, "err"); return; }
  renderTemplates();
}

function tplChips(t) {
  const textChips = t.text_fields.map(f => `<span class="chip">{{${esc(f)}}}</span>`).join("");
  const imgChips = t.img_fields.map(f => `<span class="chip chip-img">🖼️{{img:${esc(f)}}}</span>`).join("");
  return `<div class="chip-wrap">${textChips}${imgChips || ""}</div>`;
}

function renderTemplates() {
  $("#tpl-empty").style.display = tplCache.length ? "none" : "block";
  $("#tpl-grid").innerHTML = tplCache.map(t => `
    <div class="tpl-card">
      <div class="tpl-card-head">
        <div class="tpl-name">📄 ${esc(t.filename)}</div>
        <span class="tpl-type ${t.type === "docx" ? "ft-docx" : "ft-xlsx"}">${t.type.toUpperCase()}</span>
      </div>
      ${tplChips(t)}
      <div class="tpl-naming"><strong>命名规则：</strong>${t.naming_pattern ? esc(t.naming_pattern) : "（默认：模板名_日期）"}</div>
      <div class="tpl-card-actions">
        <a class="btn btn-ghost" href="/api/download/template/${t.id}">⬇️ 下载模板</a>
        <button class="btn btn-ghost" onclick="delTpl('${t.id}')">🗑️ 删除</button>
      </div>
    </div>`).join("");
}

window.delTpl = async id => {
  confirmBox("删除模板", "确定删除该模板吗？已生成的文件不受影响。", async () => {
    try { await api(`/api/templates/${id}`, { method: "DELETE" }); toast("已删除", "ok"); loadTemplates(); }
    catch (e) { toast(e.message, "err"); }
  });
};

bindDropzone("#tpl-dropzone", "#tpl-file-input", async files => {
  for (const file of files) {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("naming_pattern", $("#tpl-naming").value.trim());
    try {
      const r = await api("/api/templates/upload", { method: "POST", body: fd });
      const n = r.text_fields.length + r.img_fields.length;
      toast(`模板「${r.filename}」已上传，识别到 ${n} 个占位符`, "ok", 4000);
    } catch (e) { toast(`「${file.name}」${e.message}`, "err", 4200); }
  }
  $("#tpl-naming").value = "";
  loadTemplates();
});

/* ================= 生成向导 ================= */
let fillState = { templateId: null, preview: null };

function initFill() {
  loadFillStep1();
}

async function loadFillStep1() {
  showStep(1);
  try { tplCache = await api("/api/templates"); } catch (e) { toast(e.message, "err"); return; }
  $("#fill-tpl-empty").style.display = tplCache.length ? "none" : "block";
  $("#tpl-select-list").innerHTML = tplCache.map(t => `
    <div class="tpl-option" data-id="${t.id}" onclick="selectTpl('${t.id}')">
      <div class="tpl-radio"></div>
      <div style="flex:1">
        <div style="font-weight:700;font-size:14px">${esc(t.filename)}</div>
        <div style="margin-top:6px">${tplChips(t)}</div>
      </div>
      <span class="tpl-type ${t.type === "docx" ? "ft-docx" : "ft-xlsx"}">${t.type.toUpperCase()}</span>
    </div>`).join("");
  fillState.templateId = null;
}

window.selectTpl = async id => {
  $$(".tpl-option").forEach(o => o.classList.toggle("selected", o.dataset.id === id));
  fillState.templateId = id;
  await loadPreview(id);
};

function showStep(n) {
  [1, 2, 3].forEach(i => {
    $(`#fill-step${i}`).classList.toggle("hidden", i !== n);
    const el = $(`.step[data-step="${i}"]`);
    el.classList.toggle("active", i === n);
    el.classList.toggle("done", i < n);
  });
  $("#fill-steps").classList.remove("hidden");
}

async function loadPreview(tid) {
  const fd = new FormData();
  fd.append("template_id", tid);
  let pv;
  try { pv = await api("/api/fill/preview", { method: "POST", body: fd }); }
  catch (e) { toast(e.message, "err"); return; }
  fillState.preview = pv;
  renderPreview();
  showStep(2);
  const missing = pv.fields.filter(f => f.missing);
  if (missing.length) {
    toast(`有 ${missing.length} 项信息缺失，请在下方补充`, "warn", 4000);
  }
}

function renderPreview() {
  const pv = fillState.preview;
  const missing = pv.fields.filter(f => f.missing);
  $("#field-count").textContent = `${pv.fields.length} 个字段`;
  $("#field-sub").textContent = missing.length
    ? `其中 ${missing.length} 项缺失，请补充后再生成`
    : "所有字段均已在资料库中匹配，可直接生成";

  $("#ask-banner").classList.toggle("hidden", !missing.length);
  $("#ask-fields").innerHTML = missing.map(f => `<span class="ask-chip">${esc(f.name)}</span>`).join("");

  $("#field-tbody").innerHTML = pv.fields.map(f => {
    let tag;
    if (f.missing) tag = `<span class="tag tag-miss">❗ 待补充</span>`;
    else if (f.dynamic) tag = `<span class="tag tag-dyn">⚡ 自动生成</span>`;
    else tag = `<span class="tag tag-ok">✓ 已匹配</span>`;
    const val = f.value ?? "";
    return `
    <tr>
      <td class="entry-key">{{${esc(f.name)}}}</td>
      <td>
        <input data-field="${esc(f.name)}" value="${esc(val)}" placeholder="请输入${esc(f.name)}" class="${f.missing ? "missing-input" : ""}">
        ${f.source ? `<div class="field-src">来自：${esc(f.source)}</div>` : f.dynamic ? `<div class="field-src">生成时自动填入当前${f.name.includes("时间") ? "时间" : "日期"}</div>` : ""}
      </td>
      <td>${tag}</td>
    </tr>`;
  }).join("");

  const hasImgs = pv.images.length > 0;
  $("#img-section").classList.toggle("hidden", !hasImgs);
  if (hasImgs) {
    $("#img-tbody").innerHTML = pv.images.map(im => `
      <tr>
        <td class="entry-key">🖼️ {{img:${esc(im.name)}}}</td>
        <td>${im.matched ? `素材：<strong>${esc(im.matched)}</strong>` : `<span style="color:var(--ink-3)">未匹配到同名图片（可忽略，将留空）</span>`}</td>
        <td>${im.matched ? `<span class="tag tag-img-ok">✓ 已匹配</span>` : `<span class="tag tag-img-miss">○ 可选</span>`}</td>
      </tr>`).join("");
  }

  $("#out-filename").value = pv.suggested_filename;
  $("#filename-hint").textContent = "已按模板命名规则生成，可修改";
}

$("#btn-back-step1").onclick = () => loadFillStep1();

$("#btn-generate").onclick = async () => {
  const pv = fillState.preview;
  const values = {};
  let lack = [];
  $$("#field-tbody input[data-field]").forEach(inp => {
    const v = inp.value.trim();
    values[inp.dataset.field] = v;
    const f = pv.fields.find(x => x.name === inp.dataset.field);
    if (f && f.missing && !v) lack.push(inp.dataset.field);
  });
  if (lack.length) {
    toast(`请补充缺失信息：${lack.join("、")}`, "warn", 4200);
    const first = $(`#field-tbody input[data-field="${CSS.escape(lack[0])}"]`);
    if (first) first.focus();
    return;
  }

  const btn = $("#btn-generate");
  btn.disabled = true; btn.textContent = "⏳ 正在生成…";
  const fd = new FormData();
  fd.append("template_id", fillState.templateId);
  fd.append("filename", $("#out-filename").value.trim());
  fd.append("values", JSON.stringify(values));
  fd.append("save_to_library", $("#save-to-library").checked);
  try {
    const r = await api("/api/fill/generate", { method: "POST", body: fd });
    renderResult(r);
    showStep(3);
  } catch (e) {
    if (e.status === 422 && e.data?.missing) {
      toast(`信息缺失：${e.data.missing.join("、")}，请补充`, "warn", 4200);
    } else {
      toast(e.message, "err", 4500);
    }
  } finally {
    btn.disabled = false; btn.textContent = "🚀 生成文件";
  }
};

function renderResult(r) {
  $("#result-title").textContent = "文件已生成 🎉";
  const extra = [];
  if (r.images_inserted) extra.push(`插入图片 ${r.images_inserted} 张`);
  if (r.saved_keys?.length) extra.push(`已存入资料库：${r.saved_keys.join("、")}`);
  $("#result-sub").textContent = extra.length
    ? `格式与模板保持一致。${extra.join("；")}。`
    : "格式与模板保持一致，可直接下载使用。";
  $("#result-file").innerHTML = `📄 ${esc(r.filename)}`;
  $("#btn-download").href = r.download_url;
}

$("#btn-goto-records").onclick = () => gotoTab("records");
$("#btn-again").onclick = () => initFill();

/* ================= 生成记录 ================= */
async function loadRecords() {
  let records;
  try { records = await api("/api/records"); }
  catch (e) { toast(e.message, "err"); return; }
  $("#record-empty").style.display = records.length ? "none" : "block";
  $("#record-table").style.display = records.length ? "table" : "none";
  $("#record-tbody").innerHTML = records.map(r => `
    <tr>
      <td>${esc(r.created_at)}</td>
      <td>${esc(r.template_name)}</td>
      <td><strong>${esc(r.filename)}</strong></td>
      <td><a class="record-dl" href="/api/download/output/${encodeURIComponent(r.filename)}">⬇️ 下载</a></td>
    </tr>`).join("");
}

/* ================= 启动 ================= */
loadLibrary();
