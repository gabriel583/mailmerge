/* global Office, XLSX, Merge */
"use strict";

/*
 * Dois modos, conforme o email onde o painel é aberto:
 *  - ESCRITA (compose): substitui os marcadores no próprio email com UMA linha
 *    e preenche Para/Cc/Bcc. Várias linhas -> junta os destinatários de todas
 *    e usa os valores da primeira.
 *  - LEITURA (read): o email aberto é o modelo; abre uma nova mensagem
 *    preenchida por cada linha selecionada (displayNewMessageFormAsync).
 */

const RECIPIENTS_PER_CALL = 100; // setAsync/addAsync aceitam no máximo 100 de cada vez
const EXCHANGE_DEFAULT_MAX = 500; // destinatários por mensagem no Exchange Online (por omissão)
const MAX_FORM_RECIPIENTS = 100; // por campo em displayNewMessageFormAsync
const MAX_FORM_BODY = 32 * 1024; // limite do htmlBody em displayNewMessageForm
const MAX_SUBJECT = 255;
const MAX_TABLE_ROWS = 500;
const CONFIRM_ABOVE = 10; // pede confirmação para abrir mais de N janelas
const OPEN_DELAY_MS = 700;

const $ = (id) => document.getElementById(id);

const state = {
  mode: null, // "compose" | "read"
  workbook: null,
  headers: [],
  rows: [],
  selected: new Set(),
  previewIndex: null,
  template: null, // { subject, html }
  confirmBig: false,
  confirmSend: false, // 1.º clique em "Enviar" pede confirmação
  sender: "",
  applied: false, // modo escrita: o email já tem uma linha aplicada
  sent: new Set(), // linhas já enviadas pelo Graph nesta sessão
  stop: false,
  busy: false,
};

// ---------- Arranque ----------

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook || !Office.context.mailbox) {
    setStatus("Abre este painel dentro do Outlook.", "error");
    return;
  }
  state.mode = detectMode();
  document.body.classList.add(state.mode === "compose" ? "mode-compose" : "mode-read");
  $("modeBadge").textContent = state.mode === "compose" ? "A escrever" : "A ler";
  $("modeHint").textContent =
    state.mode === "compose"
      ? "Este email é o modelo. \"Enviar\" manda um email por linha, cada um com os seus valores. \"Aplicar a este email\" preenche só este email com uma linha."
      : "Este email é o modelo. \"Enviar\" manda um email por linha, cada um com os seus valores. \"Abrir rascunhos\" abre-os para enviares à mão.";

  bindEvents();
  renderClientIdState();
  renderColumnSelects();
  readTemplate().catch((e) => setStatus("Não consegui ler o email: " + e.message, "error"));

  // Painel fixado em modo leitura: ao mudar de email, o novo email passa a ser o modelo.
  if (state.mode === "read" && Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      if (!Office.context.mailbox.item) return;
      state.mode = detectMode();
      readTemplate().catch((e) => setStatus("Não consegui ler o email: " + e.message, "error"));
    });
  }
});

function detectMode() {
  const item = Office.context.mailbox.item;
  return item && item.subject && typeof item.subject.getAsync === "function" ? "compose" : "read";
}

function bindEvents() {
  $("fileInput").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) loadFile(file);
  });
  $("sheetSelect").addEventListener("change", (e) => loadSheet(e.target.value));
  $("ptNumbers").addEventListener("change", () => {
    if (!state.workbook) return;
    const keep = new Set(state.selected);
    loadSheet($("sheetSelect").value || state.workbook.SheetNames[0]);
    state.selected = keep; // mantém a seleção ao mudar o formato
    renderTable();
    updatePreview();
  });
  $("readTemplateBtn").addEventListener("click", () =>
    readTemplate()
      .then(() => setStatus("Modelo lido do email atual.", "ok"))
      .catch((e) => setStatus("Não consegui ler o email: " + e.message, "error"))
  );
  $("toCol").addEventListener("change", () => {
    state.confirmBig = state.confirmSend = false;
    renderTable();
    updatePreview();
  });
  $("ccCol").addEventListener("change", updatePreview);
  $("toField").addEventListener("change", updatePreview);
  $("filter").addEventListener("input", renderTable);
  $("selectAll").addEventListener("change", (e) => {
    for (const i of visibleIndexes()) {
      if (e.target.checked) state.selected.add(i);
      else state.selected.delete(i);
    }
    state.confirmBig = state.confirmSend = false;
    renderTable();
    updatePreview();
  });
  $("rowsTable").querySelector("tbody").addEventListener("click", onRowClick);
  $("applyBtn").addEventListener("click", () => guard(applyToCurrent));
  $("resetBtn").addEventListener("click", () => guard(resetTemplate));
  $("draftsBtn").addEventListener("click", () => guard(openDrafts));
  $("sendBtn").addEventListener("click", () => guard(sendAll));
  $("saveClientIdBtn").addEventListener("click", () => guard(saveClientId));
  $("stopBtn").addEventListener("click", () => {
    state.stop = true;
    $("stopBtn").disabled = true;
    setStatus("A parar depois do email atual…", "warn");
  });
}

// ---------- Office: helpers ----------

function officeAsync(fn) {
  return new Promise((resolve, reject) =>
    fn((r) =>
      r.status === Office.AsyncResultStatus.Succeeded
        ? resolve(r.value)
        : reject(new Error((r.error && r.error.message) || "Erro desconhecido do Outlook"))
    )
  );
}

async function readTemplate() {
  const item = Office.context.mailbox.item;
  const html = await officeAsync((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
  const subject = state.mode === "compose" ? await officeAsync((cb) => item.subject.getAsync(cb)) : item.subject || "";
  state.template = { html, subject };
  $("resetBtn").disabled = state.mode !== "compose";
  renderPlaceholders();
  updatePreview();
}

// ---------- Excel ----------

async function loadFile(file) {
  try {
    const buf = await file.arrayBuffer();
    // dateNF: datas com o formato "por omissão" do Excel saem como dd/mm/aaaa (e não m/d/yy)
    state.workbook = XLSX.read(buf, { type: "array", dateNF: "dd/mm/yyyy", cellNF: true });
    $("fileLabel").textContent = file.name;
    const names = state.workbook.SheetNames;
    const sel = $("sheetSelect");
    sel.innerHTML = "";
    for (const n of names) sel.add(new Option(n, n));
    $("sheetRow").hidden = names.length < 2;
    loadSheet(names[0]);
  } catch (e) {
    setStatus("Não consegui ler o ficheiro: " + e.message, "error");
  }
}

/** Aplica (ou desfaz) o formato numérico PT ao texto das células numéricas que não são datas. */
function localizeNumbers(ws, on) {
  for (const addr of Object.keys(ws)) {
    if (addr[0] === "!") continue;
    const cell = ws[addr];
    if (cell.t !== "n" || cell.w == null) continue;
    if (cell.__w === undefined) cell.__w = cell.w; // texto original
    const isDate = cell.z != null && XLSX.SSF.is_date(cell.z);
    cell.w = on && !isDate ? Merge.localizeNumberText(cell.__w) : cell.__w;
  }
}

function loadSheet(name) {
  const ws = state.workbook.Sheets[name];
  localizeNumbers(ws, $("ptNumbers").checked);
  const rows = XLSX.utils.sheet_to_json(ws, { raw: false, defval: "", blankrows: false });
  const headers = [];
  for (const r of rows) for (const k of Object.keys(r)) if (!k.startsWith("__") && !headers.includes(k)) headers.push(k);
  state.rows = rows;
  state.headers = headers;
  state.selected = new Set();
  state.sent = new Set();
  state.previewIndex = null;
  state.confirmBig = state.confirmSend = false;
  $("fileInfo").textContent = `${rows.length} linha(s) · colunas: ${headers.join(", ")}`;
  renderColumnSelects();
  renderPlaceholders();
  renderTable();
  updatePreview();
  setStatus(rows.length ? "Excel carregado. Seleciona as linhas a usar." : "A folha não tem linhas de dados.", rows.length ? "ok" : "warn");
}

function renderColumnSelects() {
  const fill = (sel, guess) => {
    const prev = sel.value;
    sel.innerHTML = "";
    sel.add(new Option("(nenhuma)", ""));
    for (const h of state.headers) sel.add(new Option(h, h));
    sel.value = state.headers.includes(prev) ? prev : guess || "";
  };
  fill($("toCol"), Merge.guessEmailColumn(state.headers, state.rows));
  fill($("ccCol"), Merge.guessCcColumn(state.headers));
}

// ---------- Marcadores ----------

function renderPlaceholders() {
  const box = $("placeholders");
  box.innerHTML = "";
  if (!state.template) return;
  const keys = Merge.findPlaceholders(state.template.html, state.template.subject);
  if (!keys.length) {
    box.innerHTML = '<span class="muted small">Nenhum marcador {{…}} encontrado neste email.</span>';
    return;
  }
  const cols = new Set(state.headers.map(Merge.normalizeKey));
  for (const k of keys) {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `{{${k}}}`;
    if (state.headers.length) {
      const ok = cols.has(Merge.normalizeKey(k));
      chip.classList.add(ok ? "ok" : "missing");
      chip.title = ok ? "Coluna encontrada no Excel" : "Não existe coluna com este nome no Excel";
    }
    box.appendChild(chip);
  }
}

// ---------- Tabela ----------

function visibleIndexes() {
  const q = Merge.normalizeKey($("filter").value);
  const out = [];
  state.rows.forEach((r, i) => {
    if (!q || state.headers.some((h) => Merge.normalizeKey(r[h]).includes(q))) out.push(i);
  });
  return out;
}

function renderTable() {
  const thead = $("rowsTable").querySelector("thead");
  const tbody = $("rowsTable").querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";
  if (!state.headers.length) {
    updateSelectionUi([]);
    return;
  }
  const toCol = $("toCol").value;
  const cols = toCol ? [toCol, ...state.headers.filter((h) => h !== toCol)] : state.headers;

  const hr = thead.insertRow();
  hr.appendChild(document.createElement("th"));
  const thNum = document.createElement("th");
  thNum.textContent = "#";
  hr.appendChild(thNum);
  for (const h of cols) {
    const th = document.createElement("th");
    th.textContent = h;
    if (h === toCol) th.className = "to-col";
    hr.appendChild(th);
  }

  const visible = visibleIndexes();
  for (const i of visible.slice(0, MAX_TABLE_ROWS)) {
    const tr = tbody.insertRow();
    tr.dataset.index = i;
    if (state.selected.has(i)) tr.classList.add("selected");
    if (i === state.previewIndex) tr.classList.add("previewing");
    if (state.sent.has(i)) tr.classList.add("sent");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = state.selected.has(i);
    cb.setAttribute("aria-label", `Selecionar linha ${excelRow(i)}`);
    tr.insertCell().appendChild(cb);
    const num = tr.insertCell();
    num.className = "rownum";
    num.textContent = (state.sent.has(i) ? "✓ " : "") + excelRow(i);
    if (state.sent.has(i)) num.title = "Já enviado";
    for (const h of cols) {
      const td = tr.insertCell();
      td.textContent = state.rows[i][h];
      td.title = state.rows[i][h];
      if (h === toCol) td.className = "to-col";
    }
  }
  if (visible.length > MAX_TABLE_ROWS) {
    const tr = tbody.insertRow();
    const td = tr.insertCell();
    td.colSpan = cols.length + 2;
    td.className = "muted";
    td.textContent = `… e mais ${visible.length - MAX_TABLE_ROWS} linhas (usa o filtro ou "Todas").`;
  }
  updateSelectionUi(visible);
}

function onRowClick(e) {
  const tr = e.target.closest("tr[data-index]");
  if (!tr) return;
  const i = Number(tr.dataset.index);
  if (e.target.matches('input[type="checkbox"]')) {
    if (e.target.checked) state.selected.add(i);
    else state.selected.delete(i);
    state.confirmBig = state.confirmSend = false;
  } else if (!state.selected.has(i)) {
    // clicar na linha seleciona-a e mostra-a na pré-visualização
    state.selected.add(i);
    state.confirmBig = state.confirmSend = false;
  }
  state.previewIndex = state.selected.has(i) ? i : null;
  renderTable();
  updatePreview();
}

function updateSelectionUi(visible) {
  const n = state.selected.size;
  $("selCount").textContent = state.rows.length ? `${n} de ${state.rows.length}` : "";
  $("selectAll").checked = visible.length > 0 && visible.every((i) => state.selected.has(i));
  const ready = n > 0 && !!state.template;
  $("applyBtn").disabled = !ready;
  $("draftsBtn").disabled = !ready;
  $("draftsBtn").textContent = n > 1 ? `Abrir ${n} rascunhos` : "Abrir rascunho";
  $("sendBtn").disabled = !ready;
  $("sendBtn").textContent = state.confirmSend
    ? `Confirmar: enviar ${n}`
    : n > 1 ? `Enviar ${n} emails` : "Enviar 1 email";

  // No modo de escrita, várias linhas + "Aplicar" = UM só email: explicar a diferença.
  const multi = state.mode === "compose" && n > 1;
  const hint = $("multiHint");
  if (hint) {
    hint.hidden = !multi;
    if (multi)
      hint.textContent =
        `"Enviar ${n} emails" manda um email a cada pessoa, cada um com os seus valores. ` +
        `"Aplicar a este email" junta as ${n} linhas num só email, com os valores da 1.ª linha.`;
  }
}

/** Número da linha no Excel (cabeçalho = linha 1). */
function excelRow(i) {
  const r = state.rows[i];
  return r && r.__rowNum__ != null ? r.__rowNum__ + 1 : i + 2;
}

function selectedIndexes() {
  return [...state.selected].sort((a, b) => a - b);
}

// ---------- Merge ----------

/** Resultado do merge para um conjunto de linhas (valores da primeira). */
function buildMessage(rowIndexes) {
  const rows = rowIndexes.map((i) => state.rows[i]);
  const valuesRow = rows[0];
  const body = Merge.mergeHtml(state.template.html, valuesRow);
  const subject = Merge.mergeText(state.template.subject, valuesRow);
  const to = Merge.collectEmails(rows, $("toCol").value);
  const cc = Merge.collectEmails(rows, $("ccCol").value);
  const missing = [...new Set([...body.missing, ...subject.missing])];
  return { html: body.html, subject: subject.text, to, cc, missing };
}

function updatePreview() {
  const box = $("preview");
  if (!state.template || !state.rows.length) {
    box.hidden = true;
    return;
  }
  const sel = selectedIndexes();
  let idx = state.previewIndex;
  if (idx == null || !state.selected.has(idx)) idx = sel.length ? sel[0] : 0;

  // Mostra o email de UMA linha, tal como sai no "Enviar" (clica noutra linha para a ver).
  const msg = buildMessage([idx]);
  box.hidden = false;
  const pos = sel.indexOf(idx);
  $("pvRowLabel").textContent = sel.length
    ? `Email da linha ${excelRow(idx)} do Excel` + (sel.length > 1 ? ` (${pos + 1} de ${sel.length}; clica noutra linha para a ver).` : ".")
    : "Sem linhas selecionadas — a mostrar a 1.ª linha.";
  $("pvTo").textContent = msg.to.valid.join("; ") || "—";
  $("pvCcLine").hidden = !$("ccCol").value;
  $("pvCc").textContent = msg.cc.valid.join("; ") || "—";
  $("pvSubject").textContent = msg.subject;
  $("pvBody").srcdoc = msg.html;
}

// ---------- Ações ----------

async function guard(fn) {
  if (state.busy) return;
  state.busy = true;
  try {
    await fn();
  } catch (e) {
    setStatus("Erro: " + e.message, "error");
  } finally {
    state.busy = false;
  }
}

/** Modo escrita: aplica ao email aberto. */
async function applyToCurrent() {
  const item = Office.context.mailbox.item;
  const sel = selectedIndexes();
  const msg = buildMessage(sel);
  const warnings = [];

  if ($("toCol").value && !msg.to.valid.length) warnings.push("Nenhum email válido na coluna de destinatários.");
  const total = msg.to.valid.length + msg.cc.valid.length;
  if (total > EXCHANGE_DEFAULT_MAX)
    warnings.push(
      `Atenção: ${total} destinatários numa só mensagem. O Exchange Online aceita por omissão até ${EXCHANGE_DEFAULT_MAX} por mensagem (o administrador pode subir até 1000); acima disso o envio é rejeitado.`
    );

  await officeAsync((cb) => item.body.setAsync(msg.html, { coercionType: Office.CoercionType.Html }, cb));
  await officeAsync((cb) => item.subject.setAsync(msg.subject.slice(0, MAX_SUBJECT), cb));
  if ($("toCol").value) await setRecipients($("toField").value === "bcc" ? item.bcc : item.to, msg.to.valid);
  if ($("ccCol").value) await setRecipients(item.cc, msg.cc.valid);

  if (msg.missing.length) warnings.push("Marcadores sem coluna no Excel (ficaram como estavam): " + msg.missing.map((k) => `{{${k}}}`).join(", "));
  const invalid = [...msg.to.invalid, ...msg.cc.invalid];
  if (invalid.length) warnings.push("Endereços ignorados por serem inválidos: " + invalid.join(", "));

  const what = sel.length > 1 ? `valores da linha ${excelRow(sel[0])}, destinatários de ${sel.length} linhas` : `linha ${excelRow(sel[0])}`;
  state.applied = true;
  setStatus(`Aplicado (${what}). Revê e envia.` + (warnings.length ? "\n" + warnings.join("\n") : ""), warnings.length ? "warn" : "ok");
}

/**
 * Preenche um campo (Para/Cc/Bcc) com qualquer número de endereços.
 * O Outlook só aceita 100 por chamada: o primeiro bloco substitui o campo
 * (setAsync) e os seguintes acrescentam (addAsync).
 */
async function setRecipients(field, list) {
  await officeAsync((cb) => field.setAsync(list.slice(0, RECIPIENTS_PER_CALL), cb));
  for (let k = RECIPIENTS_PER_CALL; k < list.length; k += RECIPIENTS_PER_CALL) {
    const chunk = list.slice(k, k + RECIPIENTS_PER_CALL);
    try {
      await officeAsync((cb) => field.addAsync(chunk, cb));
    } catch (e) {
      throw new Error(
        `Só entraram ${k} de ${list.length} endereços (${e.message}). O Outlook clássico limita cada campo a 500 ou 1000 destinatários; no Outlook na web e no novo Outlook não há esse limite.`
      );
    }
  }
}

/** Modo escrita: volta a pôr o assunto/corpo originais com os marcadores. */
async function resetTemplate() {
  const item = Office.context.mailbox.item;
  await officeAsync((cb) => item.body.setAsync(state.template.html, { coercionType: Office.CoercionType.Html }, cb));
  await officeAsync((cb) => item.subject.setAsync(state.template.subject, cb));
  state.applied = false;
  setStatus("Modelo reposto.", "ok");
}

/** Modo leitura: abre uma mensagem nova por linha selecionada. */
async function openDrafts() {
  const sel = selectedIndexes();
  if (sel.length > CONFIRM_ABOVE && !state.confirmBig) {
    state.confirmBig = true;
    $("draftsBtn").textContent = `Confirmar: abrir ${sel.length} janelas`;
    setStatus(`Vais abrir ${sel.length} janelas de mensagem. Clica outra vez no botão para confirmar.`, "warn");
    return;
  }
  state.confirmBig = state.confirmSend = false;

  const mailbox = Office.context.mailbox;
  const hasAsync = Office.context.requirements.isSetSupported("Mailbox", "1.9");
  if (!hasAsync && typeof mailbox.displayNewMessageForm !== "function")
    throw new Error("Esta versão do Outlook não permite abrir mensagens novas a partir de um suplemento.");

  const problems = [];
  let opened = 0;
  for (let n = 0; n < sel.length; n++) {
    const i = sel[n];
    const line = `Linha ${excelRow(i)}`;
    const msg = buildMessage([i]);
    if ($("toCol").value && !msg.to.valid.length) {
      problems.push(`${line}: sem email válido — ignorada.`);
      continue;
    }
    if (msg.to.valid.length > MAX_FORM_RECIPIENTS || msg.cc.valid.length > MAX_FORM_RECIPIENTS) {
      problems.push(`${line}: mais de ${MAX_FORM_RECIPIENTS} destinatários — ignorada.`);
      continue;
    }
    const html = Merge.compactHtml(msg.html, MAX_FORM_BODY);
    if (Merge.byteLength(html) > MAX_FORM_BODY) {
      problems.push(`${line}: corpo com mais de 32 KB (limite do Outlook) — ignorada.`);
      continue;
    }
    const params = { toRecipients: msg.to.valid, subject: msg.subject.slice(0, MAX_SUBJECT), htmlBody: html };
    if (msg.cc.valid.length) params.ccRecipients = msg.cc.valid;

    setStatus(`A abrir ${n + 1} de ${sel.length}…`, "ok");
    try {
      if (hasAsync) await officeAsync((cb) => mailbox.displayNewMessageFormAsync(params, cb));
      else mailbox.displayNewMessageForm(params);
      opened++;
    } catch (e) {
      problems.push(`${line}: ${e.message}`);
    }
    if (n < sel.length - 1) await new Promise((r) => setTimeout(r, OPEN_DELAY_MS));
  }

  const missing = buildMessage([sel[0]]).missing;
  if (missing.length) problems.push("Marcadores sem coluna no Excel: " + missing.map((k) => `{{${k}}}`).join(", "));
  updateSelectionUi(visibleIndexes());
  setStatus(`${opened} rascunho(s) aberto(s).` + (problems.length ? "\n" + problems.join("\n") : ""), problems.length ? "warn" : "ok");
}

/**
 * Envia um email por linha selecionada, pelo Microsoft Graph (ver graph.js).
 * 1.º clique: inicia sessão e mostra de que conta vai sair; 2.º clique: envia.
 */
async function sendAll() {
  const sel = selectedIndexes();
  if (!Graph.configured()) {
    $("settings").open = true;
    $("clientIdInput").focus();
    $("settings").scrollIntoView({ behavior: "smooth", block: "nearest" });
    setStatus('Para enviar automaticamente, cola em baixo o ID da aplicação e carrega em "Guardar" (só é preciso uma vez).', "warn");
    return;
  }
  if (!$("toCol").value) throw new Error("Escolhe primeiro a coluna dos destinatários.");

  if (!state.confirmSend) {
    // No modo escrita, relê o email para apanhar as últimas alterações (se ainda não foi aplicado)
    if (state.mode === "compose" && !state.applied) await readTemplate();
    setStatus("A iniciar sessão na conta Microsoft…", "ok");
    const user = await Graph.me();
    state.sender = user.mail || user.userPrincipalName || user.displayName || "?";
    const mailbox = (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "";
    const notes = [];
    if (mailbox && state.sender.toLowerCase() !== mailbox.toLowerCase())
      notes.push(`Atenção: a sessão iniciada (${state.sender}) não é a desta caixa de correio (${mailbox}).`);
    const again = sel.filter((i) => state.sent.has(i)).length;
    if (again) notes.push(`${again} destas linhas já foram enviadas nesta sessão e vão ser enviadas outra vez.`);
    const missing = buildMessage([sel[0]]).missing;
    if (missing.length) notes.push("Marcadores sem coluna no Excel (ficam como estão): " + missing.map((k) => `{{${k}}}`).join(", "));
    const noAddress = sel.filter((i) => !buildMessage([i]).to.valid.length).length;
    if (noAddress) notes.push(`${noAddress} linha(s) sem email válido serão ignoradas.`);
    state.confirmSend = true;
    updateSelectionUi(visibleIndexes());
    setStatus(
      `Vais enviar ${sel.length} email(s), um por linha, a partir de ${state.sender}.` +
        (notes.length ? "\n" + notes.join("\n") : "") +
        "\nClica outra vez no botão para confirmar.",
      "warn"
    );
    return;
  }

  state.confirmSend = false;
  state.stop = false;
  $("stopBtn").hidden = false;
  $("stopBtn").disabled = false;
  $("sendBtn").disabled = true;
  const problems = [];
  let ok = 0;
  let failsInRow = 0;
  try {
    for (let n = 0; n < sel.length; n++) {
      if (state.stop) {
        problems.push(`Parado por ti: ${sel.length - n} linha(s) ficaram por enviar.`);
        break;
      }
      const i = sel[n];
      const line = `Linha ${excelRow(i)}`;
      const msg = buildMessage([i]);
      if (!msg.to.valid.length) {
        problems.push(`${line}: sem email válido — ignorada.`);
        continue;
      }
      setStatus(`A enviar ${n + 1} de ${sel.length} (${msg.to.valid.join(", ")})…`, "ok");
      try {
        await Graph.sendMail({ to: msg.to.valid, cc: msg.cc.valid, subject: msg.subject.slice(0, MAX_SUBJECT), html: msg.html });
        ok++;
        failsInRow = 0;
        state.sent.add(i);
        state.selected.delete(i); // evita reenviar por engano
      } catch (e) {
        problems.push(`${line}: ${e.message}`);
        if (++failsInRow >= 3) {
          problems.push("3 erros seguidos — envio interrompido.");
          break;
        }
      }
      if (n < sel.length - 1 && !state.stop) await new Promise((r) => setTimeout(r, Graph.delayMs()));
    }
  } finally {
    $("stopBtn").hidden = true;
    renderTable();
    updatePreview();
  }
  const tail = state.mode === "compose" ? " Este rascunho (o modelo) pode ser descartado." : "";
  setStatus(
    `${ok} email(s) enviados a partir de ${state.sender}. Estão nos Itens Enviados.${tail}` + (problems.length ? "\n" + problems.join("\n") : ""),
    problems.length ? "warn" : "ok"
  );
}

// ---------- Configuração do envio automático ----------

function renderClientIdState() {
  const source = Graph.clientIdSource(); // "config" | "manifest" | "painel"
  const fixed = source !== "painel";
  const id = Graph.clientId();
  $("clientIdInput").value = id;
  $("clientIdInput").disabled = fixed;
  $("saveClientIdBtn").disabled = fixed;
  $("clientIdState").textContent =
    source === "config" ? "Definido no config.js ✓"
    : source === "manifest" ? "Definido na instalação ✓"
    : id ? "Configurado ✓" : "Ainda não configurado.";
}

async function saveClientId() {
  const v = await Graph.saveClientId($("clientIdInput").value);
  state.confirmSend = false;
  renderClientIdState();
  updateSelectionUi(visibleIndexes());
  setStatus(v ? "ID guardado. Já podes usar o botão Enviar." : "ID removido.", "ok");
}

// ---------- UI ----------

function setStatus(text, kind) {
  const el = $("status");
  el.textContent = text;
  el.className = kind || "";
  // a barra de estado é fixa no fundo: reserva espaço para não tapar os botões
  document.body.style.paddingBottom = (text ? el.offsetHeight + 16 : 80) + "px";
}
