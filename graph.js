/* global Office */
"use strict";

/*
 * Acesso ao Microsoft Graph a partir do painel:
 *  - getToken(): abre auth.html numa janela de diálogo do Office e recebe o access token
 *    (guardado em memória até perto de expirar);
 *  - me(): conta com que se vai enviar;
 *  - sendMail(): POST /me/sendMail, com novas tentativas em 429/503 e renovação do token em 401.
 */
(function () {
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const MAX_RETRIES = 3;
  let cached = null; // { token, exp }

  const config = () => window.MAILMERGE_CONFIG || {};
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  /** ID que veio no endereço do painel, definido no manifest na instalação (taskpane.html?cid=…). */
  function clientIdFromManifest() {
    const v = new URLSearchParams(location.search).get("cid") || "";
    return GUID.test(v.trim()) ? v.trim() : "";
  }

  /**
   * Client ID, por esta ordem: config.js, manifest (instalação) ou o guardado no painel
   * (roamingSettings da caixa de correio).
   */
  function clientId() {
    const fromFile = String(config().clientId || "").trim();
    if (fromFile) return fromFile;
    if (clientIdFromManifest()) return clientIdFromManifest();
    try {
      return String(Office.context.roamingSettings.get("clientId") || "").trim();
    } catch {
      return "";
    }
  }

  function saveClientId(value) {
    const v = String(value || "").trim();
    if (v && !GUID.test(v)) return Promise.reject(new Error("O ID tem o formato 00000000-0000-0000-0000-000000000000."));
    const rs = Office.context.roamingSettings;
    if (v) rs.set("clientId", v);
    else rs.remove("clientId");
    cached = null;
    return new Promise((resolve, reject) =>
      rs.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(v) : reject(new Error(r.error && r.error.message))))
    );
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function folderUrl() {
    return location.href.split("#")[0].split("?")[0].replace(/[^/]*$/, "");
  }

  function openAuthDialog() {
    const hint = (Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "";
    const url = folderUrl() + "auth.html?hint=" + encodeURIComponent(hint) + "&cid=" + encodeURIComponent(clientId());
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Não foi possível abrir a janela de login: " + (result.error && result.error.message)));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          let msg;
          try {
            msg = JSON.parse(arg.message);
          } catch {
            msg = { ok: false, error: "Resposta inesperada da janela de login." };
          }
          if (msg.ok) resolve(msg);
          else reject(new Error("Login: " + msg.error));
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          reject(new Error(arg.error === 12006 ? "A janela de login foi fechada antes de terminar." : "Erro na janela de login (" + arg.error + ")."));
        });
      });
    });
  }

  async function getToken(force) {
    if (!force && cached && cached.cid === clientId() && Date.now() < cached.exp - 2 * 60 * 1000) return cached.token;
    const res = await openAuthDialog();
    cached = { token: res.accessToken, exp: Date.now() + res.expiresIn * 1000, cid: clientId() };
    return cached.token;
  }

  async function graphError(res) {
    let code = "HTTP " + res.status;
    let message = res.statusText || "";
    try {
      const body = await res.json();
      if (body.error) {
        code = body.error.code || code;
        message = body.error.message || message;
      }
    } catch {
      /* sem corpo JSON */
    }
    if (/quota|submissionquota|limit/i.test(code + " " + message))
      message = "Limite de envio da conta atingido. Tenta mais tarde ou envia menos de cada vez. (" + message + ")";
    const err = new Error(code + ": " + message);
    err.status = res.status;
    return err;
  }

  async function graphFetch(path, options) {
    let renewed = false;
    for (let attempt = 0; ; attempt++) {
      const token = await getToken();
      const res = await fetch(GRAPH + path, {
        ...options,
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(options && options.headers) },
      });
      if (res.ok) return res;
      if (res.status === 401 && !renewed) {
        // token expirado ou revogado: pede um novo e repete uma vez
        renewed = true;
        cached = null;
        continue;
      }
      if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < MAX_RETRIES) {
        const wait = Number(res.headers.get("Retry-After")) || 10 * (attempt + 1);
        await sleep(wait * 1000);
        continue;
      }
      throw await graphError(res);
    }
  }

  async function me() {
    const res = await graphFetch("/me?$select=displayName,mail,userPrincipalName");
    return res.json();
  }

  const recipients = (list) => (list || []).map((address) => ({ emailAddress: { address } }));

  async function sendMail({ to, cc, subject, html }) {
    await graphFetch("/me/sendMail", {
      method: "POST",
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: recipients(to),
          ccRecipients: recipients(cc),
        },
        saveToSentItems: true,
      }),
    });
  }

  window.Graph = {
    configured: () => !!clientId(),
    clientId,
    clientIdSource: () =>
      String(config().clientId || "").trim() ? "config" : clientIdFromManifest() ? "manifest" : "painel",
    saveClientId,
    delayMs: () => Number(config().delayMs) || 2500,
    getToken,
    me,
    sendMail,
  };
})();
