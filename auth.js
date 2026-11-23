/* global Office */
"use strict";

/*
 * Página aberta pelo painel numa janela de diálogo do Office (displayDialogAsync).
 * Faz o login na Microsoft com o fluxo "authorization code + PKCE" (o recomendado para
 * aplicações só de browser) e devolve o access token ao painel com messageParent.
 *
 * 1.ª vez: redireciona para login.microsoftonline.com; a Microsoft volta a esta página
 *          com ?code=…, que é trocado por tokens.
 * Seguintes: usa o refresh token guardado (válido 24 h) e nem mostra o login.
 */

const CFG = window.MAILMERGE_CONFIG || {};
const AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0";
const SCOPES = "https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access openid profile";
const REDIRECT_URI = location.origin + location.pathname;
const PKCE_KEY = "mm_pkce";

function rtKey(hint) {
  return "mm_rt_" + (hint || "default").toLowerCase();
}

function show(text) {
  document.getElementById("msg").textContent = text;
}

function reply(payload) {
  Office.context.ui.messageParent(JSON.stringify(payload));
}

function base64url(bytes) {
  let s = "";
  for (const b of new Uint8Array(bytes)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function randomString(len) {
  return base64url(crypto.getRandomValues(new Uint8Array(len)));
}

async function tokenRequest(params) {
  const res = await fetch(AUTHORITY + "/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPES, ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.error || "HTTP " + res.status);
  return data;
}

function finish(tokens, hint) {
  try {
    if (tokens.refresh_token) localStorage.setItem(rtKey(hint), tokens.refresh_token);
  } catch {
    /* sem localStorage: da próxima vez pede login outra vez */
  }
  reply({ ok: true, accessToken: tokens.access_token, expiresIn: Number(tokens.expires_in) || 3600 });
}

async function startLogin(hint) {
  const verifier = randomString(48);
  const challenge = base64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  const state = randomString(16);
  sessionStorage.setItem(PKCE_KEY, JSON.stringify({ verifier, state, hint }));
  const q = new URLSearchParams({
    client_id: CFG.clientId,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    scope: SCOPES,
    code_challenge: challenge,
    code_challenge_method: "S256",
    state,
    response_mode: "query",
  });
  if (hint) q.set("login_hint", hint);
  location.assign(AUTHORITY + "/authorize?" + q.toString());
}

async function run() {
  if (!CFG.clientId) {
    reply({ ok: false, error: "Falta o Client ID no ficheiro config.js." });
    return;
  }
  const params = new URLSearchParams(location.search);

  // Regresso da Microsoft com erro (ex.: o utilizador recusou as permissões)
  if (params.has("error")) {
    reply({ ok: false, error: params.get("error_description") || params.get("error") });
    return;
  }

  // Regresso da Microsoft com o código de autorização
  if (params.has("code")) {
    const saved = JSON.parse(sessionStorage.getItem(PKCE_KEY) || "null");
    sessionStorage.removeItem(PKCE_KEY);
    if (!saved || saved.state !== params.get("state")) {
      reply({ ok: false, error: "Resposta de login inválida (state). Tenta outra vez." });
      return;
    }
    show("A concluir o início de sessão…");
    const tokens = await tokenRequest({
      grant_type: "authorization_code",
      code: params.get("code"),
      redirect_uri: REDIRECT_URI,
      code_verifier: saved.verifier,
    });
    finish(tokens, saved.hint);
    return;
  }

  // Primeira chamada: tenta o refresh token guardado; se não houver, faz login.
  const hint = params.get("hint") || "";
  let rt = null;
  try {
    rt = localStorage.getItem(rtKey(hint));
  } catch {
    rt = null;
  }
  if (rt) {
    try {
      const tokens = await tokenRequest({ grant_type: "refresh_token", refresh_token: rt });
      finish(tokens, hint);
      return;
    } catch {
      try {
        localStorage.removeItem(rtKey(hint));
      } catch {
        /* ignorar */
      }
    }
  }
  show("A abrir o login da Microsoft…");
  await startLogin(hint);
}

Office.onReady(() => {
  run().catch((e) => {
    show("Erro: " + e.message);
    try {
      reply({ ok: false, error: e.message });
    } catch {
      /* fora do Office */
    }
  });
});
