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

function rtKey(cid, hint) {
  return "mm_rt_" + cid + "_" + (hint || "default").toLowerCase();
}

/** Client ID: o do config.js, se preenchido; senão o que o painel passa no endereço (?cid=). */
function pickClientId(fromQuery) {
  return String(CFG.clientId || "").trim() || String(fromQuery || "").trim();
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

async function tokenRequest(cid, params) {
  const res = await fetch(AUTHORITY + "/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: cid, scope: SCOPES, ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.error || "HTTP " + res.status);
  return data;
}

function finish(tokens, cid, hint) {
  try {
    if (tokens.refresh_token) localStorage.setItem(rtKey(cid, hint), tokens.refresh_token);
  } catch {
    /* sem localStorage: da próxima vez pede login outra vez */
  }
  reply({ ok: true, accessToken: tokens.access_token, expiresIn: Number(tokens.expires_in) || 3600 });
}

async function startLogin(cid, hint) {
  const verifier = randomString(48);
  const challenge = base64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  const state = randomString(16);
  sessionStorage.setItem(PKCE_KEY, JSON.stringify({ verifier, state, hint, cid }));
  const q = new URLSearchParams({
    client_id: cid,
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
    const tokens = await tokenRequest(saved.cid, {
      grant_type: "authorization_code",
      code: params.get("code"),
      redirect_uri: REDIRECT_URI,
      code_verifier: saved.verifier,
    });
    finish(tokens, saved.cid, saved.hint);
    return;
  }

  // Primeira chamada: tenta o refresh token guardado; se não houver, faz login.
  const hint = params.get("hint") || "";
  const cid = pickClientId(params.get("cid"));
  if (!cid) {
    reply({ ok: false, error: "Falta o ID da aplicação (Client ID)." });
    return;
  }
  let rt = null;
  try {
    rt = localStorage.getItem(rtKey(cid, hint));
  } catch {
    rt = null;
  }
  if (rt) {
    try {
      const tokens = await tokenRequest(cid, { grant_type: "refresh_token", refresh_token: rt });
      finish(tokens, cid, hint);
      return;
    } catch {
      try {
        localStorage.removeItem(rtKey(cid, hint));
      } catch {
        /* ignorar */
      }
    }
  }
  show("A abrir o login da Microsoft…");
  await startLogin(cid, hint);
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
