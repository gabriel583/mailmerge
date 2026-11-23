/*
 * Lógica pura do mail merge (sem dependências do Office). Funciona no browser
 * (window.Merge) e em Node (module.exports), para poder ser testada.
 *
 * Marcadores: {{NomeDaColuna}}. A comparação com os cabeçalhos do Excel ignora
 * maiúsculas, acentos e espaços a mais: {{ nome }} == {{Nome}} == {{NOME}}.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.Merge = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // {{ ... }} no HTML, tolerando tags que o Outlook mete no meio (ex.: {{<span>Nome</span>}})
  const HTML_PLACEHOLDER = /\{\{((?:[^{}<]|<[^>]*>)*?)\}\}/g;
  // O mesmo, mas codificado em URL (marcadores dentro de href="...")
  const URL_PLACEHOLDER = /%7B%7B(.*?)%7D%7D/gi;
  // Texto simples (assunto)
  const TEXT_PLACEHOLDER = /\{\{([^{}]*?)\}\}/g;
  const TAG = /<[^>]*>/g;
  const EMAIL = /^[^\s@<>"',;]+@[^\s@<>"',;]+\.[^\s@<>"',;]+$/;

  function normalizeKey(s) {
    return String(s)
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  function decodeEntities(s) {
    return s
      .replace(/&nbsp;/gi, " ")
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
      .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"')
      .replace(/&#39;|&apos;/gi, "'")
      .replace(/&amp;/gi, "&");
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function keyFromHtml(inner) {
    return decodeEntities(inner.replace(TAG, "")).replace(/ /g, " ").trim();
  }

  /** Mapa chave-normalizada -> valor, a partir de uma linha do Excel. */
  function rowLookup(row) {
    const map = new Map();
    for (const [k, v] of Object.entries(row || {})) {
      if (k.startsWith("__")) continue; // colunas internas (ex.: __rowNum__)
      map.set(normalizeKey(k), v == null ? "" : String(v));
    }
    return map;
  }

  /** Lista (sem repetidos) dos marcadores encontrados no assunto e no corpo. */
  function findPlaceholders(html, subject) {
    const seen = new Map();
    const add = (k) => {
      const key = k.trim();
      if (key && !seen.has(normalizeKey(key))) seen.set(normalizeKey(key), key);
    };
    for (const m of String(html || "").matchAll(HTML_PLACEHOLDER)) add(keyFromHtml(m[1]));
    for (const m of String(html || "").matchAll(URL_PLACEHOLDER)) add(safeDecodeURI(m[1]));
    for (const m of String(subject || "").matchAll(TEXT_PLACEHOLDER)) add(m[1]);
    return [...seen.values()];
  }

  function safeDecodeURI(s) {
    try {
      return decodeURIComponent(s);
    } catch {
      return s;
    }
  }

  /**
   * Substitui os marcadores no HTML. Valores são escapados; quebras de linha
   * de uma célula viram <br>. Tags que estavam dentro do marcador são mantidas
   * (depois do valor) para não desequilibrar o HTML.
   * Devolve { html, missing: [chaves sem coluna correspondente] }.
   */
  function mergeHtml(html, row) {
    const lookup = rowLookup(row);
    const missing = new Set();
    let out = String(html || "").replace(HTML_PLACEHOLDER, (match, inner) => {
      const key = keyFromHtml(inner);
      const nk = normalizeKey(key);
      if (!lookup.has(nk)) {
        missing.add(key);
        return match;
      }
      const tags = inner.match(TAG) || [];
      return escapeHtml(lookup.get(nk)).replace(/\r?\n/g, "<br>") + tags.join("");
    });
    out = out.replace(URL_PLACEHOLDER, (match, inner) => {
      const key = safeDecodeURI(inner).trim();
      const nk = normalizeKey(key);
      if (!lookup.has(nk)) {
        missing.add(key);
        return match;
      }
      return escapeHtml(lookup.get(nk).trim());
    });
    return { html: out, missing: [...missing] };
  }

  /** Substitui os marcadores num texto simples (ex.: assunto). */
  function mergeText(text, row) {
    const lookup = rowLookup(row);
    const missing = new Set();
    const out = String(text || "").replace(TEXT_PLACEHOLDER, (match, inner) => {
      const nk = normalizeKey(inner);
      if (!lookup.has(nk)) {
        missing.add(inner.trim());
        return match;
      }
      return lookup.get(nk).replace(/\s*\r?\n\s*/g, " ");
    });
    return { text: out, missing: [...missing] };
  }

  /**
   * Extrai emails de uma célula. Aceita vários separados por ; , ou mudança
   * de linha, e o formato "Nome <email@x.pt>".
   */
  function parseEmails(cell) {
    const valid = [];
    const invalid = [];
    for (let part of String(cell || "").split(/[;,\n\r]+/)) {
      part = part.trim();
      if (!part) continue;
      const angle = part.match(/<([^>]+)>/);
      const addr = (angle ? angle[1] : part).trim().replace(/^mailto:/i, "");
      if (EMAIL.test(addr)) valid.push(addr);
      else invalid.push(part);
    }
    return { valid, invalid };
  }

  /** Junta os emails de uma coluna em várias linhas, sem repetidos. */
  function collectEmails(rows, column) {
    const seen = new Set();
    const valid = [];
    const invalid = [];
    if (!column) return { valid, invalid };
    for (const row of rows) {
      const r = parseEmails(row[column]);
      for (const e of r.valid) {
        if (!seen.has(e.toLowerCase())) {
          seen.add(e.toLowerCase());
          valid.push(e);
        }
      }
      invalid.push(...r.invalid);
    }
    return { valid, invalid };
  }

  /** Adivinha a coluna de email: pelo nome do cabeçalho ou pelo conteúdo. */
  function guessEmailColumn(headers, rows) {
    const byName = headers.find((h) => /e-?mail|correio|endere[cç]o/i.test(h) && !/^cc|bcc/i.test(h));
    if (byName) return byName;
    let best = null;
    let bestScore = 0;
    const sample = rows.slice(0, 50);
    for (const h of headers) {
      const score = sample.filter((r) => parseEmails(r[h]).valid.length > 0).length;
      if (score > bestScore) {
        best = h;
        bestScore = score;
      }
    }
    return bestScore >= Math.max(1, sample.length / 2) ? best : null;
  }

  function guessCcColumn(headers) {
    return headers.find((h) => /^\s*cc\b/i.test(h)) || null;
  }

  /**
   * Converte o texto de um número formatado à americana (1,250.50) para o
   * formato português (1 250,50). Só mexe em separadores entre dígitos.
   */
  function localizeNumberText(text) {
    return String(text).replace(/(\d)([.,])(?=\d)/g, (_, d, sep) => d + (sep === "," ? " " : ","));
  }

  /** Tamanho em bytes (UTF-8). */
  function byteLength(s) {
    return new TextEncoder().encode(s).length;
  }

  /**
   * Reduz o HTML do Outlook para caber em limites de tamanho (ex.: 32 KB do
   * displayNewMessageFormAsync). Primeiro tira comentários/XML do Word; se
   * ainda não couber, fica só com o conteúdo do <body>.
   */
  function compactHtml(html, maxBytes) {
    if (byteLength(html) <= maxBytes) return html;
    let h = html
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<xml[\s\S]*?<\/xml>/gi, "")
      .replace(/<meta[^>]*>/gi, "")
      .replace(/<o:p>\s*<\/o:p>/gi, "")
      .replace(/\n\s*\n/g, "\n");
    if (byteLength(h) <= maxBytes) return h;
    const body = h.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    if (body) h = body[1];
    return h;
  }

  return {
    normalizeKey,
    escapeHtml,
    findPlaceholders,
    mergeHtml,
    mergeText,
    parseEmails,
    collectEmails,
    guessEmailColumn,
    guessCcColumn,
    compactHtml,
    byteLength,
    localizeNumberText,
  };
});
