/*
 * Configuração do envio automático (Microsoft Graph). Não é obrigatório mexer aqui.
 * clientId: opcional. Se ficar vazio, cada pessoa cola o ID no próprio painel
 * (secção "Envio automático"), que fica guardado na caixa de correio dela.
 * Se o preencheres, vale para todos e o campo do painel deixa de ser preciso.
 */
window.MAILMERGE_CONFIG = {
  clientId: "594ff46a-cb55-45ad-ae02-959160fc6764",
  // Pausa entre envios, em milissegundos. O Exchange aceita cerca de 30 mensagens por minuto.
  delayMs: 2500,
};
