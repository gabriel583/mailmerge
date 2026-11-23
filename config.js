/*
 * Configuração do envio automático (Microsoft Graph).
 * clientId: o "ID da aplicação (cliente)" do registo no Microsoft Entra (ver README,
 * secção "Envio automático"). Enquanto estiver vazio, o botão "Enviar" explica o que falta.
 */
window.MAILMERGE_CONFIG = {
  clientId: "",
  // Pausa entre envios, em milissegundos. O Exchange aceita cerca de 30 mensagens por minuto.
  delayMs: 2500,
};
