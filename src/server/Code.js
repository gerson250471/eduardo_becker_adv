// Configuração da Versão Atual do Sistema
const VERSAO_SISTEMA = 'V 1.0.3';

/**
 * Função principal que responde às requisições GET (quando o usuário acessa o link)
 */
function doGet(e) {
  var page = e.parameter.page || 'index';
  var template;
  
  try {
    template = HtmlService.createTemplateFromFile('client/pages/' + page);
  } catch (error) {
    template = HtmlService.createTemplateFromFile('client/pages/index');
  }
  
  // Injeta a variável no template principal também
  template.versaoSistema = VERSAO_SISTEMA;
  
  return template.evaluate()
    .setTitle('Becker & Ribeiro Advocacia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função auxiliar para incluir componentes e avaliar comandos do servidor
 */
function include(filename) {
  var t = HtmlService.createTemplateFromFile('client/components/' + filename);
  t.versaoSistema = VERSAO_SISTEMA; // Compartilha a versão com os componentes
  return t.evaluate().getContent();
}