/**
 * Função principal que responde às requisições GET (quando o usuário acessa o link)
 */
function doGet(e) {
  // Pega o parâmetro 'page' da URL (ex: url-do-script/exec?page=admin)
  // Se não houver parâmetro (acesso direto ao link), define 'index' como padrão
  var page = e.parameter.page || 'index';
  
  var template;
  
  try {
    // Tenta buscar o arquivo dentro da pasta client/pages/
    template = HtmlService.createTemplateFromFile('client/pages/' + page);
  } catch (error) {
    // Se o usuário digitar uma página que não existe na URL, manda ele para a Home
    template = HtmlService.createTemplateFromFile('client/pages/index');
  }
  
  // Avalia o HTML e adiciona configurações essenciais de responsividade e título
  return template.evaluate()
    .setTitle('Becker & Ribeiro Advocacia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função auxiliar para incluir componentes e avaliar comandos do servidor
 */
function include(filename) {
  // Mudamos de createHtmlOutputFromFile para createTemplateFromFile().evaluate()
  return HtmlService.createTemplateFromFile('client/components/' + filename).evaluate().getContent();
}