// Configuração da Versão Atual do Sistema
const VERSAO_SISTEMA = 'V 1.0.4';

// ID do Banco de Dados (Google Sheets)
const ID_BANCO_DADOS = '1KrIJcYaC1G1KrRj7s-I6qut0YfqcoqQaIPaKZk7N2Ew';

/**
 * Função principal que responde às requisições GET
 */
function doGet(e) {
  var page = e.parameter.page || 'index';
  var template;
  
  try {
    template = HtmlService.createTemplateFromFile('client/pages/' + page);
  } catch (error) {
    template = HtmlService.createTemplateFromFile('client/pages/index');
  }
  
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
  t.versaoSistema = VERSAO_SISTEMA;
  return t.evaluate().getContent();
}

/**
 * Busca a lista de artigos cadastrados na aba 'Blog' da Planilha
 */
function buscarArtigosBlog() {
  try {
    const ss = SpreadsheetApp.openById(ID_BANCO_DADOS);
    const aba = ss.getSheetByName('Blog');
    
    if (!aba) return [];

    const dados = aba.getDataRange().getValues();
    const cabecalho = dados.shift(); // Remove a primeira linha (cabeçalho)

    // Mapeia as linhas para objetos JSON
    return dados.map(linha => {
      return {
        id: linha[0],
        titulo: linha[1],
        resumo: linha[2],
        conteudo: linha[3],
        data: linha[4] ? new Date(linha[4]).toLocaleDateString('pt-BR') : '',
        status: linha[5]
      };
    }).filter(artigo => artigo.status === 'Publicado'); // Retorna apenas artigos ativos

  } catch (erro) {
    Logger.log('Erro ao buscar artigos: ' + erro.toString());
    return [];
  }
}