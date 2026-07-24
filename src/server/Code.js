// =========================================================
// 1. CONFIGURAÇÕES GLOBAIS (DEVEM FICAR NO TOPO DO ARQUIVO)
// =========================================================
const VERSAO_SISTEMA = 'V 1.3.0';
const ID_BANCO_DADOS = '1KrIJcYaC1G1KrRj7s-I6qut0YfqcoqQaIPaKZk7N2Ew';

function doGet(e) {
  var page = e.parameter.page || 'index';
  var template;
  
  try {
    template = HtmlService.createTemplateFromFile('client/pages/' + page);
  } catch (error) {
    page = 'index';
    template = HtmlService.createTemplateFromFile('client/pages/index');
  }
  
  // Repassa a versão do sistema e a página atual para o template
  template.versaoSistema = VERSAO_SISTEMA;
  template.pageAtual = page;
  
  return template.evaluate()
    .setTitle('Becker & Ribeiro Advocacia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função auxiliar para incluir componentes e avaliar comandos do servidor
 */
function include(filename, pageAtual) {
  var t = HtmlService.createTemplateFromFile('client/components/' + filename);
  t.versaoSistema = VERSAO_SISTEMA;
  t.pageAtual = pageAtual || 'index'; // Repassa a página atual para o header/footer
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

/**
 * Salva os dados do formulário de contato na aba 'Contatos' da Planilha
 */
function salvarContato(dadosFormulario) {
  try {
    const ss = SpreadsheetApp.openById(ID_BANCO_DADOS);
    let aba = ss.getSheetByName('Contatos');
    
    // Se a aba não existir, cria automaticamente com os cabeçalhos
    if (!aba) {
      aba = ss.insertSheet('Contatos');
      aba.appendRow(['Data/Hora', 'Nome', 'Telefone/WhatsApp', 'E-mail', 'Cidade', 'Mensagem']);
    }

    const dataAtual = new Date().toLocaleString('pt-BR');

    aba.appendRow([
      dataAtual,
      dadosFormulario.nome,
      dadosFormulario.telefone,
      dadosFormulario.email || 'Não informado',
      dadosFormulario.cidade || 'Não informada',
      dadosFormulario.mensagem
    ]);

    return { sucesso: true, mensagem: 'Mensagem enviada com sucesso!' };

  } catch (erro) {
    Logger.log('Erro ao salvar contato: ' + erro.toString());
    return { sucesso: false, mensagem: 'Erro ao enviar mensagem. Tente novamente.' };
  }
}

/**
 * Busca a nota, total de avaliações e comentários reais do Google Meu Negócio via API do Google Places
 */
function buscarAvaliacoesGoogle() {
  const API_KEY = 'AIzaSyApbYa5LITFDA8mgLPzsC41ur4vN02g6H4'; 
  const PLACE_ID = 'ChIJXQDDy1PlG5URXeUGEmFodI8';

  const url = 'https://maps.googleapis.com/maps/api/place/details/json?place_id=' + 
              PLACE_ID + 
              '&fields=rating,user_ratings_total,reviews&language=pt-BR&key=' + 
              API_KEY;

  try {
    const resposta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(resposta.getContentText());

    if (json.status === 'OK' && json.result) {
      return {
        sucesso: true,
        notaMedia: json.result.rating || 5.0,
        totalAvaliacoes: json.result.user_ratings_total || 0,
        reviews: (json.result.reviews || []).map(r => ({
          autor: r.author_name,
          foto: r.profile_photo_url,
          nota: r.rating,
          texto: r.text,
          tempo: r.relative_time_description
        }))
      };
    } else {
      Logger.log('Erro na API do Google Places: ' + json.status);
      return { sucesso: false };
    }
  } catch (erro) {
    Logger.log('Exceção ao buscar avaliações do Google: ' + erro.toString());
    return { sucesso: false };
  }
} 