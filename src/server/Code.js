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
 * Busca a nota, total de avaliações e comentários reais do Google via Places API (New)
 */
function buscarAvaliacoesGoogle() {
  const API_KEY = 'AIzaSyBeFgWgQMhLIz6laK0vcNhkeWPDzoyIf84'; 
  const PLACE_ID = 'ChIJXQDDy1PlG5URXeUGEmFodI8';

  // Endpoint moderno da Places API (New)
  const url = 'https://places.googleapis.com/v1/places/' + PLACE_ID + '?languageCode=pt-BR';

  const options = {
    method: 'get',
    headers: {
      'X-Goog-Api-Key': API_KEY,
      'X-Goog-FieldMask': 'rating,userRatingCount,reviews'
    },
    muteHttpExceptions: true
  };

  try {
    const resposta = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(resposta.getContentText());

    if (json.rating || json.reviews) {
      return {
        sucesso: true,
        notaMedia: json.rating || 5.0,
        totalAvaliacoes: json.userRatingCount || 0,
        reviews: (json.reviews || []).map(r => ({
          autor: r.authorAttribution ? r.authorAttribution.displayName : 'Cliente Google',
          foto: r.authorAttribution ? r.authorAttribution.photoUri : '',
          nota: r.rating || 5,
          texto: r.originalText ? r.originalText.text : (r.text ? r.text.text : ''),
          tempo: r.relativePublishTimeDescription || ''
        }))
      };
    } else {
      Logger.log('Retorno do Google sem dados: ' + resposta.getContentText());
      return { sucesso: false };
    }
  } catch (erro) {
    Logger.log('Erro ao buscar avaliações do Google: ' + erro.toString());
    return { sucesso: false };
  }
}

/**
 * Valida o login de sócios/colaboradores na aba 'Usuarios'
 */
function validarLogin(usuario, senha) {
  try {
    const ss = SpreadsheetApp.openById(ID_BANCO_DADOS);
    const aba = ss.getSheetByName('Usuarios');
    
    if (!aba) {
      return { sucesso: false, mensagem: 'Aba de usuários não configurada na planilha.' };
    }

    const dados = aba.getDataRange().getValues();
    dados.shift(); // Remove cabeçalho

    const usuarioValido = dados.find(linha => {
      const userPlanilha = String(linha[0]).trim();
      const senhaPlanilha = String(linha[1]).trim();
      return (userPlanilha === usuario.trim()) && (senhaPlanilha === senha.trim());
    });

    if (usuarioValido) {
      return { sucesso: true };
    } else {
      return { sucesso: false, mensagem: 'Usuário ou senha incorretos.' };
    }
  } catch (erro) {
    Logger.log('Erro ao validar login: ' + erro.toString());
    return { sucesso: false, mensagem: 'Erro interno ao validar acesso.' };
  }
}

/**
 * Consulta a taxa média de juros mensal no BACEN para a modalidade e data informadas
 */
function consultarTaxaBacen(modalidadeCodigo, dataContratacao) {
  try {
    // Formata a data (DD/MM/AAAA) para o padrão aceito pela API do BACEN
    const partesData = dataContratacao.split('-'); // Espera AAAA-MM-DD
    const dataFormatada = partesData[2] + '/' + partesData[1] + '/' + partesData[0];

    // Endpoint do Banco Central (SGS)
    const url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + modalidadeCodigo + 
                '/dados?formato=json&dataInicial=' + dataFormatada + '&dataFinal=' + dataFormatada;

    const resposta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const dados = JSON.parse(resposta.getContentText());

    if (dados && dados.length > 0) {
      return {
        sucesso: true,
        taxaMediaBacen: parseFloat(dados[0].valor) // Ex: 1.65 (% a.m.)
      };
    } else {
      // Fallback: se a data for final de semana/feriado ou muito recente, busca o último valor disponível
      const urlUltimo = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + modalidadeCodigo + '/dados/ultimos/1?formato=json';
      const respUltimo = UrlFetchApp.fetch(urlUltimo, { muteHttpExceptions: true });
      const dadosUltimo = JSON.parse(respUltimo.getContentText());
      
      return {
        sucesso: true,
        taxaMediaBacen: parseFloat(dadosUltimo[0].valor)
      };
    }
  } catch (erro) {
    Logger.log('Erro ao consultar API BACEN: ' + erro.toString());
    return { sucesso: false, mensagem: 'Não foi possível obter a taxa do BACEN.' };
  }
}