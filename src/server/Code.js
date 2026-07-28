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
 * Consulta a taxa média do BACEN e classifica o contrato do cliente
 */
function calcularJurosAbusivos(dados) {
  try {
    const { modalidade, dataContratacao, valorFinanciado, numParcelas, valorParcela } = dados;

    // 1. Mapeamento de Séries do BACEN conforme novas modalidades
    let serieBacen = 20742; // Padrão: Veículos PF
    if (modalidade === 'veiculos') serieBacen = 20742;
    if (modalidade === 'pessoal') serieBacen = 20739;
    if (modalidade === 'consignado_inss') serieBacen = 20740; // Consignado INSS
    if (modalidade === 'consignado_clt') serieBacen = 25471;  // Consignado Privado/CLT
    if (modalidade === 'imobiliario') serieBacen = 20749;

    // 2. Data para consulta BACEN (DD/MM/AAAA)
    const partesData = dataContratacao.split('-'); // AAAA-MM-DD
    const dataInicio = `01/${partesData[1]}/${partesData[0]}`;
    const dataFim = `28/${partesData[1]}/${partesData[0]}`;

    // 3. Requisita API do Banco Central
    const urlBacen = `https://api.bcb.gov.br/dados/serie/bcdata.sgs.${serieBacen}/dados?formato=json&dataInicial=${dataInicio}&dataFinal=${dataFim}`;
    
    let taxaMediaBacen = 1.65; // Fallback caso não retorne
    try {
      const response = UrlFetchApp.fetch(urlBacen, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      if (json && json.length > 0) {
        taxaMediaBacen = parseFloat(json[json.length - 1].valor);
      }
    } catch (e) {
      Logger.log("Erro na API do BACEN: " + e.toString());
    }

    // 4. Armazena e calcula a taxa do contrato em variável (Taxa Implícita)
    let i = 0.01;
    for (let iter = 0; iter < 100; iter++) {
      let pmtCalculado = valorFinanciado * (i * Math.pow(1 + i, numParcelas)) / (Math.pow(1 + i, numParcelas) - 1);
      let diff = pmtCalculado - valorParcela;
      if (Math.abs(diff) < 0.01) break;
      i += diff > 0 ? -0.0001 : 0.0001;
    }
    const taxaContratoEncontrada = parseFloat((i * 100).toFixed(2));
    const limiteAbusivo50 = parseFloat((taxaMediaBacen * 1.5).toFixed(2));

    // 5. Classificação nos 3 cenários (A, B e C)
    let cenario = 'A';
    if (taxaContratoEncontrada > taxaMediaBacen && taxaContratoEncontrada <= limiteAbusivo50) {
      cenario = 'B';
    } else if (taxaContratoEncontrada > limiteAbusivo50) {
      cenario = 'C';
    }

    return {
      sucesso: true,
      taxaContrato: taxaContratoEncontrada, // Variável armazenada
      taxaBacen: taxaMediaBacen.toFixed(2),
      cenario: cenario
    };

  } catch (erro) {
    Logger.log("Erro no cálculo: " + erro.toString());
    return { sucesso: false, mensagem: "Não foi possível realizar a análise. Verifique os valores informados." };
  }
}