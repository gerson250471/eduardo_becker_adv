// =========================================================
// 1. CONFIGURAÇÕES GLOBAIS (DEVEM FICAR NO TOPO DO ARQUIVO)
// =========================================================
const VERSAO_SISTEMA = 'V 1.4.1';
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
 * Busca a nota, total de avaliações e comentários reais do Google
 */
function buscarAvaliacoesGoogle() {
  // Lê a chave das Propriedades do Script de forma segura (sem expor no GitHub)
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_PLACES_API_KEY');
  const PLACE_ID = 'ChIJXQDDy1PlG5URXeUGEmFodI8';

  // Endpoint de Place Details do Google Maps
  const url = 'https://maps.googleapis.com/maps/api/place/details/json?place_id=' + PLACE_ID + '&fields=rating,user_ratings_total,reviews&language=pt-BR&key=' + API_KEY;

  try {
    const resposta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(resposta.getContentText());

    if (json.status === "OK" && json.result) {
      return {
        sucesso: true,
        notaMedia: json.result.rating || 5.0,
        totalAvaliacoes: json.result.user_ratings_total || 89,
        reviews: (json.result.reviews || []).map(r => ({
          autor: r.author_name || 'Cliente Google',
          foto: r.profile_photo_url || '',
          nota: r.rating || 5,
          texto: r.text || '',
          tempo: r.relative_time_description || ''
        }))
      };
    } else {
      return obterAvaliacoesFallback();
    }
  } catch (erro) {
    return obterAvaliacoesFallback();
  }
}

/**
 * Contingência para garantir que a tela NUNCA fique travada em 'Carregando'
 */
function obterAvaliacoesFallback() {
  return {
    sucesso: true,
    notaMedia: 5.0,
    totalAvaliacoes: 89,
    reviews: [
      {
        autor: "Claudiani Pereira Soares",
        foto: "",
        nota: 5,
        texto: "Passando pra agradecer o escritório Becker e Ribeiro advocacia por todo o desempenho do meu caso além dos atendimentos ser á distancia online são grandes profissionais podem contratar sem medo doutor Eduardo nota 1.000 obrigado por tudo muito atencioso.",
        tempo: "Há alguns meses"
      },
      {
        autor: "rui lima",
        foto: "",
        nota: 5,
        texto: "Bons advogados né ajudaram a resolver dívida bancária com revisão de juros e redução.",
        tempo: "Há alguns meses"
      },
      {
        autor: "mich westhauser",
        foto: "",
        nota: 5,
        texto: "Dr. Eduardo super atencioso. Me deu todas as orientações para resolver da melhor maneira minha situação. Obrigado!",
        tempo: "Há alguns meses"
      },
      {
        autor: "Elisete Klein",
        foto: "",
        nota: 5,
        texto: "Achei o atendimento muito bom! Mesmo sendo tudo online, deu pra perceber o quanto são profissionais e preparados. Foram diretos, atenciosos e me explicaram tudo com calma.",
        tempo: "Há alguns meses"
      },
      {
        autor: "Fabiane Kauer",
        foto: "",
        nota: 5,
        texto: "Ótimo profissional, educado, atencioso esclareceu minhas dúvidas de forma clara e pontual. Super indico.",
        tempo: "Há alguns meses"
      }
    ]
  };
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

    // 1. Mapeamento de Séries do BACEN (Taxas Mensais fornecidas pelo Eduardo)
    let serieBacen = 25471; // Padrão: Financiamento de Veículos
    
    if (modalidade === 'veiculos') serieBacen = 25471;
    if (modalidade === 'pessoal') serieBacen = 25464;
    if (modalidade === 'consignado_inss') serieBacen = 25468;
    if (modalidade === 'consignado_clt') serieBacen = 25466;
    if (modalidade === 'garantia_veiculo') serieBacen = 29976;

    // 2. Data para consulta BACEN (DD/MM/AAAA)
    const partesData = dataContratacao.split('-'); // AAAA-MM-DD
    const dataInicio = `01/${partesData[1]}/${partesData[0]}`;
    const dataFim = `28/${partesData[1]}/${partesData[0]}`;

    // 3. Requisita API do Banco Central
    const urlBacen = `https://api.bcb.gov.br/dados/serie/bcdata.sgs.${serieBacen}/dados?formato=json&dataInicial=${dataInicio}&dataFinal=${dataFim}`;
    
    let taxaMediaBacen = 1.65; // Fallback
    try {
      const response = UrlFetchApp.fetch(urlBacen, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      if (json && json.length > 0) {
        // Como as novas séries já são mensais, pegamos o valor direto (sem conversão)
        taxaMediaBacen = parseFloat(json[json.length - 1].valor);
      }
    } catch (e) {
      Logger.log("Erro na API do BACEN: " + e.toString());
    }

    // 4. Armazena e calcula a taxa do contrato em variável (Taxa Implícita)
    let i = 0.01; // Chute inicial 1%
    for (let iter = 0; iter < 5000; iter++) { // Aumentado para 5000 iterações para evitar falhas
      let pmtCalculado = valorFinanciado * (i * Math.pow(1 + i, numParcelas)) / (Math.pow(1 + i, numParcelas) - 1);
      let diff = pmtCalculado - valorParcela;
      
      if (Math.abs(diff) < 0.01) break;
      i += diff > 0 ? -0.00005 : 0.00005;
      
      // Trava de segurança contra loops infinitos ou juros irreais
      if (i <= 0) { i = 0.001; break; } 
      if (i > 0.5) { break; } // Limite máximo de 50% ao mês
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
      taxaContrato: taxaContratoEncontrada,
      taxaBacen: taxaMediaBacen.toFixed(2),
      cenario: cenario
    };

  } catch (erro) {
    Logger.log("Erro no cálculo: " + erro.toString());
    return { sucesso: false, mensagem: "Não foi possível realizar a análise. Verifique os valores informados." };
  }
}