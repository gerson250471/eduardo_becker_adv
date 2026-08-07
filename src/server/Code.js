// =========================================================
// 1. CONFIGURAÇÕES GLOBAIS
// =========================================================
const VERSAO_SISTEMA = 'V 1.2.0';
const ID_PASTA_IMAGENS = '1KrO0f-nmmjCJNgb-_kfxUk-Be-VPg4os'; // Pasta 'src' do Drive

/**
 * Retorna o ID da planilha correto com base na variável de ambiente do projeto
 */
function getBancoDadosId() {
  const ambiente = PropertiesService.getScriptProperties().getProperty('AMBIENTE');
  
  if (ambiente === 'PRODUCAO') {
    return '1KrIJcYaC1G1KrRj7s-I6qut0YfqcoqQaIPaKZk7N2Ew'; // ID Planilha Oficial (Produção)
  }
  return '1dNivpBHWHEGsMTLOAI1hdJRcSphslZD4xzcorgQXsY0'; // ID Planilha Homologação (Fallback seguro)
}

// =========================================================
// 2. CONFIGURAÇÃO DE AMBIENTE (RODAR APENAS 1 VEZ EM CADA PROJETO)
// =========================================================
function setAmbienteProducao() {
  PropertiesService.getScriptProperties().setProperty('AMBIENTE', 'PRODUCAO');
  Logger.log('✅ Este projeto foi configurado como PRODUÇÃO.');
}

function setAmbienteHomologacao() {
  PropertiesService.getScriptProperties().setProperty('AMBIENTE', 'HOMOLOGACAO');
  Logger.log('✅ Este projeto foi configurado como HOMOLOGAÇÃO.');
}

function doGet(e) {
  var page = e.parameter.page || 'index';
  var idArtigo = e.parameter.id || ''; 
  
  // Identifica o ambiente para gerar os links corretos e evitar fugas de ambiente
  var ambiente = PropertiesService.getScriptProperties().getProperty('AMBIENTE');
  var hostPadrao = (ambiente === 'PRODUCAO') 
    ? 'https://beckereribeiro.com/' 
    : 'https://beckereribeiroadv.com/homologacao/';
    
  var hostUrl = e.parameter.host || hostPadrao; 
  
  var template;
  try {
    template = HtmlService.createTemplateFromFile('client/pages/' + page);
  } catch (error) {
    page = 'index';
    template = HtmlService.createTemplateFromFile('client/pages/index');
  }
  
  // Repassa as variáveis para as telas (Front-end)
  template.versaoSistema = VERSAO_SISTEMA;
  template.pageAtual = page;
  template.hostUrl = hostUrl; 
  template.idArtigo = idArtigo; // Injeta o ID da URL
  
  return template.evaluate()
    .setTitle('Becker & Ribeiro Advocacia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função auxiliar para incluir componentes e avaliar comandos do servidor
 */
function include(filename, pageAtual, hostUrl) {
  var t = HtmlService.createTemplateFromFile('client/components/' + filename);
  t.versaoSistema = VERSAO_SISTEMA;
  t.pageAtual = pageAtual || 'index'; 
  t.hostUrl = hostUrl || 'https://beckereribeiro.com/'; // Repassa o host para o header/footer
  return t.evaluate().getContent();
}

/**
 * Salva os dados do formulário de contato na aba 'Contatos' da Planilha
 */
function salvarContato(dadosFormulario) {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
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

// =========================================================
// MÓDULO DE SEGURANÇA E AUTENTICAÇÃO DE USUÁRIOS
// =========================================================

/**
 * Função auxiliar que junta o Usuário + Senha e cria o Hash irreversível (SHA-256)
 */
function gerarHashBase(usuario, senha) {
  const textoParaCriptografar = String(usuario).trim().toLowerCase() + String(senha).trim();
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, textoParaCriptografar)
    .map(function(chr){return (chr+256).toString(16).slice(-2)})
    .join('');
}

/**
 * 1. SCRIPT DE USO ÚNICO: Gerar os primeiros hashes na planilha
 * Selecione esta função no painel superior e clique em "Executar"
 */
function gerarHashesIniciais() {
  const ss = SpreadsheetApp.openById(getBancoDadosId());
  const aba = ss.getSheetByName('Usuarios');
  if (!aba) {
    Logger.log("Aba 'Usuarios' não encontrada!");
    return;
  }
  
  const dados = aba.getDataRange().getValues();
  const senhaPadrao = "123456";
  
  // O loop começa em 1 para pular o cabeçalho
  for (let i = 1; i < dados.length; i++) {
    const usuario = String(dados[i][0]).trim();
    const senhaHashAtual = String(dados[i][1]).trim();
    
    // Se existe usuário escrito, mas o hash está vazio, ele calcula e grava
    if (usuario !== "" && senhaHashAtual === "") {
      const hash = gerarHashBase(usuario, senhaPadrao);
      aba.getRange(i + 1, 2).setValue(hash); // Grava na Coluna B
      Logger.log(`Hash gerado para: ${usuario}`);
    }
  }
  Logger.log("✅ Hashes iniciais gerados com sucesso!");
}

/**
 * 2. Valida o login conferindo o Hash e a Situação da conta
 */
function validarLogin(usuario, senha) {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const aba = ss.getSheetByName('Usuarios');
    if (!aba) return { sucesso: false, mensagem: 'Aba de usuários não configurada.' };

    const dados = aba.getDataRange().getValues();
    dados.shift(); // Remove cabeçalho

    const hashTentativa = gerarHashBase(usuario, senha);

    const usuarioEncontrado = dados.find(linha => {
      const userPlanilha = String(linha[0]).trim().toLowerCase();
      const hashPlanilha = String(linha[1]).trim();
      const situacao = String(linha[3]).trim().toLowerCase(); // Coluna D
      
      return (userPlanilha === String(usuario).trim().toLowerCase()) && 
             (hashPlanilha === hashTentativa) && 
             (situacao === 'ativo');
    });

    if (usuarioEncontrado) {
      return { 
        sucesso: true, 
        nomeUsuario: usuarioEncontrado[0],
        exigeTroca: usuarioEncontrado[5] // Coluna F (Trocar_Senha)
      };
    } else {
      return { sucesso: false, mensagem: 'Usuário/senha incorretos ou conta inativa.' };
    }
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro interno ao validar acesso.' };
  }
}

/**
 * 3. Salva a nova senha definitiva quando o usuário for forçado a trocar
 */
function atualizarSenhaObrigatoria(usuario, novaSenha) {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const aba = ss.getSheetByName('Usuarios');
    const dados = aba.getDataRange().getValues();
    
    const userBusca = String(usuario).trim().toLowerCase();
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim().toLowerCase() === userBusca) {
        const novoHash = gerarHashBase(usuario, novaSenha);
        aba.getRange(i + 1, 2).setValue(novoHash); // Atualiza Coluna B (Hash)
        aba.getRange(i + 1, 6).setValue('Nao');    // Atualiza Coluna F (Trocar_Senha)
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Usuário não encontrado no banco de dados.' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao gravar a nova senha.' };
  }
}

/**
 * 4. Recuperação de Senha Automatizada via Gmail
 */
function recuperarSenhaEmail(email) {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const aba = ss.getSheetByName('Usuarios');
    const dados = aba.getDataRange().getValues();
    
    const emailBusca = String(email).trim().toLowerCase();
    
    for (let i = 1; i < dados.length; i++) {
      const emailPlanilha = String(dados[i][2]).trim().toLowerCase(); // Coluna C
      
      if (emailPlanilha === emailBusca) {
        const usuario = String(dados[i][0]).trim();
        const senhaProvisoria = Math.random().toString(36).slice(-8); // Gera 8 caracteres aleatórios
        const novoHash = gerarHashBase(usuario, senhaProvisoria);
        
        aba.getRange(i + 1, 2).setValue(novoHash); // Substitui a senha antiga pelo novo Hash
        aba.getRange(i + 1, 6).setValue('Sim');    // Força a troca no próximo login
        
        MailApp.sendEmail({
          to: emailPlanilha,
          subject: "Recuperação de Senha - Área Exclusiva",
          htmlBody: `Olá,<br><br>Sua senha provisória foi gerada com sucesso.<br><br>
                     <b>Usuário:</b> ${usuario}<br>
                     <b>Senha Provisória:</b> ${senhaProvisoria}<br><br>
                     Ao fazer o login, o sistema exigirá a criação de uma nova senha definitiva por motivos de segurança.`
        });
        
        return { sucesso: true };
      }
    }
    return { sucesso: false, mensagem: 'Este e-mail não foi encontrado na base de usuários.' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao tentar enviar o e-mail de recuperação.' };
  }
}

// =========================================================
// MÓDULO DO BLOG (Gestão de Artigos)
// =========================================================

/**
 * Recebe os dados do front-end, salva a imagem no Drive e os dados na Planilha
 */
function salvarArtigoBlog(pacoteArtigo, dadosImagem) {
  try {
    // 1. Tratamento da Imagem de Capa (Salva no Drive)
    const pastaDestino = DriveApp.getFolderById(ID_PASTA_IMAGENS);
    
    // Converte a string Base64 de volta para um arquivo binário (Blob)
    const blob = Utilities.newBlob(Utilities.base64Decode(dadosImagem.conteudoBase64), dadosImagem.tipo, dadosImagem.nome);
    const arquivoCapa = pastaDestino.createFile(blob);
    
    // Força a permissão de visualização pública para garantir que a imagem apareça no site
    arquivoCapa.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Captura a URL pública da imagem recém-criada
    const urlCapa = arquivoCapa.getUrl(); 

    // 2. Preparação dos Dados para a Planilha
    const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    const idUnico = 'ART-' + new Date().getTime(); // Gera um ID único baseado no timestamp atual
    
    // 3. Conexão com o Banco de Dados
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const abaBlog = ss.getSheetByName('Blog');
    
    if (!abaBlog) {
      return { sucesso: false, mensagem: "Aba 'Blog' não encontrada no banco de dados." };
    }
    
    // 4. Salva a nova linha com os dados exatos nas 8 colunas que você criou
    abaBlog.appendRow([
      idUnico,                // Coluna A (ID)
      dataAtual,              // Coluna B (Data_Publicacao)
      pacoteArtigo.autor,     // Coluna C (Autor)
      pacoteArtigo.titulo,    // Coluna D (Titulo)
      pacoteArtigo.resumo,    // Coluna E (Resumo)
      urlCapa,                // Coluna F (Capa_URL)
      pacoteArtigo.conteudoHTML, // Coluna G (Conteudo_HTML - Texto do Quill)
      pacoteArtigo.status     // Coluna H (Status)
    ]);
    
    return { sucesso: true, mensagem: "Artigo publicado com sucesso!" };
    
  } catch (erro) {
    Logger.log("Erro no Blog: " + erro.message);
    return { sucesso: false, mensagem: "Erro ao salvar o artigo: " + erro.message };
  }
}

/**
 * =========================================================
 * BUSCA DE ARTIGOS PARA A VITRINE DO BLOG
 * =========================================================
 */
/**
 * =========================================================
 * BUSCA DE ARTIGOS PARA A VITRINE DO BLOG
 * =========================================================
 */
function buscarArtigosBlog() {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const aba = ss.getSheetByName('Blog');
    if (!aba) return [];

    const dados = aba.getDataRange().getValues();
    dados.shift(); 

    const artigos = [];
    
    for (let i = 0; i < dados.length; i++) {
      const status = String(dados[i][7]).trim(); 
      
      if (status === 'Publicado') {
        
        let urlCapa = String(dados[i][5]).trim();
        const match = urlCapa.match(/\/d\/(.+?)\//);
        if (match && match[1]) {
          // Usa a nova rota oficial do Google para evitar bloqueios de CORS
          urlCapa = 'https://lh3.googleusercontent.com/d/' + match[1];
        }

        // PROTEÇÃO: Converte Datas para texto simples para não quebrar o script
        let dataFormatada = dados[i][1];
        if (dataFormatada instanceof Date) {
          dataFormatada = Utilities.formatDate(dataFormatada, "GMT-3", "dd/MM/yyyy HH:mm:ss");
        } else {
          dataFormatada = String(dataFormatada);
        }

        artigos.push({
          id: String(dados[i][0]),
          data: dataFormatada,
          autor: String(dados[i][2]),
          titulo: String(dados[i][3]),
          resumo: String(dados[i][4]),
          capa: urlCapa
        });
      }
    }
    
    return artigos.reverse();
    
  } catch (erro) {
    Logger.log("Erro ao buscar artigos: " + erro.message);
    return [];
  }
}

/**
 * =========================================================
 * BUSCA UM ARTIGO COMPLETO PELO ID (TELA DE LEITURA)
 * =========================================================
 */
function buscarArtigoPorId(idArtigo) {
  try {
    const ss = SpreadsheetApp.openById(getBancoDadosId());
    const aba = ss.getSheetByName('Blog');
    if (!aba) return null;

    const dados = aba.getDataRange().getValues();
    dados.shift(); 

    for (let i = 0; i < dados.length; i++) {
      const id = String(dados[i][0]).trim(); 
      const status = String(dados[i][7]).trim(); 

      if (id === idArtigo && status === 'Publicado') {
        
        let urlCapa = String(dados[i][5]).trim();
        const match = urlCapa.match(/\/d\/(.+?)\//);
        if (match && match[1]) {
          // Usa a nova rota oficial do Google para evitar bloqueios de CORS
          urlCapa = 'https://lh3.googleusercontent.com/d/' + match[1];
        }
        
        // PROTEÇÃO: Converte Datas para texto simples para não quebrar o script
        let dataFormatada = dados[i][1];
        if (dataFormatada instanceof Date) {
          dataFormatada = Utilities.formatDate(dataFormatada, "GMT-3", "dd/MM/yyyy HH:mm:ss");
        } else {
          dataFormatada = String(dataFormatada);
        }

        return {
          id: id,
          data: dataFormatada,
          autor: String(dados[i][2]),
          titulo: String(dados[i][3]),
          resumo: String(dados[i][4]),
          capa: urlCapa,
          conteudo: String(dados[i][6]) 
        };
      }
    }
    return null; 
  } catch (erro) {
    Logger.log("Erro ao buscar artigo por ID: " + erro.message);
    return null;
  }
}

/**
 * =========================================================
 * FUNÇÃO DE AUTORIZAÇÃO (RODAR ANTES DO DEPLOY)
 * =========================================================
 * Esta função serve apenas para forçar o Google a solicitar 
 * todas as permissões necessárias para o sistema funcionar.
 */
function autorizarPermissoes() {
  try {
    // 1. Força a permissão de leitura/escrita no Google Sheets
    SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. Força a permissão de acesso à internet (APIs BACEN e Google Maps)
    UrlFetchApp.fetch('https://www.google.com.br', { muteHttpExceptions: true });
    
    // 3. Força a leitura de propriedades do script (API Key)
    PropertiesService.getScriptProperties().getProperty('TESTE');
    
    Logger.log('✅ Permissões concedidas com sucesso! O sistema está pronto para o Deploy.');
  } catch (e) {
    Logger.log('⚠️ As permissões foram solicitadas. Verifique se aceitou todas.');
  }
}