function criarBancoDeDados() {
    // ======================================================
  // PROTEÇÃO CONTRA CRIAÇÃO ACIDENTAL DE UM NOVO BANCO
  // ======================================================

  const propriedades = PropertiesService.getScriptProperties();

  const idBancoExistente = propriedades.getProperty('ID_PLANILHA');

  if (idBancoExistente) {

    try {

      const bancoExistente = SpreadsheetApp.openById(idBancoExistente);

      throw new Error(
        'OPERAÇÃO CANCELADA: Já existe um banco de dados configurado: "' +
        bancoExistente.getName() +
        '". Nenhuma nova planilha foi criada e seus dados permanecem seguros.'
      );

    } catch (erro) {

      // Se for nosso erro de proteção, repassa o erro
      if (
        erro.message &&
        erro.message.includes('OPERAÇÃO CANCELADA')
      ) {
        throw erro;
      }

      // Se o ID existir, mas a planilha não puder ser aberta,
      // também bloqueamos por segurança.
      throw new Error(
        'OPERAÇÃO CANCELADA: Existe um ID de banco configurado, ' +
        'mas a planilha não pôde ser acessada. ' +
        'Nenhum novo banco foi criado para evitar perda ou troca acidental de dados.'
      );
    }
  }
  // Cria a planilha que será usada como banco de dados
  const planilha = SpreadsheetApp.create('Sistema Lua de Sorvete - Banco de Dados');
  const idPlanilha = planilha.getId();

  // Salva o ID para que todo o sistema possa encontrar o banco depois
  PropertiesService.getScriptProperties()
    .setProperty('ID_PLANILHA', idPlanilha);

  // Configuração das abas e colunas
  const estrutura = {
    
    'Vendas': [
      'ID_VENDA',
      'DATA_HORA',
      'CLIENTE',
      'TIPO_VENDA',
      'SUBTOTAL',
      'DESCONTO',
      'TOTAL',
      'FORMA_PAGAMENTO',
      'STATUS',
      'OBSERVACOES'
    ],

    'Itens_Venda': [
      'ID_ITEM',
      'ID_VENDA',
      'ID_PRODUTO',
      'PRODUTO',
      'TIPO',
      'QUANTIDADE',
      'PESO_KG',
      'PRECO_UNITARIO',
      'PRECO_KG',
      'TOTAL_ITEM'
    ],

    'Comandas': [
      'ID_COMANDA',
      'NUMERO_COMANDA',
      'MESA',
      'RESPONSAVEL',
      'DATA_ABERTURA',
      'DATA_FECHAMENTO',
      'TOTAL',
      'STATUS',
      'OBSERVACOES'
    ],

    'Itens_Comanda': [
      'ID_ITEM',
      'ID_COMANDA',
      'ID_PRODUTO',
      'PRODUTO',
      'TIPO',
      'QUANTIDADE',
      'PESO_KG',
      'PRECO_UNITARIO',
      'PRECO_KG',
      'TOTAL_ITEM',
      'OBSERVACOES'
    ],

   'Financeiro': [
  'ID_CONTA',
  'GRUPO_ID',
  'DESCRICAO',
  'FORNECEDOR',
  'CATEGORIA',
  'TIPO_DOCUMENTO',
  'NUMERO_NOTA',
  'DATA_EMISSAO',
  'FORMA_PAGAMENTO',
  'VALOR_TOTAL',
  'PARCELA_ATUAL',
  'TOTAL_PARCELAS',
  'VALOR_PARCELA',
  'VENCIMENTO',
  'DATA_PAGAMENTO',
  'STATUS',
  'OBSERVACOES',
  'CRIADO_EM',
  'ATUALIZADO_EM',
  'USUARIO',
  'ATIVO'
],

    'Produtos': [
      'ID_PRODUTO',
      'NOME',
      'CATEGORIA',
      'TIPO_VENDA',
      'PRECO',
      'PRECO_KG',
      'ATIVO'
    ],

    'Configuracoes': [
      'CHAVE',
      'VALOR',
      'DESCRICAO'
    ],

    'Usuarios': [
      'EMAIL',
      'NOME',
      'PERFIL',
      'ATIVO'
    ]
  };

  // Remove a aba padrão e cria as abas do sistema
  const abaPadrao = planilha.getSheets()[0];

  Object.keys(estrutura).forEach(nomeAba => {
    const aba = planilha.insertSheet(nomeAba);
    const cabecalhos = estrutura[nomeAba];

    aba.getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    aba.getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight('bold')
      .setBackground('#1F2937')
      .setFontColor('#FFFFFF');

    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, cabecalhos.length);
  });

  planilha.deleteSheet(abaPadrao);

  // Configurações iniciais
  const abaConfig = planilha.getSheetByName('Configuracoes');

  abaConfig.getRange(2, 1, 4, 3).setValues([
    ['NOME_EMPRESA', 'Lua de Sorvete', 'Nome da empresa'],
    ['MOEDA', 'BRL', 'Moeda utilizada'],
    ['PRECO_KG_PADRAO', '0', 'Preço padrão do sorvete por kg'],
    ['FUSO_HORARIO', 'America/Sao_Paulo', 'Fuso horário do sistema']
  ]);

  // Registra o proprietário
  const emailUsuario = Session.getEffectiveUser().getEmail();

  planilha.getSheetByName('Usuarios')
    .appendRow([
      emailUsuario,
      'Administrador',
      'ADMIN',
      true
    ]);

  Logger.log('Banco de dados criado com sucesso!');
  Logger.log('URL da planilha: ' + planilha.getUrl());
}



function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Sistema Lua de Sorvete')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ======================================================
// CONEXÃO COM O BANCO DE DADOS
// ======================================================

function getBancoDeDados() {

  const idPlanilha = PropertiesService
    .getScriptProperties()
    .getProperty('ID_PLANILHA');

  if (!idPlanilha) {
    throw new Error(
      'O ID da planilha do banco de dados não foi encontrado.'
    );
  }

  return SpreadsheetApp.openById(idPlanilha);
}

// ======================================================
// MÓDULO DE PRODUTOS
// ======================================================


// Converte valores brasileiros como "79,90" para número
function converterValorParaNumero(valor) {

  if (valor === null || valor === undefined || valor === '') {
    return 0;
  }

  if (typeof valor === 'number') {
    return valor;
  }

  let texto = String(valor)
    .trim()
    .replace(/\s/g, '')
    .replace('R$', '');

  // Exemplo: 1.299,90
  if (texto.includes(',') && texto.includes('.')) {
    texto = texto
      .replace(/\./g, '')
      .replace(',', '.');

  // Exemplo: 79,90
  } else if (texto.includes(',')) {
    texto = texto.replace(',', '.');
  }

  const numero = Number(texto);

  return isNaN(numero) ? 0 : numero;
}


// Lista todos os produtos
function listarProdutos() {

  const planilha = getBancoDeDados();
  const aba = planilha.getSheetByName('Produtos');

  if (!aba) {
    throw new Error('A aba Produtos não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 7)
    .getValues();

  return dados.map(linha => ({

    id: String(linha[0] || ''),

    nome: String(linha[1] || ''),

    categoria: String(linha[2] || ''),

    tipoVenda: String(linha[3] || ''),

    preco: converterValorParaNumero(linha[4]),

    precoKg: converterValorParaNumero(linha[5]),

    ativo:
      linha[6] === true ||
      String(linha[6]).toUpperCase() === 'TRUE'

  }));
}


// Salva um produto novo ou edita um existente
function salvarProduto(produto) {

  if (!produto) {
    throw new Error('Dados do produto não recebidos.');
  }

  const nome = String(produto.nome || '').trim();
  const categoria = String(produto.categoria || '').trim();
  const tipoVenda = String(produto.tipoVenda || '').trim().toUpperCase();

  const preco = converterValorParaNumero(produto.preco);
  const precoKg = converterValorParaNumero(produto.precoKg);

  if (!nome) {
    throw new Error('Informe o nome do produto.');
  }

  if (!['UNIDADE', 'PESO'].includes(tipoVenda)) {
    throw new Error('Selecione um tipo de venda válido.');
  }

  if (tipoVenda === 'UNIDADE' && preco <= 0) {
    throw new Error('Informe um preço unitário maior que zero.');
  }

  if (tipoVenda === 'PESO' && precoKg <= 0) {
    throw new Error('Informe um preço por kg maior que zero.');
  }

  const planilha = getBancoDeDados();
  const aba = planilha.getSheetByName('Produtos');

  if (!aba) {
    throw new Error('A aba Produtos não foi encontrada.');
  }


  // EDIÇÃO DE PRODUTO EXISTENTE
  if (produto.id) {

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      throw new Error('Produto não encontrado para edição.');
    }

    const ids = aba
      .getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat();

    const indice = ids.findIndex(
      id => String(id) === String(produto.id)
    );

    if (indice === -1) {
      throw new Error('Produto não encontrado para edição.');
    }

    const linhaPlanilha = indice + 2;
        const dadosAnteriores = {
      id: String(aba.getRange(linhaPlanilha, 1).getValue() || ''),
      nome: String(aba.getRange(linhaPlanilha, 2).getValue() || ''),
      categoria: String(aba.getRange(linhaPlanilha, 3).getValue() || ''),
      tipoVenda: String(aba.getRange(linhaPlanilha, 4).getValue() || ''),
      preco: converterValorParaNumero(
        aba.getRange(linhaPlanilha, 5).getValue()
      ),
      precoKg: converterValorParaNumero(
        aba.getRange(linhaPlanilha, 6).getValue()
      ),
      ativo: aba.getRange(linhaPlanilha, 7).getValue() === true
    };

    const statusAtual = aba
      .getRange(linhaPlanilha, 7)
      .getValue();

    aba
      .getRange(linhaPlanilha, 1, 1, 7)
      .setValues([[
        produto.id,
        nome,
        categoria,
        tipoVenda,
        tipoVenda === 'UNIDADE' ? preco : 0,
        tipoVenda === 'PESO' ? precoKg : 0,
        statusAtual
      ]]);
          const dadosNovos = {
      id: produto.id,
      nome: nome,
      categoria: categoria,
      tipoVenda: tipoVenda,
      preco: tipoVenda === 'UNIDADE' ? preco : 0,
      precoKg: tipoVenda === 'PESO' ? precoKg : 0,
      ativo: statusAtual === true
    };

    registrarAuditoria(
      'Produtos',
      'EDIÇÃO',
      'Produto alterado: ' + nome,
      produto.id,
      dadosAnteriores,
      dadosNovos
    );

    return {
      sucesso: true,
      mensagem: 'Produto atualizado com sucesso!',
      id: produto.id
    };
  }


  // NOVO PRODUTO
  const idProduto =
    'PROD-' +
    Utilities
      .getUuid()
      .substring(0, 8)
      .toUpperCase();

  aba.appendRow([
    idProduto,
    nome,
    categoria,
    tipoVenda,
    tipoVenda === 'UNIDADE' ? preco : 0,
    tipoVenda === 'PESO' ? precoKg : 0,
    true
  ]);
    registrarAuditoria(
    'Produtos',
    'CADASTRO',
    'Novo produto cadastrado: ' + nome,
    idProduto,
    null,
    {
      nome: nome,
      categoria: categoria,
      tipoVenda: tipoVenda,
      preco: tipoVenda === 'UNIDADE' ? preco : 0,
      precoKg: tipoVenda === 'PESO' ? precoKg : 0,
      ativo: true
    }
  );

  return {
    sucesso: true,
    mensagem: 'Produto cadastrado com sucesso!',
    id: idProduto
  };
}


// Ativa ou desativa produto
function alterarStatusProduto(idProduto) {

  if (!idProduto) {
    throw new Error('ID do produto não informado.');
  }

  const planilha = getBancoDeDados();
  const aba = planilha.getSheetByName('Produtos');

  if (!aba) {
    throw new Error('A aba Produtos não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error('Nenhum produto cadastrado.');
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 7)
    .getValues();

  for (let i = 0; i < dados.length; i++) {

    if (String(dados[i][0]) === String(idProduto)) {

      const linhaPlanilha = i + 2;

      const statusAtual =
        dados[i][6] === true ||
        String(dados[i][6]).toUpperCase() === 'TRUE';

      const novoStatus = !statusAtual;

      aba
        .getRange(linhaPlanilha, 7)
        .setValue(novoStatus);
              registrarAuditoria(
        'Produtos',
        novoStatus ? 'ATIVAÇÃO' : 'DESATIVAÇÃO',
        (novoStatus
          ? 'Produto ativado: '
          : 'Produto desativado: '
        ) + String(dados[i][1] || ''),
        idProduto,
        {
          ativo: statusAtual
        },
        {
          ativo: novoStatus
        }
      );

      return {
        sucesso: true,
        novoStatus: novoStatus,
        mensagem: novoStatus
          ? 'Produto ativado com sucesso!'
          : 'Produto desativado com sucesso!'
      };
    }
  }

  throw new Error('Produto não encontrado.');
}
  


  // ======================================================
// SISTEMA DE BACKUP AUTOMÁTICO
// ======================================================

const NOME_PASTA_BACKUPS = 'Backups - Sistema Lua de Sorvete';
const QUANTIDADE_MAXIMA_BACKUPS = 30;


// Cria um backup completo do banco de dados
function criarBackupBancoDeDados() {

  const idPlanilha = PropertiesService
    .getScriptProperties()
    .getProperty('ID_PLANILHA');

  if (!idPlanilha) {
    throw new Error(
      'Não foi possível encontrar o ID do banco de dados.'
    );
  }

  // Localiza o arquivo original no Google Drive
  const arquivoOriginal = DriveApp.getFileById(idPlanilha);

  // Localiza ou cria a pasta de backups
  const pastaBackup = obterPastaDeBackups();

  // Define data e hora do backup
  const agora = new Date();

  const dataFormatada = Utilities.formatDate(
    agora,
    'America/Sao_Paulo',
    'yyyy-MM-dd_HH-mm-ss'
  );

  const nomeBackup =
    'Backup Lua de Sorvete - ' + dataFormatada;

  // Cria uma cópia completa da planilha
  const arquivoBackup = arquivoOriginal.makeCopy(
    nomeBackup,
    pastaBackup
  );

  // Remove backups antigos excedentes
  limparBackupsAntigos();

  console.log(
    'Backup criado com sucesso: ' + arquivoBackup.getName()
  );

  return {
    sucesso: true,
    mensagem: 'Backup criado com sucesso!',
    nomeArquivo: arquivoBackup.getName()
  };
}


// Localiza ou cria a pasta dos backups
function obterPastaDeBackups() {

  const propriedades = PropertiesService.getScriptProperties();

  const idPastaSalvo = propriedades.getProperty(
    'ID_PASTA_BACKUPS'
  );

  // Tenta usar a pasta já cadastrada
  if (idPastaSalvo) {

    try {
      return DriveApp.getFolderById(idPastaSalvo);
    } catch (erro) {
      // Se a pasta foi excluída, criaremos outra
    }
  }

  // Procura uma pasta existente pelo nome
  const pastas = DriveApp.getFoldersByName(
    NOME_PASTA_BACKUPS
  );

  let pasta;

  if (pastas.hasNext()) {

    pasta = pastas.next();

  } else {

    pasta = DriveApp.createFolder(
      NOME_PASTA_BACKUPS
    );
  }

  // Guarda o ID para não depender apenas do nome
  propriedades.setProperty(
    'ID_PASTA_BACKUPS',
    pasta.getId()
  );

  return pasta;
}


// Mantém apenas os backups mais recentes
function limparBackupsAntigos() {

  const pastaBackup = obterPastaDeBackups();

  const arquivos = pastaBackup.getFiles();

  const backups = [];

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();

    backups.push({
      arquivo: arquivo,
      dataCriacao: arquivo.getDateCreated()
    });
  }

  // Mais recentes primeiro
  backups.sort(function(a, b) {
    return b.dataCriacao - a.dataCriacao;
  });

  // Mantém somente os 30 mais recentes
  if (backups.length > QUANTIDADE_MAXIMA_BACKUPS) {

    const backupsParaExcluir = backups.slice(
      QUANTIDADE_MAXIMA_BACKUPS
    );

    backupsParaExcluir.forEach(function(item) {
      item.arquivo.setTrashed(true);
    });
  }
}


// Cria o acionador que executará o backup diariamente
function configurarBackupAutomatico() {

  // Evita criar vários acionadores iguais
  const acionadores = ScriptApp.getProjectTriggers();

  acionadores.forEach(function(acionador) {

    if (
      acionador.getHandlerFunction() ===
      'criarBackupBancoDeDados'
    ) {
      ScriptApp.deleteTrigger(acionador);
    }
  });

  // Cria novo acionador diário
  ScriptApp
    .newTrigger('criarBackupBancoDeDados')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  return {
    sucesso: true,
    mensagem:
      'Backup automático configurado para executar diariamente.'
  };
}



// ======================================================
// MÓDULO DE AUDITORIA
// ======================================================

/**
 * Garante que a aba Auditoria exista.
 * Se não existir, cria automaticamente.
 */
function garantirAbaAuditoria() {

  const planilha = getBancoDeDados();

  let aba = planilha.getSheetByName('Auditoria');

  if (!aba) {

    aba = planilha.insertSheet('Auditoria');

    const cabecalhos = [
      'ID_AUDITORIA',
      'DATA_HORA',
      'MODULO',
      'ACAO',
      'DESCRICAO',
      'ID_REGISTRO',
      'USUARIO',
      'DADOS_ANTERIORES',
      'DADOS_NOVOS'
    ];

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight('bold')
      .setBackground('#1F2937')
      .setFontColor('#FFFFFF');

    aba.setFrozenRows(1);

    aba.autoResizeColumns(
      1,
      cabecalhos.length
    );
  }

  return aba;
}


/**
 * Registra uma operação importante na auditoria.
 */
function registrarAuditoria(
  modulo,
  acao,
  descricao,
  idRegistro,
  dadosAnteriores,
  dadosNovos
) {

  try {

    const aba = garantirAbaAuditoria();

    const idAuditoria =
      'AUD-' +
      Utilities
        .getUuid()
        .substring(0, 8)
        .toUpperCase();

    const agora = new Date();

    let usuario = '';

    try {

      usuario =
        Session
          .getActiveUser()
          .getEmail() || '';

    } catch (erro) {

      usuario = '';
    }

    if (!usuario) {

      try {

        usuario =
          Session
            .getEffectiveUser()
            .getEmail() || '';

      } catch (erro) {

        usuario = 'Não identificado';
      }
    }


    const anterioresTexto =
      dadosAnteriores
        ? JSON.stringify(dadosAnteriores)
        : '';


    const novosTexto =
      dadosNovos
        ? JSON.stringify(dadosNovos)
        : '';


    aba.appendRow([
      idAuditoria,
      agora,
      String(modulo || ''),
      String(acao || ''),
      String(descricao || ''),
      String(idRegistro || ''),
      usuario,
      anterioresTexto,
      novosTexto
    ]);


    return {
      sucesso: true,
      id: idAuditoria
    };

  } catch (erro) {

    /*
      A auditoria não deve impedir uma operação principal
      como salvar um produto ou finalizar uma venda.

      Por isso, em caso de falha, registramos o erro no log,
      mas não interrompemos a operação principal.
    */

    console.error(
      'Erro ao registrar auditoria: ' +
      erro.message
    );

    return {
      sucesso: false,
      erro: erro.message
    };
  }
}


/**
 * Lista registros da auditoria.
 * Pode receber filtros opcionais.
 */
function listarAuditoria(filtros) {

  const aba = garantirAbaAuditoria();

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }


  const dados = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      9
    )
    .getValues();


  filtros = filtros || {};


  let registros = dados.map(function(linha) {

    return {

      id: String(linha[0] || ''),

      dataHora:
        linha[1] instanceof Date
          ? linha[1].toISOString()
          : String(linha[1] || ''),

      modulo: String(linha[2] || ''),

      acao: String(linha[3] || ''),

      descricao: String(linha[4] || ''),

      idRegistro: String(linha[5] || ''),

      usuario: String(linha[6] || ''),

      dadosAnteriores: String(linha[7] || ''),

      dadosNovos: String(linha[8] || '')
    };
  });


  // Filtro por módulo
  if (filtros.modulo) {

    registros = registros.filter(function(item) {

      return item.modulo === filtros.modulo;

    });
  }


  // Filtro por ação
  if (filtros.acao) {

    registros = registros.filter(function(item) {

      return item.acao === filtros.acao;

    });
  }


  // Pesquisa livre
  if (filtros.pesquisa) {

    const termo =
      String(filtros.pesquisa)
        .trim()
        .toLowerCase();


    registros = registros.filter(function(item) {

      return (
        item.descricao
          .toLowerCase()
          .includes(termo) ||

        item.idRegistro
          .toLowerCase()
          .includes(termo) ||

        item.usuario
          .toLowerCase()
          .includes(termo)
      );
    });
  }


  // Mais recentes primeiro
  registros.reverse();


  // Limite inicial para não sobrecarregar a tela
  return registros.slice(0, 500);
}

// ======================================================
// MÓDULO DE CONFIGURAÇÕES
// ======================================================

function obterConfiguracoesSistema() {

  const planilha = getBancoDeDados();
  const aba = planilha.getSheetByName('Configuracoes');

  if (!aba) {
    throw new Error('A aba Configuracoes não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  const resultado = {};

  if (ultimaLinha >= 2) {

    const dados = aba
      .getRange(2, 1, ultimaLinha - 1, 3)
      .getValues();

    dados.forEach(function(linha) {

      const chave = String(linha[0] || '');

      if (chave) {
        resultado[chave] = linha[1];
      }
    });
  }

  return {
    nomeEmpresa: resultado.NOME_EMPRESA || 'Lua de Sorvete',
    moeda: resultado.MOEDA || 'BRL',
    precoKgPadrao: converterValorParaNumero(
      resultado.PRECO_KG_PADRAO
    ),
    fusoHorario:
      resultado.FUSO_HORARIO || 'America/Sao_Paulo'
  };
}


// Salva as configurações gerais
function salvarConfiguracoesSistema(configuracoes) {

  if (!configuracoes) {
    throw new Error('Configurações não recebidas.');
  }

  const planilha = getBancoDeDados();
  const aba = planilha.getSheetByName('Configuracoes');

  if (!aba) {
    throw new Error('A aba Configuracoes não foi encontrada.');
  }

  const configuracoesAnteriores =
    obterConfiguracoesSistema();

  const nomeEmpresa =
    String(configuracoes.nomeEmpresa || '').trim();

  const precoKgPadrao =
    converterValorParaNumero(
      configuracoes.precoKgPadrao
    );

  if (!nomeEmpresa) {
    throw new Error('Informe o nome da empresa.');
  }

  const novosValores = {
    NOME_EMPRESA: nomeEmpresa,
    MOEDA: 'BRL',
    PRECO_KG_PADRAO: precoKgPadrao,
    FUSO_HORARIO: 'America/Sao_Paulo'
  };

  const descricoes = {
    NOME_EMPRESA: 'Nome da empresa',
    MOEDA: 'Moeda utilizada',
    PRECO_KG_PADRAO: 'Preço padrão do sorvete por kg',
    FUSO_HORARIO: 'Fuso horário do sistema'
  };

  Object.keys(novosValores).forEach(function(chave) {

    atualizarOuCriarConfiguracao(
      aba,
      chave,
      novosValores[chave],
      descricoes[chave]
    );
  });

  const configuracoesNovas =
    obterConfiguracoesSistema();

  registrarAuditoria(
    'Configurações',
    'ALTERAÇÃO',
    'Configurações gerais do sistema foram alteradas.',
    'CONFIG-GERAL',
    configuracoesAnteriores,
    configuracoesNovas
  );

  return {
    sucesso: true,
    mensagem: 'Configurações salvas com sucesso!'
  };
}


// Atualiza uma configuração existente ou cria uma nova
function atualizarOuCriarConfiguracao(
  aba,
  chave,
  valor,
  descricao
) {

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha >= 2) {

    const chaves = aba
      .getRange(2, 1, ultimaLinha - 1, 1)
      .getValues()
      .flat();

    const indice = chaves.findIndex(function(item) {
      return String(item) === String(chave);
    });

    if (indice !== -1) {

      const linha = indice + 2;

      aba
        .getRange(linha, 2)
        .setValue(valor);

      aba
        .getRange(linha, 3)
        .setValue(descricao);

      return;
    }
  }

  aba.appendRow([
    chave,
    valor,
    descricao
  ]);
}


// Retorna informações sobre os backups
function obterInformacoesBackup() {

  const pasta = obterPastaDeBackups();

  const arquivos = pasta.getFiles();

  const backups = [];

  while (arquivos.hasNext()) {

    const arquivo = arquivos.next();

    backups.push({
      nome: arquivo.getName(),
      dataCriacao: arquivo.getDateCreated().toISOString()
    });
  }

  backups.sort(function(a, b) {

    return new Date(b.dataCriacao) -
           new Date(a.dataCriacao);
  });

  const acionadores = ScriptApp.getProjectTriggers();

  const backupAutomaticoAtivo =
    acionadores.some(function(acionador) {

      return (
        acionador.getHandlerFunction() ===
        'criarBackupBancoDeDados'
      );
    });

  return {
    backupAutomaticoAtivo: backupAutomaticoAtivo,
    quantidadeBackups: backups.length,
    ultimoBackup:
      backups.length > 0
        ? backups[0]
        : null,
    limiteBackups: QUANTIDADE_MAXIMA_BACKUPS
  };
}


// Cria backup manual pela interface
function criarBackupManual() {

  const resultado = criarBackupBancoDeDados();

  registrarAuditoria(
    'Backup',
    'BACKUP MANUAL',
    'Backup manual do banco de dados criado.',
    resultado.nomeArquivo || '',
    null,
    {
      nomeArquivo: resultado.nomeArquivo || ''
    }
  );

  return resultado;
}

// ======================================================
// MÓDULO FINANCEIRO - NOVA VERSÃO
// Estrutura de 21 colunas
// ======================================================


// ------------------------------------------------------
// 1. GARANTIR A ESTRUTURA DA ABA FINANCEIRO
// ------------------------------------------------------

function garantirAbaFinanceiro() {

  const planilha = getBancoDeDados();

  let aba = planilha.getSheetByName('Financeiro');

  const cabecalhos = [
    'ID_CONTA',
    'GRUPO_ID',
    'DESCRICAO',
    'FORNECEDOR',
    'CATEGORIA',
    'TIPO_DOCUMENTO',
    'NUMERO_NOTA',
    'DATA_EMISSAO',
    'FORMA_PAGAMENTO',
    'VALOR_TOTAL',
    'VALOR_PARCELA',
    'PARCELA_ATUAL',
    'TOTAL_PARCELAS',
    'VENCIMENTO',
    'DATA_PAGAMENTO',
    'STATUS',
    'OBSERVACOES',
    'DATA_CADASTRO',
    'DATA_ATUALIZACAO',
    'USUARIO',
    'ATIVO'
  ];

  if (!aba) {

    aba = planilha.insertSheet('Financeiro');

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight('bold')
      .setBackground('#1F2937')
      .setFontColor('#FFFFFF');

    aba.setFrozenRows(1);

    aba
      .getRange('J:K')
      .setNumberFormat('R$ #,##0.00');

    aba
      .getRange('H:H')
      .setNumberFormat('dd/MM/yyyy');

    aba
      .getRange('N:O')
      .setNumberFormat('dd/MM/yyyy');

    aba
      .getRange('R:S')
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');

    aba.autoResizeColumns(
      1,
      cabecalhos.length
    );
  }

  return aba;
}


// ------------------------------------------------------
// 2. FUNÇÕES AUXILIARES DE DATA
// ------------------------------------------------------

function converterDataFinanceiro(valor) {

  if (!valor) {
    return null;
  }

  if (valor instanceof Date) {

    if (isNaN(valor.getTime())) {
      return null;
    }

    return new Date(
      valor.getFullYear(),
      valor.getMonth(),
      valor.getDate()
    );
  }

  const texto = String(valor).trim();

  // Formato recebido do input HTML: AAAA-MM-DD
  const partes = texto.split('-');

  if (partes.length === 3) {

    const ano = Number(partes[0]);
    const mes = Number(partes[1]) - 1;
    const dia = Number(partes[2]);

    const data = new Date(
      ano,
      mes,
      dia
    );

    if (!isNaN(data.getTime())) {
      return data;
    }
  }

  const data = new Date(valor);

  if (isNaN(data.getTime())) {
    return null;
  }

  return new Date(
    data.getFullYear(),
    data.getMonth(),
    data.getDate()
  );
}


function formatarDataFinanceiroISO(valor) {

  if (!valor) {
    return '';
  }

  const data =
    valor instanceof Date
      ? valor
      : new Date(valor);

  if (isNaN(data.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone() ||
      'America/Sao_Paulo',
    'yyyy-MM-dd'
  );
}


function obterHojeFinanceiro() {

  const agora = new Date();

  return new Date(
    agora.getFullYear(),
    agora.getMonth(),
    agora.getDate()
  );
}


// ------------------------------------------------------
// 3. FUNÇÕES AUXILIARES DE VALOR
// ------------------------------------------------------

function converterValorFinanceiro(valor) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {
    return 0;
  }

  if (typeof valor === 'number') {

    return isNaN(valor)
      ? 0
      : valor;
  }

  let texto = String(valor)
    .trim()
    .replace(/\s/g, '')
    .replace('R$', '');

  // Exemplo brasileiro: 1.250,50
  if (
    texto.includes(',') &&
    texto.includes('.')
  ) {

    texto = texto
      .replace(/\./g, '')
      .replace(',', '.');

  } else if (texto.includes(',')) {

    texto = texto.replace(',', '.');
  }

  const numero = Number(texto);

  return isNaN(numero)
    ? 0
    : numero;
}


// ------------------------------------------------------
// 4. GERAR IDENTIFICADORES ÚNICOS
// ------------------------------------------------------

function gerarIdContaFinanceira() {

  return (
    'CONT-' +
    Utilities
      .getUuid()
      .substring(0, 8)
      .toUpperCase()
  );
}


function gerarGrupoIdFinanceiro() {

  return (
    'GRP-' +
    Utilities
      .getUuid()
      .substring(0, 8)
      .toUpperCase()
  );
}


// ------------------------------------------------------
// 5. OBTER USUÁRIO ATUAL
// ------------------------------------------------------

function obterUsuarioFinanceiro() {

  try {

    const email =
      Session
        .getActiveUser()
        .getEmail();

    return email || 'USUARIO';

  } catch (erro) {

    return 'USUARIO';
  }
}


// ------------------------------------------------------
// 6. STATUS AUTOMÁTICO
// ------------------------------------------------------

function calcularStatusFinanceiro(
  formaPagamento,
  statusSalvo,
  vencimento,
  dataPagamento
) {

  const forma =
    String(formaPagamento || '')
      .trim()
      .toUpperCase();

  const status =
    String(statusSalvo || '')
      .trim()
      .toUpperCase();


  // Se já possui pagamento, sempre é PAGA
  if (dataPagamento) {
    return 'PAGA';
  }


  // Pagamentos imediatos
  if (
    forma === 'DINHEIRO' ||
    forma === 'PIX' ||
    forma === 'DEBITO'
  ) {
    return 'PAGA';
  }


  // Se estiver explicitamente paga
  if (status === 'PAGA') {
    return 'PAGA';
  }


  // Boleto depende do vencimento
  if (forma === 'BOLETO') {

    const dataVencimento =
      converterDataFinanceiro(vencimento);

    if (!dataVencimento) {
      return 'A PAGAR';
    }

    if (
      dataVencimento <
      obterHojeFinanceiro()
    ) {
      return 'VENCIDA';
    }

    return 'A PAGAR';
  }


  /*
    Crédito:
    nesta primeira versão não há vínculo
    com banco, cartão, fechamento ou fatura.

    As parcelas são geradas automaticamente,
    sem exigir vencimento.
  */
  if (forma === 'CREDITO') {

    if (status === 'PAGA') {
      return 'PAGA';
    }

    return 'A PAGAR';
  }


  // Outro: respeita o status informado
  if (forma === 'OUTRO') {

    if (
      status === 'VENCIDA' ||
      status === 'PAGA' ||
      status === 'A PAGAR'
    ) {
      return status;
    }

    return 'A PAGAR';
  }


  return status || 'A PAGAR';
}


// ------------------------------------------------------
// 7. CONVERTER LINHA DA PLANILHA EM OBJETO
// ------------------------------------------------------

function converterLinhaFinanceiroEmObjeto(linha) {

  const statusCalculado =
    calcularStatusFinanceiro(
      linha[8],   // FORMA_PAGAMENTO
      linha[15],  // STATUS
      linha[13],  // VENCIMENTO
      linha[14]   // DATA_PAGAMENTO
    );

  return {

    id: String(linha[0] || ''),

    grupoId: String(linha[1] || ''),

    descricao: String(linha[2] || ''),

    fornecedor: String(linha[3] || ''),

    categoria: String(linha[4] || ''),

    tipoDocumento:
      String(linha[5] || ''),

    numeroNota:
      String(linha[6] || ''),

    dataEmissao:
      formatarDataFinanceiroISO(
        linha[7]
      ),

    formaPagamento:
      String(linha[8] || ''),

    valorTotal:
      converterValorFinanceiro(
        linha[9]
      ),

    valorParcela:
      converterValorFinanceiro(
        linha[10]
      ),

    parcelaAtual:
      Number(linha[11] || 1),

    totalParcelas:
      Number(linha[12] || 1),

    vencimento:
      formatarDataFinanceiroISO(
        linha[13]
      ),

    dataPagamento:
      formatarDataFinanceiroISO(
        linha[14]
      ),

    status: statusCalculado,

    observacoes:
      String(linha[16] || ''),

    dataCadastro:
      linha[17] instanceof Date
        ? linha[17].toISOString()
        : String(linha[17] || ''),

    dataAtualizacao:
      linha[18] instanceof Date
        ? linha[18].toISOString()
        : String(linha[18] || ''),

    usuario:
      String(linha[19] || ''),

    ativo:
      linha[20] === true ||
      String(linha[20])
        .toUpperCase() === 'TRUE'
  };
}


// ------------------------------------------------------
// 8. LISTAR CONTAS FINANCEIRAS
// ------------------------------------------------------


// ======================================================
// DASHBOARD - RESUMO ATUAL
// ======================================================

// ======================================================
// COMANDAS - MESAS, COMANDAS E ITENS
// ======================================================
const COMANDAS_CABECALHOS = [
  'ID_COMANDA','MESA','NOME_CLIENTE','STATUS','CRIADO_EM',
  'ATUALIZADO_EM','USUARIO','ATIVO'
];

const ITENS_COMANDA_CABECALHOS = [
  'ID_ITEM','ID_COMANDA','ID_PRODUTO','NOME_PRODUTO','TIPO_VENDA',
  'QUANTIDADE','PESO_KG','PRECO_UNITARIO','SUBTOTAL','CRIADO_EM','ATIVO'
];

function garantirEstruturaComandas() {
  const planilha = getBancoDeDados();

  let abaComandas = planilha.getSheetByName('Comandas');
  if (!abaComandas) {
    abaComandas = planilha.insertSheet('Comandas');
    abaComandas.getRange(1, 1, 1, COMANDAS_CABECALHOS.length)
      .setValues([COMANDAS_CABECALHOS]);
    abaComandas.setFrozenRows(1);
  }

  let abaItens = planilha.getSheetByName('Itens_Comanda');
  if (!abaItens) {
    abaItens = planilha.insertSheet('Itens_Comanda');
    abaItens.getRange(1, 1, 1, ITENS_COMANDA_CABECALHOS.length)
      .setValues([ITENS_COMANDA_CABECALHOS]);
    abaItens.setFrozenRows(1);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOTAL_MESAS')) {
    props.setProperty('TOTAL_MESAS', '10');
  }

  return {
    abaComandas: abaComandas,
    abaItens: abaItens,
    totalMesas: Number(props.getProperty('TOTAL_MESAS') || 10)
  };
}

function listarPainelComandas() {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaComandas;
  const ultimaLinha = aba.getLastRow();
  let comandas = [];

  if (ultimaLinha >= 2) {
    const dados = aba.getRange(2, 1, ultimaLinha - 1, COMANDAS_CABECALHOS.length).getValues();
    comandas = dados.map(function(l) {
      return {
        id: String(l[0] || ''),
        mesa: Number(l[1] || 0),
        nomeCliente: String(l[2] || ''),
        status: String(l[3] || 'ABERTA').toUpperCase(),
        criadoEm: l[4] instanceof Date ? l[4].toISOString() : String(l[4] || ''),
        atualizadoEm: l[5] instanceof Date ? l[5].toISOString() : String(l[5] || ''),
        ativo: l[7] === true || String(l[7]).toUpperCase() === 'TRUE'
      };
    }).filter(function(c) { return c.ativo && c.status !== 'FINALIZADA'; });
  }

  const itens = listarTodosItensComanda_();
  const totais = {};
  itens.forEach(function(item) {
    totais[item.idComanda] = (totais[item.idComanda] || 0) + Number(item.subtotal || 0);
  });

  comandas.forEach(function(c) {
    c.total = totais[c.id] || 0;
  });

  const mesas = [];
  for (let numero = 1; numero <= estrutura.totalMesas; numero++) {
    const daMesa = comandas.filter(function(c) { return c.mesa === numero; });
    const abertas = daMesa.filter(function(c) { return c.status === 'ABERTA'; }).length;
    const emPagamento = daMesa.filter(function(c) { return c.status === 'AGUARDANDO PAGAMENTO'; }).length;
    let statusMesa = 'LIVRE';
    if (abertas > 0 && emPagamento > 0) statusMesa = 'MISTA';
    else if (abertas > 0) statusMesa = 'OCUPADA';
    else if (emPagamento > 0) statusMesa = 'AGUARDANDO PAGAMENTO';

    mesas.push({
      numero: numero,
      quantidadeComandas: daMesa.length,
      total: daMesa.reduce(function(soma, c) { return soma + Number(c.total || 0); }, 0),
      abertas: abertas,
      emPagamento: emPagamento,
      status: statusMesa
    });
  }

  return { mesas: mesas, comandas: comandas };
}

function adicionarMesaComanda() {
  const estrutura = garantirEstruturaComandas();
  const novoTotal = estrutura.totalMesas + 1;
  PropertiesService.getScriptProperties().setProperty('TOTAL_MESAS', String(novoTotal));
  return { sucesso: true, totalMesas: novoTotal };
}

function criarNovaComanda(mesa, nomeCliente) {
  const estrutura = garantirEstruturaComandas();
  mesa = Number(mesa || 0);

  if (!mesa || mesa < 1 || mesa > estrutura.totalMesas) {
    throw new Error('Selecione uma mesa válida.');
  }

  const id = Utilities.getUuid();
  const agora = new Date();
  const usuario = Session.getActiveUser().getEmail() || 'Sistema';

  estrutura.abaComandas.appendRow([
    id, mesa, String(nomeCliente || '').trim(), 'ABERTA',
    agora, agora, usuario, true
  ]);

  if (typeof registrarAuditoria === 'function') {
    registrarAuditoria('Comandas', 'CADASTRO', id,
      'Nova comanda aberta na Mesa ' + String(mesa).padStart(2, '0'),
      '', JSON.stringify({ mesa: mesa, nomeCliente: String(nomeCliente || '').trim() }));
  }

  return { sucesso: true, id: id };
}

function obterDetalhesComanda(idComanda) {
  const painel = listarPainelComandas();
  const comanda = painel.comandas.find(function(c) { return c.id === String(idComanda); });
  if (!comanda) throw new Error('Comanda não encontrada.');

  const itens = listarTodosItensComanda_().filter(function(i) {
    return i.idComanda === String(idComanda) && i.ativo;
  });

  comanda.itens = itens;
  comanda.total = itens.reduce(function(soma, item) {
    return soma + Number(item.subtotal || 0);
  }, 0);

  return comanda;
}

function listarProdutosAtivosParaComanda() {
  return listarProdutos().filter(function(p) { return p.ativo; });
}

function adicionarItemComanda(dados) {
  if (!dados || !dados.idComanda || !dados.idProduto) {
    throw new Error('Dados do item incompletos.');
  }

  const comanda = obterDetalhesComanda(dados.idComanda);
  if (comanda.status !== 'ABERTA') {
    throw new Error('Esta comanda não está aberta para novos lançamentos.');
  }

  const produto = listarProdutos().find(function(p) {
    return String(p.id) === String(dados.idProduto) && p.ativo;
  });
  if (!produto) throw new Error('Produto não encontrado ou inativo.');

  const tipo = String(produto.tipoVenda || '').toUpperCase();
  let quantidade = 0;
  let pesoKg = 0;
  let precoUnitario = 0;
  let subtotal = 0;

  if (tipo === 'PESO') {
    pesoKg = Number(dados.pesoKg || 0);
    if (pesoKg <= 0) throw new Error('Informe um peso maior que zero.');
    precoUnitario = Number(produto.precoKg || 0);
    subtotal = pesoKg * precoUnitario;
  } else {
    quantidade = Math.max(1, Number(dados.quantidade || 1));
    precoUnitario = Number(produto.preco || 0);
    subtotal = quantidade * precoUnitario;
  }

  const estrutura = garantirEstruturaComandas();
  const agora = new Date();
  estrutura.abaItens.appendRow([
    Utilities.getUuid(),
    String(dados.idComanda),
    String(produto.id),
    String(produto.nome),
    tipo,
    quantidade,
    pesoKg,
    precoUnitario,
    Math.round(subtotal * 100) / 100,
    agora,
    true
  ]);

  atualizarDataComanda_(dados.idComanda);
  return obterDetalhesComanda(dados.idComanda);
}

function alterarQuantidadeItemComanda(idItem, delta) {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaItens;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) throw new Error('Item não encontrado.');

  const dados = aba.getRange(2, 1, ultimaLinha - 1, ITENS_COMANDA_CABECALHOS.length).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idItem) && (dados[i][10] === true || String(dados[i][10]).toUpperCase() === 'TRUE')) {
      if (String(dados[i][4]).toUpperCase() === 'PESO') {
        throw new Error('Para produto por peso, remova e adicione novamente com o peso correto.');
      }
      const novaQtd = Math.max(1, Number(dados[i][5] || 1) + Number(delta || 0));
      const preco = Number(dados[i][7] || 0);
      aba.getRange(i + 2, 6).setValue(novaQtd);
      aba.getRange(i + 2, 9).setValue(Math.round(novaQtd * preco * 100) / 100);
      atualizarDataComanda_(dados[i][1]);
      return obterDetalhesComanda(dados[i][1]);
    }
  }
  throw new Error('Item não encontrado.');
}

function editarItemComanda(idItem, alteracoes) {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaItens;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) throw new Error('Item não encontrado.');
  const dados = aba.getRange(2, 1, ultimaLinha - 1, ITENS_COMANDA_CABECALHOS.length).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idItem) && (dados[i][10] === true || String(dados[i][10]).toUpperCase() === 'TRUE')) {
      const tipo = String(dados[i][4]).toUpperCase();
      const preco = Number(dados[i][7] || 0);
      if (tipo === 'PESO') {
        const peso = Number(alteracoes.pesoKg || 0);
        if (peso <= 0) throw new Error('Informe um peso válido.');
        aba.getRange(i + 2, 7).setValue(peso);
        aba.getRange(i + 2, 9).setValue(Math.round(peso * preco * 100) / 100);
      } else {
        const qtd = Number(alteracoes.quantidade || 0);
        if (!Number.isInteger(qtd) || qtd < 1) throw new Error('Informe uma quantidade válida.');
        aba.getRange(i + 2, 6).setValue(qtd);
        aba.getRange(i + 2, 9).setValue(Math.round(qtd * preco * 100) / 100);
      }
      atualizarDataComanda_(dados[i][1]);
      return obterDetalhesComanda(dados[i][1]);
    }
  }
  throw new Error('Item não encontrado.');
}

function removerItemComanda(idItem) {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaItens;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) throw new Error('Item não encontrado.');

  const dados = aba.getRange(2, 1, ultimaLinha - 1, ITENS_COMANDA_CABECALHOS.length).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idItem)) {
      aba.getRange(i + 2, 11).setValue(false);
      atualizarDataComanda_(dados[i][1]);
      return obterDetalhesComanda(dados[i][1]);
    }
  }
  throw new Error('Item não encontrado.');
}



function transferirItemEntreComandas(dados) {
  if (!dados || !dados.idItem || !dados.idComandaDestino) throw new Error('Dados da transferência incompletos.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const estrutura = garantirEstruturaComandas();
    const abaItens = estrutura.abaItens;
    const ultimaLinha = abaItens.getLastRow();
    if (ultimaLinha < 2) throw new Error('Item não encontrado.');

    const qtdColunas = ITENS_COMANDA_CABECALHOS.length;
    const faixa = abaItens.getRange(2, 1, ultimaLinha - 1, qtdColunas);
    const valores = faixa.getValues();
    let indice = -1;
    let item = null;

    for (let i = 0; i < valores.length; i++) {
      const ativo = valores[i][10] === true || String(valores[i][10]).toUpperCase() === 'TRUE';
      if (String(valores[i][0]) === String(dados.idItem) && ativo) {
        indice = i;
        item = valores[i].slice();
        break;
      }
    }
    if (!item) throw new Error('Item não encontrado ou já removido.');

    const idOrigem = String(item[1]);
    const idDestino = String(dados.idComandaDestino);
    if (idOrigem === idDestino) throw new Error('Escolha outra comanda como destino.');

    const origemAntes = obterDetalhesComanda(idOrigem);
    const destinoAntes = obterDetalhesComanda(idDestino);
    if (Number(origemAntes.mesa) !== Number(destinoAntes.mesa)) throw new Error('A transferência só é permitida entre comandas da mesma mesa.');
    if (origemAntes.status !== 'ABERTA' || destinoAntes.status !== 'ABERTA') throw new Error('As duas comandas precisam estar abertas para transferir itens.');

    const tipo = String(item[4] || '').toUpperCase();
    const quantidadeAtual = Number(item[5] || 0);
    const preco = Number(item[7] || 0);
    let quantidadeTransferida = quantidadeAtual;
    let transferenciaParcial = false;

    if (tipo !== 'PESO' && dados.quantidade !== null && dados.quantidade !== undefined && dados.quantidade !== '') {
      quantidadeTransferida = Number(dados.quantidade);
      if (!Number.isInteger(quantidadeTransferida) || quantidadeTransferida < 1 || quantidadeTransferida > quantidadeAtual) {
        throw new Error('Quantidade inválida para transferência.');
      }
      transferenciaParcial = quantidadeTransferida < quantidadeAtual;
    }

    let idItemDestino = String(item[0]);

    if (tipo === 'PESO' || !transferenciaParcial) {
      // V4.3: grava a linha inteira de uma vez, alterando somente o ID_COMANDA.
      // Isso evita uma atualização parcial da linha e facilita a confirmação real da transferência.
      item[1] = idDestino;
      abaItens.getRange(indice + 2, 1, 1, qtdColunas).setValues([item]);
    } else {
      const restante = quantidadeAtual - quantidadeTransferida;
      const linhaOrigem = item.slice();
      linhaOrigem[5] = restante;
      linhaOrigem[8] = Math.round(restante * preco * 100) / 100;
      abaItens.getRange(indice + 2, 1, 1, qtdColunas).setValues([linhaOrigem]);

      idItemDestino = Utilities.getUuid();
      const linhaDestino = [
        idItemDestino, idDestino, String(item[2]), String(item[3]), tipo,
        quantidadeTransferida, 0, preco,
        Math.round(quantidadeTransferida * preco * 100) / 100,
        new Date(), true
      ];
      abaItens.appendRow(linhaDestino);
    }

    atualizarDataComanda_(idOrigem);
    atualizarDataComanda_(idDestino);
    SpreadsheetApp.flush();

    // Confirma diretamente na planilha, sem depender de cache ou resumo de painel.
    const ultimaLinhaDepois = abaItens.getLastRow();
    const dadosDepois = abaItens.getRange(2, 1, ultimaLinhaDepois - 1, qtdColunas).getValues();
    const linhaConfirmadaDestino = dadosDepois.find(function(l) {
      const ativo = l[10] === true || String(l[10]).toUpperCase() === 'TRUE';
      return ativo && String(l[0]) === String(idItemDestino) && String(l[1]) === idDestino;
    });
    if (!linhaConfirmadaDestino) {
      throw new Error('A transferência não foi gravada na comanda de destino.');
    }

    const origemDepois = obterDetalhesComanda(idOrigem);
    const destinoDepois = obterDetalhesComanda(idDestino);

    if (typeof registrarAuditoria === 'function') {
      registrarAuditoria('Comandas', 'TRANSFERENCIA', String(dados.idItem),
        'Item transferido entre comandas da Mesa ' + String(origemAntes.mesa).padStart(2, '0'),
        JSON.stringify({ idComanda: idOrigem, nomeCliente: origemAntes.nomeCliente, quantidade: quantidadeAtual }),
        JSON.stringify({ idComanda: idDestino, nomeCliente: destinoAntes.nomeCliente, quantidadeTransferida: tipo === 'PESO' ? null : quantidadeTransferida, pesoKg: tipo === 'PESO' ? Number(item[6] || 0) : null, produto: String(item[3]) })
      );
    }

    return {
      sucesso: true,
      confirmado: true,
      idOrigem: idOrigem,
      idDestino: idDestino,
      idItemDestino: idItemDestino,
      origem: origemDepois,
      destino: destinoDepois
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelarComanda(idComanda) {
  const estrutura = garantirEstruturaComandas();
  const abaComandas = estrutura.abaComandas;
  const abaItens = estrutura.abaItens;
  const ultimaLinha = abaComandas.getLastRow();
  if (ultimaLinha < 2) throw new Error('Comanda não encontrada.');

  const dados = abaComandas.getRange(2, 1, ultimaLinha - 1, COMANDAS_CABECALHOS.length).getValues();
  let linhaComanda = -1;
  let mesa = 0;
  let nomeCliente = '';

  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idComanda)) {
      linhaComanda = i + 2;
      mesa = Number(dados[i][1] || 0);
      nomeCliente = String(dados[i][2] || '');
      const status = String(dados[i][3] || '').toUpperCase();
      if (status !== 'ABERTA') throw new Error('Somente comandas abertas podem ser canceladas.');
      break;
    }
  }
  if (linhaComanda < 0) throw new Error('Comanda não encontrada.');

  // Cancela a comanda sem apagar o histórico.
  abaComandas.getRange(linhaComanda, 4).setValue('CANCELADA');
  abaComandas.getRange(linhaComanda, 6).setValue(new Date());
  abaComandas.getRange(linhaComanda, 8).setValue(false);

  // Desativa todos os itens da comanda em uma única gravação em lote.
  const ultimaLinhaItens = abaItens.getLastRow();
  if (ultimaLinhaItens >= 2) {
    const qtd = ultimaLinhaItens - 1;
    const dadosItens = abaItens.getRange(2, 1, qtd, ITENS_COMANDA_CABECALHOS.length).getValues();
    let alterou = false;
    for (let i = 0; i < dadosItens.length; i++) {
      if (String(dadosItens[i][1]) === String(idComanda) &&
          (dadosItens[i][10] === true || String(dadosItens[i][10]).toUpperCase() === 'TRUE')) {
        dadosItens[i][10] = false;
        alterou = true;
      }
    }
    if (alterou) {
      abaItens.getRange(2, 1, qtd, ITENS_COMANDA_CABECALHOS.length).setValues(dadosItens);
    }
  }

  if (typeof registrarAuditoria === 'function') {
    registrarAuditoria(
      'Comandas', 'CANCELAMENTO', String(idComanda),
      'Comanda cancelada e mesa liberada quando não houver outras comandas ativas.',
      JSON.stringify({ mesa: mesa, nomeCliente: nomeCliente, status: 'ABERTA' }),
      JSON.stringify({ mesa: mesa, nomeCliente: nomeCliente, status: 'CANCELADA' })
    );
  }

  return { sucesso: true, idComanda: String(idComanda), mesa: mesa };
}

function enviarComandaParaPagamento(idComanda) {
  // V4.4: usa o mesmo lock da transferência. Se uma transferência ainda estiver
  // sendo persistida, o pagamento espera sua conclusão e lê os itens já gravados.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    SpreadsheetApp.flush();
    const estrutura = garantirEstruturaComandas();
    const aba = estrutura.abaComandas;
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) throw new Error('Comanda não encontrada.');

    const detalhes = obterDetalhesComanda(idComanda);
    if (!detalhes.itens.length) throw new Error('Adicione pelo menos um produto antes de ir para pagamento.');

    const dados = aba.getRange(2, 1, ultimaLinha - 1, COMANDAS_CABECALHOS.length).getValues();
    for (let i = 0; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idComanda)) {
        aba.getRange(i + 2, 4).setValue('AGUARDANDO PAGAMENTO');
        aba.getRange(i + 2, 6).setValue(new Date());
        SpreadsheetApp.flush();

        // Releitura final do servidor: o PDV recebe somente os dados persistidos.
        const confirmado = obterDetalhesComanda(idComanda);
        return {
          sucesso: true,
          mensagem: 'Comanda enviada ao PDV para pagamento.',
          comanda: confirmado
        };
      }
    }
    throw new Error('Comanda não encontrada.');
  } finally {
    lock.releaseLock();
  }
}

// ======================================================
// PDV - VENDAS, ITENS E PAGAMENTOS
// ======================================================
const VENDAS_CABECALHOS = ['ID_VENDA','DATA_HORA','ORIGEM','ID_COMANDA','MESA','NOME_CLIENTE','TOTAL','STATUS','USUARIO'];
const ITENS_VENDA_CABECALHOS = ['ID_ITEM','ID_VENDA','ID_PRODUTO','NOME_PRODUTO','TIPO_VENDA','QUANTIDADE','PESO_KG','PRECO_UNITARIO','SUBTOTAL'];
const PAGAMENTOS_VENDA_CABECALHOS = ['ID_PAGAMENTO','ID_VENDA','FORMA','VALOR','DATA_HORA'];

function garantirEstruturaVendas_(){
  const p=getBancoDeDados();
  function aba(nome,cab){let a=p.getSheetByName(nome);if(!a){a=p.insertSheet(nome);a.getRange(1,1,1,cab.length).setValues([cab]);a.setFrozenRows(1);}return a;}
  return {vendas:aba('Vendas',VENDAS_CABECALHOS),itens:aba('Itens_Venda',ITENS_VENDA_CABECALHOS),pagamentos:aba('Pagamentos_Venda',PAGAMENTOS_VENDA_CABECALHOS)};
}

function finalizarVendaPDV(dados){
  if(!dados||!dados.itens||!dados.itens.length) throw new Error('A venda não possui itens.');
  if(!dados.pagamentos||!dados.pagamentos.length) throw new Error('Informe a forma de pagamento.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const total=dados.itens.reduce(function(s,i){return s+Number(i.subtotal||0);},0);
    const informado=dados.pagamentos.reduce(function(s,p){return s+Number(p.valor||0);},0);
    const temDinheiro=dados.pagamentos.some(function(p){return String(p.forma)==='DINHEIRO';});
    if(informado+0.005<total) throw new Error('O valor pago é menor que o total da venda.');
    if(informado>total+0.005&&!temDinheiro) throw new Error('Valor acima do total só é permitido quando há pagamento em dinheiro para troco.');
    let comanda=null;
    if(dados.idComanda){ comanda=obterDetalhesComanda(dados.idComanda); if(comanda.status!=='AGUARDANDO PAGAMENTO') throw new Error('A comanda não está aguardando pagamento.'); }
    const e=garantirEstruturaVendas_(), agora=new Date(), id='V'+Utilities.formatDate(agora,Session.getScriptTimeZone(),'yyyyMMddHHmmss')+Math.floor(Math.random()*900+100);
    e.vendas.appendRow([id,agora,comanda?'COMANDA':'BALCAO',comanda?comanda.id:'',comanda?comanda.mesa:'',comanda?comanda.nomeCliente:'',total,'FINALIZADA','']);
    const linhasItens=dados.itens.map(function(i){return [Utilities.getUuid(),id,String(i.idProduto||''),String(i.nome||i.nomeProduto||''),String(i.tipoVenda||''),Number(i.quantidade||0),Number(i.pesoKg||0),Number(i.precoUnitario||0),Number(i.subtotal||0)];});
    e.itens.getRange(e.itens.getLastRow()+1,1,linhasItens.length,ITENS_VENDA_CABECALHOS.length).setValues(linhasItens);
    const linhasPag=dados.pagamentos.map(function(p){return [Utilities.getUuid(),id,String(p.forma||''),Number(p.valor||0),agora];});
    e.pagamentos.getRange(e.pagamentos.getLastRow()+1,1,linhasPag.length,PAGAMENTOS_VENDA_CABECALHOS.length).setValues(linhasPag);
    if(comanda){
      const ec=garantirEstruturaComandas(), ult=ec.abaComandas.getLastRow(), vals=ec.abaComandas.getRange(2,1,ult-1,COMANDAS_CABECALHOS.length).getValues();
      for(let i=0;i<vals.length;i++) if(String(vals[i][0])===String(comanda.id)){ec.abaComandas.getRange(i+2,4).setValue('PAGA');ec.abaComandas.getRange(i+2,6).setValue(agora);ec.abaComandas.getRange(i+2,8).setValue(false);break;}
      // Desativa os itens da comanda paga para que não permaneçam visíveis como itens ativos.
      const ultItens = ec.abaItens.getLastRow();
      if (ultItens >= 2) {
        const dadosItens = ec.abaItens.getRange(2, 1, ultItens - 1, ITENS_COMANDA_CABECALHOS.length).getValues();
        for (let j = 0; j < dadosItens.length; j++) {
          if (String(dadosItens[j][1]) === String(comanda.id)) {
            ec.abaItens.getRange(j + 2, 11).setValue(false);
          }
        }
      }
    }
    registrarAuditoria('PDV','VENDA',id,'Venda finalizada'+(comanda?' a partir da comanda '+comanda.id:' no balcão'),'',JSON.stringify({total:total,pagamentos:dados.pagamentos}));
    return {sucesso:true,idVenda:id,total:total,troco:Math.max(0,informado-total)};
  } finally { lock.releaseLock(); }
}

function obterResumoVendas_(){
  const e=garantirEstruturaVendas_(), ult=e.vendas.getLastRow(); if(ult<2)return {vendasHoje:0,faturamentoHoje:0,faturamentoMes:0};
  const vals=e.vendas.getRange(2,1,ult-1,VENDAS_CABECALHOS.length).getValues(), hoje=new Date();
  let qtd=0,dia=0,mes=0; vals.forEach(function(r){if(String(r[7])!=='FINALIZADA')return;const d=new Date(r[1]),v=Number(r[6]||0);if(d.getFullYear()===hoje.getFullYear()&&d.getMonth()===hoje.getMonth()){mes+=v;if(d.getDate()===hoje.getDate()){qtd++;dia+=v;}}});
  return {vendasHoje:qtd,faturamentoHoje:dia,faturamentoMes:mes};
}

function listarTodosItensComanda_() {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaItens;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return [];

  return aba.getRange(2, 1, ultimaLinha - 1, ITENS_COMANDA_CABECALHOS.length)
    .getValues()
    .map(function(l) {
      return {
        id: String(l[0] || ''),
        idComanda: String(l[1] || ''),
        idProduto: String(l[2] || ''),
        nomeProduto: String(l[3] || ''),
        tipoVenda: String(l[4] || ''),
        quantidade: Number(l[5] || 0),
        pesoKg: Number(l[6] || 0),
        precoUnitario: Number(l[7] || 0),
        subtotal: Number(l[8] || 0),
        ativo: l[10] === true || String(l[10]).toUpperCase() === 'TRUE'
      };
    }).filter(function(i) { return i.ativo; });
}

function atualizarDataComanda_(idComanda) {
  const estrutura = garantirEstruturaComandas();
  const aba = estrutura.abaComandas;
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  const ids = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idComanda)) {
      aba.getRange(i + 2, 6).setValue(new Date());
      return;
    }
  }
}

function contarComandasAbertas() {
  return listarPainelComandas().comandas.filter(function(c) {
    return c.status === 'ABERTA' || c.status === 'AGUARDANDO PAGAMENTO';
  }).length;
}


function obterResumoDashboard() {
  const contas = listarContasFinanceiras({}) || [];

  let totalAPagar = 0;
  let totalVencidas = 0;

  contas.forEach(function(conta) {
    const status = String(conta.status || '').trim().toUpperCase();
    const valor = Number(conta.valorParcela || 0);

    if (status === 'A PAGAR') {
      totalAPagar += valor;
    } else if (status === 'VENCIDA') {
      totalVencidas += valor;
    }
  });

  const resumoVendas = obterResumoVendas_();

  return {
    vendasHoje: resumoVendas.vendasHoje,
    faturamentoHoje: resumoVendas.faturamentoHoje,
    faturamentoMes: resumoVendas.faturamentoMes,
    comandasAbertas: contarComandasAbertas(),
    contasAPagar: totalAPagar,
    contasVencidas: totalVencidas
  };
}

function listarContasFinanceiras(filtros) {

  const aba = garantirAbaFinanceiro();

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  filtros = filtros || {};

  const dados = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      21
    )
    .getValues();

  let contas = dados
    .map(converterLinhaFinanceiroEmObjeto)
    .filter(function(conta) {
      return conta.ativo;
    });


  // STATUS
  if (filtros.status) {

    const statusFiltro =
      String(filtros.status)
        .trim()
        .toUpperCase();

    contas = contas.filter(
      function(conta) {

        return (
          conta.status ===
          statusFiltro
        );
      }
    );
  }


  // CATEGORIA
  if (filtros.categoria) {

    const categoriaFiltro =
      String(filtros.categoria)
        .trim()
        .toLowerCase();

    contas = contas.filter(
      function(conta) {

        return (
          conta.categoria
            .trim()
            .toLowerCase() ===
          categoriaFiltro
        );
      }
    );
  }


  // FORMA DE PAGAMENTO
  if (filtros.formaPagamento) {

    const formaFiltro =
      String(filtros.formaPagamento)
        .trim()
        .toUpperCase();

    contas = contas.filter(
      function(conta) {

        return (
          conta.formaPagamento
            .trim()
            .toUpperCase() ===
          formaFiltro
        );
      }
    );
  }


  // DATA INICIAL
  if (filtros.dataInicial) {

    const dataInicial =
      converterDataFinanceiro(
        filtros.dataInicial
      );

    if (dataInicial) {

      contas = contas.filter(
        function(conta) {

          const dataReferencia =
            converterDataFinanceiro(
              conta.vencimento ||
              conta.dataEmissao
            );

          return (
            dataReferencia &&
            dataReferencia >= dataInicial
          );
        }
      );
    }
  }


  // DATA FINAL
  if (filtros.dataFinal) {

    const dataFinal =
      converterDataFinanceiro(
        filtros.dataFinal
      );

    if (dataFinal) {

      contas = contas.filter(
        function(conta) {

          const dataReferencia =
            converterDataFinanceiro(
              conta.vencimento ||
              conta.dataEmissao
            );

          return (
            dataReferencia &&
            dataReferencia <= dataFinal
          );
        }
      );
    }
  }


  // PESQUISA LIVRE
  if (filtros.pesquisa) {

    const termo =
      String(filtros.pesquisa)
        .trim()
        .toLowerCase();

    contas = contas.filter(
      function(conta) {

        return (
          conta.descricao
            .toLowerCase()
            .includes(termo) ||

          conta.fornecedor
            .toLowerCase()
            .includes(termo) ||

          conta.categoria
            .toLowerCase()
            .includes(termo) ||

          conta.numeroNota
            .toLowerCase()
            .includes(termo) ||

          conta.formaPagamento
            .toLowerCase()
            .includes(termo) ||

          conta.observacoes
            .toLowerCase()
            .includes(termo) ||

          conta.id
            .toLowerCase()
            .includes(termo)
        );
      }
    );
  }


  // ORDENAÇÃO
  contas.sort(function(a, b) {

    const dataA =
      converterDataFinanceiro(
        a.vencimento ||
        a.dataEmissao
      );

    const dataB =
      converterDataFinanceiro(
        b.vencimento ||
        b.dataEmissao
      );

    if (!dataA && !dataB) {
      return 0;
    }

    if (!dataA) {
      return 1;
    }

    if (!dataB) {
      return -1;
    }

    return dataA - dataB;
  });


  return contas;
}


// ------------------------------------------------------
// 9. RESUMO FINANCEIRO
// ------------------------------------------------------

function obterResumoFinanceiro(filtros) {

  const contas =
    listarContasFinanceiras(
      filtros || {}
    );

  let totalAPagar = 0;
  let totalVencido = 0;
  let totalPago = 0;

  let quantidadeAPagar = 0;
  let quantidadeVencidas = 0;
  let quantidadePagas = 0;


  contas.forEach(function(conta) {

    const valor =
      converterValorFinanceiro(
        conta.valorParcela
      );

    if (conta.status === 'A PAGAR') {

      totalAPagar += valor;
      quantidadeAPagar++;
    }

    if (conta.status === 'VENCIDA') {

      totalVencido += valor;
      quantidadeVencidas++;
    }

    if (conta.status === 'PAGA') {

      totalPago += valor;
      quantidadePagas++;
    }
  });


  return {

    totalAPagar: totalAPagar,

    totalVencido: totalVencido,

    totalPago: totalPago,

    quantidadeAPagar:
      quantidadeAPagar,

    quantidadeVencidas:
      quantidadeVencidas,

    quantidadePagas:
      quantidadePagas,

    quantidadeTotal:
      contas.length
  };
}


// ------------------------------------------------------
// 10. OPÇÕES PARA FILTROS E FORMULÁRIOS
// ------------------------------------------------------

function obterOpcoesFinanceiro() {

  const contas =
    listarContasFinanceiras({});

  const categorias = [];

  const fornecedores = [];


  contas.forEach(function(conta) {

    if (
      conta.categoria &&
      !categorias.includes(
        conta.categoria
      )
    ) {

      categorias.push(
        conta.categoria
      );
    }


    if (
      conta.fornecedor &&
      !fornecedores.includes(
        conta.fornecedor
      )
    ) {

      fornecedores.push(
        conta.fornecedor
      );
    }
  });


  categorias.sort(function(a, b) {

    return a.localeCompare(
      b,
      'pt-BR'
    );
  });


  fornecedores.sort(function(a, b) {

    return a.localeCompare(
      b,
      'pt-BR'
    );
  });


  return {

    categorias: categorias,

    fornecedores: fornecedores,

    formasPagamento: [
      'DINHEIRO',
      'PIX',
      'DEBITO',
      'CREDITO',
      'BOLETO',
      'OUTRO'
    ],

    tiposDocumento: [
      'SEM DOCUMENTO',
      'NOTA FISCAL',
      'RECIBO',
      'FATURA',
      'OUTRO'
    ]
  };
}

// ======================================================
// MÓDULO FINANCEIRO - BLOCO 2
// Cadastro, parcelamento, edição e pagamentos
// ======================================================


// ------------------------------------------------------
// 11. DIVIDIR VALOR EM PARCELAS
// Garante que a soma seja exatamente igual ao total.
// ------------------------------------------------------

function dividirValorEmParcelasFinanceiro(
  valorTotal,
  quantidadeParcelas
) {

  const totalCentavos = Math.round(
    converterValorFinanceiro(valorTotal) * 100
  );

  const quantidade = Number(quantidadeParcelas);

  if (
    !Number.isInteger(quantidade) ||
    quantidade < 1
  ) {
    throw new Error(
      'Quantidade de parcelas inválida.'
    );
  }

  const valorBaseCentavos = Math.floor(
    totalCentavos / quantidade
  );

  let restanteCentavos =
    totalCentavos -
    (valorBaseCentavos * quantidade);

  const parcelas = [];

  for (let i = 0; i < quantidade; i++) {

    let valorCentavos = valorBaseCentavos;

    if (restanteCentavos > 0) {
      valorCentavos++;
      restanteCentavos--;
    }

    parcelas.push(
      valorCentavos / 100
    );
  }

  return parcelas;
}


// ------------------------------------------------------
// 12. NORMALIZAR FORMA DE PAGAMENTO
// ------------------------------------------------------

function normalizarFormaPagamentoFinanceiro(valor) {

  const forma = String(valor || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

  const formasPermitidas = [
    'DINHEIRO',
    'PIX',
    'DEBITO',
    'CREDITO',
    'BOLETO',
    'OUTRO'
  ];

  if (!formasPermitidas.includes(forma)) {
    throw new Error(
      'Selecione uma forma de pagamento válida.'
    );
  }

  return forma;
}


// ------------------------------------------------------
// 13. NORMALIZAR TIPO DE DOCUMENTO
// ------------------------------------------------------

function normalizarTipoDocumentoFinanceiro(valor) {

  const tipo = String(
    valor || 'SEM DOCUMENTO'
  )
    .trim()
    .toUpperCase();

  const tiposPermitidos = [
    'SEM DOCUMENTO',
    'NOTA FISCAL',
    'RECIBO',
    'FATURA',
    'OUTRO'
  ];

  return tiposPermitidos.includes(tipo)
    ? tipo
    : 'OUTRO';
}


// ------------------------------------------------------
// 14. VALIDAR E PREPARAR PARCELAS RECEBIDAS
// ------------------------------------------------------

function prepararParcelasFinanceiro(
  conta,
  formaPagamento,
  valorTotal,
  totalParcelas
) {

  const valoresAutomaticos =
    dividirValorEmParcelasFinanceiro(
      valorTotal,
      totalParcelas
    );

  const parcelasRecebidas =
    Array.isArray(conta.parcelas)
      ? conta.parcelas
      : [];

  const parcelas = [];

  for (
    let i = 0;
    i < totalParcelas;
    i++
  ) {

    const parcelaRecebida =
      parcelasRecebidas[i] || {};

    let valorParcela =
      converterValorFinanceiro(
        parcelaRecebida.valor
      );

    // Se não foi informado valor individual,
    // usa a divisão automática.
    if (valorParcela <= 0) {
      valorParcela = valoresAutomaticos[i];
    }

    let vencimento = null;

    if (formaPagamento === 'BOLETO') {

      vencimento =
        converterDataFinanceiro(
          parcelaRecebida.vencimento
        );

      if (!vencimento) {
        throw new Error(
          'Informe o vencimento da parcela ' +
          (i + 1) +
          ' do boleto.'
        );
      }
    }

    if (
      formaPagamento === 'OUTRO' &&
      parcelaRecebida.vencimento
    ) {

      vencimento =
        converterDataFinanceiro(
          parcelaRecebida.vencimento
        );

      if (!vencimento) {
        throw new Error(
          'O vencimento informado na parcela ' +
          (i + 1) +
          ' é inválido.'
        );
      }
    }

    parcelas.push({
      numero: i + 1,
      valor: valorParcela,
      vencimento: vencimento
    });
  }

  // Validação da soma dos valores
  const somaCentavos = parcelas.reduce(
    function(total, parcela) {

      return total + Math.round(
        parcela.valor * 100
      );
    },
    0
  );

  const valorTotalCentavos = Math.round(
    valorTotal * 100
  );

  if (somaCentavos !== valorTotalCentavos) {

    throw new Error(
      'A soma das parcelas é diferente do valor total. ' +
      'Revise os valores antes de salvar.'
    );
  }

  return parcelas;
}


// ------------------------------------------------------
// 15. SALVAR NOVO LANÇAMENTO FINANCEIRO
// ------------------------------------------------------

function salvarContaFinanceira(conta) {

  if (!conta) {
    throw new Error(
      'Os dados do lançamento não foram recebidos.'
    );
  }

  /*
    Se existe ID, trata como edição individual.
  */
  if (conta.id) {
    return editarLancamentoFinanceiro(conta);
  }


  // DADOS BÁSICOS

  const descricao = String(
    conta.descricao || ''
  ).trim();

  const fornecedor = String(
    conta.fornecedor || ''
  ).trim();

  const categoria = String(
    conta.categoria || ''
  ).trim();

  const tipoDocumento =
    normalizarTipoDocumentoFinanceiro(
      conta.tipoDocumento
    );

  const numeroNota = String(
    conta.numeroNota || ''
  ).trim();

  const dataEmissao =
    conta.dataEmissao
      ? converterDataFinanceiro(
          conta.dataEmissao
        )
      : null;

  const formaPagamento =
    normalizarFormaPagamentoFinanceiro(
      conta.formaPagamento
    );

  const valorTotal =
    converterValorFinanceiro(
      conta.valorTotal
    );

  const observacoes = String(
    conta.observacoes || ''
  ).trim();


  // VALIDAÇÕES

  if (!descricao) {
    throw new Error(
      'Informe a descrição do lançamento.'
    );
  }

  if (valorTotal <= 0) {
    throw new Error(
      'Informe um valor total maior que zero.'
    );
  }


  // QUANTIDADE DE PARCELAS

  let totalParcelas = 1;

  if (
    formaPagamento === 'CREDITO' ||
    formaPagamento === 'BOLETO'
  ) {

    totalParcelas = Number(
      conta.totalParcelas || 1
    );

    if (
      !Number.isInteger(totalParcelas) ||
      totalParcelas < 1 ||
      totalParcelas > 120
    ) {
      throw new Error(
        'Informe uma quantidade de parcelas entre 1 e 120.'
      );
    }
  }


  if (formaPagamento === 'OUTRO') {

    const quantidadeInformada = Number(
      conta.totalParcelas || 1
    );

    totalParcelas =
      Number.isInteger(quantidadeInformada) &&
      quantidadeInformada >= 1 &&
      quantidadeInformada <= 120
        ? quantidadeInformada
        : 1;
  }


  // PREPARA AS PARCELAS

  const parcelas =
    prepararParcelasFinanceiro(
      conta,
      formaPagamento,
      valorTotal,
      totalParcelas
    );


  const aba = garantirAbaFinanceiro();

  const agora = new Date();

  const usuario =
    obterUsuarioFinanceiro();

  const grupoId =
    totalParcelas > 1
      ? gerarGrupoIdFinanceiro()
      : '';

  const linhas = [];

  const contasCriadas = [];


  parcelas.forEach(function(parcela) {

    const idConta =
      gerarIdContaFinanceira();

    let statusInicial = 'A PAGAR';

    let dataPagamento = '';


    // Pagamentos imediatos
    if (
      formaPagamento === 'DINHEIRO' ||
      formaPagamento === 'PIX' ||
      formaPagamento === 'DEBITO'
    ) {

      statusInicial = 'PAGA';

      dataPagamento =
        dataEmissao ||
        obterHojeFinanceiro();
    }


    // Boleto: vencido ou a pagar
    if (formaPagamento === 'BOLETO') {

      statusInicial =
        calcularStatusFinanceiro(
          formaPagamento,
          'A PAGAR',
          parcela.vencimento,
          null
        );
    }


    // Crédito sem banco/cartão vinculado
    if (formaPagamento === 'CREDITO') {
      statusInicial = 'A PAGAR';
    }


    // Outro pode receber status escolhido
    if (formaPagamento === 'OUTRO') {

      const statusInformado = String(
        conta.status || 'A PAGAR'
      )
        .trim()
        .toUpperCase();

      if (
        [
          'A PAGAR',
          'VENCIDA',
          'PAGA'
        ].includes(statusInformado)
      ) {
        statusInicial = statusInformado;
      }

      if (statusInicial === 'PAGA') {

        dataPagamento =
          conta.dataPagamento
            ? converterDataFinanceiro(
                conta.dataPagamento
              )
            : obterHojeFinanceiro();
      }
    }


    linhas.push([
      idConta,                    // A ID_CONTA
      grupoId,                    // B GRUPO_ID
      descricao,                  // C DESCRICAO
      fornecedor,                 // D FORNECEDOR
      categoria,                  // E CATEGORIA
      tipoDocumento,              // F TIPO_DOCUMENTO
      numeroNota,                 // G NUMERO_NOTA
      dataEmissao || '',          // H DATA_EMISSAO
      formaPagamento,             // I FORMA_PAGAMENTO
      valorTotal,                 // J VALOR_TOTAL
      parcela.valor,              // K VALOR_PARCELA
      parcela.numero,             // L PARCELA_ATUAL
      totalParcelas,              // M TOTAL_PARCELAS
      parcela.vencimento || '',   // N VENCIMENTO
      dataPagamento || '',        // O DATA_PAGAMENTO
      statusInicial,              // P STATUS
      observacoes,                // Q OBSERVACOES
      agora,                      // R DATA_CADASTRO
      agora,                      // S DATA_ATUALIZACAO
      usuario,                    // T USUARIO
      true                        // U ATIVO
    ]);


    contasCriadas.push({
      id: idConta,
      parcelaAtual: parcela.numero,
      totalParcelas: totalParcelas,
      valorParcela: parcela.valor,
      vencimento:
        parcela.vencimento
          ? formatarDataFinanceiroISO(
              parcela.vencimento
            )
          : '',
      status: statusInicial
    });
  });


  // GRAVA TODAS AS LINHAS DE UMA VEZ

  aba
    .getRange(
      aba.getLastRow() + 1,
      1,
      linhas.length,
      21
    )
    .setValues(linhas);


  // AUDITORIA

  registrarAuditoria(
    'Financeiro',
    'CADASTRO',
    'Novo lançamento financeiro: ' +
      descricao,
    grupoId || contasCriadas[0].id,
    null,
    {
      descricao: descricao,
      fornecedor: fornecedor,
      categoria: categoria,
      tipoDocumento: tipoDocumento,
      numeroNota: numeroNota,
      dataEmissao:
        dataEmissao
          ? formatarDataFinanceiroISO(
              dataEmissao
            )
          : '',
      formaPagamento: formaPagamento,
      valorTotal: valorTotal,
      totalParcelas: totalParcelas,
      parcelas: contasCriadas
    }
  );


  return {
    sucesso: true,
    mensagem:
      totalParcelas === 1
        ? 'Lançamento salvo com sucesso!'
        : totalParcelas +
          ' parcelas geradas com sucesso!',
    grupoId: grupoId,
    quantidadeCriada: totalParcelas,
    contas: contasCriadas
  };
}


// ------------------------------------------------------
// 16. LOCALIZAR LANÇAMENTO PELO ID
// ------------------------------------------------------

function localizarLancamentoFinanceiro(
  aba,
  idConta
) {

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const ids = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      1
    )
    .getValues()
    .flat();

  const indice = ids.findIndex(
    function(id) {

      return (
        String(id) ===
        String(idConta)
      );
    }
  );

  if (indice === -1) {
    return null;
  }

  const numeroLinha = indice + 2;

  const linha = aba
    .getRange(
      numeroLinha,
      1,
      1,
      21
    )
    .getValues()[0];

  return {
    numeroLinha: numeroLinha,
    linha: linha,
    conta:
      converterLinhaFinanceiroEmObjeto(
        linha
      )
  };
}


// ------------------------------------------------------
// 17. EDITAR UM LANÇAMENTO INDIVIDUAL
// ------------------------------------------------------

function editarLancamentoFinanceiro(conta) {

  if (!conta || !conta.id) {
    throw new Error(
      'ID do lançamento não informado.'
    );
  }

  const aba = garantirAbaFinanceiro();

  const localizado =
    localizarLancamentoFinanceiro(
      aba,
      conta.id
    );

  if (!localizado) {
    throw new Error(
      'Lançamento não encontrado.'
    );
  }

  const anterior = localizado.conta;

  if (!anterior.ativo) {
    throw new Error(
      'Não é possível editar um lançamento cancelado.'
    );
  }


  const descricao = String(
    conta.descricao !== undefined
      ? conta.descricao
      : anterior.descricao
  ).trim();

  const fornecedor = String(
    conta.fornecedor !== undefined
      ? conta.fornecedor
      : anterior.fornecedor
  ).trim();

  const categoria = String(
    conta.categoria !== undefined
      ? conta.categoria
      : anterior.categoria
  ).trim();

  const tipoDocumento =
    normalizarTipoDocumentoFinanceiro(
      conta.tipoDocumento !== undefined
        ? conta.tipoDocumento
        : anterior.tipoDocumento
    );

  const numeroNota = String(
    conta.numeroNota !== undefined
      ? conta.numeroNota
      : anterior.numeroNota
  ).trim();

  const dataEmissao =
    conta.dataEmissao
      ? converterDataFinanceiro(
          conta.dataEmissao
        )
      : (
          anterior.dataEmissao
            ? converterDataFinanceiro(
                anterior.dataEmissao
              )
            : null
        );

  const formaPagamento =
    normalizarFormaPagamentoFinanceiro(
      conta.formaPagamento !== undefined
        ? conta.formaPagamento
        : anterior.formaPagamento
    );

  const valorTotal =
    conta.valorTotal !== undefined &&
    conta.valorTotal !== ''
      ? converterValorFinanceiro(
          conta.valorTotal
        )
      : anterior.valorTotal;

  const valorParcela =
    conta.valorParcela !== undefined &&
    conta.valorParcela !== ''
      ? converterValorFinanceiro(
          conta.valorParcela
        )
      : anterior.valorParcela;

  const observacoes = String(
    conta.observacoes !== undefined
      ? conta.observacoes
      : anterior.observacoes
  ).trim();


  if (!descricao) {
    throw new Error(
      'Informe a descrição do lançamento.'
    );
  }

  if (valorTotal <= 0) {
    throw new Error(
      'O valor total deve ser maior que zero.'
    );
  }

  if (valorParcela <= 0) {
    throw new Error(
      'O valor da parcela deve ser maior que zero.'
    );
  }


  // VENCIMENTO

  let vencimento =
    anterior.vencimento
      ? converterDataFinanceiro(
          anterior.vencimento
        )
      : null;

  if (
    conta.vencimento !== undefined
  ) {

    vencimento =
      conta.vencimento
        ? converterDataFinanceiro(
            conta.vencimento
          )
        : null;
  }


  if (
    formaPagamento === 'BOLETO' &&
    !vencimento
  ) {
    throw new Error(
      'Informe o vencimento do boleto.'
    );
  }


  // PRESERVA DADOS INTERNOS

  const linhaAnterior =
    localizado.linha;

  const grupoId =
    linhaAnterior[1];

  const parcelaAtual =
    linhaAnterior[11];

  const totalParcelas =
    linhaAnterior[12];

  const dataPagamento =
    linhaAnterior[14];

  const dataCadastro =
    linhaAnterior[17];

  const usuario =
    obterUsuarioFinanceiro();

  const ativo =
    linhaAnterior[20];


  // RECALCULA STATUS

  const novoStatus =
    calcularStatusFinanceiro(
      formaPagamento,
      linhaAnterior[15],
      vencimento,
      dataPagamento
    );


  const novaLinha = [
    conta.id,
    grupoId,
    descricao,
    fornecedor,
    categoria,
    tipoDocumento,
    numeroNota,
    dataEmissao || '',
    formaPagamento,
    valorTotal,
    valorParcela,
    parcelaAtual,
    totalParcelas,
    vencimento || '',
    dataPagamento || '',
    novoStatus,
    observacoes,
    dataCadastro,
    new Date(),
    usuario,
    ativo
  ];


  aba
    .getRange(
      localizado.numeroLinha,
      1,
      1,
      21
    )
    .setValues([novaLinha]);


  const novaConta =
    converterLinhaFinanceiroEmObjeto(
      novaLinha
    );


  registrarAuditoria(
    'Financeiro',
    'EDIÇÃO',
    'Lançamento financeiro alterado: ' +
      descricao,
    conta.id,
    anterior,
    novaConta
  );


  return {
    sucesso: true,
    mensagem:
      'Lançamento atualizado com sucesso!',
    conta: novaConta
  };
}


// ------------------------------------------------------
// 18. MARCAR COMO PAGO
// ------------------------------------------------------

function marcarContaComoPaga(
  idConta,
  dataPagamento
) {

  if (!idConta) {
    throw new Error(
      'ID do lançamento não informado.'
    );
  }

  const aba = garantirAbaFinanceiro();

  const localizado =
    localizarLancamentoFinanceiro(
      aba,
      idConta
    );

  if (!localizado) {
    throw new Error(
      'Lançamento não encontrado.'
    );
  }

  const anterior = localizado.conta;

  if (!anterior.ativo) {
    throw new Error(
      'Não é possível pagar um lançamento cancelado.'
    );
  }

  if (anterior.status === 'PAGA') {
    throw new Error(
      'Este lançamento já está pago.'
    );
  }


  const data =
    dataPagamento
      ? converterDataFinanceiro(
          dataPagamento
        )
      : obterHojeFinanceiro();

  if (!data) {
    throw new Error(
      'Data de pagamento inválida.'
    );
  }


  aba
    .getRange(
      localizado.numeroLinha,
      15
    )
    .setValue(data);

  aba
    .getRange(
      localizado.numeroLinha,
      16
    )
    .setValue('PAGA');

  aba
    .getRange(
      localizado.numeroLinha,
      19
    )
    .setValue(new Date());

  aba
    .getRange(
      localizado.numeroLinha,
      20
    )
    .setValue(
      obterUsuarioFinanceiro()
    );


  const novaLinha = aba
    .getRange(
      localizado.numeroLinha,
      1,
      1,
      21
    )
    .getValues()[0];

  const novaConta =
    converterLinhaFinanceiroEmObjeto(
      novaLinha
    );


  registrarAuditoria(
    'Financeiro',
    'PAGAMENTO',
    'Lançamento marcado como pago: ' +
      anterior.descricao,
    idConta,
    anterior,
    novaConta
  );


  return {
    sucesso: true,
    mensagem:
      'Lançamento marcado como pago com sucesso!'
  };
}


// ------------------------------------------------------
// 19. DESFAZER PAGAMENTO
// ------------------------------------------------------

function reabrirContaFinanceira(idConta) {

  if (!idConta) {
    throw new Error(
      'ID do lançamento não informado.'
    );
  }

  const aba = garantirAbaFinanceiro();

  const localizado =
    localizarLancamentoFinanceiro(
      aba,
      idConta
    );

  if (!localizado) {
    throw new Error(
      'Lançamento não encontrado.'
    );
  }

  const anterior = localizado.conta;

  if (!anterior.ativo) {
    throw new Error(
      'O lançamento está cancelado.'
    );
  }

  if (anterior.status !== 'PAGA') {
    throw new Error(
      'Somente um lançamento pago pode ter o pagamento desfeito.'
    );
  }


  const novoStatus =
    calcularStatusFinanceiro(
      anterior.formaPagamento,
      'A PAGAR',
      anterior.vencimento,
      null
    );


  aba
    .getRange(
      localizado.numeroLinha,
      15
    )
    .clearContent();

  aba
    .getRange(
      localizado.numeroLinha,
      16
    )
    .setValue(novoStatus);

  aba
    .getRange(
      localizado.numeroLinha,
      19
    )
    .setValue(new Date());

  aba
    .getRange(
      localizado.numeroLinha,
      20
    )
    .setValue(
      obterUsuarioFinanceiro()
    );


  const novaLinha = aba
    .getRange(
      localizado.numeroLinha,
      1,
      1,
      21
    )
    .getValues()[0];

  const novaConta =
    converterLinhaFinanceiroEmObjeto(
      novaLinha
    );


  registrarAuditoria(
    'Financeiro',
    'REABERTURA',
    'Pagamento desfeito: ' +
      anterior.descricao,
    idConta,
    anterior,
    novaConta
  );


  return {
    sucesso: true,
    mensagem:
      'Pagamento desfeito com sucesso!'
  };
}


// ------------------------------------------------------
// 20. CANCELAR LANÇAMENTO
// ------------------------------------------------------

function cancelarContaFinanceira(idConta) {

  if (!idConta) {
    throw new Error(
      'ID do lançamento não informado.'
    );
  }

  const aba = garantirAbaFinanceiro();

  const localizado =
    localizarLancamentoFinanceiro(
      aba,
      idConta
    );

  if (!localizado) {
    throw new Error(
      'Lançamento não encontrado.'
    );
  }

  const anterior = localizado.conta;

  if (!anterior.ativo) {
    throw new Error(
      'Este lançamento já está cancelado.'
    );
  }

  if (anterior.status === 'PAGA') {
    throw new Error(
      'Um lançamento pago não pode ser cancelado. Desfaça o pagamento primeiro.'
    );
  }


  aba
    .getRange(
      localizado.numeroLinha,
      21
    )
    .setValue(false);

  aba
    .getRange(
      localizado.numeroLinha,
      19
    )
    .setValue(new Date());

  aba
    .getRange(
      localizado.numeroLinha,
      20
    )
    .setValue(
      obterUsuarioFinanceiro()
    );


  const novaLinha = aba
    .getRange(
      localizado.numeroLinha,
      1,
      1,
      21
    )
    .getValues()[0];

  const novaConta =
    converterLinhaFinanceiroEmObjeto(
      novaLinha
    );


  registrarAuditoria(
    'Financeiro',
    'CANCELAMENTO',
    'Lançamento financeiro cancelado: ' +
      anterior.descricao,
    idConta,
    anterior,
    novaConta
  );


  return {
    sucesso: true,
    mensagem:
      'Lançamento cancelado com sucesso!'
  };
}

// ======================================================
// FUNÇÃO TEMPORÁRIA - RECRIAR SOMENTE A ABA FINANCEIRO
// Executar apenas uma vez para migrar para a nova estrutura.
// ATENÇÃO: apaga todos os lançamentos existentes no Financeiro.
// ======================================================

function recriarAbaFinanceiroNovaEstrutura() {

  const planilha = getBancoDeDados();

  // Procura a aba Financeiro existente
  const abaExistente = planilha.getSheetByName('Financeiro');

  // Remove somente a aba Financeiro
  if (abaExistente) {
    planilha.deleteSheet(abaExistente);
  }

  // Cria uma nova aba Financeiro
  const aba = planilha.insertSheet('Financeiro');

  // Estrutura definitiva do novo módulo Financeiro
  const cabecalhos = [
    'ID_CONTA',
    'GRUPO_ID',
    'DESCRICAO',
    'FORNECEDOR',
    'CATEGORIA',
    'TIPO_DOCUMENTO',
    'NUMERO_NOTA',
    'DATA_EMISSAO',
    'FORMA_PAGAMENTO',
    'VALOR_TOTAL',
    'VALOR_PARCELA',
    'PARCELA_ATUAL',
    'TOTAL_PARCELAS',
    'VENCIMENTO',
    'DATA_PAGAMENTO',
    'STATUS',
    'OBSERVACOES',
    'DATA_CADASTRO',
    'DATA_ATUALIZACAO',
    'USUARIO',
    'ATIVO'
  ];

  // Insere os cabeçalhos
  aba
    .getRange(1, 1, 1, cabecalhos.length)
    .setValues([cabecalhos]);

  // Formata os cabeçalhos
  aba
    .getRange(1, 1, 1, cabecalhos.length)
    .setFontWeight('bold')
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');

  // Congela a primeira linha
  aba.setFrozenRows(1);

  // Formata as colunas de valores
  aba
    .getRange('J:K')
    .setNumberFormat('R$ #,##0.00');

  // Formata as colunas de datas simples
  aba
    .getRange('H:H')
    .setNumberFormat('dd/MM/yyyy');

  aba
    .getRange('N:O')
    .setNumberFormat('dd/MM/yyyy');

  // Formata data e hora de cadastro e atualização
  aba
    .getRange('R:S')
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Ajusta automaticamente a largura das colunas
  aba.autoResizeColumns(1, cabecalhos.length);

  // Registra a mudança na auditoria
  registrarAuditoria(
    'Financeiro',
    'REESTRUTURAÇÃO',
    'A aba Financeiro foi recriada com a nova estrutura de 21 colunas.',
    'FINANCEIRO-ESTRUTURA',
    null,
    {
      quantidadeColunas: cabecalhos.length,
      estrutura: cabecalhos
    }
  );

  return {
    sucesso: true,
    mensagem: 'Nova estrutura do Financeiro criada com sucesso!',
    quantidadeColunas: cabecalhos.length
  };
}


