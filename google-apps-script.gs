// ===== CODIGO ATUALIZADO DO GOOGLE APPS SCRIPT =====
// Suporta: Entrega/Devolucao (aba KitMojo), Solicitacoes (aba Solicitacoes), Edicao, Renovacao e Exclusao
// + Formulario Vmix (aba "Formulario Vmix") - entrega/atualizacao de localizacao do equipamento VMix
//
// IMPORTANTE: Apos colar, faca NOVA IMPLANTACAO

// >>>>>> Configuracao do Formulario Vmix <<<<<<
var VMIX_SHEET_NAME = 'Formulario Vmix';
var VMIX_FOLDER_ID = '1wpQE3_qfK06oWR5C5lSEAJcHb3jGqMN0';

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById('1ggGHnB1zleQPff-AGj2OAei_bUnaZsO8KLiFnzGyrI0');

    // === DELETAR POR SOLICITACAO ID (limpeza de entradas indesejadas) ===
    if (dados.acao === 'DELETAR_POR_SOL_ID') {
      return handleDeletarPorSolId(ss, dados);
    }

    // === SOLICITACAO ===
    if (dados.acao === 'SOLICITACAO') {
      return handleSolicitacao(ss, dados);
    }

    // === ATENDER SOLICITACAO ===
    if (dados.acao === 'ATENDER_SOLICITACAO') {
      return handleAtenderSolicitacao(ss, dados);
    }

    // === FINALIZAR SOLICITACAO ===
    if (dados.acao === 'FINALIZAR_SOLICITACAO') {
      return handleFinalizarSolicitacao(ss, dados);
    }

    // === RENOVAR KIT (atualiza projeto, mantém datas) ===
    if (dados.acao === 'RENOVAR_KIT') {
      return handleRenovarKit(ss, dados);
    }

    // === EXCLUIR SOLICITACAO ===
    if (dados.acao === 'EXCLUIR_SOLICITACAO') {
      return handleExcluirSolicitacao(ss, dados);
    }

    // === EDITAR REGISTRO ===
    if (dados.action === 'editRegistro') {
      return handleEditRegistro(ss, dados);
    }

    // === FORMULARIO VMIX (entrega/atualizacao de localizacao do equipamento VMix) ===
    // Identificado pelo payload do index.html/dashboard.html do formulario-vmix:
    // { empresa, recebedor, vmix, fotos, origem?, observacao? } — sem acao/action/tipo.
    if (dados.acao === undefined && dados.action === undefined && dados.tipo === undefined && dados.vmix !== undefined) {
      return handleFormularioVmix(ss, dados);
    }

    // === ENTREGA / DEVOLUCAO (KitMojo) ===
    return handleKitMojo(ss, dados);

  } catch (err) {
    Logger.log('ERRO: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// === HANDLER: FORMULARIO VMIX (entrega/atualizacao) ===
function handleFormularioVmix(ss, dados) {
  var sheet = ss.getSheetByName(VMIX_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(VMIX_SHEET_NAME);
    sheet.appendRow(['Data/Hora', 'Empresa', 'Recebedor', 'VMix', 'Fotos', 'Observação']);
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#ffffff');
  }

  var empresa    = dados.empresa    || '';
  var recebedor  = dados.recebedor  || '';
  var vmix       = dados.vmix       || '';
  var fotos      = dados.fotos      || [];
  var origem     = dados.origem     || 'formulario';
  var observacao = dados.observacao || '';

  var timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd_HH-mm-ss');
  var fotoLinks = [];

  if (fotos.length > 0 && VMIX_FOLDER_ID) {
    try {
      var folder = DriveApp.getFolderById(VMIX_FOLDER_ID);

      fotos.forEach(function(base64, index) {
        var base64Data = base64.replace(/^data:image\/\w+;base64,/, '');
        var blob = Utilities.newBlob(
          Utilities.base64Decode(base64Data),
          'image/jpeg',
          empresa.replace(/[^a-zA-Z0-9]/g, '_') + '_' + timestamp + '_foto' + (index + 1) + '.jpg'
        );

        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoLinks.push(file.getUrl());
      });
    } catch (driveError) {
      Logger.log('Erro ao salvar fotos do VMix: ' + driveError.toString());
      fotoLinks.push('ERRO: Não foi possível salvar as fotos - Verifique VMIX_FOLDER_ID');
    }
  }

  var dataHora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

  var observacaoFinal = origem === 'dashboard'
    ? ('[Atualizado pelo dashboard] ' + observacao).trim()
    : observacao;

  sheet.appendRow([
    dataHora,
    empresa,
    recebedor,
    vmix,
    fotoLinks.join('\n'),
    observacaoFinal
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', fotos: fotoLinks.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: DADOS DO FORMULARIO VMIX (usado pelo dashboard.html) ===
function handleFormularioVmixDados(ss) {
  var sheet = ss.getSheetByName(VMIX_SHEET_NAME);

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', records: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var values = sheet.getDataRange().getValues();
  var records = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[3]) continue; // sem VMix preenchido, ignora a linha

    records.push({
      dataHora: formatCellValueVmix(row[0]),
      empresa: row[1] || '',
      recebedor: row[2] || '',
      vmix: row[3] || '',
      fotos: row[4] || '',
      observacao: row[5] || ''
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', records: records }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatCellValueVmix(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  }
  return value || '';
}

// === HANDLER: EDITAR REGISTRO ===
function handleEditRegistro(ss, dados) {
  var sheet = ss.getSheetByName('KitMojo');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Aba KitMojo nao encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getDisplayValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    var dataHoraCelula = String(data[i][0]).trim();
    var dataHoraOriginal = String(dados.dataHoraOriginal || '').trim();

    if (dataHoraCelula === dataHoraOriginal) {
      // Colunas A-L (1-12) + N (14) para Local; coluna M (13) = Foto, nao mexe
      sheet.getRange(i + 1, 1).setValue(dados.dataHora || data[i][0]);
      sheet.getRange(i + 1, 2).setValue(dados.tipo || data[i][1]);
      sheet.getRange(i + 1, 3).setValue(dados.nome1 || data[i][2]);
      sheet.getRange(i + 1, 4).setValue(dados.nome2 || data[i][3]);
      sheet.getRange(i + 1, 5).setValue(dados.tipoKit || data[i][4]);
      sheet.getRange(i + 1, 6).setValue(dados.iphone || data[i][5]);
      sheet.getRange(i + 1, 7).setValue(dados.android || data[i][6]);
      sheet.getRange(i + 1, 8).setValue(dados.portSrt || data[i][7]);
      sheet.getRange(i + 1, 9).setValue(dados.tipoChip || data[i][8]);
      sheet.getRange(i + 1, 10).setValue(dados.itensConferidos || data[i][9]);
      sheet.getRange(i + 1, 11).setValue(dados.itensFaltando || data[i][10]);
      sheet.getRange(i + 1, 12).setValue(dados.observacao !== undefined ? dados.observacao : data[i][11]);
      // col 13 = Foto (URL) — nao altera
      sheet.getRange(i + 1, 14).setValue(dados.local !== undefined ? dados.local : (data[i][13] || ''));
      // col 15 = SolicitacaoID — nao altera durante edicao
      sheet.getRange(i + 1, 16).setValue(dados.numeroIphone !== undefined ? dados.numeroIphone : (data[i][15] || ''));
      sheet.getRange(i + 1, 17).setValue(dados.numeroAndroid !== undefined ? dados.numeroAndroid : (data[i][16] || ''));
      found = true;
      break;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: found ? 'ok' : 'not_found' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: DELETAR POR SOLICITACAO ID ===
// Remove todas as linhas da aba KitMojo onde col15 (SolicitacaoID) = dados.solicitacaoId
function handleDeletarPorSolId(ss, dados) {
  var sheet = ss.getSheetByName('KitMojo');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Aba KitMojo nao encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var solId = String(dados.solicitacaoId || '').trim();
  if (!solId) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'solicitacaoId ausente' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var removidos = 0;

  // Itera de trás pra frente para não deslocar índices ao deletar
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][14]).trim() === solId) {   // col15 (índice 14) = SolicitacaoID
      sheet.deleteRow(i + 1);
      removidos++;
    }
  }

  Logger.log('DELETAR_POR_SOL_ID: ' + solId + ' → ' + removidos + ' linhas removidas');

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', removidos: removidos }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: EXCLUIR SOLICITACAO ===
// Remove linha da aba Solicitacoes onde col10 (ID) = dados.solicitacaoId
function handleExcluirSolicitacao(ss, dados) {
  var sheet = ss.getSheetByName('Solicitacoes');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Aba Solicitacoes nao encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var solId = String(dados.solicitacaoId || '').trim();
  if (!solId) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'solicitacaoId ausente' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var removidos = 0;

  // Itera de trás pra frente para não deslocar índices ao deletar
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][9]).trim() === solId) {  // col10 (índice 9) = ID
      sheet.deleteRow(i + 1);
      removidos++;
    }
  }

  Logger.log('EXCLUIR_SOLICITACAO: ' + solId + ' → ' + removidos + ' linha(s) removida(s)');

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', removidos: removidos }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: ENTREGA / DEVOLUCAO ===
function handleKitMojo(ss, dados) {
  var sheet = ss.getSheetByName('KitMojo');

  if (!sheet) {
    sheet = ss.insertSheet('KitMojo');
    sheet.appendRow([
      'Data/Hora', 'Tipo', 'Nome 1', 'Nome 2', 'Tipo KitMojo',
      'Modelo iPhone', 'Modelo Android', 'Port SRT', 'Tipo Chip',
      'Itens Conferidos', 'Itens Faltando', 'Observacao', 'Foto (URL)', 'Local', 'SolicitacaoID',
      'Numero iPhone', 'Numero Android'
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 17);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#ffffff');
  }

  var fotoUrl = '';
  if (dados.foto && dados.foto.indexOf('data:') === 0) {
    try {
      var folder = DriveApp.getFolderById('1PIcLHBuoPjDH77atG2pcHN2Ws3gxzujV');
      var base64Data = dados.foto.split(',')[1];
      var decoded = Utilities.base64Decode(base64Data);
      var nomeArquivo = 'KitMojo_' + dados.tipo + '_' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd_HHmmss') + '.jpg';
      var blob = Utilities.newBlob(decoded, 'image/jpeg', nomeArquivo);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoUrl = file.getUrl();
    } catch (fotoErr) {
      fotoUrl = 'ERRO_FOTO: ' + fotoErr.toString();
    }
  }

  sheet.appendRow([
    dados.dataHora || new Date().toLocaleString('pt-BR'),
    dados.tipo || '',
    dados.nome1 || '',
    dados.nome2 || '',
    dados.tipoKit || '',
    dados.iphone || '',
    dados.android || '',
    dados.portSrt || '',
    dados.tipoChip || '',
    dados.itensConferidos || '',
    dados.itensFaltando || '',
    dados.observacao || '',
    fotoUrl,
    dados.local || '',
    dados.solicitacaoId || '',
    dados.numeroIphone || '',
    dados.numeroAndroid || ''
  ]);

  if (dados.tipo === 'ENTREGA' && dados.solicitacaoId) {
    try {
      marcarSolicitacaoAtendida(ss, dados.solicitacaoId);
    } catch (e) {
      Logger.log('Erro ao atender solicitacao: ' + e.toString());
    }
  }

  // Devolucao com solicitacaoId vinculado → finaliza automaticamente
  if ((dados.tipo === 'DEVOLUÇÃO' || dados.tipo === 'DEVOLUCAO') && dados.solicitacaoId) {
    try {
      marcarSolicitacaoFinalizada(ss, dados.solicitacaoId);
    } catch (e) {
      Logger.log('Erro ao finalizar solicitacao: ' + e.toString());
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', fotoUrl: fotoUrl }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: SOLICITACAO ===
// Estrutura da aba Solicitacoes (10 colunas):
// col1=Data Solicitação, col2=Quando, col3=Onde, col4=Quem, col5=Projeto,
// col6=Observação, col7=Status (Pendente/Atendida), col8=Data Entrega,
// col9=Data Devolução, col10=ID
function handleSolicitacao(ss, dados) {
  var sheet = ss.getSheetByName('Solicitacoes');

  if (!sheet) {
    sheet = ss.insertSheet('Solicitacoes');
    sheet.appendRow([
      'Data Solicitação', 'Quando', 'Onde', 'Quem', 'Projeto',
      'Observação', 'Status (Pendente/Atendida)', 'Data Entrega', 'Data Devolução', 'ID'
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 10);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#ffffff');
  }

  var id = 'SOL-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMddHHmmss');

  sheet.appendRow([
    dados.dataHora || new Date().toLocaleString('pt-BR'), // col1 = Data Solicitação
    dados.quando || '',                                    // col2 = Quando
    dados.onde || '',                                      // col3 = Onde
    dados.quem || '',                                      // col4 = Quem
    dados.projeto || '',                                   // col5 = Projeto
    dados.observacao || '',                                // col6 = Observação
    'Pendente',                                            // col7 = Status
    '',                                                    // col8 = Data Entrega
    '',                                                    // col9 = Data Devolução
    id                                                     // col10 = ID
  ]);

  // Notifica o Slack
  try {
    var slackMsg = '*🎯 Nova Solicitação de KitMojo*\n'
      + '*👤 Quem:* ' + (dados.quem || '—') + '\n'
      + '*📅 Quando:* ' + (dados.quando || '—') + '\n'
      + '*📍 Onde:* ' + (dados.onde || '—') + '\n'
      + '*🎬 Projeto:* ' + (dados.projeto || '—') + '\n'
      + (dados.observacao ? '*📝 Obs:* ' + dados.observacao + '\n' : '')
      + '*🕐 Solicitado em:* ' + (dados.dataHora || new Date().toLocaleString('pt-BR'));

    var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
    if (!webhookUrl) return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', id: id }))
      .setMimeType(ContentService.MimeType.JSON);
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: slackMsg })
    });
  } catch (slackErr) {
    Logger.log('Erro Slack: ' + slackErr.toString());
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: ATENDER SOLICITACAO ===
function handleAtenderSolicitacao(ss, dados) {
  marcarSolicitacaoAtendida(ss, dados.solicitacaoId);

  // Registra entrada minima no KitMojo para aparecer no dashboard como "Em Uso"
  var sheet = ss.getSheetByName('KitMojo');
  if (sheet && dados.quem) {
    var obs = (dados.projeto || '') + (dados.quando ? ' · ' + dados.quando : '');
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy, HH:mm:ss'),
      'ENTREGA',
      'Confirmado via Solicitacao',
      dados.quem || '',
      'KitMojo',
      '', '', '', '',
      '', '',
      obs,
      '',
      dados.onde || '',
      dados.solicitacaoId || ''
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: FINALIZAR SOLICITACAO ===
function handleFinalizarSolicitacao(ss, dados) {
  marcarSolicitacaoFinalizada(ss, dados.solicitacaoId);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === HANDLER: RENOVAR KIT ===
// Marca a nova solicitação como Atendida e atualiza o projeto na solicitação ativa.
// NÃO altera a data de entrega original — o tempo "Fora há" continua acumulando.
function handleRenovarKit(ss, dados) {
  var sheet = ss.getSheetByName('Solicitacoes');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Aba Solicitacoes nao encontrada' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy, HH:mm:ss');

  // 1. Marcar nova solicitação como Atendida (sem mexer na data de entrega original do kit)
  if (dados.solicitacaoIdNova) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][9] === dados.solicitacaoIdNova) {
        sheet.getRange(i + 1, 7).setValue('Atendida');  // col7 = Status
        sheet.getRange(i + 1, 8).setValue(agora);       // col8 = Data Entrega
        break;
      }
    }
  }

  // 2. Atualizar o projeto na solicitação ativa (a que está vinculada ao kit em uso)
  if (dados.solicitacaoIdAtiva && dados.novoProjeto !== undefined) {
    for (var j = 1; j < data.length; j++) {
      if (data[j][9] === dados.solicitacaoIdAtiva) {
        sheet.getRange(j + 1, 5).setValue(dados.novoProjeto); // col5 = Projeto
        break;
      }
    }
  }

  // 3. Atualizar observação na aba KitMojo para refletir o novo projeto
  if (dados.quem && dados.novoProjeto) {
    var kitSheet = ss.getSheetByName('KitMojo');
    if (kitSheet) {
      var kitData = kitSheet.getDataRange().getValues();
      // Encontrar a última ENTREGA para essa pessoa (mais recente = última linha)
      var lastRow = -1;
      for (var k = 1; k < kitData.length; k++) {
        if (kitData[k][1] === 'ENTREGA' && kitData[k][3] && kitData[k][3].toString().toLowerCase().trim() === dados.quem.toLowerCase().trim()) {
          lastRow = k;
        }
      }
      if (lastRow >= 0) {
        var obsAtual = kitData[lastRow][11] || '';
        var novaObs = dados.novoProjeto + (obsAtual ? ' (anterior: ' + obsAtual + ')' : '');
        kitSheet.getRange(lastRow + 1, 12).setValue(novaObs); // col12 = Observacao
      }
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Kit renovado com sucesso' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === MARCAR SOLICITACAO COMO ATENDIDA ===
function marcarSolicitacaoAtendida(ss, solicitacaoId) {
  var sheet = ss.getSheetByName('Solicitacoes');
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === solicitacaoId) {          // col10 (índice 9) = ID
      sheet.getRange(i + 1, 7).setValue('Atendida');                        // col7 = Status
      sheet.getRange(i + 1, 8).setValue(new Date().toLocaleString('pt-BR')); // col8 = Data Entrega
      break;
    }
  }
}

// === MARCAR SOLICITACAO COMO FINALIZADA ===
function marcarSolicitacaoFinalizada(ss, solicitacaoId) {
  var sheet = ss.getSheetByName('Solicitacoes');
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === solicitacaoId) {          // col10 (índice 9) = ID
      sheet.getRange(i + 1, 7).setValue('Finalizada');                       // col7 = Status
      sheet.getRange(i + 1, 9).setValue(new Date().toLocaleString('pt-BR')); // col9 = Data Devolução
      break;
    }
  }
}

// === HANDLER: KITS EM USO (para pedidos-cobertura) ===
// Lê a aba KitMojo, processa pares ENTREGA/DEVOLUÇÃO e retorna apenas kits em uso.
function handleKitsEmUso(ss) {
  var sheet = ss.getSheetByName('KitMojo');
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var kits = {};

  // Processa linha a linha em ordem cronológica (planilha já está em ordem)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var tipo       = String(row[1] || '').trim().toUpperCase();
    var nome1      = String(row[2] || '').trim();  // entregador / devolvedor
    var nome2      = String(row[3] || '').trim();  // portador
    var tipoKit    = String(row[4] || 'Kit').trim();
    var iphone     = String(row[5] || '?').trim();
    var android    = String(row[6] || '?').trim();
    var observacao = String(row[11] || '').trim();
    var local      = String(row[13] || '').trim();
    var solId      = String(row[14] || '').trim();
    var dataHora   = String(row[0] || '').trim();

    if (tipo === 'ENTREGA') {
      var key = solId || (tipoKit + '|' + nome2);
      kits[key] = { comQuem: nome2, tipoKit: tipoKit, iphone: iphone, android: android,
                    local: local, saiu: dataHora, obs: observacao, solId: solId };
    } else if (tipo === 'DEVOLUÇÃO' || tipo === 'DEVOLUCAO') {
      // Casa pelo solicitacaoId primeiro, depois pela chave de specs+pessoa
      var matched = null;
      if (solId && kits[solId]) {
        matched = solId;
      } else {
        var keyD = tipoKit + '|' + nome1;
        if (kits[keyD]) matched = keyD;
      }
      if (matched) delete kits[matched];
    }
  }

  var resultado = Object.keys(kits).map(function(k) { return kits[k]; });
  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

// === doGet ===
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById('1ggGHnB1zleQPff-AGj2OAei_bUnaZsO8KLiFnzGyrI0');
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

    if (action === 'kitsEmUso') {
      return handleKitsEmUso(ss);
    }

    if (action === 'solicitacoesPendentes') {
      var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
      var sheet = ss.getSheetByName('Solicitacoes');
      if (!sheet) {
        var emptyJson = JSON.stringify([]);
        return ContentService
          .createTextOutput(callback ? callback + '(' + emptyJson + ')' : emptyJson)
          .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
      }
      var data = sheet.getDataRange().getDisplayValues();
      var resultado = [];
      for (var i = 1; i < data.length; i++) {
        resultado.push({
          id: data[i][9],               // col10 = ID
          dataSolicitacao: data[i][0],  // col1  = Data Solicitação
          quando: data[i][1],           // col2  = Quando
          onde: data[i][2],             // col3  = Onde
          quem: data[i][3],             // col4  = Quem
          projeto: data[i][4],          // col5  = Projeto
          observacao: data[i][5],       // col6  = Observação
          status: data[i][6],           // col7  = Status
          dataEntrega: data[i][7],      // col8  = Data Entrega
          dataDevolucao: data[i][8]     // col9  = Data Devolução
        });
      }
      var json = JSON.stringify(resultado);
      return ContentService
        .createTextOutput(callback ? callback + '(' + json + ')' : json)
        .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
    }

    // === FORMULARIO VMIX (dados para o dashboard.html) ===
    if (action === 'formularioVmix') {
      return handleFormularioVmixDados(ss);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testeManual() {
  var resultado = doPost({
    postData: {
      contents: JSON.stringify({
        tipo: 'ENTREGA', nome1: 'Teste', nome2: 'Teste2',
        tipoKit: 'LiveMode', portSrt: '9000', tipoChip: 'Fisico',
        iphone: 'iPhone 16', android: 'Samsung',
        itensConferidos: 'SmartRig Completo', itensFaltando: 'Espuma',
        observacao: 'Teste', local: 'Estudio', foto: '',
        dataHora: new Date().toLocaleString('pt-BR')
      })
    }
  });
  Logger.log(resultado.getContent());
}

function testeSolicitacao() {
  var resultado = doPost({
    postData: {
      contents: JSON.stringify({
        acao: 'SOLICITACAO',
        quando: 'Sabado, 14/03',
        onde: 'Sao Paulo',
        quem: 'Alan (Fala, Porco)',
        projeto: 'Brasileirao',
        observacao: '',
        dataHora: new Date().toLocaleString('pt-BR')
      })
    }
  });
  Logger.log(resultado.getContent());
}

function testeFormularioVmix() {
  var resultado = doPost({
    postData: {
      contents: JSON.stringify({
        empresa: 'Multivídeo',
        recebedor: 'Teste',
        vmix: 'VMix 1',
        fotos: []
      })
    }
  });
  Logger.log(resultado.getContent());
}

function testeSlack() {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
  Logger.log('Webhook URL: ' + webhookUrl);

  if (!webhookUrl) {
    Logger.log('ERRO: propriedade SLACK_WEBHOOK nao encontrada!');
    return;
  }

  try {
    var resp = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '✅ Teste do KitMojo Bot — está funcionando!' })
    });
    Logger.log('Status: ' + resp.getResponseCode());
    Logger.log('Resposta: ' + resp.getContentText());
  } catch (err) {
    Logger.log('ERRO: ' + err.toString());
  }
}

function autorizarDrive() {
  DriveApp.getRootFolder().getName();
  var t = DriveApp.createFile('teste.txt', 'teste', 'text/plain');
  t.setTrashed(true);
  var folder = DriveApp.getFolderById('1PIcLHBuoPjDH77atG2pcHN2Ws3gxzujV');
  var t2 = folder.createFile('teste2.txt', 'teste', 'text/plain');
  t2.setTrashed(true);
  Logger.log('Drive OK!');
}
