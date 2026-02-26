// =====================================================================
// INSTRUÇÕES DE CONFIGURAÇÃO
// =====================================================================
//
// 1. Acesse https://script.google.com e crie um novo projeto
//
// 2. Cole TODO este código no editor (apague o conteúdo padrão)
//
// 3. No topo do código abaixo, altere SPREADSHEET_ID e FOLDER_ID:
//    - Crie uma Google Sheets nova → copie o ID da URL
//      (ex: https://docs.google.com/spreadsheets/d/ESTE_É_O_ID/edit)
//    - Crie uma pasta no Google Drive para as fotos → copie o ID da URL
//      (ex: https://drive.google.com/drive/folders/ESTE_É_O_ID)
//
// 4. Na planilha, crie os cabeçalhos na primeira linha:
//    A1: Data/Hora | B1: Empresa | C1: Recebedor | D1: VMix | E1: Fotos
//
// 5. Clique em "Implantar" → "Nova implantação"
//    - Tipo: "App da Web"
//    - Executar como: "Eu" (sua conta)
//    - Quem tem acesso: "Qualquer pessoa"
//    - Clique em "Implantar" e copie a URL gerada
//
// 6. Cole a URL no arquivo index.html na variável SCRIPT_URL
//
// 7. Compartilhe a PLANILHA com quem precisa visualizar os registros
//    (as pessoas que preenchem o formulário NÃO terão acesso)
//
// =====================================================================

// >>>>>> ALTERE ESTES VALORES <<<<<<
const SPREADSHEET_ID = '1ggGHnB1zleQPff-AGj2OAei_bUnaZsO8KLiFnzGyrI0';
const SHEET_NAME = 'Formulario Vmix';
const FOLDER_ID = '1wpQE3_qfK06oWR5C5lSEAJcHb3jGqMN0';

/**
 * Recebe os dados do formulário via POST
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const empresa   = data.empresa   || '';
    const recebedor = data.recebedor || '';
    const vmix      = data.vmix      || '';
    const fotos     = data.fotos     || [];

    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd_HH-mm-ss');
    const fotoLinks = [];

    // Salvar fotos no Google Drive (se FOLDER_ID estiver configurado)
    if (FOLDER_ID && FOLDER_ID !== 'COLE_O_ID_DA_PASTA_DO_DRIVE_AQUI') {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        
        fotos.forEach(function(base64, index) {
          // Remove o prefixo "data:image/jpeg;base64,"
          const base64Data = base64.replace(/^data:image\/\w+;base64,/, '');
          const blob = Utilities.newBlob(
            Utilities.base64Decode(base64Data),
            'image/jpeg',
            empresa.replace(/[^a-zA-Z0-9]/g, '_') + '_' + timestamp + '_foto' + (index + 1) + '.jpg'
          );

          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fotoLinks.push(file.getUrl());
        });
      } catch (driveError) {
        Logger.log('Erro ao salvar fotos: ' + driveError.toString());
        fotoLinks.push('ERRO: Não foi possível salvar as fotos - Verifique FOLDER_ID');
      }
    } else {
      fotoLinks.push('AVISO: FOLDER_ID não configurado - Fotos não foram salvas');
    }

    // Salvar na planilha
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Aba "' + SHEET_NAME + '" não encontrada na planilha!');
    }

    const dataHora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

    sheet.appendRow([
      dataHora,
      empresa,
      recebedor,
      vmix,
      fotoLinks.join('\n')
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', fotos: fotoLinks.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('ERRO COMPLETO: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Necessário para que o script aceite requisições GET (teste)
 */
function doGet(e) {
  return ContentService
    .createTextOutput('Formulário VMix - Endpoint ativo!')
    .setMimeType(ContentService.MimeType.TEXT);
}
