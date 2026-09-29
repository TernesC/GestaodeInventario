// ID da sua planilha no Google Sheets:
var ID_DA_PLANILHA = "";

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(ID_DA_PLANILHA);
  } catch (e) {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) return activeSs;
    throw new Error("Erro ao acessar a planilha.");
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setTitle('Gestão de Inventário')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Retorna os itens incluindo o campo de Observação
function getInventario(categoria) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var isOutros = (categoria === 'OUTROS');
  var numCols = isOutros ? 4 : 7; // Garante a leitura até a coluna G (7)
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][0] !== null) {
      
      // Trata a observação para nunca ser 'undefined'
      var obsBruta = isOutros ? data[i][3] : data[i][6];
      var obsTratada = (obsBruta !== undefined && obsBruta !== null) ? String(obsBruta) : '';
      if (obsTratada.toLowerCase() === 'undefined') obsTratada = '';

      result.push({
        rowIndex: i + 2,
        nome: String(data[i][0]),
        quantidade: Number(data[i][1]) || 0,
        localizacao: String(data[i][2] || '-'),
        serial: isOutros ? '-' : String(data[i][3] || '-'),
        patrimonio: isOutros ? '-' : String(data[i][4] || '-'),
        danificado: isOutros ? 'N/A' : String(data[i][5] || 'NÃO'),
        observacao: obsTratada
      });
    }
  }
  return result;
}

// Altera a quantidade (+1 ou -1)
function alterarQuantidade(categoria, rowIndex, delta) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  if (!sheet) return { success: false };
  var cell = sheet.getRange(rowIndex, 2);
  var val = (Number(cell.getValue()) || 0) + delta;
  cell.setValue(val < 0 ? 0 : val);
  return { success: true };
}

// Altera a localização
function alterarLocalizacao(categoria, rowIndex, novaLoc) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  if (!sheet) return { success: false };
  sheet.getRange(rowIndex, 3).setValue(novaLoc);
  return { success: true };
}

// Altera o status danificado
function alterarDanificado(categoria, rowIndex, novoStatus) {
  if (categoria === 'OUTROS') return { success: false };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  if (!sheet) return { success: false };
  sheet.getRange(rowIndex, 6).setValue(novoStatus);
  return { success: true };
}

// FUNÇÃO NOVA: Grava a alteração da observação diretamente na célula da planilha
function alterarObservacao(categoria, rowIndex, novaObs) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  if (!sheet) return { success: false };
  
  // Grava na Coluna D (4) para a aba OUTROS ou Coluna G (7) para as restantes abas
  var colObs = (categoria === 'OUTROS') ? 4 : 7; 
  sheet.getRange(rowIndex, colObs).setValue(novaObs);
  return { success: true };
}

// Adiciona novo item incluindo o parâmetro de observação
function adicionarItem(categoria, nome, quantidade, localizacao, serial, patrimonio, danificado, observacao) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(categoria) || (categoria === 'DESKTOPS' ? ss.getSheetByName('DESKTOP') : null);
  
  if (!sheet) {
    sheet = ss.insertSheet(categoria);
    sheet.appendRow(categoria === 'OUTROS' ? 
      ['NOME DO ITEM', 'QUANTIDADE', 'CAIXA/LOCALIZAÇÃO', 'OBSERVAÇÃO'] : 
      ['NOME DO ITEM', 'QUANTIDADE', 'CAIXA/LOCALIZAÇÃO', 'SERIAL', 'PATRIMÔNIO', 'DANIFICADO', 'OBSERVAÇÃO']);
  }
  
  if (categoria === 'OUTROS') {
    sheet.appendRow([nome, Number(quantidade), localizacao, observacao || '']);
  } else {
    sheet.appendRow([nome, Number(quantidade), localizacao || '-', serial || '-', patrimonio || '-', danificado || 'NÃO', observacao || '']);
  }
  
  return { success: true };
}