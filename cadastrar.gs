function informarDados() {
  var html = HtmlService
    .createHtmlOutputFromFile('cadastrarDados')
    .setTitle('Cadastrar dados do aluno')
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi() // Ordena a abertura da caixa de diálogo
    .showModalDialog(html, 'Cadastrar dados do aluno');
}
function gravarDados1(ra, nome, curso, concedente) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Base de Dados');
  var colunaF = planilha.getRange("G:G").getValues();
  var colunaE = planilha.getRange("F:F").getValues();
  var proximaLinhaG = 1;
  for (var i = 0; i < colunaF.length; i++) {
    if (!colunaF[i][0]) { // Se estiver vazia
      proximaLinhaG = i + 1;
      break;
    }
  }
  if (concedente != "") {
    var proximaLinhaF = 1;
    for (var i = 0; i < colunaE.length; i++) {
      if (!colunaE[i][0]) { // Se estiver vazia
        proximaLinhaF = i + 1;
        break;
      }
    }
    planilha.getRange(proximaLinhaF, 5).setValue(concedente);
  }
  planilha.getRange(proximaLinhaG, 7).setValue(ra);
  planilha.getRange(proximaLinhaG, 8).setValue(nome);
  planilha.getRange(proximaLinhaG, 9).setValue(curso);
  Browser.msgBox('Dados Cadastrados');
  return;
}
