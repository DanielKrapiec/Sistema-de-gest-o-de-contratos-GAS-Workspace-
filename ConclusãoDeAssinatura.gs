function buscarEAdicionar() { //Este codigo gera um pop-up para informar a data de conclusão de assinatura de um contrato
  var html = HtmlService
    .createHtmlOutputFromFile('conclusãoDeAssinatura')
    .setTitle('Inserir Código e Data')
    .setWidth(600)
    .setHeight(180);
  SpreadsheetApp.getUi() // Ordena a abertura da caixa de diálogo
    .showModalDialog(html, 'Inserir Dados');
}
function processarDados(codigo, data) {
  var planilha = SpreadsheetApp.getActive().getSheetByName("Arquivo");
  // Busca o código na coluna F
  var dados = planilha.getDataRange().getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][6] == codigo) { // Índice 6 corresponde à coluna G
      // Adiciona a data na coluna R (índice 17)
      planilha.getRange(i+1, 18).setValue(data);
      Browser.msgBox('✅Código encontrado e data adicionada!✅');
      return;
    }
  }
  // Se o código não for encontrado
  Browser.msgBox('❌Código não encontrado.❌');
}
