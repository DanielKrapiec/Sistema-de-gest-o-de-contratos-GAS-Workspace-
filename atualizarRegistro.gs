function atualizarRegistro() {
  var html = HtmlService
    .createHtmlOutputFromFile('painelAtualizarRegistro')
    .setTitle('inserir o codigo que deseja recuperar')
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi() // Ordena a abertura da caixa de diálogo
    .showModalDialog(html, 'Insira o contrato que deseja atualizar');
}
function buscarContrato1(codigoCerto, codigo) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var menu = planilha.getSheetByName('Menu');
  var arquivo = planilha.getSheetByName('Arquivo');
  var dados = arquivo.getDataRange().getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][6] == codigoCerto) { // Índice 6 corresponde à coluna G
      var valorColunaC = dados[i][2]; // Coluna C = índice 2
      //var valorColunaG = codigo;
      var valoresHN = dados[i].slice(7, 14);
      var valoresHNVertical = valoresHN.map(function(valor) {
        return [valor];
      });
      var valoreColunaP = dados[i][15]
      var dataInicio = dados[i][16]
      menu.getRange('C14').setValue(dataInicio);
      menu.getRange('B1').setValue(valorColunaC);
      //menu.getRange('B5').setValue(valorColunaG) //o codigo é gerado automatiamente;
      menu.getRange('B6:B12').setValues(valoresHNVertical);
      menu.getRange('B14').setValue(valoreColunaP);
      menu.getRange('C1').setValue(new Date);
      if (dataInicio) {
        menu.getRange('B15').setValue("SIM");
      };
      Browser.msgBox('Informações localizadas');
      return;
    }
  }
  Browser.msgBox('Contrato não encontrado');
}
