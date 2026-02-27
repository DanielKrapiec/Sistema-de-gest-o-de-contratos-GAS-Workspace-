function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Relatório').addItem('Relatório', 'gerarRelatorio').addToUi();
}

function gerarRelatorio() {
  var html = HtmlService
    .createHtmlOutputFromFile('gerarRelatorio')
    .setTitle('Gerar relatório do periodo')
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Inserir dados');
}

function procurarDados(dataInicio, dataFim) {
  var planilha = SpreadsheetApp.openById("[ID a planilha]").getSheetByName('Arquivo');
  var colunaB = planilha.getRange("B:B").getValues();

  var dataFimFormatada = new Date(dataFim);
  dataFimFormatada = new Date(dataFimFormatada.getTime() + 1000 * 60 * 60 * 3);

  var dataInicioFormatada = new Date(dataInicio);
  dataInicioFormatada = new Date(dataInicioFormatada.getTime() + 1000 * 60 * 60 * 3);

  if (dataFimFormatada < dataInicioFormatada) {
    Browser.msgBox("A data final não pode ser anterior a data inicial.");
  } else {

    var indice = colunaB.findIndex(row => row[0] == Utilities.formatDate(dataFimFormatada, Session.getScriptTimeZone(), "dd/MM/yyyy"));

    var diasParaSubtrair = 1;

    while (indice === -1 && diasParaSubtrair <= 30) {
      var dataAnterior = new Date(dataFimFormatada);
      dataAnterior.setDate(dataAnterior.getDate() - diasParaSubtrair);
      indice = colunaB.findIndex(row => row[0] == Utilities.formatDate(dataAnterior, Session.getScriptTimeZone(), "dd/MM/yyyy"));
      diasParaSubtrair++;
    };

    var linhaInicial = indice + 1;

    colunaB.reverse();

    var indice1 = colunaB.findIndex(row => row[0] == Utilities.formatDate(dataInicioFormatada, Session.getScriptTimeZone(), "dd/MM/yyyy"));

    var diasParaAdicionar = 1;

    while (indice1 === -1 && diasParaAdicionar <= 30) {
      var dataPosterior = new Date(dataInicioFormatada);
      dataPosterior.setDate(dataPosterior.getDate() + diasParaAdicionar);
      indice1 = colunaB.findIndex(row => row[0] == Utilities.formatDate(dataPosterior, Session.getScriptTimeZone(), "dd/MM/yyyy"));
      diasParaAdicionar++;
    };

    var linhaFinal = planilha.getLastRow() - indice1;

    SpreadsheetApp.flush();

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guia = planilha.getSheetByName('Arquivo');

    var range = guia.getRange(linhaInicial + 1, 1, linhaFinal - linhaInicial + 1, guia.getLastColumn());

    var gid = guia.getSheetId();

    var printRange = objectToQueryString({
      'c1': range.getColumn() - 1,
      'r1': range.getRow() - 1,
      'c2': range.getColumn() + range.getWidth() - 1,
      'r2': range.getRow() + range.getHeight() - 1
    });

    var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

    var htmlTemplate = HtmlService.createTemplateFromFile('Abrirpdf');
    htmlTemplate.url = url;
    console.log(url);
    SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setHeight(10).setWidth(100), 'Gerando PDF');
  };
}

var PRINT_OPTIONS = {
  'size': 0,               // Tamanho do papel. 0 = carta, 1 = tablóide, 2 = Ofício, 3 = declaração, 4 = executivo, 5 = fólio, 6 = A3, 7 = A4, 8 = A5, 9 = B4, 10 = B
  'fzr': false,            // repetir cabeçalhos de linha
  'portrait': false,        // falso = paisagem
  'fitw': true,            // ajustar a janela ou tamanho real
  'gridlines': false,      // mostrar linhas de grade
  'printtitle': false,
  'sheetnames': false,
  'pagenum': 'UNDEFINED',  // CENTRO = mostrar números de página / UNDEFINED = não mostrar
  'attachment': false
}

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function objectToQueryString(obj) {
  return Object.keys(obj).map(function (key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
}






