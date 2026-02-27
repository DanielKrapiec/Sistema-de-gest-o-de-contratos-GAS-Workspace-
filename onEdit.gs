var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
function onEdit(e) {
  var cellAddress = e.range.getA1Notation();
  
  if (cellAddress === 'B1') {
    inicio();
    return;
  }
  if (cellAddress === 'B8') {
    alertMessage(); // Chama a função que lê o valor da célula C5
    return;
  }
  if (cellAddress === 'B7') {
    alertMessage1(); // Aqui, o valor editado é passado diretamente
    return;
  }
}
function inicio() {
  spreadsheet.getRange('C1').setValue(new Date()); // Atualiza C1 com a data e hora atuais
}
function alertMessage() {
  var cellCode = spreadsheet.getRange('C5').getValue(); // Agora lê o valor de C5 diretamente
  if (cellCode === "a") {
    showToast("⭕Codigo duplicado, verifique antes de continuar.⭕");
  }
}
function alertMessage1() {
var cellDocument = spreadsheet.getRange('B6').getValue();
  if (cellDocument === "TCE-E") {
    showToast("⚠️Verificar a documentação do estágio remunerado!⚠️");
  }
}
function showToast(message) {
  spreadsheet.toast(message, "Atenção", 10);
}
