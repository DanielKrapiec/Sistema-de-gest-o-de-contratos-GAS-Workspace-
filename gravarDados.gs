function gravarDados() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menu = spreadsheet.getSheetByName('Menu');
  var arquivo = spreadsheet.getSheetByName('Arquivo');

  var codigoContrato = menu.getRange('C6').getValue();  // Pega o código do contrato
  var codigoControle = menu.getRange('C5').getValue();  // Pega o valor de C5

  var copia = menu.getRange('B15').getValue();
  var tipoDoc = menu.getRange('B6').getValue();

  var final = new Date();
  var inicio = menu.getRange('C1').getValue();
  var delta1 = final - inicio; // diferença em milissegundos

  // Cálculo real
  var totalSegundos = Math.floor(delta1 / 1000);
  var minutos = Math.floor(totalSegundos / 60);
  var segundos = totalSegundos % 60;

  // Formatando como mm:ss
  var delta = "00" + ":" + minutos.toString().padStart(2, '0') + ":" + segundos.toString().padStart(2, '0');

  // Se C5 for igual a "a", procura a linha correspondente ao código do contrato
  if (codigoControle === "a") {
    var dadosArquivo = arquivo.getDataRange().getValues(); // Pega todos os dados da aba "Arquivo"
    var linhaEncontrada = -1; // Variável para armazenar a linha do código encontrado

    // Percorre os dados para encontrar o código
    for (var i = 0; i < dadosArquivo.length; i++) {
      if (dadosArquivo[i][6] == codigoContrato) { // Assume que o código do contrato está na coluna G (índice 6)
        linhaEncontrada = i + 1; // Linha encontrada, a partir de 1 no Google Sheets
        break;
      }
    }

    if (linhaEncontrada !== -1) {
      // Atualiza os valores na linha encontrada
      arquivo.getRange(linhaEncontrada, 3).setValue(menu.getRange('B1').getValue()); // RA
      arquivo.getRange(linhaEncontrada, 4).setValue(menu.getRange('B2').getValue()); // Nome
      arquivo.getRange(linhaEncontrada, 5).setValue(menu.getRange('B3').getValue()); // Unidade
      arquivo.getRange(linhaEncontrada, 6).setValue(menu.getRange('B4').getValue()); // Curso
      arquivo.getRange(linhaEncontrada, 7).setValue(menu.getRange('C6').getValue()); // Código contrato
      arquivo.getRange(linhaEncontrada, 8, 1, 9).setValues([menu.getRange('B6:B14').getValues()]); // Dados do termo
      arquivo.getRange(linhaEncontrada, 17).setValue(menu.getRange('C15').getValue()); // Status de assinatura
      arquivo.getRange(linhaEncontrada, 19).setValue(delta);

      spreadsheet.getRangeList(['Menu!B1', 'Menu!B6:B12', 'Menu!B14:B15', 'Menu!C7','Menu!C14']).clear({ contentsOnly: true, skipFilteredRows: true });
      spreadsheet.getRange('Menu!B1').activate();

      var final1 = new Date();
      final1.setDate(final1.getDate() + 14); // Modifica o objeto 'final' diretamente

      // Agora sim, formata a data para 'dd/mm/aaaa'
      var dia1 = String(final1.getDate()).padStart(2, '0');
      var mes1 = String(final1.getMonth() + 1).padStart(2, '0');
      var ano1 = final1.getFullYear();

      if (copia === "SIM" && tipoDoc !== "ATA" ) {
        var texto = "Olá, saudações. \n\nObrigado pelo envio dos documentos. \n\nRecebemos a sua solicitação de estágio obrigatório e iniciamos o processo de assinatura eletrônica. \n\nApós assinatura da Coordenação de Curso, o aluno e a empresa, receberão um e-mail para realizar a assinatura. \n\nO PRAZO ESTIPULADO É: " + `${dia1}/${mes1}/${ano1}` + "\n" + codigoContrato +"\n\nEstamos à disposição.";
        SpreadsheetApp.getUi().alert("" + texto);
      }
    } else {
      Browser.msgBox("Contrato não encontrado para atualização");
    }
  } else {
    // Caso contrário, insere uma nova linha na posição 4
    arquivo.insertRowsBefore(4, 1);
    var hoje = new Date();
    var dia = hoje.getDate().toString();
    if (dia.length == 1) dia = "0" + dia;
    var mes = (hoje.getMonth() + 1).toString();
    if (mes.length == 1) mes = "0" + mes;
    var ano = hoje.getFullYear();
    var hojeFormatada = dia + '/' + mes + '/' + ano;

    // Preenche a nova linha com os dados
    arquivo.getRange('B4').setValue(hojeFormatada);
    arquivo.getRange('C4:F4').setValues([menu.getRange('B1:B4').getValues()]);
    arquivo.getRange('G4').setValue(menu.getRange('C6').getValue()); // Código do contrato
    arquivo.getRange('H4:P4').setValues([menu.getRange('B6:B14').getValues()]);
    arquivo.getRange('Q4').setValue(menu.getRange('C15').getValue());
    arquivo.getRange('S4').setValue(delta);

    // Limpa os dados na aba Menu
    menu.getRangeList(['B1', 'B6:B12', 'B14:B15', 'C14']).clear({ contentsOnly: true, skipFilteredRows: true });
    menu.getRange('B1').activate();

    var final1 = new Date();
    final1.setDate(final1.getDate() + 14); // Modifica o objeto 'final' diretamente

    // Agora sim, formata a data para 'dd/mm/aaaa'
    var dia1 = String(final1.getDate()).padStart(2, '0');
    var mes1 = String(final1.getMonth() + 1).padStart(2, '0');
    var ano1 = final1.getFullYear();

    if (copia === "SIM" && tipoDoc !== "ATA") {
      var texto = "Olá, saudações. \n\nObrigado pelo envio dos documentos. \n\nRecebemos a sua solicitação de estágio obrigatório e iniciamos o processo de assinatura eletrônica. \n\nApós assinatura da Coordenação de Curso, do Professor Orientador e da Parte Concedente, o(a) aluno(a), receberá um e-mail com o link para realizar a assinatura. \n\nO PRAZO ESTIPULADO É: " + `${dia1}/${mes1}/${ano1}` + "\n" + codigoContrato +"\n\nEstamos à disposição.";
      SpreadsheetApp.getUi().alert("" + texto);
    }
  }
}
