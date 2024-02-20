function enviarEmails(apenasPrimeiraLinha, linhaInicio) {
  Logger.log("Iniciando a função enviarEmails");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var disparoEmailsSheet = spreadsheet.getSheetByName("Disparo de E-mails");
  var rpaSheet = spreadsheet.getSheetByName("RPA");
  var folderId = "1UmcGv4VWB1D78_HTiVehF-9ULMXrQ7lq";
  var emailTemplatePrimeiraFatura = getTemplateFromDrive(folderId, 'E-mail para Primeira Fatura.html');
  var emailTemplateRPA = getTemplateFromDrive(folderId, 'E-mail para RPA.html');
  var emailTemplateNF = getTemplateFromDrive(folderId, 'E-mail para NF.html');
  var linhaInicial = linhaInicio || 2; // Se linhaInicio não for fornecido, usa 2 como padrão
  var numLinhas = apenasPrimeiraLinha ? 1 : disparoEmailsSheet.getLastRow() - linhaInicial + 1;
  var data = disparoEmailsSheet.getRange(linhaInicial, 1, numLinhas, 8).getValues();
  
  Logger.log("Dados da planilha carregados: " + data.length + " linhas");
  
  function salvarComoExcel(sheet, folderId, fileName, isRPA) {
  if (isRPA) {
    // Converte as fórmulas em valores estáticos nas células específicas
    var rangeValoresCalculados = sheet.getRange("F23:F26");
    Logger.log("Antes da conversão: " + rangeValoresCalculados.getFormulas());
    rangeValoresCalculados.copyTo(rangeValoresCalculados, {contentsOnly: true});
    Logger.log("Após a conversão: " + rangeValoresCalculados.getValues());

  }

  var url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + '/export?exportFormat=xlsx&gid=' + sheet.getSheetId();
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var blob = response.getBlob().setName(fileName + '.xlsx');
  var file = DriveApp.getFolderById(folderId).createFile(blob);

  if (isRPA) {
    // Reverte as células para fórmulas após a exportação
    sheet.getRange("F23").setFormula('=Imposto_de_Renda!B1-F20');
    sheet.getRange("F24").setFormula('=Imposto_de_Renda!G10');
    sheet.getRange("F25").setFormula('=Imposto_de_Renda!H10');
    sheet.getRange("F26").setFormula('=Imposto_de_Renda!B2');
    Logger.log("Fórmulas revertidas após a exportação.");

  }

  return file;
}


    function salvarComoPdfEEnviarEmail(sheet, folderId, fileName, emailBody, email, subject) {
    var url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + 
              '/export?exportFormat=pdf&gid=' + sheet.getSheetId() + 
              '&size=letter' +  // Tamanho da página (pode ser A4, letter, etc.)
              '&portrait=true' + // Orientação da página (true para retrato, false para paisagem)
              '&fitw=true' +     // Ajustar largura da página ao conteúdo
              '&top_margin=0.20' +   // Margem superior mínima
              '&bottom_margin=0.20' + // Margem inferior mínima
              '&left_margin=0.20' +   // Margem esquerda mínima
              '&right_margin=0.20' +  // Margem direita mínima
              '&horizontal_alignment=CENTER' + // Alinhamento horizontal
              '&vertical_alignment=TOP';       // Alinhamento vertical

    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    var blob = response.getBlob().setName(fileName + '.pdf');
    var file = DriveApp.getFolderById(folderId).createFile(blob);

    var attachments = [file.getAs(MimeType.PDF), salvarComoExcel(sheet, folderId, fileName, perfil === "RPA").getAs(MimeType.MICROSOFT_EXCEL)];
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: emailBody,
      attachments: attachments,
      name: "Amanda do Financeiro 99Hunters"
    });
  }
  
// Checando se a planilha está preenchida corretamente

for (var i = 0; i < data.length; i++) {
    // Verifica se alguma das colunas obrigatórias (A-F) está vazia
    var linhaIncompleta = false;
    for (var j = 0; j < 6; j++) { // Altere o 6 para o número de colunas obrigatórias
        if (data[i][j] === "") {
            linhaIncompleta = true;
            break;
        }
      }

      if (linhaIncompleta) {
          Logger.log("Linha " + (i + 2) + " está incompleta");
          throw "É necessário completar todas as informações da planilha para executar a automação";
    }
}


  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var nome = row[0];
    var email = row[2];
    var perfil = row[3];
    var valorFatura = row[4]
    var demonstrativo = row[5].replace(/\n/g, '<br>');

    // Formata o demonstrativo para aparecer em itálico e como um quote
    var demonstrativoFormatado = "<blockquote>" + demonstrativo.replace(/\n/g, '<br>') + "</blockquote>";

    // Adiciona a linha centralizada e em negrito após o demonstrativo
    var linhaValorTotal = "<div style='text-align: center; font-weight: bold;'>Valor total do recebimento: " + "R$ " + valorFatura + ",00" + "</div>";

    var dadosPagamento = row[6] || "";
    var remetenteEmail = Session.getActiveUser().getEmail(); // Endereço de e-mail do remetente
    var dataHoraExecucao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    var logMessage = ""; // Inicializa a mensagem de log

    var mensagemDadosPagamento;
    var pediuConfirmacaoDadosBancarios = "Não";
    if (dadosPagamento.trim() === "") {
        mensagemDadosPagamento = "É importante para nós garantir que os dados de pagamento estão atualizados para enviarmos seu repasse para a conta correta. Consegue responder confirmando novamente, por gentileza, seus dados de PIX ou transferência?";
        pediuConfirmacaoDadosBancarios = "Sim";

    } else {
        // Formata os dados de pagamento para aparecerem em itálico e como um quote
        mensagemDadosPagamento = "Ah, e a checagem de sempre: Seus dados de transferência continuam os mesmos?<br><blockquote><i>" + dadosPagamento.replace(/\n/g, '<br>') + "</i></blockquote>" + "Se tiver mudado, favor enviar os dados atualizados em sua resposta.";        
    }


    var emailBody;
    switch (perfil) {
      case "Primeira Fatura":
        emailBody = emailTemplatePrimeiraFatura.replace("{Nome}", nome).replace("{Demonstrativo}", demonstrativoFormatado + linhaValorTotal);
        Logger.log("Enviando e-mail para o perfil A: " + email);
        GmailApp.sendEmail(email, "Seu repasse 99Hunters", "", {htmlBody: emailBody, name: "Amanda do Financeiro 99Hunters"});
        logMessage = "Primeira Fatura enviada por " + remetenteEmail + " " + "em " + dataHoraExecucao;

        break;
      case "RPA":
        var cellValue = rpaSheet.getRange("F12").getValue();
        Logger.log("Valor original da célula F12: " + cellValue);

        var novoValor = row[4];
        rpaSheet.getRange("F12").setValue(novoValor);
        Logger.log("Novo valor definido para a célula F12: " + novoValor);

        // Aguarde a planilha ser atualizada antes de continuar
        SpreadsheetApp.flush();

        var rpaFileName = "RPA_" + nome;
        emailBody = emailTemplateRPA.replace("{Nome}", nome).replace("{Demonstrativo}", demonstrativoFormatado + linhaValorTotal).replace("{Dados de Pagamento}", mensagemDadosPagamento);
        Logger.log("Enviando e-mail para o perfil B: " + email);
        salvarComoPdfEEnviarEmail(rpaSheet, "1ucxksOJ9XkgSNBrEJvPZTEm_viYvsFAc", rpaFileName, emailBody, email, "Seu repasse 99Hunters");
        logMessage = "RPA enviada " + (pediuConfirmacaoDadosBancarios === "Sim" ? "com pedido de dados bancários " : "com dados bancários inclusos ") + "por " + remetenteEmail + " " + "em " + dataHoraExecucao + ".";

        break;

      case "NF":
        emailBody = emailTemplateNF.replace("{Nome}", nome).replace("{Demonstrativo}", demonstrativoFormatado + linhaValorTotal).replace("{Dados de Pagamento}", mensagemDadosPagamento);
        Logger.log("Enviando e-mail para o perfil C: " + email);
        GmailApp.sendEmail(email, "Seu repasse 99Hunters", "", {htmlBody: emailBody, name: "Amanda do Financeiro 99Hunters"});
        logMessage = "NF enviada " + (pediuConfirmacaoDadosBancarios === "Sim" ? "com pedido de dados bancários " : "com dados bancários inclusos ") + "por " + remetenteEmail + " " + "em " + dataHoraExecucao + ".";

        break;
      default:
        Logger.log("Perfil desconhecido: " + perfil);
        continue;
    }
    if (logMessage) {
        disparoEmailsSheet.getRange(linhaInicial + i, 8).setValue(logMessage); //
        SpreadsheetApp.flush(); // Garante que a planilha seja atualizada imediatamente
    }

  }  
  Logger.log("Função enviarEmails concluída");
  Browser.msgBox("Envio concluído com Sucesso", Browser.Buttons.OK);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automação de E-mails')
      .addItem('Testar com Primeira Linha', 'testePrimeiraLinha')
      .addItem('Disparar E-mails', 'enviarEmails')
      .addItem('Continuar de uma linha específica', 'continuarDeUmaLinhaEspecifica')
      .addItem('Fazer Backup e Limpar Planilha', 'limparPlanilha')
      .addToUi();
}

function continuarDeUmaLinhaEspecifica() {
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.prompt('Continuar a partir de qual linha?', ui.ButtonSet.OK_CANCEL);

  if (resposta.getSelectedButton() == ui.Button.OK) {
    var linha = parseInt(resposta.getResponseText());
    if (!isNaN(linha) && linha > 1) {
      enviarEmails(false, linha);
      ui.alert('Execução concluída a partir da linha ' + linha);
    } else {
      ui.alert('Por favor, insira um número de linha válido maior que 1.');
    }
  } else {
    ui.alert('Operação cancelada.');
  }
}


function limparPlanilha() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Você tem certeza que quer limpar todos os dados e salvar um backup?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var disparoEmailsSheet = spreadsheet.getSheetByName("Disparo de E-mails");
    var folderId = "10MSdA4qo7_g_V3h_g64Hl2Kfo-tJRQYu";
    var fileName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_dd HH:mm:ss") + "_repasses_hunters";

    // Exporta a aba "Disparo de E-mails" como Excel
    var file = exportarAbaComoExcel(disparoEmailsSheet, folderId, fileName);

    // Limpa os dados da aba "Disparo de E-mails", preservando o cabeçalho
    var lastRow = disparoEmailsSheet.getLastRow();
    if (lastRow > 1) {
      disparoEmailsSheet.getRange(2, 1, lastRow - 1, disparoEmailsSheet.getLastColumn()).clearContent();
    }

    // Notificação de conclusão
    ui.alert('Backup concluído e salvo na pasta "Backups" com o nome ' + file.getName());
  } else {
    ui.alert('Operação cancelada.');
  }
}

function exportarAbaComoExcel(sheet, folderId, fileName) {
  var url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + '/export?exportFormat=xlsx&gid=' + sheet.getSheetId();
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var blob = response.getBlob().setName(fileName + '.xlsx');
  var file = DriveApp.getFolderById(folderId).createFile(blob);
  return file;
}


function testePrimeiraLinha() {
  enviarEmails(true); // Chama enviarEmails para processar apenas a primeira linha
  Logger.log("Teste concluído.");
  Browser.msgBox("Teste concluído.", Browser.Buttons.OK);
}

function getTemplateFromDrive(folderId, fileName) {
  Logger.log("Buscando template: " + fileName);
  var files = DriveApp.getFolderById(folderId).getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    return file.getBlob().getDataAsString();
  }
  Logger.log("Template não encontrado: " + fileName);
  return null;
}


