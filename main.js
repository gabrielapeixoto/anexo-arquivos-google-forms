FormApp.getActiveForm();

function onFormSubmit(e) {
  var itemResponses = e.response.getItemResponses();
  var razaoSocial = "";
  var certidao = null;
  var balanco = null;
  var dre = null;
  var certificacao = null;

  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var question = itemResponse.getItem().getTitle();
    var response = itemResponse.getResponse();

    // Logger.log("Pergunta: " + question); // Log da pergunta
    // Logger.log("Resposta: " + response); // Log da resposta

    if (question == "Razão Social:") {
      razaoSocialQuestion = JSON.stringify(question);
      razaoSocial = JSON.stringify(response);
    } else if (question == "Certidão Negativa de Débitos (CND)") {
      certidaoQuestion = JSON.stringify(question);
      certidao = JSON.stringify(response);
    } else if (question == "Balanço Patrimonial") {
      balancoQuestion = JSON.stringify(question);
      balanco = JSON.stringify(response);
    } else if (question == "DRE") {
      dreQuestion = JSON.stringify(question);
      dre = JSON.stringify(response);
    } else if (question == "Se a resposta anterior for sim, anexar Certificações:") {
      certificacaoQuestion = JSON.stringify(question);
      certificacao = JSON.stringify(response);
    }
  }

  // Logger.log("Razão Social: " + razaoSocial);
  // Logger.log("Certidão CND URL: " + certidao);
  // Logger.log("Balanço Patrimonial: " + balanco);
  // Logger.log("DRE URL: " + dre);
  // Logger.log("DRE URL: " + certificacao);

  var pastaRaiz = DriveApp.getFolderById("1jA6j3r1tP983ihkaKtTHRMwGQVsqinJl"); // ID da sua pasta raiz

  // Crie a pasta com a Razão Social, se não existir
  var pastas = pastaRaiz.getFoldersByName(razaoSocial);
  var pastaEmpresa = pastas.hasNext() ? pastas.next() : pastaRaiz.createFolder(razaoSocial.replace(/"/g, ''));


  if (certidao != null) {
    var certidoes = certidao.split(",");
    var pastasCampo = pastaEmpresa.getFoldersByName(certidaoQuestion);
    var subpastaEmpresa = pastasCampo.hasNext() ? pastasCampo.next() : pastaEmpresa.createFolder(certidaoQuestion.replace(/"/g, ''));

    for (var i = 0; i < certidoes.length; i++) {
      var idArquivo = certidoes[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, subpastaEmpresa);
    }
  }

  if (balanco != null) {
    var balancos = balanco.split(",");
    var pastasCampo = pastaEmpresa.getFoldersByName(balancoQuestion);
    var subpastaEmpresa = pastasCampo.hasNext() ? pastasCampo.next() : pastaEmpresa.createFolder(balancoQuestion.replace(/"/g, ''));

    for (var i = 0; i < balancos.length; i++) {
      var idArquivo = balancos[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, subpastaEmpresa);
    }
  }

  if (dre != null) {
    var dres = dre.split(",");
    var pastasCampo = pastaEmpresa.getFoldersByName(dreQuestion);
    var subpastaEmpresa = pastasCampo.hasNext() ? pastasCampo.next() : pastaEmpresa.createFolder(dreQuestion.replace(/"/g, ''));

    for (var i = 0; i < dres.length; i++) {
      var idArquivo = dres[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, subpastaEmpresa);
    }
  }

  if (certificacao != null) {
    var certificacoes = certificacao.split(",");
    var pastasCampo = pastaEmpresa.getFoldersByName(certificacaoQuestion);
    var subpastaEmpresa = pastasCampo.hasNext() ? pastasCampo.next() : pastaEmpresa.createFolder(certificacaoQuestion.replace(/"/g, ''));

    for (var i = 0; i < certificacoes.length; i++) {
      var idArquivo = certificacoes[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, subpastaEmpresa);
    }
  }
}

function salvarArquivo(idArquivo, pastaDestino) {
  if (typeof idArquivo === 'string') {
    try {
      // 1. Obtém o arquivo pelo ID
      var file = DriveApp.getFileById(idArquivo);

      // 2. Obtém o Blob do arquivo
      var blob = file.getBlob();

      // 3. Obtém o nome original do arquivo
      var nomeArquivo = file.getName();
      Logger.log("Nome do arquivo: " + nomeArquivo); // Log do nome do arquivo

      // 4. Define o nome do Blob
      blob.setName(nomeArquivo);

      // 5. Cria o arquivo na pasta de destino
      pastaDestino.createFile(blob);

    } catch (error) {
      Logger.log("Erro ao salvar arquivo " + nomeArquivo + ": " + error);

      MailApp.sendEmail({
        to: "email@ipnet.cloud",
        subject: "Erro ao salvar arquivo",
        body: "Ocorreu um erro ao salvar o arquivo " + nomeArquivo + ": " + error
      });
    }
  } else {
    Logger.log("ID do arquivo inválido: " + idArquivo);

    MailApp.sendEmail({
      to: "email@ipnet.cloud",
      subject: "ID do arquivo inválido",
      body: "Ocorreu um erro ao salvar o arquivo " + nomeArquivo + ": " + error
    });
  }
}
