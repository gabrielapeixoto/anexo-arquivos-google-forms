FormApp.getActiveForm();

function onFormSubmit(e) {
  var itemResponses = e.response.getItemResponses();
  var razaoSocial = "";
  var certidao = null;
  var balanco = null;
  var dre = null;
  var cartificacao = null;

  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var question = itemResponse.getItem().getTitle();
    var response = itemResponse.getResponse();

    // Logger.log("Pergunta: " + question); // Log da pergunta
    // Logger.log("Resposta: " + response); // Log da resposta

    if (question == "Razão Social:") {
      razaoSocial = JSON.stringify(response);
    } else if (question == "Certidão Negativa de Débitos (CND)") {
      certidao = JSON.stringify(response);
    } else if (question == "Balanço Patrimonial") {
      balancoQuestion = JSON.stringify(question);
      balanco = JSON.stringify(response);
    } else if (question == "DRE") {
      dre = JSON.stringify(response);
    } else if (question == "Se a resposta anterior for sim, anexar Certificações:") {
      cartificacao = JSON.stringify(response);
    }
  }

  // Logger.log("Razão Social: " + razaoSocial);
  // Logger.log("Certidão CND URL: " + certidao);
  // Logger.log("Balanço Patrimonial: " + balanco);
  // Logger.log("DRE URL: " + dre);
  // Logger.log("DRE URL: " + cartificacao);

  var pastaRaiz = DriveApp.getFolderById("ID_PASTA_RAIZ"); // ID da sua pasta raiz

  // Crie a pasta com a Razão Social, se não existir
  var pastas = pastaRaiz.getFoldersByName(razaoSocial);
  if (pastas.hasNext()) {
    pastaEmpresa = pastas.next();
  } else {
    pastaEmpresa = pastaRaiz.createFolder(razaoSocial.replace(/"/g, ''));
  }

  if (certidao != null) {
    var certidoes = certidao.split(",");
    for (var i = 0; i < certidoes.length; i++) {
      var idArquivo = certidoes[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, pastaEmpresa);
    }
  }

  if (balanco != null) {
    var balancos = balanco.split(",");
    for (var i = 0; i < balancos.length; i++) {
      var idArquivo = balancos[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, pastaEmpresa);
    }
  }

  if (dre != null) {
    var dres = dre.split(",");
    for (var i = 0; i < dres.length; i++) {
      var idArquivo = dres[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, pastaEmpresa);
    }
  }

  if (cartificacao != null) {
    var certificacoes = cartificacao.split(",");
    for (var i = 0; i < certificacoes.length; i++) {
      var idArquivo = certificacoes[i].replace(/[\[\]"]/g, '');
      Logger.log("Balanço Patrimonial: " + idArquivo);
      salvarArquivo(idArquivo, pastaEmpresa);
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

      // 5. Cria o arquivo na pasta de destino e verifica se arquivo já existe
      var arquivosExistentes = pastaDestino.getFilesByName(nomeArquivo);
      if (!arquivosExistentes.hasNext()) {
        pastaDestino.createFile(blob);
      } else {
        Logger.log("Arquivo " + nomeArquivo + " já existe na pasta.");
      }

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
