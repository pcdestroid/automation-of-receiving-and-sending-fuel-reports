function export_relatorio_combustiveis() {

  //Exportar relatório de comsbustível do email e salvar no Drive

  const id_plan_relatorio_combustiveis = PropertiesService.getScriptProperties().getProperty("id_plan_relatorio_combustiveis")

  const id_pasta_anexo_temporario = PropertiesService.getScriptProperties().getProperty("id_pasta_anexo_temporario")

  const assuntoAProcurar = 'RELATÓRIO MENSAL DE COMBUSTÍVEIS';
  const minhaPasta = DriveApp.getFolderById(id_pasta_anexo_temporario);

  // Procurar apenas uma thread com o assunto específico.
  const threads = GmailApp.search(`subject:"${assuntoAProcurar}"`, 0, 1);

  if (threads.length > 0 && threads[0].isUnread()) { // Se houver pelo menos uma thread encontrada e não foi lido.

    threads[0].markRead() // Marca com lido

    const messages = threads[0].getMessages();
    const attachments = messages[messages.length - 1].getAttachments(); // Obter anexos da última mensagem da thread.
    for (let j = 0; j < attachments.length; j++) {
      const anexo = attachments[j];

      // remover o arquivo existente, se houver
      const arquivosAntigos = minhaPasta.getFilesByName('RELATORIO_COMBUSTIVEL.XLSX');
      while (arquivosAntigos.hasNext()) {
        const arquivoAntigo = arquivosAntigos.next();
        arquivoAntigo.setTrashed(true);
      }

      // salvar o novo anexo na pasta especificada
      let arquivo = minhaPasta.createFile(anexo).setName('RELATORIO_COMBUSTIVEL.XLSX');

      // Converta o arquivo para o formato Google Sheets
      let arquivoConvertido = Drive.Files.copy({}, arquivo.getId(), { convert: true });

      // Abra a planilha convertida
      let planilha = SpreadsheetApp.openById(arquivoConvertido.id);

      // Acesse a primeira guia da planilha convertida
      let guia = planilha.getSheets()[0];

      // Obtenha os dados da planilha convertida
      let dados = guia.getDataRange().getValues();

      // Filtra as linhas baseando-se nas condições dadas
      let dados_filtrados = dados.slice(1).filter(linha => {
        // Define as variáveis para as condições
        let unidade_negocio = linha[1];
        let codigo_produto = linha[14];

        // Retorna verdadeiro para linhas que devem ser incluídas
        return unidade_negocio !== "" &&
          unidade_negocio !== "Unidade de Negócio" &&
          codigo_produto !== "E0801" &&
          codigo_produto !== "E1552";
      });

      // Planilha de Relatório de combustíveis
      let spreadsheetId = id_plan_relatorio_combustiveis;
      let relacao_combustiveis = SpreadsheetApp.openById(spreadsheetId);
      let guiaRecebimento = relacao_combustiveis.getSheetByName('Dados');
      let guiaResumo = relacao_combustiveis.getSheetByName('Resumo');

      let tipos_combustiveis = relacao_combustiveis.getSheetByName("Combustíveis");

      // Criar mapa com códigos e descrição do combustível
      let mapa = {}
      let valores = tipos_combustiveis.getRange("A:B").getValues();
      valores.slice(1).map(produto => { mapa[produto[0]] = produto[1] })

      let competencia = ""
      let fullDate = ""

      // Transforma os dados filtrados para o formato desejado
      let dados_planilha = dados_filtrados.map(linha => {
        let produto = linha[14]
        if (produto === "E0333") produto = "P0058"
        if (competencia == "") {
          var dataObjeto = new Date(linha[10]); // Converte a string em objeto de data
          var mes = dataObjeto.getMonth() + 1; // getMonth() retorna mês de 0-11
          var ano = dataObjeto.getFullYear(); // retorna o ano

          // Assegurar que o mês seja sempre dois dígitos
          var mesFormatado = mes < 10 ? '0' + mes : mes;

          // Formata a competência como MM/AAAA
          competencia = mesFormatado + "/" + ano;
          fullDate = new Date(mesFormatado + "/" + "01" + "/" + ano)
        }

        return [
          linha[0], // codigo
          linha[1], // unidade_negocio
          linha[2], // codigo_fornecedor
          linha[3], // razao_social
          linha[8], // nota_fiscal
          linha[9], // emissao_nota
          linha[10], // entrada_nota
          linha[11], // ordem_compra
          linha[12], // quantidade
          linha[13], // valor
          produto, // codigo_produto
          linha[15], // valor_bruto
          mapa[produto], // descricao
          linha[18], // conta_financeira
          linha[19], // conta_reduzida
          linha[20], // centro_custo
          competencia
        ];
      });

      // Verificar se a competencia já foi registrada alguma vez
      let periodos_registrados = guiaRecebimento.getRange(guiaRecebimento.getLastRow(), 17).getValue()

      if (new Date(periodos_registrados).toString() === new Date(fullDate).toString()) {
        Logger.log("Período já foi registrado.")
        return
      }

      // Localiza a última linha com dados na guia 'Dados'
      let ultimaLinha = guiaRecebimento.getLastRow();

      // Se houver dados a inserir
      if (dados_planilha.length > 0) {
        // Calcula o range necessário para os novos dados
        let numeroLinhas = dados_planilha.length;
        let numeroColunas = dados_planilha[0].length; // assume que todas as linhas têm o mesmo número de colunas
        let range = guiaRecebimento.getRange(ultimaLinha + 1, 1, numeroLinhas, numeroColunas);

        // Insere todos os dados de uma vez
        range.setValues(dados_planilha);
        Logger.log('Relatório atualizado')
      }

      let emailComprador = CalendarApp.getId();
      let dadosComprador = getUserByEmail(emailComprador)
      let assinatura = `
  
  <table style="color: rgb(136, 136, 136); white-space: normal; background-color: rgb(255, 255, 255); font-size: 12px; font-family: Arial; width: 513px;">
  <tbody>
  <tr>
  <td style="margin: 0px; max-width: 100px; width: 100px;" valign="top"><img class="CToWUd" style="border-radius: 4px; width: 113px; height: 118px;" src="https://ci6.googleusercontent.com/proxy/iEOrmj7yb3xF4gNx9mDtAxGFuveeaACz1ZAt-RYPn78dDAGhfRJDLwhu34hjsDUlpuRooQ_aF2V0DcCRjx4IDp3O=s0-d-e1-ft#https://i.ibb.co/fG73FxC/Logo-BR-Assinatura.png" alt="Sua foto ou logo" data-bit="iit"></td>
  <td style="margin: 0px; padding: 0px 8px; width: 414px;">
  <table style="color: rgb(128, 128, 128); height: 108px; width: 414px;">
  <tbody>
  <tr style="height: 18px;">
  <td style="margin: 0px; font-size: 14px; font-weight: bold; color: rgb(14, 43, 141); height: 18px; width: 404px;">${dadosComprador[0]}</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">Compras</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">BR Marinas</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">${dadosComprador[5]}&nbsp;&nbsp;<a aria-label="Chat on WhatsApp" href="https://wa.me/${dadosComprador[6]}"><img class="CToWUd" src="https://ci3.googleusercontent.com/proxy/kJDPaPYcNQs64k_qKrGFp6XYuYXrA0FkVNTRpvuAzRv7COIpd65R8420WBcSTG4QRUPKdM_7DUQ=s0-d-e1-ft#https://i.ibb.co/Bjn8dcj/whatsapp.png" alt="Whatsapp" width="16" height="16" data-bit="iit"></a></td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;"><a style="color: rgb(17, 85, 204);" href="mailto:${emailComprador}" target="_blank" rel="noopener">${emailComprador}</a></td>
  </tr>
  </tbody>
  </table>
  </td>
  </tr>
  </tbody>
  </table>
  `

      // Converta os dados filtrados de volta para um novo arquivo XLSX
      const newSpreadsheet = SpreadsheetApp.create('RELATORIO_COMBUSTIVEL');
      const sheet = newSpreadsheet.getSheets()[0];
      // Verifique se há dados para inserir
      if (dados_planilha.length > 0 && dados_planilha[0].length > 0) {
        sheet.appendRow(['Cod', 'Unidade de Negócio', 'Forn.', 'Razão Social', 'NF', 'Emissão', 'Entrada', 'OC', 'QTD', 'Vlr', 'Produto', 'Vlr. Bruto', 'Descrição', 'Conta Financeira', 'Conta reduzida', 'Centro de Custo', 'Mês'])
        // Aplicar negrito, cor de fundo e cor do texto na primeira linha
        sheet.getRange(1, 1, 1, dados_planilha[0].length)
          .setFontWeight("bold")
          .setBackground("#073763")
          .setFontColor("#FFFFFF");

        // Alinha todo o texto da planilha à esquerda
        sheet.getDataRange().setHorizontalAlignment("left");

        sheet.getRange(2, 1, dados_planilha.length, dados_planilha[0].length).setValues(dados_planilha);
      } else {
        Logger.log("Sem dados para inserir na planilha.");
        return; // Sair da função se não houver dados
      }

      // Garante que todas as alterações pendentes na planilha sejam aplicadas
      SpreadsheetApp.flush();

      // Obtém os dados da planilha
      var dadosResumo = sheet.getDataRange().getValues();
      var resumo = {};

      // Itera por cada linha de dados (excluindo o cabeçalho)
      for (var i = 1; i < dadosResumo.length; i++) {
        var linha = dadosResumo[i];
        var marina = linha[1]; // Coluna B
        var descricao = linha[12]; // Coluna M
        var quantidade = linha[8]; // Coluna I
        var valorUnitario = linha[9]; // Coluna J

        // Se não existir, inicializa o objeto para a marina e combustível
        if (!resumo[marina]) resumo[marina] = {};
        if (!resumo[marina][descricao]) resumo[marina][descricao] = { quantidade: 0, valorUnitario: 0, valorTotal: 0 };

        // Agrega os dados
        var grupo = resumo[marina][descricao];
        grupo.quantidade += quantidade;
        grupo.valorTotal += quantidade * valorUnitario;
        grupo.valorUnitario = grupo.valorTotal / grupo.quantidade; // Calcula o valor unitário médio
      }

      let html_resumo = gerarHTML(resumo);

      // Primeiro, obtenha o URL de exportação
      const file = DriveApp.getFileById(newSpreadsheet.getId());
      const url = file.getUrl().replace(/\/edit.*$/, '');

      // Prepare os parâmetros para a solicitação de exportação
      const exportUrl = url + '/export?exportFormat=xlsx&format=xlsx';

      // Use o Google Drive API para exportar o arquivo como XLSX
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: {
          'Authorization': 'Bearer ' + token
        }
      });

      // Crie o Blob a partir da resposta da API
      const blob = response.getBlob().setName('RELATORIO_COMBUSTIVEL.XLSX');

      // Crie o rascunho do email com anexo
      const subject = `Combustível - ${competencia} `;

      let link_plan = PropertiesService.getScriptProperties().getProperty("link_plan")

      let emailTemp = HtmlService.createHtmlOutput(`<span>${bomDia()} Alice!</span><br><span>Anexo relatório de combustível ref. ${competencia}.</span><br><span>Segue abaixo o resumo das aquisições de combustíveis realizadas no mês por Marina.</span><br><br>Consulte todo período aqui: <a href="${link_plan}">Combustíveis - BR Marinas</a><br>${html_resumo}<br><br>Atenciosamente,</span><br><br>${assinatura}<br>`).setTitle('Email');

      let htmlMessage = emailTemp.getContent();

      const email_para_envio = PropertiesService.getScriptProperties().getProperty("email_para_envio")

      const email_copia = PropertiesService.getScriptProperties().getProperty("email_copia");

      GmailApp.sendEmail(
        email_para_envio,
        subject,
        "Your email doesn't support HTML.",
        {
          name: 'Pedido de orçamento',
          htmlBody: htmlMessage,
          bcc: email_copia,
          attachments: [blob]
        }
      );

      // Apagar a planilha temporária criada para o relatório
      DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);

      // Log de conclusão
      Logger.log('Rascunho de email com relatório de combustível criado');

      // Feche a planilha convertida
      DriveApp.getFileById(arquivoConvertido.id).setTrashed(true);

    }

  } else {
    Logger.log('Nenhum relatório encontrado no email')
    return false;
  }

}

function getUserByEmail(email) {
  //Pegar nome do usuário
  const id_plan_dados_usuarios = PropertiesService.getScriptProperties().getProperty("id_plan_dados_usuarios")
  let users = SpreadsheetApp.openById(id_plan_dados_usuarios).getSheetByName('usuarios');
  let dadosUsers = users.getRange(3, 1, users.getLastRow(), 7).getValues();
  //Pegar nome do comprador
  let user = dadosUsers.find(dado => dado[2] === email);
  return user ? user : '';
}

function bomDia() {
  // Retorna a frase, bom dia quando é de dia, boa tarde quando é de tarde e boa noite quando é de noite.
  let data = new Date();
  let hora = data.getHours();
  switch (true) {
    case (hora >= 5 && hora < 12):
      return 'Bom dia';
    case (hora >= 12 && hora < 18):
      return 'Boa tarde';
    case (hora >= 18 && hora <= 24 || hora >= 0 && hora <= 4):
      return 'Boa noite';
  }
}

function gerarHTML(resumo) {
  var html = "<div style='max-height: 75%; max-width: 75%; overflow: auto;'>"; // Div para controlar a altura máxima

  for (var marina in resumo) {
    // Título com o nome da marina
    html += "<h3>" + marina + "</h3>";

    // Inicia uma nova tabela com borda sólida
    html += "<table style='border-collapse: collapse; width: 75%;'>";

    // Cabeçalho da tabela com estilo
    html += "<tr style='background-color: #073763; color: white;'>";
    html += "<th style='border: 1px solid black; padding: 5px; text-align: center;'>Combustível</th>";
    html += "<th style='border: 1px solid black; padding: 5px; text-align: center;'>Qtd. Total</th>";
    html += "<th style='border: 1px solid black; padding: 5px; text-align: center;'>Preço Médio</th>";
    html += "<th style='border: 1px solid black; padding: 5px; text-align: center;'>Valor Total</th>";
    html += "</tr>";

    for (var combustivel in resumo[marina]) {
      var grupo = resumo[marina][combustivel];
      // Dados do grupo específico com borda e duas casas decimais para quantidade
      html += "<tr>";
      html += "<td style='border: 1px solid black; padding: 5px; text-align: center;'>" + combustivel + "</td>";
      html += "<td style='border: 1px solid black; padding: 5px; text-align: center;'>" + grupo.quantidade.toFixed(2) + "</td>";
      html += "<td style='border: 1px solid black; padding: 5px; text-align: center;'>R$ " + grupo.valorUnitario.toFixed(2) + "</td>";
      html += "<td style='border: 1px solid black; padding: 5px; text-align: center;'>R$ " + grupo.valorTotal.toFixed(2) + "</td>";
      html += "</tr>";
    }

    html += "</table><br>"; // Fecha a tabela da marina atual
  }

  html += "</div>"; // Fecha a div de controle de altura

  return html;
}
