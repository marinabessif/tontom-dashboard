/**
 * Recebe as requisições GET (JSONP) do painel Tontom Dashboard.
 * Versão Consolidada: Audiências, Cancelamento, Serprec (ERRO / VINC.) e Migração.
 */
function doGet(e) {
  var callback = e.parameter.callback || "callback";
  var action = e.parameter.action;
  
  try {
    // Abre a planilha Consolidadora correta
    var ss = null;
    var spreadsheetId = "1blaDiVHRBXor-T9cG9shnOXkSnPJFIyEU-kXQL_p5hk"; // ID da planilha de Gestores
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (err) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    if (!ss) throw new Error("Não foi possível conectar à planilha.");

    // ==========================================
    // 1. ROTAS DE AUDIÊNCIAS CENTRAL
    // ==========================================
    
    // Rota: action=read (Leitura das Audiências)
    if (action === 'read') {
      var sheet = obterAbaResiliente(ss, 'REMESSA CENTRAL AUDIÊNCIA');
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'REMESSA CENTRAL AUDIÊNCIA' não encontrada." });
      }
      var values = sheet.getDataRange().getDisplayValues();
      return retornarJSONP(callback, { status: "ok", data: values });
    }
    
    // Rota: action=update (Gravação das Audiências)
    if (action === 'update') {
      var dataJson = e.parameter.data;
      if (!dataJson) throw new Error("Parâmetro 'data' ausente.");
      var updates = JSON.parse(decodeURIComponent(dataJson));
      
      var sheet = obterAbaResiliente(ss, 'REMESSA CENTRAL AUDIÊNCIA');
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'REMESSA CENTRAL AUDIÊNCIA' não encontrada." });
      }
      
      var values = sheet.getDataRange().getValues();
      var count = 0;
      
      for (var i = 1; i < values.length; i++) {
        var npuMaster = String(values[i][2]).replace(/\D/g, "");
        
        for (var npuKey in updates) {
          var cleanKey = npuKey.replace(/\D/g, "");
          if (npuMaster === cleanKey) {
            var situacao = updates[npuKey].situacao;
            var justificativa = updates[npuKey].justificativa;
            
            if (situacao === "APAGAR") situacao = "";
            if (justificativa === "APAGAR") justificativa = "";
            
            sheet.getRange(i + 1, 5).setValue(situacao);     // Coluna E
            sheet.getRange(i + 1, 6).setValue(justificativa); // Coluna F
            count++;
          }
        }
      }
      return retornarJSONP(callback, { status: "success", message: "Gravação efetuada: " + count + " linha(s) de audiência atualizada(s)." });
    }

    // ==========================================
    // 2. ROTAS DE CANCELAMENTO DE CONCLUSÕES
    // ==========================================

    // Rota: action=readCancelamento (Leitura dos Cancelamentos)
    if (action === "readCancelamento") {
      var sheet = obterAbaResiliente(ss, "CANCELAMENTO");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'CANCELAMENTO' não encontrada." });
      }
      var values = sheet.getDataRange().getDisplayValues();
      return retornarJSONP(callback, { status: "ok", data: values });
    }

    // Rota: action=updateCancelamento (Gravação dos Cancelamentos)
    if (action === "updateCancelamento") {
      var dataJson = e.parameter.data;
      if (!dataJson) throw new Error("Parâmetro 'data' ausente.");
      var updates = JSON.parse(decodeURIComponent(dataJson));
      
      var sheet = obterAbaResiliente(ss, "CANCELAMENTO");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'CANCELAMENTO' não encontrada." });
      }
      
      var allData = sheet.getDataRange().getValues();
      var count = 0;
      
      for (var chave in updates) {
        var partes = chave.split("_");
        var processo = partes[0];
        var servidora = partes.slice(1).join("_");
        
        for (var i = 1; i < allData.length; i++) {
          var processoSheet = String(allData[i][4] || "").trim();  // Coluna E
          var servidoraSheet = String(allData[i][9] || "").trim(); // Coluna J
          
          if (processoSheet === processo && servidoraSheet === servidora) {
            var cancelamento = updates[chave].cancelamento;
            var observacao = updates[chave].observacao;
            
            if (cancelamento === "APAGAR") cancelamento = "";
            if (observacao === "APAGAR") observacao = "";
            
            sheet.getRange(i + 1, 23).setValue(cancelamento); // Coluna W
            sheet.getRange(i + 1, 24).setValue(observacao);   // Coluna X
            count++;
            break;
          }
        }
      }
      return retornarJSONP(callback, { status: "success", message: "Gravação efetuada: " + count + " linha(s) de cancelamento atualizada(s)." });
    }

    // ==========================================
    // 3. ROTAS DE ERRO E VINCULAÇÃO (SERPREC)
    // ==========================================

    // Rota: action=readSerprec (Leitura dos Erros/Vinculações)
    if (action === "readSerprec") {
      var sheet = obterAbaResiliente(ss, "ERRO / VINC.");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'ERRO / VINC.' não encontrada." });
      }
      var values = sheet.getDataRange().getDisplayValues();
      return retornarJSONP(callback, { status: "ok", data: values });
    }

    // Rota: action=updateSerprec (Gravação dos Erros/Vinculações)
    if (action === "updateSerprec") {
      var dataJson = e.parameter.data;
      if (!dataJson) throw new Error("Parâmetro 'data' ausente.");
      var updates = JSON.parse(decodeURIComponent(dataJson));
      
      var sheet = obterAbaResiliente(ss, "ERRO / VINC.");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'ERRO / VINC.' não encontrada." });
      }
      
      var allData = sheet.getDataRange().getValues();
      var headers = allData[0].map(function(h) { return String(h || "").toUpperCase().trim(); });
      
      var idxDataRetorno = headers.indexOf("DATA DE RETORNO");
      if (idxDataRetorno < 0) idxDataRetorno = headers.indexOf("RETORNO");
      if (idxDataRetorno < 0) {
        idxDataRetorno = headers.findIndex(function(h) { return h.indexOf("RETORNO") !== -1 || h.indexOf("DATA") !== -1; });
      }
      if (idxDataRetorno < 0) idxDataRetorno = 6;
      var colNumDataRet = idxDataRetorno + 1;
      
      var count = 0;
      for (var npuKey in updates) {
        var cleanNpuKey = npuKey.replace(/\D/g, "");
        
        for (var i = 1; i < allData.length; i++) {
          var cleanNpuSheet = String(allData[i][1] || "").replace(/\D/g, "");
          
          if (cleanNpuSheet === cleanNpuKey) {
            var situacao = updates[npuKey].situacao;
            var analisado = updates[npuKey].analisado;
            var adicionarObs = updates[npuKey].adicionarObs;
            var dataRetorno = updates[npuKey].dataRetorno;
            
            if (situacao !== undefined) {
              if (situacao === "APAGAR") situacao = "";
              sheet.getRange(i + 1, 5).setValue(situacao);
            }
            
            if (analisado !== undefined) {
              if (analisado === "APAGAR") analisado = "";
              sheet.getRange(i + 1, 6).setValue(analisado);
            }
            
            if (dataRetorno !== undefined) {
              if (dataRetorno === "APAGAR") dataRetorno = "";
              sheet.getRange(i + 1, colNumDataRet).setValue(dataRetorno);
            }
            
            if (adicionarObs !== undefined && adicionarObs.trim() !== "") {
              var cellRange = sheet.getRange(i + 1, 9);
              var currentVal = String(cellRange.getValue() || "").trim();
              var newVal = currentVal ? currentVal + "\n" + adicionarObs : adicionarObs;
              cellRange.setValue(newVal);
            }
            
            count++;
            break;
          }
        }
      }
      return retornarJSONP(callback, { status: "success", message: "Gravação efetuada: " + count + " linha(s) da Serprec atualizada(s)." });
    }

    // Rota: action=apagarAnalisados (Apagar a coluna Analisado da planilha Gestores)
    if (action === "apagarAnalisados") {
      var supervisor = e.parameter.supervisor;
      var dataJson = e.parameter.data;
      var sheet = obterAbaResiliente(ss, "ERRO / VINC.");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'ERRO / VINC.' não encontrada." });
      }
      
      var allData = sheet.getDataRange().getValues();
      var headers = allData[0].map(function(h) { return String(h || "").toUpperCase().trim(); });
      var idxSupervisor = headers.indexOf("SUPERVISOR");
      if (idxSupervisor < 0) idxSupervisor = 3;
      var idxAnalisado = headers.indexOf("ANALISADO");
      if (idxAnalisado < 0) idxAnalisado = 5;

      var count = 0;
      var npusList = null;
      if (dataJson) {
        try { npusList = JSON.parse(decodeURIComponent(dataJson)); } catch(err){}
      }

      for (var i = 1; i < allData.length; i++) {
        var cleanNpuSheet = String(allData[i][1] || "").replace(/\D/g, "");
        var supervisorSheet = String(allData[i][idxSupervisor] || "").trim();
        
        var match = false;
        if (npusList && Array.isArray(npusList) && npusList.length > 0) {
          if (npusList.indexOf(cleanNpuSheet) !== -1) match = true;
        } else if (supervisor) {
          if (verificarOverlapSupervisor(supervisorSheet, supervisor)) match = true;
        } else {
          match = true;
        }

        if (match) {
          sheet.getRange(i + 1, idxAnalisado + 1).setValue("");
          count++;
        }
      }
      return retornarJSONP(callback, { status: "success", message: "Análises zeradas com sucesso: " + count + " linha(s) limpa(s)." });
    }

    // ==========================================
    // 4. ROTA DE LEITURA DOS GESTORES E SERVIDORES
    // ==========================================
    if (action === "readGestoresServidores") {
      var sheet = obterAbaResiliente(ss, "Lista Gestores Servidores");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'Lista Gestores Servidores' não encontrada." });
      }
      var values = sheet.getDataRange().getDisplayValues();
      return retornarJSONP(callback, { status: "ok", data: values });
    }

    // ==========================================
    // 5. ROTA DE MIGRAÇÃO (ERRO / VINCULAÇÃO)
    // ==========================================
    if (action === "migrarErroVinculacao") {
      var dataJson = e.parameter.data;
      if (!dataJson) throw new Error("Parâmetro 'data' ausente.");
      
      var registros = JSON.parse(decodeURIComponent(dataJson));
      var sheet = obterAbaResiliente(ss, "ERRO / VINC.");
      
      if (!sheet) {
        var nomesAbas = ss.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
        throw new Error("Aba 'ERRO / VINC.' não encontrada na planilha. Abas existentes no arquivo: " + nomesAbas);
      }
      
      registros.forEach(function(reg) {
        var tipo = reg.tipo;
        if (!tipo) {
          tipo = (reg.obsDestino && reg.obsDestino.trim() !== "") ? "ERRO DE FLUXO" : "SISCONDJ";
        }
        
        var origemFormatada = "";
        var obsFormatada = "";
        
        if (tipo === "ERRO DE FLUXO") {
          origemFormatada = "ERRO DE FLUXO/SISTEMA (informar nº do chamado na obs)";
          obsFormatada = "ERRO DE FLUXO (Nº do Chamado) - " + (reg.obsDestino || "");
        } else {
          origemFormatada = "SISCONDJ - VINCULAÇÃO DE CONTA";
          obsFormatada = "SISCONDJ - VINCULAÇÃO DE CONTA";
        }
        
        sheet.appendRow([
          reg.vara || "",                   // Col A (VARA)
          reg.processo || "",               // Col B (PROCESSO)
          origemFormatada,                  // Col C (ORIGEM)
          reg.duplaGestores || "",          // Col D (DUPLA GESTORES)
          "",                               // Col E (SITUAÇÃO)
          "",                               // Col F (ANALISADO)
          "",                               // Col G (DATA DE RETORNO)
          reg.servidor || "",               // Col H (SERVIDOR QUE CUMPRIU)
          obsFormatada,                     // Col I (OBS)
          "=VLOOKUP(B:B, xdias!B:D, 3, 0)", // Col J (TAREFA - Fórmula)
          "=VLOOKUP(B:B, xdias!B:F, 5, 0)", // Col K (DIAS - Fórmula)
          reg.dataInclusao || ""            // Col L (DATA DE INCLUSÃO)
        ]);
      });
      
      return retornarJSONP(callback, { status: "success", message: registros.length + " registros migrados com sucesso." });
    }

    // ====================================================
    // 6. MIGRAR CENTRAL DE AGILIZAÇÃO / PRAZO ABERTO / SUSPENSOS / ARQUIVADO PROVISÓRIO
    // ====================================================
    if (action === "migrarCentralAgilizacaoSuspensos") {
      var targetSpreadsheetId = e.parameter.spreadsheetId || "1yOAupHFi0ERglZPTJ4yrXHPwBP8oemOfLAKFJ9zT9Po";
      var dataStr = e.parameter.data;

      if (!dataStr) {
        throw new Error("Parâmetro de dados ausente (data).");
      }

      var payload = JSON.parse(decodeURIComponent(dataStr));
      var ssTarget = SpreadsheetApp.openById(targetSpreadsheetId);

      var countAguardar = 0;
      var countSuspensos = 0;

      // Cache das abas e suas próximas linhas disponíveis baseadas na Coluna B (Processo)
      var cacheProximaLinha = {};

      for (var j = 0; j < payload.length; j++) {
        var item = payload[j];
        var destAba = item.destAba;
        var sheet = obterAbaResiliente(ssTarget, destAba);

        if (!sheet) {
          throw new Error("Aba '" + destAba + "' não encontrada na planilha de controle.");
        }

        // Se ainda não calculou a primeira linha disponível na Coluna B (2) desta aba
        if (!cacheProximaLinha[destAba]) {
          cacheProximaLinha[destAba] = obterProximaLinhaVaziaPorColuna(sheet, 2);
        }

        var proximaLinha = cacheProximaLinha[destAba];

        // Grava especificamente nas Colunas A a E (1 a 5) daquela linha sem interferir nas colunas G, J, K...
        var rangeAlvo = sheet.getRange(proximaLinha, 1, 1, 5);
        try {
          rangeAlvo.clearDataValidations();
        } catch (eVal) {}
        
        rangeAlvo.setValues([[
          item.vara || "",
          item.processo || "",
          item.situacao || "",
          item.dataRetorno || "",
          item.obs || ""
        ]]);

        // Incrementa o ponteiro para o próximo item
        cacheProximaLinha[destAba] = proximaLinha + 1;

        if (destAba.indexOf("AGUARDAR") !== -1) {
          countAguardar++;
        } else {
          countSuspensos++;
        }
      }

      return retornarJSONP(callback, {
        status: "success",
        message: "Processos migrados com sucesso! Aba Aguardar: " + countAguardar + ", Aba Suspensos/Arq: " + countSuspensos
      });
    }

    // ==========================================
    // 7. ROTA DE MIGRAÇÃO (RECJUD / LEILÃO)
    // ==========================================
    if (action === "migrarRecjudLeilao") {
      var targetSpreadsheetId = e.parameter.spreadsheetId;
      var dataStr = e.parameter.data;
      
      if (!targetSpreadsheetId) {
        throw new Error("ID da planilha não especificado (spreadsheetId).");
      }
      if (!dataStr) {
        throw new Error("Parâmetro de dados ausente (data).");
      }

      var payload = JSON.parse(decodeURIComponent(dataStr));
      var ssTarget = SpreadsheetApp.openById(targetSpreadsheetId);
      
      var countRecjud = 0;
      var countLeilao = 0;

      var sheetRecjud = obterAbaResiliente(ssTarget, "Recjud");
      var sheetLeilao = obterAbaResiliente(ssTarget, "Leilão");

      var hoje = new Date();
      var hojeFormatado = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

      for (var i = 0; i < payload.length; i++) {
        var item = payload[i];
        var tipo = String(item.tipo).toUpperCase();

        // Se for RECJUD ou AMBOS, adiciona na aba 'Recjud'
        if (tipo === "RECJUD" || tipo === "AMBOS") {
          if (!sheetRecjud) {
            throw new Error("Aba 'Recjud' não encontrada na planilha especificada.");
          }
          sheetRecjud.appendRow([
            item.vara || "",
            item.processo || "",
            item.observacao || "",
            hojeFormatado
          ]);
          countRecjud++;
        }

        // Se for LEILÃO ou AMBOS, adiciona na aba 'Leilão'
        if (tipo === "LEILÃO" || tipo === "LEILAO" || tipo === "AMBOS") {
          if (!sheetLeilao) {
            throw new Error("Aba 'Leilão' não encontrada na planilha especificada.");
          }
          sheetLeilao.appendRow([
            item.vara || "",
            item.processo || ""
          ]);
          countLeilao++;
        }
      }

      return retornarJSONP(callback, { 
        status: "success", 
        message: "Processos migrados com sucesso! Recjud: " + countRecjud + ", Leilão: " + countLeilao 
      });
    }

    // ==========================================
    // 8. REGISTRO E LEITURA DE DECISÕES (NÃO MIGRAR / DISPENSA)
    // ==========================================
    if (action === "registrarDecisaoDispensa") {
      var npu = e.parameter.npu;
      var card = e.parameter.card;
      var data = e.parameter.data;
      var sheetDecisoes = ss.getSheetByName("HISTORICO_DECISOES");
      if (!sheetDecisoes) {
        sheetDecisoes = ss.insertSheet("HISTORICO_DECISOES");
        sheetDecisoes.appendRow(["NPU_BASE", "ACAO", "DATA", "CARD"]);
      }
      sheetDecisoes.appendRow([npu, "NAO_MIGRAR", data, card]);
      return retornarJSONP(callback, { status: "success", message: "Decisão registrada." });
    }

    if (action === "removerDecisaoDispensa") {
      var npu = e.parameter.npu;
      var sheetDecisoes = ss.getSheetByName("HISTORICO_DECISOES");
      if (sheetDecisoes && sheetDecisoes.getLastRow() > 1) {
        var values = sheetDecisoes.getRange(2, 1, sheetDecisoes.getLastRow() - 1, 4).getValues();
        for (var k = values.length - 1; k >= 0; k--) {
          if (String(values[k][0]) === String(npu)) {
            sheetDecisoes.deleteRow(k + 2);
          }
        }
      }
      return retornarJSONP(callback, { status: "success", message: "Decisão removida." });
    }

    if (action === "lerDecisoesDispensa") {
      var sheetDecisoes = ss.getSheetByName("HISTORICO_DECISOES");
      var list = [];
      if (sheetDecisoes && sheetDecisoes.getLastRow() > 1) {
        var values = sheetDecisoes.getRange(2, 1, sheetDecisoes.getLastRow() - 1, 4).getValues();
        for (var k = 0; k < values.length; k++) {
          list.push({
            npu13: String(values[k][0]),
            acao: String(values[k][1]),
            data: String(values[k][2]),
            card: String(values[k][3])
          });
        }
      }
      return retornarJSONP(callback, { status: "success", data: list });
    }

    throw new Error("Ação '" + action + "' não reconhecida.");
    
  } catch (error) {
    return retornarJSONP(callback, { status: "error", message: error.toString() });
  }
}

/**
 * Encontra a primeira linha vazia de uma aba com base em uma coluna específica (ex: Coluna B = 2),
 * evitando ignorar linhas por conta de fórmulas em colunas distantes (ex: Colunas G, J, K...).
 */
function obterProximaLinhaVaziaPorColuna(sheet, colIndex) {
  var colValues = sheet.getRange(1, colIndex, sheet.getLastRow() + 50, 1).getValues();
  for (var i = 0; i < colValues.length; i++) {
    var val = String(colValues[i][0] || "").trim();
    if (val === "" && i > 0) { // i > 0 pula o cabeçalho na linha 1
      return i + 1;
    }
  }
  return colValues.length + 1;
}

/**
 * Função auxiliar para buscar a aba de forma resiliente, ignorando espaços e pontos extras.
 */
function obterAbaResiliente(ss, targetName) {
  var sheet = ss.getSheetByName(targetName);
  if (sheet) return sheet;
  
  var normalizar = function(name) {
    if (!name) return "";
    return name.trim()
               .toUpperCase()
               .replace(/\s+/g, "")
               .replace(/\.$/, "");
  };
  
  var targetNormalized = normalizar(targetName);
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (normalizar(allSheets[i].getName()) === targetNormalized) {
      return allSheets[i];
    }
  }
  return null;
}

/**
 * Retorna os dados formatados em JSONP
 */
function retornarJSONP(callback, objeto) {
  var outputText = callback + "(" + JSON.stringify(objeto) + ");";
  return ContentService.createTextOutput(outputText).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
