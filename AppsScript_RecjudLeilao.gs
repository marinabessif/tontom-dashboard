/**
 * Recebe as requisições GET (JSONP) do painel Tontom Dashboard.
 * Versão Consolidada: Audiências, Cancelamento, Serprec (ERRO / VINC.) e Migração.
 */
function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) || "callback";
  var action = (e && e.parameter && e.parameter.action) || "";
  
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
      var sheet = obterAbaResiliente(ss, "ERRO / VINC./ TED DEV.") || obterAbaResiliente(ss, "ERRO / VINC.");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'ERRO / VINC./ TED DEV.' ou 'ERRO / VINC.' não encontrada." });
      }
      var values = sheet.getDataRange().getDisplayValues();
      return retornarJSONP(callback, { status: "ok", data: values });
    }

    // Rota: action=updateSerprec (Gravação dos Erros/Vinculações)
    if (action === "updateSerprec") {
      var dataJson = e.parameter.data;
      if (!dataJson) throw new Error("Parâmetro 'data' ausente.");
      var updates = JSON.parse(decodeURIComponent(dataJson));
      
      var sheet = obterAbaResiliente(ss, "ERRO / VINC./ TED DEV.") || obterAbaResiliente(ss, "ERRO / VINC.");
      if (!sheet) {
        return retornarJSONP(callback, { status: "error", message: "Aba 'ERRO / VINC./ TED DEV.' ou 'ERRO / VINC.' não encontrada." });
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
      var sheet = obterAbaResiliente(ss, "ERRO / VINC./ TED DEV.") || obterAbaResiliente(ss, "ERRO / VINC / TED DEV.") || obterAbaResiliente(ss, "ERRO / VINC.");
      
      if (!sheet) {
        var nomesAbas = ss.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
        throw new Error("Aba 'ERRO / VINC./ TED DEV.' não encontrada na planilha. Abas existentes no arquivo: " + nomesAbas);
      }
      
      registros.forEach(function(reg) {
        var tipo = reg.tipo;
        if (!tipo) {
          tipo = (reg.obsDestino && reg.obsDestino.trim() !== "") ? "ERRO DE FLUXO" : "SISCONDJ";
        }
        
        var origemFormatada = "";
        var obsBase = "";
        
        if (tipo === "ERRO DE FLUXO") {
          origemFormatada = "ERRO DE FLUXO/SISTEMA (informar nº do chamado na obs)";
          obsBase = "ERRO DE FLUXO (Nº do Chamado) - " + (reg.obsDestino || "");
        } else {
          origemFormatada = "SISCONDJ - VINCULAÇÃO DE CONTA";
          obsBase = "SISCONDJ - VINCULAÇÃO DE CONTA";
        }

        // Tag de migração Tontom no topo da observação
        var tagMigracao = reg.tagMigracao || "";
        var textoObsCompleto = tagMigracao ? (obsBase ? tagMigracao + "\n" + obsBase : tagMigracao) : obsBase;
        
        sheet.appendRow([
          reg.vara || "",                   // Col A (VARA)
          reg.processo || "",               // Col B (PROCESSO)
          origemFormatada,                  // Col C (ORIGEM)
          "",                               // Col D (SUPERVISOR - deixada vazia para a fórmula da planilha atuar)
          "",                               // Col E (SITUAÇÃO)
          "",                               // Col F (ANALISADO)
          "",                               // Col G (DATA DE RETORNO)
          reg.servidor || "",               // Col H (SERVIDOR QUE CUMPRIU)
          textoObsCompleto,                 // Col I (OBS)
          "=VLOOKUP(B:B, xdias!B:D, 3, 0)", // Col J (TAREFA - Fórmula)
          "=VLOOKUP(B:B, xdias!B:F, 5, 0)", // Col K (DIAS - Fórmula)
          reg.dataInclusao || ""            // Col L (DATA DE INCLUSÃO)
        ]);

        var newRow = sheet.getLastRow();

        // 1. Limpa Coluna D para permitir expansão de fórmula da planilha
        try {
          sheet.getRange(newRow, 4).clearContent();
        } catch (eD) {}

        // 2. Formata Coluna I com RichText para a parte 2 (nome planilha e data de modificação) vir com fonte menor
        var part1 = reg.tagMigracaoParte1 || "";
        var part2 = reg.tagMigracaoParte2 || "";
        if (part1 && part2 && textoObsCompleto.indexOf(part1) === 0) {
          try {
            var startIndex = part1.length;
            var endIndex = part1.length + part2.length;
            var styleMenor = SpreadsheetApp.newTextStyle().setFontSize(8).build();
            var richText = SpreadsheetApp.newRichTextValue()
              .setText(textoObsCompleto)
              .setTextStyle(startIndex, endIndex, styleMenor)
              .build();
            sheet.getRange(newRow, 9).setRichTextValue(richText);
          } catch (eRich) {
            // Se falhar o RichText, o texto já foi gravado pelo appendRow
          }
        }
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

        // Garante que a linha e as 5 colunas existam na aba antes de gravar
        garantirDimensoes(sheet, proximaLinha, 5);

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

    // ==========================================
    // 9. ROTAS DE CONFIGURAÇÃO DE EQUIPES E SUPERVISORES (NUVEM)
    // ==========================================
    if (action === "saveEquipesConfig") {
      var eqJson = e.parameter.equipes;
      var supJson = e.parameter.supervisores;
      
      var sheet = ss.getSheetByName("CONFIG_EQUIPES");
      if (!sheet) {
        sheet = ss.insertSheet("CONFIG_EQUIPES");
        sheet.appendRow(["KEY", "VALUE"]);
      }
      
      // Limpa dados antigos (mantém cabeçalho)
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
      }
      
      if (eqJson) {
        sheet.getRange(2, 1).setValue("equipes");
        sheet.getRange(2, 2).setValue(decodeURIComponent(eqJson));
      }
      if (supJson) {
        sheet.getRange(3, 1).setValue("supervisores");
        sheet.getRange(3, 2).setValue(decodeURIComponent(supJson));
      }
      
      return retornarJSONP(callback, { status: "success", message: "Configuração de equipes salva na nuvem." });
    }

    if (action === "readEquipesConfig") {
      var sheet = ss.getSheetByName("CONFIG_EQUIPES");
      if (!sheet || sheet.getLastRow() < 2) {
        return retornarJSONP(callback, { status: "ok", equipes: null, supervisores: null });
      }
      
      var data = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 2), 2).getValues();
      var result = { status: "ok", equipes: null, supervisores: null };
      
      for (var i = 0; i < data.length; i++) {
        var key = String(data[i][0]).trim();
        var val = String(data[i][1]).trim();
        if (key === "equipes" && val) {
          try { result.equipes = JSON.parse(val); } catch(err) {}
        }
        if (key === "supervisores" && val) {
          try { result.supervisores = JSON.parse(val); } catch(err) {}
        }
      }
      
      return retornarJSONP(callback, result);
    }

    // ==========================================
    // 10. ROTAS DE CONFIGURAÇÃO DE FERIADOS (NUVEM)
    // ==========================================
    if (action === "saveFeriadosConfig") {
      var feriadosText = e.parameter.feriados || "";
      
      var sheet = ss.getSheetByName("CONFIG_FERIADOS");
      if (!sheet) {
        sheet = ss.insertSheet("CONFIG_FERIADOS");
        sheet.appendRow(["KEY", "VALUE"]);
      }
      
      // Limpa dados antigos (mantém cabeçalho)
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
      }
      
      sheet.getRange(2, 1).setValue("feriados");
      sheet.getRange(2, 2).setValue(decodeURIComponent(feriadosText));
      
      return retornarJSONP(callback, { status: "success", message: "Feriados salvos na nuvem." });
    }

    if (action === "readFeriadosConfig") {
      var sheet = ss.getSheetByName("CONFIG_FERIADOS");
      if (!sheet || sheet.getLastRow() < 2) {
        return retornarJSONP(callback, { status: "ok", feriados: null });
      }
      
      var val = String(sheet.getRange(2, 2).getValue() || "").trim();
      return retornarJSONP(callback, { status: "ok", feriados: val });
    }

    // ==========================================
    // 11. MIGRAÇÃO CONTROLE SERPREC (PRECATÓRIO / RPV)
    // ==========================================
    if (action === "migrarControleSerprec") {
      var targetSpreadsheetId = e.parameter.spreadsheetId || "1yOAupHFi0ERglZPTJ4yrXHPwBP8oemOfLAKFJ9zT9Po";
      var dataStr = e.parameter.data;

      if (!dataStr) {
        throw new Error("Parâmetro de dados ausente (data).");
      }

      var payload = JSON.parse(decodeURIComponent(dataStr));
      var ssTarget = SpreadsheetApp.openById(targetSpreadsheetId);
      var sheet = obterAbaResiliente(ssTarget, "SERPREC");

      if (!sheet) {
        var nomesAbas = ssTarget.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
        throw new Error("Aba 'SERPREC' não encontrada. Abas existentes: " + nomesAbas);
      }

      // Encontra a próxima linha vazia com base na Coluna B (Processo),
      // ignorando linhas que só têm fórmulas em G-J
      var proximaLinha = obterProximaLinhaVaziaPorColuna(sheet, 2);

      var hoje = new Date();
      var hojeFormatado = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");

      for (var j = 0; j < payload.length; j++) {
        var item = payload[j];

        garantirDimensoes(sheet, proximaLinha, 11);

        // Coluna A (1): Vara
        sheet.getRange(proximaLinha, 1).setValue(item.vara || "");
        // Coluna B (2): Processo
        sheet.getRange(proximaLinha, 2).setValue(item.processo || "");
        // Coluna K (11): data_Tontom (data da migração)
        sheet.getRange(proximaLinha, 11).setValue(hojeFormatado);

        proximaLinha++;
      }

      return retornarJSONP(callback, {
        status: "success",
        message: payload.length + " processo(s) migrado(s) para a aba SERPREC com sucesso."
      });
    }

    // ==========================================
    // 12. MIGRAÇÃO PRAZO PARA PLANILHA SALDO (PRAZO_TONTOM)
    // ==========================================
    if (action === "migrarPrazoSaldo") {
      var targetSpreadsheetId = e.parameter.spreadsheetId || "1toiSlIHFAZX0l_6JRPh4xQHQWJ1ZBXzDJP3wzDErVBA";
      var nomePlanilhaOrigem = e.parameter.nomePlanilhaOrigem || "";
      var dataStr = e.parameter.data;

      if (!dataStr) {
        throw new Error("Parâmetro de dados ausente (data).");
      }

      var payload = JSON.parse(decodeURIComponent(dataStr));
      var ssTarget = SpreadsheetApp.openById(targetSpreadsheetId);
      var sheet = obterAbaResiliente(ssTarget, "PRAZO_TONTOM");

      if (!sheet) {
        var nomesAbas = ssTarget.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
        throw new Error("Aba 'PRAZO_TONTOM' não encontrada. Abas existentes: " + nomesAbas);
      }

      // Encontra a próxima linha vazia com base na Coluna B (Processo)
      var proximaLinha = obterProximaLinhaVaziaPorColuna(sheet, 2);

      var hoje = new Date();
      var hojeFormatado = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
      var hojeCompleto = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yy");
      var horaCompleta = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "HH:mm");

      for (var j = 0; j < payload.length; j++) {
        var item = payload[j];

        garantirDimensoes(sheet, proximaLinha, 6);

        // Coluna A (1): Vara
        sheet.getRange(proximaLinha, 1).setValue(item.vara || "");
        // Coluna B (2): Processo
        sheet.getRange(proximaLinha, 2).setValue(item.processo || "");
        // Coluna C (3): Data de Retorno
        sheet.getRange(proximaLinha, 3).setValue(item.dataRetorno || "");
        // Colunas D e E: FÓRMULAS — não tocar
        // Coluna F (6): data_Tontom (data da migração)
        sheet.getRange(proximaLinha, 6).setValue(hojeFormatado);

        proximaLinha++;
      }

      // Atualiza A1 com a informação da última atualização
      var infoA1 = "Última atualização da planilha " + nomePlanilhaOrigem + " em " + hojeCompleto + " às " + horaCompleta;
      sheet.getRange(1, 1).setValue(infoA1);

      return retornarJSONP(callback, {
        status: "success",
        message: payload.length + " processo(s) migrado(s) para a aba PRAZO_TONTOM com sucesso."
      });
    }

    // ==========================================
    // 7. ROTAS DE X-DIAS (ACOMPANHAMENTO DE ASSINATURAS)
    // ==========================================
    // 3.5. ROTA DE ALIMENTAÇÃO SIMAP (PLANILHA SOBRESCRITA A CADA 30 MIN)
    // ==========================================
    if (action === "obterSimapDrive") {
      var idPlanilhaSimap = (e && e.parameter && e.parameter.sheetId) || "1zZohMKhAUHvhYhCp4W-YPgEm7wj8eBQ-MEsMmVw6VJo";
      var gidSimap = (e && e.parameter && e.parameter.gid) || "1608349017";
      
      var ssSimap = null;
      try {
        ssSimap = SpreadsheetApp.openById(idPlanilhaSimap);
      } catch (errOpen) {
        return retornarJSONP(callback, { status: "error", message: "Não foi possível abrir a planilha SIMAP no Drive: " + errOpen.message });
      }
      
      var sheet = null;
      var sheets = ssSimap.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        if (String(sheets[s].getSheetId()) === String(gidSimap)) {
          sheet = sheets[s];
          break;
        }
      }
      if (!sheet) sheet = sheets[0];
      
      var valores = sheet.getDataRange().getDisplayValues();
      var ultimaMod = new Date().getTime();
      var nomeArq = ssSimap.getName();
      try {
        var fileObj = DriveApp.getFileById(idPlanilhaSimap);
        ultimaMod = fileObj.getLastUpdated().getTime();
        nomeArq = fileObj.getName();
      } catch (eFile) {}
      
      var sheetNome = sheet.getName();

      return retornarJSONP(callback, {
        status: "ok",
        success: true,
        nomeArquivo: nomeArq,
        sheetNome: sheetNome,
        ultimaDataArq: ultimaMod,
        matriz: valores,
        linhas: valores,
        totalLinhas: valores.length > 1 ? (valores.length - 1) : 0
      });
    }

    // ==========================================
    // 4. ROTAS DE X-DIAS / SIMAP AUTO (PASTA xdias_auto)
    // ==========================================
    if (action === "obterXdias" || action === "obterSimapAuto") {
      var pastaNome = "xdias_auto";
      var folderId = "1HYtNMyvy8gJykFXDehXtcZYKDJvZdJ9l";
      var pasta = null;
      try {
        pasta = DriveApp.getFolderById(folderId);
      } catch (err) {
        var pastas = DriveApp.getFoldersByName(pastaNome);
        if (pastas.hasNext()) pasta = pastas.next();
      }
      
      if (!pasta) {
        return retornarJSONP(callback, { status: "error", message: "Pasta de X-Dias (xdias_auto) não encontrada no Google Drive." });
      }
      
      var arquivos = pasta.getFiles();
      var ultimoArquivo = null;
      var ultimaDataArq = 0;
      
      while (arquivos.hasNext()) {
        var arquivo = arquivos.next();
        var tUpdate = arquivo.getLastUpdated().getTime();
        var tNome = extrairTimestampNomeArquivo(arquivo.getName());
        var tScore = Math.max(tUpdate, tNome);
        if (tScore > ultimaDataArq) {
          ultimaDataArq = tScore;
          ultimoArquivo = arquivo;
        }
      }
      
      if (!ultimoArquivo) {
        return retornarJSONP(callback, { status: "error", message: "Nenhum arquivo X-Dias/SIMAP encontrado na pasta xdias_auto." });
      }
      
      var nomeArq = ultimoArquivo.getName();
      var matrizDados = [];
      var base64Dados = "";
      
      try {
        var planilhaExterna = SpreadsheetApp.openById(ultimoArquivo.getId()).getSheets()[0];
        matrizDados = planilhaExterna.getDataRange().getValues();
      } catch (errOpen) {
        // Se não for Google Sheets nativo (ex: .xlsx, .csv)
        var mime = ultimoArquivo.getMimeType();
        if (mime === MimeType.CSV || nomeArq.toLowerCase().endsWith(".csv")) {
          try {
            var csvStr = ultimoArquivo.getBlob().getDataAsString("ISO-8859-1");
            matrizDados = Utilities.parseCsv(csvStr);
          } catch (eCsv) {
            var csvStr2 = ultimoArquivo.getBlob().getDataAsString("UTF-8");
            matrizDados = Utilities.parseCsv(csvStr2);
          }
        } else {
          // Arquivo Excel (.xlsx / .xls) ou binário: envia em base64 para decodificação rápida no cliente com SheetJS
          base64Dados = Utilities.base64Encode(ultimoArquivo.getBlob().getBytes());
        }
      }
      
      var varasTempInfo = extrairDistribuicaoVarasTemporarias();

      return retornarJSONP(callback, {
        status: "ok",
        success: true,
        nomeArquivo: nomeArq,
        ultimaDataArq: ultimaDataArq,
        ultimaData: ultimaDataArq,
        matriz: matrizDados,
        linhas: matrizDados,
        base64: base64Dados,
        totalLinhas: matrizDados.length > 1 ? (matrizDados.length - 1) : 0,
        varasTemporariasInfo: varasTempInfo,
        amarelinhoInfo: extrairEscalaAmarelinho()
      });
    }

    if (action === "obterVarasPermanentes" || action === "obterNucleoProcessamento") {
      var idPlanilhaOrganograma = "1Rq_FIlUmpXyDK0alQ21_Z5b1bHMT_HlzU66QUGdr1LQ";
      var ssOrg = null;
      try {
        ssOrg = SpreadsheetApp.openById(idPlanilhaOrganograma);
      } catch (errOrg) {
        return retornarJSONP(callback, { status: "error", message: "Não foi possível abrir a planilha de organograma: " + errOrg.message });
      }
      var sheet = ssOrg.getSheetByName("NÚCLEO DE PROCESSAMENTO") || ssOrg.getSheets()[0];
      var dados = sheet.getDataRange().getValues();
      var mapaEquipes = {};
      if (dados && dados.length > 1) {
        var cab = dados[0];
        // As varas permanentes ficam estritamente entre as linhas 2 e 9 da planilha (índices 1 a 8)
        var maxLinha = Math.min(dados.length, 9);
        for (var c = 0; c < cab.length; c++) {
          var cel = String(cab[c] || "").trim();
          var mEq = cel.match(/Equipe\s*(\d+)/i);
          if (mEq) {
            var chaveEq = "Equipe " + mEq[1];
            var listaVaras = [];
            for (var r = 1; r < maxLinha; r++) {
              var vCod = String(dados[r][c] || "").trim().toUpperCase();
              // Apenas códigos válidos de vara (ex: 2A, 11A, 28B), excluindo nomes de servidores
              if (vCod && /^\d{1,2}[A-B]$/.test(vCod)) {
                listaVaras.push(vCod);
              }
            }
            mapaEquipes[chaveEq] = listaVaras;
          }
        }
      }
      var amarelinhoData = extrairEscalaAmarelinho(ssOrg);
      return retornarJSONP(callback, { status: "ok", success: true, dados: dados, matriz: dados, equipes: mapaEquipes, amarelinhoInfo: amarelinhoData });
    }

    if (action === "obterVarasTemporarias") {
      var distInfo = extrairDistribuicaoVarasTemporarias();
      return retornarJSONP(callback, distInfo);
    }

    if (action === "obterSupervisorAmarelinho") {
      var amarelinhoResp = extrairEscalaAmarelinho();
      return retornarJSONP(callback, amarelinhoResp);
    }

    if (action === "obterAusencias") {
      var ausenciasResp = extrairAusenciasOrganograma();
      return retornarJSONP(callback, ausenciasResp);
    }

    throw new Error("Ação '" + action + "' não reconhecida.");
    
  } catch (error) {
    return retornarJSONP(callback, { status: "error", message: error.toString() });
  }
}

/**
 * Garante que a aba possua dimensões suficientes para a linha e coluna solicitadas.
 */
function garantirDimensoes(sheet, minRow, minCol) {
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if (minRow > maxRows) {
    sheet.insertRowsAfter(maxRows, Math.max(minRow - maxRows, 20));
  }
  if (minCol > maxCols) {
    sheet.insertColumnsAfter(maxCols, minCol - maxCols);
  }
}

/**
 * Encontra a primeira linha vazia de uma aba com base em uma coluna específica (ex: Coluna B = 2),
 * evitando ignorar linhas por conta de fórmulas em colunas distantes (ex: Colunas G, J, K...).
 * Lê apenas as linhas que já existem no intervalo da planilha, evitando erro de coordenadas fora das dimensões.
 */
function obterProximaLinhaVaziaPorColuna(sheet, colIndex) {
  var lastRow = sheet.getLastRow();
  var maxCols = sheet.getMaxColumns();
  
  if (colIndex > maxCols) {
    sheet.insertColumnsAfter(maxCols, colIndex - maxCols);
  }

  if (lastRow <= 1) {
    return 2;
  }

  // Lê apenas as linhas existentes até lastRow (nunca ultrapassa getMaxRows)
  var colValues = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  for (var i = 1; i < colValues.length; i++) { // i = 1 pula o cabeçalho na linha 1
    var val = String(colValues[i][0] || "").trim();
    if (val === "") {
      return i + 1;
    }
  }
  return lastRow + 1;
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


/**
 * Extrai a distribuição de varas temporárias por subperíodo
 * da planilha de distribuição de ausências (aba imediatamente após 'Config_Duplas').
 */
function extrairDistribuicaoVarasTemporarias() {
  try {
    var idPlanilhaDist = "14Nsr0d5bQOPI58-Wzzp8xrDK__7kqopsR6bAHgw2QL0";
    var ssDist = SpreadsheetApp.openById(idPlanilhaDist);
    var sheets = ssDist.getSheets();
    var idxConfig = -1;
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().trim().toLowerCase() === "config_duplas") {
        idxConfig = s;
        break;
      }
    }

    var targetSheet = null;
    if (idxConfig !== -1 && idxConfig + 1 < sheets.length) {
      targetSheet = sheets[idxConfig + 1];
    } else if (sheets.length > 1) {
      targetSheet = sheets[1];
    } else {
      targetSheet = sheets[0];
    }

    var nomeAba = targetSheet.getName();
    var dados = targetSheet.getDataRange().getValues();

    var subperiodos = [];
    var currentSub = null;
    var readingTable = false;
    var colAcumuladas = 6;

    for (var r = 0; r < dados.length; r++) {
      var row = dados[r];
      if (!row) continue;
      var colA = String(row[0] || "").trim();
      var colB = String(row[1] || "").trim();

      // Identifica início de subperíodo (ex: "BLOCO 1 - SUBPERÍODO:")
      if (colA.toUpperCase().indexOf("SUBPER") !== -1) {
        if (currentSub) {
          subperiodos.push(currentSub);
        }

        var dtInicio = null;
        var dtFim = null;
        var datesMatch = colB.match(/(\d{2})[\/\.-](\d{2})[\/\.-](\d{2,4})/g);
        if (datesMatch && datesMatch.length > 0) {
          var p1 = datesMatch[0].split(/[\/\.-]/);
          var y1 = p1[2].length === 2 ? ("20" + p1[2]) : p1[2];
          dtInicio = y1 + "-" + p1[1] + "-" + p1[0];

          var p2 = datesMatch[datesMatch.length - 1].split(/[\/\.-]/);
          var y2 = p2[2].length === 2 ? ("20" + p2[2]) : p2[2];
          dtFim = y2 + "-" + p2[1] + "-" + p2[0];
        }

        currentSub = {
          bloco: colA,
          textoPeriodo: colB,
          dataInicio: dtInicio,
          dataFim: dtFim,
          ausentes: "",
          equipes: {}
        };
        readingTable = false;
        continue;
      }

      if (currentSub && colA.toUpperCase().indexOf("AUSENTE") !== -1) {
        currentSub.ausentes = colB;
        continue;
      }

      if (currentSub && (colA.toUpperCase().indexOf("DUPLA") !== -1 || colA.toUpperCase().indexOf("ID") !== -1)) {
        for (var c = 0; c < row.length; c++) {
          if (String(row[c] || "").toUpperCase().indexOf("ACUMULADA") !== -1) {
            colAcumuladas = c;
            break;
          }
        }
        readingTable = true;
        continue;
      }

      if (currentSub && readingTable) {
        var mId = colA.match(/^(?:EQUIPE\s*)?(\d{1,2}|SCGA)$/i);
        if (mId) {
          var idVal = mId[1].toUpperCase();
          var chaveEquipe = (idVal !== "SCGA") ? ("Equipe " + idVal) : "SCGA";

          var unidadesTxt = String(row[colAcumuladas] || "").trim();
          var varas = [];
          if (unidadesTxt && unidadesTxt.toUpperCase() !== "NENHUMA" && unidadesTxt !== "-" && unidadesTxt.toUpperCase() !== "NENHUM") {
            var tokens = unidadesTxt.split(/[,;\s]+/);
            for (var t = 0; t < tokens.length; t++) {
              var tk = tokens[t].trim().toUpperCase();
              if (/^\d{1,2}[A-B]$/.test(tk)) {
                varas.push(tk);
              }
            }
          }
          currentSub.equipes[chaveEquipe] = varas;
        }
      }
    }

    if (currentSub) {
      subperiodos.push(currentSub);
    }

    return {
      status: "ok",
      success: true,
      nomeAba: nomeAba,
      subperiodos: subperiodos
    };
  } catch (err) {
    return {
      status: "error",
      message: err.toString(),
      subperiodos: []
    };
  }
}


/**
 * Extrai a escala de Supervisores Amarelinhos da aba 'Amarelinho e Redistribuições'
 * da planilha de Organograma (1Rq_FIlUmpXyDK0alQ21_Z5b1bHMT_HlzU66QUGdr1LQ).
 */
function extrairEscalaAmarelinho(ssOrg) {
  try {
    var idPlanilhaOrg = "1Rq_FIlUmpXyDK0alQ21_Z5b1bHMT_HlzU66QUGdr1LQ";
    if (!ssOrg) {
      ssOrg = SpreadsheetApp.openById(idPlanilhaOrg);
    }
    var sheet = ssOrg.getSheetByName("Amarelinho e Redistribuições") || ssOrg.getSheetByName("Amarelinho");
    if (!sheet) return { status: "error", message: "Aba 'Amarelinho e Redistribuições' não encontrada.", escala: {} };

    var valores = sheet.getDataRange().getValues();
    var escala = {};
    var lista = [];

    for (var i = 0; i < valores.length; i++) {
      var cellA = valores[i][0];
      var sup = String(valores[i][1] || "").trim();
      if (!cellA || !sup) continue;

      var dt = null;
      if (cellA instanceof Date && !isNaN(cellA.getTime())) {
        dt = cellA;
      } else {
        var strA = String(cellA).trim();
        var match = strA.match(/(\d{1,2})[\/\.-](\d{1,2})(?:[\/\.-](\d{2,4}))?/);
        if (match) {
          var dia = parseInt(match[1], 10);
          var mes = parseInt(match[2], 10) - 1;
          var ano = match[3] ? (match[3].length === 2 ? 2000 + parseInt(match[3], 10) : parseInt(match[3], 10)) : 2026;
          dt = new Date(ano, mes, dia);
        }
      }

      if (dt && !isNaN(dt.getTime())) {
        var anoStr = String(dt.getFullYear());
        var mesStr = String(dt.getMonth() + 1);
        if (mesStr.length === 1) mesStr = "0" + mesStr;
        var diaStr = String(dt.getDate());
        if (diaStr.length === 1) diaStr = "0" + diaStr;

        var chaveIso = anoStr + "-" + mesStr + "-" + diaStr;
        escala[chaveIso] = sup;
        lista.push({
          dataIso: chaveIso,
          dataFormatada: diaStr + "/" + mesStr + "/" + anoStr,
          supervisor: sup
        });
      }
    }

    var hoje = new Date();
    var yH = hoje.getFullYear();
    var mH = String(hoje.getMonth() + 1);
    if (mH.length === 1) mH = "0" + mH;
    var dH = String(hoje.getDate());
    if (dH.length === 1) dH = "0" + dH;
    var hojeIso = yH + "-" + mH + "-" + dH;

    var amarelinhoHoje = null;
    if (escala[hojeIso]) {
      amarelinhoHoje = {
        dataIso: hojeIso,
        dataFormatada: dH + "/" + mH + "/" + yH,
        supervisor: escala[hojeIso],
        isHoje: true
      };
    } else {
      for (var k = 0; k < lista.length; k++) {
        if (lista[k].dataIso >= hojeIso && lista[k].supervisor) {
          amarelinhoHoje = {
            dataIso: lista[k].dataIso,
            dataFormatada: lista[k].dataFormatada,
            supervisor: lista[k].supervisor,
            isHoje: (lista[k].dataIso === hojeIso)
          };
          break;
        }
      }
    }

    return {
      status: "ok",
      success: true,
      escala: escala,
      lista: lista,
      amarelinhoHoje: amarelinhoHoje
    };
  } catch (err) {
    return { status: "error", message: err.toString(), escala: {}, lista: [], amarelinhoHoje: null };
  }
}

/**
 * Extrai a lista de ausências (afastamentos curtos e longos) da aba 'Ausências'
 * da planilha de Organograma (1Rq_FIlUmpXyDK0alQ21_Z5b1bHMT_HlzU66QUGdr1LQ).
 */
function extrairAusenciasOrganograma(ssOrg) {
  try {
    var idPlanilhaOrg = "1Rq_FIlUmpXyDK0alQ21_Z5b1bHMT_HlzU66QUGdr1LQ";
    if (!ssOrg) {
      ssOrg = SpreadsheetApp.openById(idPlanilhaOrg);
    }
    var sheet = ssOrg.getSheetByName("Ausências") || ssOrg.getSheetByName("Ausencias");
    if (!sheet) {
      var allSheets = ssOrg.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        if (String(allSheets[s].getSheetId()) === "1834927774") {
          sheet = allSheets[s];
          break;
        }
      }
    }
    if (!sheet) {
      return { status: "error", message: "Aba 'Ausências' (gid 1834927774) não encontrada na planilha de Organograma.", dados: [] };
    }

    var valores = sheet.getDataRange().getDisplayValues();
    return {
      status: "ok",
      success: true,
      dados: valores,
      totalLinhas: valores.length
    };
  } catch (err) {
    return { status: "error", message: err.toString(), dados: [] };
  }
}

/**
 * Extrai timestamp numérico a partir do nome do arquivo SIMAP/X-Dias
 * Exemplos aceitos:
 *  - "xdias_13-11-2026_10h33_A" -> 13/11/2026 10:33
 *  - "SIMAP - ATUAL (NOV 02) 06-11-2026_21h30_A.xlsx" -> 06/11/2026 21:30
 */
function extrairTimestampNomeArquivo(nome) {
  if (!nome) return 0;
  var mData = nome.match(/(\d{2})[-/](\d{2})[-/](\d{4})/);
  var mHora = nome.match(/(\d{1,2})h(\d{2})/i) || nome.match(/(\d{1,2}):(\d{2})/);
  if (mData && mHora) {
    var dia = parseInt(mData[1], 10);
    var mes = parseInt(mData[2], 10) - 1;
    var ano = parseInt(mData[3], 10);
    var hora = parseInt(mHora[1], 10);
    var min = parseInt(mHora[2], 10);
    var dt = new Date(ano, mes, dia, hora, min);
    if (!isNaN(dt.getTime())) return dt.getTime();
  }
  return 0;
}
