function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function obterListasFiscais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Lendo os Municípios da aba 'lista'
  var abaLista = ss.getSheetByName('lista');
  var dadosMunicipios = abaLista.getRange('A2:A167').getValues();
  var listaCidades = dadosMunicipios.flat().filter(String);
  
  // 2. Lendo os Códigos de Serviço da aba 'alíquota'
  var abaAliq = ss.getSheetByName('alíquota');
  var dadosCodigos = abaAliq.getRange('A2:A216').getValues();
  var listaCodigos = dadosCodigos
    .flat()
    .filter(String)
    .map(item => String(item).replace('.', ','));
  
  return {
    municipios: listaCidades,
    codigos: listaCodigos
  };
}


// Esta função joga as escolhas do site nas células C2 e C3 e lê as respostas solicitadas
// Esta função joga as escolhas do site nas células C2 e C3 e lê as respostas solicitadas
function processarSimulacao(municipio, codigo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('painel_Consulta');
  
  // Insere as escolhas do usuário
  aba.getRange('C2').setValue(municipio);
  aba.getRange('C3').setValue(codigo);
    
  SpreadsheetApp.flush(); // Força as fórmulas da planilha a calcularem
  
  // Capturando a célula D3 de forma isolada e limpa
  var valorD3 = aba.getRange('D3').getValue();
  
  // Lendo as colunas completas da linha 4 até a 21 (TUDO ALINHADO NA LINHA 4)
  var valoresB = aba.getRange('B4:B21').getValues();
  var valoresC = aba.getRange('C4:C21').getValues();
  var valoresD = aba.getRange('D4:D21').getValues();
  
  // Lendo os hiperlinks especificamente para as linhas 16 a 21 na coluna D
  var richTextD = aba.getRange('D16:D21').getRichTextValues();
  
  // FUNÇÃO AUXILIAR CORRIGIDA: Retorna o valor exato [0] da célula para não dar undefined
  function pegarLinha(matriz, numeroLinha) {
    return matriz[numeroLinha - 4][0];
  }
  
  // Função auxiliar para extrair a URL do link
  function pegarLink(matrizRich, numeroLinha) {
    var celula = matrizRich[numeroLinha - 16][0];
    return celula ? celula.getLinkUrl() : null;
  }
  
  // Montando o pacote de dados limpo e organizado linha por linha
  var resultado = {
    d3: valorD3,
    
    c4: pegarLinha(valoresC, 4), d4: pegarLinha(valoresD, 4),
    b5: pegarLinha(valoresB, 5),
    b7: pegarLinha(valoresB, 7), d7: pegarLinha(valoresD, 7),
    b8: pegarLinha(valoresB, 8), d8: pegarLinha(valoresD, 8),
    b9: pegarLinha(valoresB, 9), d9: pegarLinha(valoresD, 9),
    b10: pegarLinha(valoresB, 10), d10: pegarLinha(valoresD, 10),
    b11: pegarLinha(valoresB, 11), d11: pegarLinha(valoresD, 11),
    b12: pegarLinha(valoresB, 12), d12: pegarLinha(valoresD, 12),
    b13: pegarLinha(valoresB, 13), d13: pegarLinha(valoresD, 13),
    b14: pegarLinha(valoresB, 14), d14: pegarLinha(valoresD, 14),
    
    // Linhas com Texto + Link
    b16: pegarLinha(valoresB, 16), d16: pegarLinha(valoresD, 16), linkD16: pegarLink(richTextD, 16),
    b17: pegarLinha(valoresB, 17), d17: pegarLinha(valoresD, 17), linkD17: pegarLink(richTextD, 17),
    b18: pegarLinha(valoresB, 18), d18: pegarLinha(valoresD, 18), linkD18: pegarLink(richTextD, 18),
    b19: pegarLinha(valoresB, 19), d19: pegarLinha(valoresD, 19), linkD19: pegarLink(richTextD, 19),
    b20: pegarLinha(valoresB, 20), d20: pegarLinha(valoresD, 20), linkD20: pegarLink(richTextD, 20),
    b21: pegarLinha(valoresB, 21), d21: pegarLinha(valoresD, 21), linkD21: pegarLink(richTextD, 21)
  };
  
  return resultado;
}


