// Função para transferir dados da página "PERÍODO x VENDAS x ESTOQUE" para "SUGESTÃO"
function atualizarSugestao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName("PERIODO x VENDAS x ESTOQUE");
  var destino = ss.getSheetByName("SUGESTÃO");

  // Obtém os dados desde a segunda linha até a última linha com dados na planilha origem
  var dadosOrigem = origem.getRange("A2:H" + origem.getLastRow()).getValues();
  var dadosDestino = [];

  dadosOrigem.forEach(function(row) {
    var codigo = row[0];          // Código do produto
    var produto = row[1];         // Descrição do produto
    var pecasVendidas = row[5];   // Peças vendidas no período, está na coluna F da origem
    var estoqueAtual = row[7];    // Estoque atual na data, está na coluna H da origem

    // Preparar cada linha de dados ajustada para o destino
    dadosDestino.push([codigo, produto, '', pecasVendidas, estoqueAtual]); // Insere espaço vazio em 'Cobrir por quanto tempo (C)'
  });

  // Assegura que não está tentando escrever uma matriz vazia
  if (dadosDestino.length > 0) {
    destino.getRange(2, 1, dadosDestino.length, dadosDestino[0].length).setValues(dadosDestino);
  }
}

// ===================================================================================
// FUNÇÃO CORRIGIDA
// ===================================================================================
// Função para transferir dados da página "SUGESTÃO" para "FORNECEDOR x PEDIDO FINAL"
function transferirDadosParaPedidoFinal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSugestao = ss.getSheetByName('SUGESTÃO');
  var sheetPedidoFinal = ss.getSheetByName('FORNECEDOR x PEDIDO FINAL');

  var rangeSugestao = sheetSugestao.getRange('A2:I' + sheetSugestao.getLastRow());
  var dadosSugestao = rangeSugestao.getValues();
  
  var rangePedidoFinal = sheetPedidoFinal.getRange('A2:I' + sheetPedidoFinal.getLastRow());
  var dadosPedidoFinal = rangePedidoFinal.getValues();

  var mapaPedidosSKU = new Map();
  var mapaPedidosTitulo = new Map();

  dadosPedidoFinal.forEach(function(row, index) {
    // CORREÇÃO: Verifica se a célula não está vazia, converte para String e depois usa o .trim()
    var sku = row[0] ? String(row[0]).trim() : "";
    var titulo = row[1] ? String(row[1]).trim() : "";
    
    if (sku) {
      mapaPedidosSKU.set(sku, index + 2); // SKU na coluna A
    }
    if (titulo) {
      mapaPedidosTitulo.set(titulo, index + 2); // Título na coluna B
    }
  });

  dadosSugestao.forEach(function(rowSugestao) {
    // CORREÇÃO: Verifica se a célula não está vazia, converte para String e depois usa o .trim()
    var sku = rowSugestao[0] ? String(rowSugestao[0]).trim() : "";
    var titulo = rowSugestao[1] ? String(rowSugestao[1]).trim() : "";
    var quantidadeSugestao = rowSugestao[8];

    // Só processa se tiver um SKU ou Título e uma quantidade
    if ((sku || titulo) && (quantidadeSugestao !== null && quantidadeSugestao !== '')) {
      var linhaPedido = mapaPedidosSKU.get(sku) || mapaPedidosTitulo.get(titulo);

      if (linhaPedido) {
        sheetPedidoFinal.getRange('I' + linhaPedido).setValue(quantidadeSugestao);
      } else {
        console.log('Não encontrado por SKU ou Título:', sku, titulo);
      }
    }
  });
}

function limparCelulasEspecificasPagina2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SUGESTÃO");  // Substitua "SUGESTÃO" pelo nome exato da sua página 2, se for diferente.

  // Limpa todas as células exceto as linhas de cabeçalho A1:I1 e as células G2:I2
  var rangeParaLimpar = [
    'A2:F',  // Limpa da coluna A até a coluna F da linha 2 até o fim
    'J2:J'   // Limpa a coluna J da linha 2 até o fim (se necessário)
  ];

  rangeParaLimpar.forEach(function(range) {
    var fullRange = range + sheet.getLastRow();  // Define o limite da limpeza até a última linha com dados
    sheet.getRange(fullRange).clearContent();  // Limpa o conteúdo das células
  });

  // Adicionalmente, limpa todas as linhas abaixo da 3 exceto as colunas G, H, I
  var lastRow = sheet.getLastRow();
  if (lastRow > 3) {
    sheet.getRange('A3:F' + lastRow).clearContent();  // Limpa da linha 3 até a última linha nas colunas de A até F
    sheet.getRange('J3:J' + lastRow).clearContent();  // Limpa da linha 3 até a última linha na coluna J
  }
}

/**
 * Função para adicionar um menu personalizado na planilha ao abrir.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Automações de Pedidos')
      .addItem('1. Atualizar Sugestão (PERÍODO -> SUGESTÃO)', 'atualizarSugestao')
      .addItem('2. Limpar Células da Sugestão', 'limparCelulasEspecificasPagina2')
      .addItem('3. Transferir Sugestão para Pedido Final (SUGESTÃO -> PEDIDO FINAL)', 'transferirDadosParaPedidoFinal')
      .addSeparator()
      .addItem('4. Projetar Quantidades em Tabela do Fornecedor (PEDIDO FINAL -> TABELA DINÂMICA)', 'projetarQuantidadesNaTabelaFornecedor') // Nome do item de menu atualizado
      .addToUi();
}


// ===================================================================================
// NOVA FUNÇÃO (MODIFICADA) PARA PROJETAR QUANTIDADES NA TABELA DO FORNECEDOR (DINÂMICA)
// ===================================================================================
function projetarQuantidadesNaTabelaFornecedor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // --- 1. Obter dados da origem ("FORNECEDOR x PEDIDO FINAL") ---
  var sheetPedidoFinal = ss.getSheetByName('FORNECEDOR x PEDIDO FINAL');
  if (!sheetPedidoFinal) {
    ui.alert('Erro', 'A aba "FORNECEDOR x PEDIDO FINAL" não foi encontrada.');
    return;
  }

  var ultimaLinhaPedidoFinal = sheetPedidoFinal.getLastRow();
  if (ultimaLinhaPedidoFinal < 2) {
    ui.alert('Aviso', 'Não há dados na aba "FORNECEDOR x PEDIDO FINAL" para projetar.');
    return;
  }

  // Lemos as colunas G (Código), H (Descrição) e I (Quantidade)
  var rangeDadosPedido = sheetPedidoFinal.getRange('G2:I' + ultimaLinhaPedidoFinal);
  var dadosPedidoFinal = rangeDadosPedido.getValues();

  // Mapa para busca rápida de quantidades e para rastrear o que foi encontrado
  var mapaProdutosDoPedido = new Map();
  dadosPedidoFinal.forEach(function(row) {
    var codigo = row[0] ? String(row[0]).trim() : "";
    var descricao = row[1];
    var quantidade = row[2];

    if (codigo && (quantidade !== null && quantidade !== "")) {
      mapaProdutosDoPedido.set(codigo, {
        descricao: descricao,
        quantidade: quantidade,
        encontrado: false // Começa como 'não encontrado'
      });
    }
  });

  if (mapaProdutosDoPedido.size === 0) {
    ui.alert('Aviso', 'Nenhum código/quantidade válido encontrado na aba "FORNECEDOR x PEDIDO FINAL".');
    return;
  }

  // --- 2. Obter informações do usuário sobre a planilha de destino ---
  var promptNomeAba = ui.prompt('Nome da Aba do Fornecedor', 'Digite o nome EXATO da aba do fornecedor:', ui.ButtonSet.OK_CANCEL);
  if (promptNomeAba.getSelectedButton() != ui.Button.OK || !promptNomeAba.getResponseText()) return;
  var nomeAbaFornecedor = promptNomeAba.getResponseText().trim();

  var promptColunaCodigo = ui.prompt('Coluna dos Códigos', 'Digite a LETRA da coluna dos códigos na aba "' + nomeAbaFornecedor + '":', ui.ButtonSet.OK_CANCEL);
  if (promptColunaCodigo.getSelectedButton() != ui.Button.OK || !promptColunaCodigo.getResponseText()) return;
  var letraColunaCodigo = promptColunaCodigo.getResponseText().trim().toUpperCase();
  var numColunaCodigo = colLetraParaNum(letraColunaCodigo);
  if (numColunaCodigo === 0) {
    ui.alert('Erro', 'Letra da coluna de códigos inválida.');
    return;
  }

  var promptColunaQuantidade = ui.prompt('Coluna das Quantidades', 'Digite a LETRA da coluna para inserir as quantidades na aba "' + nomeAbaFornecedor + '":', ui.ButtonSet.OK_CANCEL);
  if (promptColunaQuantidade.getSelectedButton() != ui.Button.OK || !promptColunaQuantidade.getResponseText()) return;
  var letraColunaQuantidade = promptColunaQuantidade.getResponseText().trim().toUpperCase();
  var numColunaQuantidade = colLetraParaNum(letraColunaQuantidade);
  if (numColunaQuantidade === 0) {
    ui.alert('Erro', 'Letra da coluna de quantidade inválida.');
    return;
  }

  // --- 3. Processar a aba do fornecedor e marcar itens encontrados ---
  var sheetFornecedor = ss.getSheetByName(nomeAbaFornecedor);
  if (!sheetFornecedor) {
    ui.alert('Erro', 'A aba "' + nomeAbaFornecedor + '" não foi encontrada.');
    return;
  }

  var ultimaLinhaFornecedor = sheetFornecedor.getLastRow();
  if (ultimaLinhaFornecedor < 2) {
    ui.alert('Aviso', 'A aba do fornecedor está vazia.');
  }

  // Limpa a coluna de quantidade antes de preencher
  sheetFornecedor.getRange(2, numColunaQuantidade, ultimaLinhaFornecedor, 1).clearContent();

  var rangeCodigosFornecedor = sheetFornecedor.getRange(2, numColunaCodigo, ultimaLinhaFornecedor - 1, 1);
  var codigosFornecedor = rangeCodigosFornecedor.getValues();
  var quantidadesParaProjetar = [];

  codigosFornecedor.forEach(function(row) {
    var codigoFornecedor = row[0] ? String(row[0]).trim() : "";
    var quantidadeEncontrada = ""; // Padrão vazio

    if (codigoFornecedor && mapaProdutosDoPedido.has(codigoFornecedor)) {
      var produto = mapaProdutosDoPedido.get(codigoFornecedor);
      quantidadeEncontrada = produto.quantidade;
      produto.encontrado = true; // MARCA O PRODUTO COMO ENCONTRADO!
      mapaProdutosDoPedido.set(codigoFornecedor, produto); // Atualiza o mapa
    }
    quantidadesParaProjetar.push([quantidadeEncontrada]);
  });

  // --- 4. Escrever as quantidades na aba do fornecedor ---
  if (quantidadesParaProjetar.length > 0) {
    sheetFornecedor.getRange(2, numColunaQuantidade, quantidadesParaProjetar.length, 1).setValues(quantidadesParaProjetar);
  }

  // --- 5. Gerar o relatório de itens não encontrados ---
  var itensNaoEncontrados = [];
  mapaProdutosDoPedido.forEach(function(produto, codigo) {
    if (!produto.encontrado) {
      itensNaoEncontrados.push([codigo, produto.descricao, produto.quantidade]);
    }
  });

  var nomeAbaRelatorio = "Relatório de Itens Faltantes";
  var sheetRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (sheetRelatorio) {
    sheetRelatorio.clear(); // Limpa a aba se já existir
  } else {
    sheetRelatorio = ss.insertSheet(nomeAbaRelatorio); // Cria a aba se não existir
  }

  // Formata o cabeçalho do relatório
  var cabecalho = [
    ["Código não Encontrado", "Descrição do Produto", "Quantidade do Pedido"]
  ];
  sheetRelatorio.getRange("A1:C1").setValues(cabecalho).setFontWeight("bold").setBackground("#d9ead3");
  sheetRelatorio.setColumnWidths(1, 3, 200); // Ajusta a largura das colunas

  if (itensNaoEncontrados.length > 0) {
    // Escreve os dados dos itens faltantes no relatório
    sheetRelatorio.getRange(2, 1, itensNaoEncontrados.length, 3).setValues(itensNaoEncontrados);
    ss.setActiveSheet(sheetRelatorio); // Ativa a aba do relatório para o usuário ver
    ui.alert('Processo Concluído', 'As quantidades foram projetadas. ATENÇÃO: ' + itensNaoEncontrados.length + ' item(ns) do seu pedido não foram encontrados na tabela do fornecedor. Verifique a aba "' + nomeAbaRelatorio + '" para detalhes.', ui.ButtonSet.OK);
  } else {
    // Se não houver itens faltantes, pode até apagar a aba de relatório para não poluir
    ss.deleteSheet(sheetRelatorio);
    ui.alert('Sucesso!', 'Todas as quantidades foram projetadas e todos os itens do pedido foram encontrados na tabela do fornecedor.', ui.ButtonSet.OK);
  }
}

/**
 * Função auxiliar para converter letra da coluna (e.g., 'A', 'B', 'AA') para seu número (1-based index).
 * @param {string} letra A letra da coluna.
 * @return {number} O número da coluna (1 para A, 2 para B, etc.) ou 0 se inválido.
 */
function colLetraParaNum(letra) {
  if (!letra || typeof letra !== 'string') return 0;
  letra = letra.toUpperCase();
  let coluna = 0, length = letra.length;
  for (let i = 0; i < length; i++) {
    let charCode = letra.charCodeAt(i);
    if (charCode < 65 || charCode > 90) { // Verifica se é uma letra de A-Z
        return 0; // Letra inválida
    }
    coluna += (charCode - 64) * Math.pow(26, length - i - 1);
  }
  return coluna;
}
