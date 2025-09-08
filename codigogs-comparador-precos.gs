/**************  Comparador de Preços — multi-fornecedor (v2.1) ****************
 * - Sidebar em HTML (arquivo "Sidebar.html") para mapear N fornecedores.
 * - Base de Itens (Código, Marca, Quantidade) define o escopo da comparação.
 * - Match por Código + Marca, com normalização opcional.
 * - Relatório "Comparação": Preços, fornecedor mais barato, códigos, etc.
 * - Relatório "… - Itens Únicos": itens exclusivos de um fornecedor.
 * - Função para aplicar as quantidades sugeridas do relatório "Comparação".
 * - NOVO (v2.1): Menu `onOpen` reestruturado para separar a geração do
 *   relatório da aplicação de quantidades, com uma função para rodar a
 *   comparação diretamente com a última configuração salva.
 ************************************************************************/

/**
 * Cria o menu personalizado na planilha ao abri-la.
 * O menu agora segue um fluxo de trabalho lógico: Configurar, Gerar, Aplicar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Comparador de Preços')
    .addItem('1. Configurar Comparador (Abrir Sidebar)', 'showSidebar')
    .addSeparator()
    .addItem('2. Gerar Relatório (usando última config.)', 'runComparisonFromLastConfig')
    .addItem('3. Aplicar Quantidades Sugeridas', 'applySuggestedQuantities')
    .addSeparator()
    .addItem('Comparar (modo antigo – 2 fornecedores)', 'showComparisonDialog')
    .addToUi();
}


/**
 * NOVO: Executa a comparação usando a última configuração salva no PropertiesService.
 * Permite que o usuário gere o relatório diretamente pelo menu, sem abrir a sidebar.
 */
function runComparisonFromLastConfig() {
  const ui = SpreadsheetApp.getUi();
  try {
    const userProps = PropertiesService.getUserProperties();
    const lastConfigJSON = userProps.getProperty('cmp:lastConfigJSON');

    if (!lastConfigJSON) {
      ui.alert('Nenhuma configuração salva.', 'Por favor, use a opção "1. Configurar Comparador" e execute uma comparação pela primeira vez para salvar as configurações.', ui.ButtonSet.OK);
      return;
    }

    const config = JSON.parse(lastConfigJSON);

    SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando a comparação com a última configuração salva...', 'Status', 10);

    const result = compareFromConfig(config);

    const message = `Relatórios "${result.reportSheet}" e "${result.uniqueSheet}" foram gerados/atualizados.\n\n` +
                    `Itens Comparados: ${result.totalItens}\n` +
                    `Itens Únicos Encontrados: ${result.totalUniques}`;
    ui.alert('Comparação Concluída!', message, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ui.alert('Erro ao Gerar Relatório', e.message, ui.ButtonSet.OK);
  }
}


/* ======================== FLUXO (SIDEBAR) ======================== */

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Comparador de Preços')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// dados para preencher a UI
function getInitialData() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets().map(s => s.getName());
  const props = PropertiesService.getUserProperties();
  const lastConfig = props.getProperty('cmp:lastConfigJSON');
  return { sheets, lastConfig: lastConfig ? JSON.parse(lastConfig) : null };
}

// ponto de entrada chamado pelo HTML
function compareFromConfig(config) {
  if (!config || !Array.isArray(config.suppliers) || config.suppliers.length < 2) {
    throw new Error('Adicione pelo menos 2 fornecedores.');
  }
  if (!config.baseItems || !config.baseItems.sheetName || !config.baseItems.codeCol || !config.baseItems.brandCol || !config.baseItems.qtyCol) {
    throw new Error('Mapeie a Base de Itens (Aba / Código / Marca / Qtd).');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportName = (config.reportName && String(config.reportName).trim()) || 'Comparação';
  const uniqueName = reportName + ' - Itens Únicos';
  const normalizeCodeFlag  = !!config.normalizeCode;
  const normalizeBrandFlag = !!config.normalizeBrand;
  const suppliers = config.suppliers;

  // salva último config
  PropertiesService.getUserProperties().setProperty('cmp:lastConfigJSON', JSON.stringify(config));

  // ===== 1) Carrega a Base de Itens (apenas estes serão comparados) =====
  const base = config.baseItems;
  const baseSheet = ss.getSheetByName(base.sheetName);
  if (!baseSheet) throw new Error(`Aba da Base de Itens não encontrada: ${base.sheetName}`);

  const baseCodeIdx  = colToIndex(base.codeCol)  - 1;
  const baseBrandIdx = colToIndex(base.brandCol) - 1;
  const baseQtyIdx   = colToIndex(base.qtyCol)   - 1;
  if ([baseCodeIdx, baseBrandIdx, baseQtyIdx].some(i => isNaN(i) || i < 0)) {
    throw new Error('Mapeamento inválido na Base de Itens. Verifique as letras/números das colunas de Código/Marca/Qtd.');
  }

  const baseValues = baseSheet.getDataRange().getValues();
  const baseMap = {}; // key -> { codeDisp, brandDisp, qty }
  const baseKeys = new Set();

  for (let r = 1; r < baseValues.length; r++) {
    const codeRaw  = baseValues[r][baseCodeIdx];
    const brandRaw = baseValues[r][baseBrandIdx];
    const qtyRaw   = baseValues[r][baseQtyIdx];

    const codeNorm  = normalizeCodeFlag  ? normalizeCode(codeRaw)  : strSafe(codeRaw).toUpperCase();
    const brandNorm = normalizeBrandFlag ? normalizeBrand(brandRaw) : strSafe(brandRaw).toUpperCase();
    if (!codeNorm || !brandNorm) continue;

    const key = codeNorm + '||' + brandNorm;
    const qty = parseQty(qtyRaw);

    if (!baseMap[key]) {
      baseMap[key] = { codeDisp: strDisplay(codeRaw), brandDisp: strDisplay(brandRaw), qty: isNaN(qty) ? '' : qty };
    } else {
      // Se houver duplicados na base, somamos a quantidade
      const prev = parseQty(baseMap[key].qty);
      baseMap[key].qty = (isNaN(prev) ? 0 : prev) + (isNaN(qty) ? 0 : qty);
    }
    baseKeys.add(key);
  }

  if (!baseKeys.size) {
    throw new Error('A Base de Itens não possui linhas válidas (verifique Código/Marca/Quantidade).');
  }

  // ===== 2) Coleta preços por fornecedor, mas só para itens da Base =====
  const supplierDataByKey = {}; // { sheetName: { key: { price:number, id:string } } }

  suppliers.forEach(sup => {
    const sh = ss.getSheetByName(sup.sheetName);
    if (!sh) throw new Error(`Aba não encontrada: ${sup.sheetName}`);

    const codeIdx  = colToIndex(sup.codeCol)  - 1;
    const priceIdx = colToIndex(sup.priceCol) - 1;
    const brandIdx = colToIndex(sup.brandCol) - 1;
    const idIdx    = sup.idCol ? (colToIndex(sup.idCol) - 1) : -1;

    if ([codeIdx, priceIdx, brandIdx].some(i => isNaN(i) || i < 0)) {
      throw new Error(`Mapeamento inválido em "${sup.sheetName}". Verifique as letras/números das colunas.`);
    }

    const values = sh.getDataRange().getValues();
    const map = {};
    for (let r = 1; r < values.length; r++) {
      const codeRaw  = values[r][codeIdx];
      const priceRaw = values[r][priceIdx];
      const brandRaw = values[r][brandIdx];
      const idRaw    = idIdx >= 0 ? values[r][idIdx] : '';

      const price = parsePrice(priceRaw);
      if (isNaN(price)) continue;

      const codeNorm  = normalizeCodeFlag  ? normalizeCode(codeRaw)  : strSafe(codeRaw).toUpperCase();
      const brandNorm = normalizeBrandFlag ? normalizeBrand(brandRaw) : strSafe(brandRaw).toUpperCase();
      if (!codeNorm || !brandNorm) continue;

      const key = codeNorm + '||' + brandNorm;
      if (!baseKeys.has(key)) continue; // só itens da Base

      const current = map[key];
      const idDisp  = strDisplay(idRaw);
      if (!current || price < current.price) {
        map[key] = { price, id: idDisp };
      }
    }
    supplierDataByKey[sup.sheetName] = map;
  });

  // ===== 3) Monta os relatórios =====
  const baseColsCount = 3; // Código, Marca, Qtd Sugerida
  const header = ['Código', 'Marca', 'Qtd Sugerida'];
  suppliers.forEach(s => header.push(`Preço ${s.sheetName}`));
  header.push('Fornecedor Mais Barato', 'Código Fornecedor (Mais Barato)', 'Diferença (R$)', 'Diferença (%)');

  const rows = [];
  const uniquesBySupplier = Object.fromEntries(suppliers.map(s => [s.sheetName, []]));

  // === LOOP: garante que TODO item da Base entre no relatório ===
  Array.from(baseKeys).forEach(key => {
    const baseItem = baseMap[key]; // { codeDisp, brandDisp, qty }

    const perSupplier = suppliers.map(s => {
      const entry = supplierDataByKey[s.sheetName][key];
      return (entry && typeof entry.price === 'number') ? entry.price : '';
    });

    const pairs = suppliers
      .map(s => {
        const entry = supplierDataByKey[s.sheetName][key];
        return { supplier: s.sheetName, price: entry?.price, id: entry?.id };
      })
      .filter(p => typeof p.price === 'number')
      .sort((a, b) => a.price - b.price);

    let cheapestSupplier = '';
    let cheapestId = '';
    let diff  = '';
    let diffPct = '';

    if (pairs.length >= 2) {
      const cheapest = pairs[0];
      const second   = pairs[1];
      cheapestSupplier = cheapest.supplier;
      cheapestId       = cheapest.id || '';
      diff  = second.price - cheapest.price;
      diffPct = diff / second.price;
    } else if (pairs.length === 1) {
      const only = pairs[0];
      cheapestSupplier = only.supplier;
      cheapestId       = only.id || '';
      uniquesBySupplier[only.supplier].push([baseItem.codeDisp, baseItem.brandDisp, only.price]);
    }

    rows.push([
      baseItem.codeDisp,
      baseItem.brandDisp,
      baseItem.qty,
      ...perSupplier,
      cheapestSupplier,
      cheapestId,
      diff,
      diffPct
    ]);
  });

  writeComparisonSheet_(ss, reportName, header, rows, suppliers.length, baseColsCount);
  writeUniquesSheet_(ss, uniqueName, uniquesBySupplier);

  return {
    reportSheet: reportName,
    uniqueSheet: uniqueName,
    totalItens: rows.length,
    totalUniques: Object.values(uniquesBySupplier).reduce((a, arr) => a + arr.length, 0),
  };
}

/***** === APLICAR QUANTIDADES: COMPARAÇÃO → FORNECEDOR === *****/

/**
 * Lê a aba "Comparação" e escreve a Quantidade Sugerida no fornecedor vencedor,
 * fazendo match pelo "Código Fornecedor (Mais Barato)" na aba do fornecedor.
 */
function applySuggestedQuantities(options) {
  const userProps = PropertiesService.getUserProperties();
  const lastConfig = safeParseJSON_(userProps.getProperty('cmp:lastConfigJSON')) || {};
  const reportName = (options && options.reportName) || lastConfig.reportName || 'Comparação';

  const REPORT_HEADERS = {
    supplierSheet: 'Fornecedor Mais Barato',
    supplierCode:  'Código Fornecedor (Mais Barato)',
    qty:           'Qtd Sugerida'
  };

  const NORMALIZE_CODE = (typeof lastConfig.normalizeCode === 'boolean') ? lastConfig.normalizeCode : true;

  const ss = SpreadsheetApp.getActive();
  const report = ss.getSheetByName(reportName);
  if (!report) {
    SpreadsheetApp.getUi().alert(`Aba de relatório "${reportName}" não encontrada.`);
    return;
  }

  const lastRow = report.getLastRow();
  const lastCol = report.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert(`Aba "${reportName}" está vazia.`);
    return;
  }

  const header = report.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxSupplierSheet = findHeaderIndex_(header, REPORT_HEADERS.supplierSheet);
  const idxSupplierCode  = findHeaderIndex_(header, REPORT_HEADERS.supplierCode);
  const idxQty           = findHeaderIndex_(header, REPORT_HEADERS.qty);

  if (!idxSupplierSheet || !idxSupplierCode || !idxQty) {
    throw new Error(`Aba "${reportName}" precisa ter as colunas: "${REPORT_HEADERS.supplierSheet}", "${REPORT_HEADERS.supplierCode}" e "${REPORT_HEADERS.qty}".`);
  }

  const rows = report.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // --- Agrupar por fornecedor vencedor
  const bySupplier = {};
  rows.forEach(r => {
    const sheetName = (r[idxSupplierSheet - 1] || '').toString().trim();
    const code      = (r[idxSupplierCode  - 1] || '').toString().trim();
    const qty       = toNumberOrNull_(r[idxQty - 1]);

    if (!sheetName || !code || qty == null) return;
    if (!bySupplier[sheetName]) bySupplier[sheetName] = [];
    bySupplier[sheetName].push({ code, qty });
  });

  const supplierConfigs = Array.isArray(lastConfig.suppliers) ? lastConfig.suppliers : [];
  const summary = [];

  // --- Processar cada fornecedor
  Object.keys(bySupplier).forEach(supplierSheetName => {
    const sh = ss.getSheetByName(supplierSheetName);
    if (!sh) {
      summary.push({ fornecedor: supplierSheetName, atualizados: 0, ignorados: bySupplier[supplierSheetName].length, motivo: 'Aba do fornecedor não encontrada' });
      return;
    }

    const conf = supplierConfigs.find(s => (s.sheetName || '') === supplierSheetName) || {};
    
    // (1) Descobrir a coluna de CÓDIGO (usa config com fallback)
    let codeColIndex = null;
    if (conf.codeCol) codeColIndex = colToIndex(conf.codeCol);
    if (!codeColIndex) {
      const shLastCol = sh.getLastColumn();
      const shHeader = sh.getRange(1, 1, 1, shLastCol).getValues()[0];
      const candidates = ['Código', 'Codigo', 'CÓDIGO', 'SKU', 'ID', 'Cód. Fornecedor', 'Código Fornecedor'];
      for (const c of candidates) {
        const idx = findHeaderIndex_(shHeader, c);
        if (idx) { codeColIndex = idx; break; }
      }
      if (!codeColIndex) throw new Error(`Não encontrei a coluna de código na aba "${supplierSheetName}". Mapeie-a na Sidebar.`);
    }

    // (2) Descobrir a coluna de QUANTIDADE a partir da configuração da Sidebar
    if (!conf.qtyCol) {
      summary.push({ fornecedor: supplierSheetName, atualizados: 0, ignorados: bySupplier[supplierSheetName].length, motivo: 'Coluna de Quantidade não foi mapeada na Sidebar.' });
      return;
    }
    const qtyColIndex = colToIndex(conf.qtyCol);
    if (!qtyColIndex || qtyColIndex < 1) {
      summary.push({ fornecedor: supplierSheetName, atualizados: 0, ignorados: bySupplier[supplierSheetName].length, motivo: `Coluna de Quantidade "${conf.qtyCol}" é inválida.` });
      return;
    }

    // (3) Indexar os códigos do fornecedor
    const lastRowSupplier = sh.getLastRow();
    if (lastRowSupplier < 2) {
      summary.push({ fornecedor: supplierSheetName, atualizados: 0, ignorados: bySupplier[supplierSheetName].length, motivo: 'Fornecedor sem linhas de dados' });
      return;
    }

    const supplierCodes = sh.getRange(2, codeColIndex, lastRowSupplier - 1, 1).getValues().map(v => (v[0] || '').toString());
    const codeIndexMap = new Map();
    supplierCodes.forEach((c, i) => {
      const key = NORMALIZE_CODE ? normalizeCode(c) : strSafe(c).toUpperCase();
      if (!key) return;
      if (!codeIndexMap.has(key)) codeIndexMap.set(key, i + 2);
    });

    // (4) Preparar vetor com as quantidades atuais
    let qtyColumnValues;
    try {
      qtyColumnValues = sh.getRange(2, qtyColIndex, lastRowSupplier - 1, 1).getValues();
    } catch (e) {
      // Se a coluna não existir, cria um array vazio para preencher
      qtyColumnValues = Array(lastRowSupplier - 1).fill(null).map(() => ['']);
    }
    
    // (5) Aplicar atualizações
    let updated = 0, notFound = 0;
    bySupplier[supplierSheetName].forEach(item => {
      const sought = NORMALIZE_CODE ? normalizeCode(item.code) : strSafe(item.code).toUpperCase();
      const rowNumber = codeIndexMap.get(sought);
      if (!rowNumber) {
        notFound++;
        return;
      }
      const arrIndex = rowNumber - 2;
      if (arrIndex >= 0 && arrIndex < qtyColumnValues.length) {
        qtyColumnValues[arrIndex][0] = item.qty;
        updated++;
      }
    });

    // (6) Gravar de volta
    if (updated > 0) {
      sh.getRange(2, qtyColIndex, qtyColumnValues.length, 1).setValues(qtyColumnValues);
    }
    summary.push({ fornecedor: supplierSheetName, atualizados: updated, ignorados: notFound, motivo: notFound ? 'Códigos não encontrados' : '' });
  });

  const totalAtualizados = summary.reduce((a, s) => a + s.atualizados, 0);
  const totalIgnorados   = summary.reduce((a, s) => a + s.ignorados, 0);
  const msg =
    `Aplicação concluída a partir de "${reportName}".\n` +
    `Linhas atualizadas: ${totalAtualizados}.\n` +
    `Não encontradas/ignoradas: ${totalIgnorados}.`;

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
  return { ok: true, resumo: summary, mensagem: msg };
}


/* ============================= HELPERS ============================= */

function findHeaderIndex_(headerArr, targetName) {
  const want = norm_(targetName);
  for (let i = 0; i < headerArr.length; i++) {
    if (norm_(headerArr[i]) === want) return i + 1;
  }
  return 0;
}

function norm_(s) {
  return (s || '').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ').trim().toUpperCase();
}

function toNumberOrNull_(v) {
  if (v === null || v === '' || typeof v === 'undefined') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function safeParseJSON_(txt) {
  try { return JSON.parse(txt); } catch(e) { return null; }
}

function colToIndex(input) {
  if (typeof input === 'number') return input;
  let s = String(input || '').trim().toUpperCase();
  if (!s) return NaN;
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function parsePrice(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return NaN;
  let s = String(v).trim();
  if (!s) return NaN;
  s = s.replace(/[^\d,.\-]+/g, '');
  s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

function parseQty(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return NaN;
  let s = String(v).trim();
  if (!s) return NaN;
  s = s.replace(/[^\d,.\-]+/g, '');
  s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

function strSafe(v)     { return (v === null || v === undefined) ? '' : String(v).trim(); }
function strDisplay(v)  { const s = strSafe(v); return s.length ? s : ''; }

function normalizeCode(v) {
  return strSafe(v).toUpperCase().replace(/[^\p{L}\p{N}]/gu, '');
}
function normalizeBrand(v) {
  let s = strSafe(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return s.toUpperCase().replace(/[^\p{L}\p{N}]/gu, '');
}

function writeComparisonSheet_(ss, name, header, rows, supplierCount, baseColsCount) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const existingFilter = sh.getFilter();
  if (existingFilter) existingFilter.remove();
  sh.clear();

  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);

  const numRows = rows.length;
  if (numRows > 0) {
    const priceStartCol = baseColsCount + 1;
    sh.getRange(2, priceStartCol, numRows, supplierCount).setNumberFormat('R$ #,##0.00');
    const diffCol = header.length - 2;
    sh.getRange(2, diffCol, numRows, 1).setNumberFormat('R$ #,##0.00');
    sh.getRange(2, diffCol + 1, numRows, 1).setNumberFormat('0.00%');
  }
  sh.getRange(1, 1, Math.max(1, rows.length + 1), header.length).createFilter();
}

function writeUniquesSheet_(ss, name, uniquesBySupplier) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();

  let row = 1;
  const suppliers = Object.keys(uniquesBySupplier);
  suppliers.forEach(sup => {
    const data = uniquesBySupplier[sup];
    if (!data || !data.length) return;
    sh.getRange(row, 1).setValue(`Itens Exclusivos de ${sup}`).setFontWeight('bold');
    row++;
    sh.getRange(row, 1, 1, 3).setValues([['Código', 'Marca', `Preço ${sup}`]]).setFontWeight('bold');
    row++;
    sh.getRange(row, 1, data.length, 3).setValues(data);
    sh.getRange(row, 3, data.length, 1).setNumberFormat('R$ #,##0.00');
    row += data.length + 1;
  });
  if (row === 1) sh.getRange(1, 1).setValue('Não há itens exclusivos em nenhum fornecedor.');
  sh.autoResizeColumns(1, 3);
}

/* ========================= COMPATIBILIDADE ======================== */

function showComparisonDialog() {
  const ui = SpreadsheetApp.getUi();

  const r1 = ui.prompt('Configuração da Comparação', 'Fornecedor 1 - Aba, Coluna Código, Coluna Preço, Coluna Marca (ex: Tabela Montanna,A,B,C):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return ui.alert('Ação cancelada.');
  const [sheetName1, codeColumn1, priceColumn1, brandColumn1] = r1.getResponseText().split(',').map(s => s.trim());

  const r2 = ui.prompt('Configuração da Comparação', 'Fornecedor 2 - Aba, Coluna Código, Coluna Preço, Coluna Marca (ex: Tabela LM,A,B,C):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return ui.alert('Ação cancelada.');
  const [sheetName2, codeColumn2, priceColumn2, brandColumn2] = r2.getResponseText().split(',').map(s => s.trim());

  const r3 = ui.prompt('Nome da Aba de Relatório', 'Digite o nome para a aba de relatório principal (ex: Comparação):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK || !r3.getResponseText()) return ui.alert('Ação cancelada.');
  const reportSheetName = r3.getResponseText().trim();

  const config = {
    reportName: reportSheetName || 'Comparação',
    normalizeCode: true,
    normalizeBrand: true,
    baseItems: { sheetName: sheetName1, codeCol: codeColumn1, brandCol: brandColumn1, qtyCol: priceColumn1 },
    suppliers: [
      { sheetName: sheetName1, codeCol: codeColumn1, priceCol: priceColumn1, brandCol: brandColumn1 },
      { sheetName: sheetName2, codeCol: codeColumn2, priceCol: priceColumn2, brandCol: brandColumn2 },
    ],
  };
  const res = compareFromConfig(config);
  SpreadsheetApp.getUi().alert(`Relatórios gerados: "${res.reportSheet}" e "${res.uniqueSheet}".`);
}
