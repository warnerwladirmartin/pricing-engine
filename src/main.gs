// ============================================================
// Pricing & Margin Engine v3.0 — "Motor de Margem"
// Google Apps Script — Lubricant manufacturer price tables
// ============================================================
// SANITIZED: no real sheet IDs, no real customer data.
// Replace all <PLACEHOLDER> values with your actual IDs.
// ============================================================
//
// ⚠️  RUNTIME WARNING: This script runs on the legacy Rhino
//     engine. Google has deprecated Rhino; migration to V8 is
//     required before the deprecation deadline. No syntax
//     changes have been made to facilitate the migration yet.
//
// ARCHITECTURE OVERVIEW
// ─────────────────────
// Six external spreadsheets feed this motor:
//   1. EP_CLIENTES          — customer master (CNPJ, rep, state, billing %)
//   2. EP_BASE_VENDAS       — sales history (last price, date, volume)
//   3. EP_TABELAS_REF       — benchmark prices (PADRAO, RJ, LEME, CONSUMO, VAREJO)
//   4. EP_BASE_CUSTOS /     — production costs (BOM cascade: MP→PE→PA + OLUC 12%)
//      CostStructure_v2
//   5. EP_PARAMETROS_MARGEM — freight, commission, taxes, margin policy, strategic
//   6. EP_MOTOR itself      — CONFIG, REAJUSTES, CLASSIFICACAO, TABELA_NOVA
//
// SECTION MAP
// ───────────
// [A] Constants & COL mapping
// [B] Lifecycle hooks: onOpen, onEdit, atualizarDropdownB2_
// [C] Utility: parsePreco_, removerAcentos_, lerClassificacoes_, lerUnidades_,
//              calcPrazoMedio_
// [D] Data loaders: lerCustos_, lerParametros_, buscarMargem_,
//                   calcComissao_, lerReajustes_
// [E] Main workflows: carregarCliente, carregarTabelaPadrao_,
//                     corrigirFormulas (Recalcular Margens)
// [F] Export: exportarTabelaAtual, exportarTodasTabelas,
//             montarItens_, montarItensVarejo_, criarArquivoTabela_
// ============================================================

// ── [A] CONSTANTS & COLUMN MAP ────────────────────────────────

// External spreadsheet IDs — replace with your actual IDs.
var ID_MOTOR        = '<ID_MOTOR_PRECIFICACAO>';
var ID_CLIENTES     = '<ID_CLIENTES>';
var ID_BASE_VENDAS  = '<ID_BASE_VENDAS>';
var ID_BASE_CUSTOS  = '<ID_BASE_CUSTOS>';         // CostStructure_v2
var ID_TABELAS_REF  = '<ID_TABELAS_REF>';
var ID_PARAMETROS   = '<ID_PARAMETROS_MARGEM>';

// Sheet/tab names inside EP_MOTOR
var ABA_CONFIG      = 'CONFIG';
var ABA_REAJUSTES   = 'REAJUSTES';
var ABA_NOVA        = 'TABELA_NOVA';
var ABA_CLASSIF     = 'CLASSIFICACAO';

// Standard tables available in EP_TABELAS_REF
var STD_TABS = {
  'PADRAO GERAL':    'PADRAO',
  'PADRAO RJ':       'RJ',
  'PADRAO CONSUMO':  'CONSUMO',
  'PADRAO VAREJO':   'VAREJO'
};
// NOTE: customers with TABELA_REF = PADRAO use the LEME tab as source,
// not the PADRAO tab. This is intentional (LEME = commercial reference).

// Brand palette (not sensitive)
var COR_LARANJA       = '#CA4F24';
var COR_LARANJA_ESC   = '#222221';
var COR_LARANJA_CLARO = '#FFF3EB';
var COR_BRANCO        = '#FFFFFF';

// TABELA_NOVA column map — 1-indexed, 28 columns
var COL = {
  SKU:            1,   // A — internal ERP product code
  PRODUTO:        2,   // B — product description
  FAMILIA:        3,   // C — product family grouping
  CLASSIF:        4,   // D — commercial classification (PREMIUM, STANDARD, etc.)
  UNID:           5,   // E — unit of measure (BD, TB, CX, ...)

  CUSTO_PROD:     6,   // F — production cost from BOM cascade (incl. OLUC)
  FRETE:          7,   // G — freight cost as % of price (by state)
  COMISSAO:       8,   // H — sales commission % (by rep / CNPJ exception)
  IMPOSTOS:       9,   // I — tax burden % (by state)
  CUSTO_FIN:     10,   // J — financial carrying cost % (monthly rate)
  CUSTO_TOTAL:   11,   // K — sum of all variable cost %s (G+H+I+J)

  MARGEM_MIN:    12,   // L — minimum acceptable margin %
  MARGEM_ALVO:   13,   // M — target margin % (from margin policy)

  PRECO_MIN:     14,   // N — floor price: custo / (1 - custos% - margem_min%)
  PRECO_SUGERIDO: 15,  // O — suggested price: custo / (1 - custos% - margem_alvo%)

  PRECO_REF:     16,   // P — benchmark price from EP_TABELAS_REF
  ULT_PRECO:     17,   // Q — last invoiced price from sales history
  TEM_HIST:      18,   // R — flag: TRUE if customer has sales history for this SKU

  PRECO_BASE:    19,   // S — base price (ULT_PRECO if TEM_HIST, else PRECO_REF)
  REAJUSTE:      20,   // T — adjustment % from REAJUSTES tab (global/classif/family)
  NOVO_PRECO:    21,   // U — final proposed price after adjustment

  MARGEM_REAL_PCT: 22, // V — actual margin % at NOVO_PRECO
  MARGEM_REAL_RS:  23, // W — actual margin R$ at NOVO_PRECO
  ALERTA:          24, // X — alert: ABAIXO MINIMO / ABAIXO ALVO / OK / PREMIUM

  DT_ULT_VENDA:  25,  // Y — date of last sale (from EP_BASE_VENDAS)
  QTD_HIST:      26,  // Z — total volume in history period
  N_VENDAS:      27,  // AA — number of invoices in history period
  FONTE:         28   // AB — data source flag (DAGDA / HISTORICO / SEM HIST)
};


// ── [B] LIFECYCLE HOOKS ───────────────────────────────────────

/**
 * onOpen — creates the custom menu "EP Precificacao" and refreshes the
 * customer dropdown in B2.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('EP Precificacao')
    .addItem('Carregar Cliente / Tabela', 'carregarCliente')
    .addItem('Recalcular Margens',        'corrigirFormulas')
    .addSeparator()
    .addItem('Exportar Tabela Atual',     'exportarTabelaAtual')
    .addItem('Exportar Todas as Tabelas', 'exportarTodasTabelas')
    .addSeparator()
    .addItem('Atualizar Dropdown B2',     'atualizarDropdownB2_')
    .addToUi();

  atualizarDropdownB2_();
}

/**
 * onEdit — auto-loads the selected customer whenever cell B2 changes.
 */
function onEdit(e) {
  var range = e.range;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsCfg = ss.getSheetByName(ABA_CONFIG);
  if (!wsCfg) return;
  if (range.getSheet().getName() !== ABA_CONFIG) return;
  if (range.getA1Notation() !== 'B2') return;
  carregarCliente();
}

/**
 * atualizarDropdownB2_ — populates cell B2 with a data-validation list of
 * customers + standard tables.
 * Primary source: EP_CLIENTES sheet (column with customer names).
 * Fallback: column S of CONFIG tab in the motor itself.
 */
function atualizarDropdownB2_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsCfg = ss.getSheetByName(ABA_CONFIG);
  if (!wsCfg) return;

  var nomes = [];

  // Try EP_CLIENTES external spreadsheet first
  try {
    var ssClientes = SpreadsheetApp.openById(ID_CLIENTES);
    var wsClientes = ssClientes.getSheets()[0];
    var dados = wsClientes.getDataRange().getValues();
    // Row 1 = header; column 0 = customer name
    for (var i = 1; i < dados.length; i++) {
      var nome = String(dados[i][0]).trim();
      if (nome) nomes.push(nome);
    }
  } catch (err) {
    Logger.log('EP_CLIENTES indisponivel, usando fallback col S: ' + err.message);
    // Fallback: column S of the CONFIG tab
    var colS = wsCfg.getRange(2, 19, wsCfg.getLastRow(), 1).getValues();
    colS.forEach(function(r) {
      var v = String(r[0]).trim();
      if (v) nomes.push(v);
    });
  }

  // Prepend standard tables
  var opcoes = Object.keys(STD_TABS).concat(nomes);

  if (opcoes.length === 0) {
    SpreadsheetApp.getUi().alert('Nenhum cliente encontrado. Verifique EP_CLIENTES.');
    return;
  }

  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(opcoes, true)
    .build();
  wsCfg.getRange('B2').setDataValidation(regra);
  Logger.log('Dropdown atualizado: ' + opcoes.length + ' opcoes.');
}


// ── [C] UTILITY FUNCTIONS ─────────────────────────────────────

/**
 * parsePreco_ — converts PT-BR price strings to numbers.
 * Handles: "R$ 1.234,56", "1.234,56", 1234.56, "", null.
 */
function parsePreco_(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var s = String(val).replace(/R\$\s*/g, '').trim();
  s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * removerAcentos_ — strips diacritics for accent-insensitive comparison.
 */
function removerAcentos_(str) {
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

/**
 * lerClassificacoes_ — reads SKU → classification mapping from the
 * CLASSIFICACAO tab. Returns an object keyed by SKU.
 */
function lerClassificacoes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CLASSIF);
  var mapa = {};
  if (!ws) return mapa;
  var dados = ws.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var sku    = String(dados[i][0]).trim();
    var classif = String(dados[i][1]).trim();
    if (sku) mapa[sku] = classif;
  }
  return mapa;
}

/**
 * lerUnidades_ — reads unit-of-measure (BD/TB/CX) from EP_BASE_CUSTOS.
 * Returns an object keyed by SKU.
 */
function lerUnidades_(wsCfg) {
  var unidades = {};
  try {
    var ssCustos = SpreadsheetApp.openById(ID_BASE_CUSTOS);
    // The units live on the first sheet of EP_BASE_CUSTOS (not RESULTADO)
    var wsUnid = ssCustos.getSheets()[0];
    var dados  = wsUnid.getDataRange().getValues();
    var hdr    = dados[0].map(function(c) { return removerAcentos_(c); });
    var iSku   = hdr.indexOf('SKU');
    var iUnid  = hdr.indexOf('UNID');
    if (iSku < 0 || iUnid < 0) return unidades;
    for (var i = 1; i < dados.length; i++) {
      var sku  = String(dados[i][iSku]).trim();
      var unid = String(dados[i][iUnid]).trim();
      if (sku) unidades[sku] = unid;
    }
  } catch (err) {
    Logger.log('lerUnidades_: ' + err.message);
  }
  return unidades;
}

/**
 * calcPrazoMedio_ — calculates weighted average payment term for a customer
 * from their sales history. Used to estimate financial carrying cost.
 * Returns number of days (float).
 */
function calcPrazoMedio_(historicoCliente) {
  var totalValor = 0;
  var totalPeso  = 0;
  historicoCliente.forEach(function(row) {
    var valor = parsePreco_(row.valor) || 0;
    var prazo = Number(row.prazo)      || 0;
    totalValor += valor;
    totalPeso  += valor * prazo;
  });
  return totalValor > 0 ? totalPeso / totalValor : 0;
}


// ── [D] DATA LOADERS ─────────────────────────────────────────

/**
 * lerCustos_ — reads production costs from the RESULTADO tab of
 * CostStructure_v2 (same spreadsheet as EP_BASE_CUSTOS).
 *
 * Returns an object keyed by SKU: { custo: number }
 *
 * Column H of RESULTADO = custo_c_oluc (production cost including OLUC 12%).
 * 537 PAs have a calculated cost; 75 PAs have no BOM and return 0.
 */
function lerCustos_() {
  var custos = {};
  try {
    var ssCustos   = SpreadsheetApp.openById(ID_BASE_CUSTOS);
    var wsResultado = ssCustos.getSheetByName('RESULTADO');
    if (!wsResultado) throw new Error('Aba RESULTADO nao encontrada em EP_BASE_CUSTOS');

    var dados = wsResultado.getDataRange().getValues();
    // Row 0 = header. Assumed columns: A=SKU, H=custo_c_oluc (col index 7)
    var hdr   = dados[0].map(function(c) { return removerAcentos_(c); });
    var iSku  = hdr.indexOf('SKU');
    // custo_c_oluc is column H (index 7); fallback to header search
    var iCusto = hdr.indexOf('CUSTO_C_OLUC');
    if (iCusto < 0) iCusto = 7;
    if (iSku < 0)   iSku   = 0;

    for (var i = 1; i < dados.length; i++) {
      var sku   = String(dados[i][iSku]).trim();
      var custo = parsePreco_(dados[i][iCusto]);
      if (sku) custos[sku] = custo;
    }
    Logger.log('lerCustos_: ' + Object.keys(custos).length + ' SKUs carregados.');
  } catch (err) {
    Logger.log('lerCustos_: ' + err.message);
  }
  return custos;
}

/**
 * lerParametros_ — reads all margin parameters from EP_PARAMETROS_MARGEM.
 *
 * Returns an object with the following structure:
 * {
 *   frete:            { UF: pct },           // e.g. { SP: 2.5, RJ: 3.0 }
 *   comissao:         { repCode: pct },       // default by rep code
 *   comissaoExcecoes: { cnpj: pct },          // CNPJ-level exceptions
 *   comissaoUF:       { UF: pct },            // special rules by state
 *   impostos:         { UF: pct },            // total tax burden by state
 *   margemPolitica:   [ { familia, classif, quadrante, min, alvo } ],
 *   custosAdicionais: { custoFin: pct, custoAdmin: pct },
 *   classifEstrategica: { sku: quadrante }    // strategic quadrant mapping
 * }
 */
function lerParametros_() {
  var params = {
    frete: {}, comissao: {}, comissaoExcecoes: {}, comissaoUF: {},
    impostos: {}, margemPolitica: [], custosAdicionais: {},
    classifEstrategica: {}
  };

  try {
    var ssParam = SpreadsheetApp.openById(ID_PARAMETROS);

    // FRETE tab: col A = UF, col B = pct
    var wsFrete = ssParam.getSheetByName('FRETE');
    if (wsFrete) {
      var dFrete = wsFrete.getDataRange().getValues();
      for (var i = 1; i < dFrete.length; i++) {
        var uf  = String(dFrete[i][0]).trim().toUpperCase();
        var pct = parseFloat(dFrete[i][1]) || 0;
        if (uf) params.frete[uf] = pct;
      }
    }

    // COMISSAO tab: col A = type (REP/CNPJ/UF), col B = key, col C = pct
    var wsComis = ssParam.getSheetByName('COMISSAO');
    if (wsComis) {
      var dComis = wsComis.getDataRange().getValues();
      for (var i = 1; i < dComis.length; i++) {
        var tipo = String(dComis[i][0]).trim().toUpperCase();
        var key  = String(dComis[i][1]).trim();
        var pct  = parseFloat(dComis[i][2]) || 0;
        if (tipo === 'REP')  params.comissao[key] = pct;
        if (tipo === 'CNPJ') params.comissaoExcecoes[key] = pct;
        if (tipo === 'UF')   params.comissaoUF[key.toUpperCase()] = pct;
      }
    }

    // IMPOSTOS tab: col A = UF, col B = total tax burden %
    var wsImp = ssParam.getSheetByName('IMPOSTOS');
    if (wsImp) {
      var dImp = wsImp.getDataRange().getValues();
      for (var i = 1; i < dImp.length; i++) {
        var uf  = String(dImp[i][0]).trim().toUpperCase();
        var pct = parseFloat(dImp[i][1]) || 0;
        if (uf) params.impostos[uf] = pct;
      }
    }

    // MARGEM_POLITICA tab: familia, classificacao, quadrante_estrategico, min%, alvo%
    var wsMargem = ssParam.getSheetByName('MARGEM_POLITICA');
    if (wsMargem) {
      var dMargem = wsMargem.getDataRange().getValues();
      for (var i = 1; i < dMargem.length; i++) {
        params.margemPolitica.push({
          familia:    String(dMargem[i][0]).trim().toUpperCase(),
          classif:    String(dMargem[i][1]).trim().toUpperCase(),
          quadrante:  String(dMargem[i][2]).trim().toUpperCase(),
          min:        parseFloat(dMargem[i][3]) || 0,
          alvo:       parseFloat(dMargem[i][4]) || 0
        });
      }
    }

    // CUSTOS_ADICIONAIS tab: custo financeiro mensal %, custo administrativo %
    var wsCustosAd = ssParam.getSheetByName('CUSTOS_ADICIONAIS');
    if (wsCustosAd) {
      var dCustosAd = wsCustosAd.getDataRange().getValues();
      dCustosAd.slice(1).forEach(function(row) {
        var chave = removerAcentos_(row[0]);
        var valor = parseFloat(row[1]) || 0;
        if (chave.indexOf('FIN') >= 0)   params.custosAdicionais.custoFin   = valor;
        if (chave.indexOf('ADMIN') >= 0) params.custosAdicionais.custoAdmin = valor;
      });
    }

    // CLASSIF_ESTRATEGICA tab: col A = SKU, col B = quadrante
    var wsEstr = ssParam.getSheetByName('CLASSIF_ESTRATEGICA');
    if (wsEstr) {
      var dEstr = wsEstr.getDataRange().getValues();
      for (var i = 1; i < dEstr.length; i++) {
        var sku = String(dEstr[i][0]).trim();
        var quad = String(dEstr[i][1]).trim().toUpperCase();
        if (sku) params.classifEstrategica[sku] = quad;
      }
    }

    Logger.log('lerParametros_: todos os parametros carregados.');
  } catch (err) {
    Logger.log('lerParametros_: ' + err.message);
  }
  return params;
}

/**
 * buscarMargem_ — hierarchical margin lookup.
 *
 * Priority (highest to lowest):
 *   1. Strategic quadrant match (sku-level via classifEstrategica)
 *   2. Classification match (col D = CLASSIF)
 *   3. Family match (col C = FAMILIA)
 *   4. Default: min=15%, alvo=25%
 *
 * Returns { min: number, alvo: number }
 */
function buscarMargem_(sku, familia, classif, params) {
  var politica = params.margemPolitica;
  var quad     = params.classifEstrategica[sku] || '';

  // 1. Strategic quadrant
  if (quad) {
    for (var i = 0; i < politica.length; i++) {
      if (politica[i].quadrante === quad &&
          politica[i].familia   === '' &&
          politica[i].classif   === '') {
        return { min: politica[i].min, alvo: politica[i].alvo };
      }
    }
  }

  // 2. Classification exact match
  var normClassif = removerAcentos_(classif);
  for (var i = 0; i < politica.length; i++) {
    if (removerAcentos_(politica[i].classif) === normClassif && normClassif !== '') {
      return { min: politica[i].min, alvo: politica[i].alvo };
    }
  }

  // 3. Family match
  var normFamilia = removerAcentos_(familia);
  for (var i = 0; i < politica.length; i++) {
    if (removerAcentos_(politica[i].familia) === normFamilia && normFamilia !== '') {
      return { min: politica[i].min, alvo: politica[i].alvo };
    }
  }

  // 4. Default
  return { min: 15, alvo: 25 };
}

/**
 * calcComissao_ — determines commission % for a given customer/product.
 *
 * Priority chain:
 *   1. CNPJ-level exception (params.comissaoExcecoes[cnpj])
 *   2. Special UF rule (params.comissaoUF[uf])
 *   3. Representative default (params.comissao[repCode])
 *   4. Fallback: 5%
 */
function calcComissao_(cnpj, uf, repCode, params) {
  if (cnpj && params.comissaoExcecoes[cnpj] !== undefined) {
    return params.comissaoExcecoes[cnpj];
  }
  var ufUp = String(uf).trim().toUpperCase();
  if (ufUp && params.comissaoUF[ufUp] !== undefined) {
    return params.comissaoUF[ufUp];
  }
  if (repCode && params.comissao[repCode] !== undefined) {
    return params.comissao[repCode];
  }
  return 5; // default 5%
}

/**
 * lerReajustes_ — reads adjustment percentages from the REAJUSTES tab.
 *
 * Supports three levels (applied additively):
 *   - Global: applies to all SKUs
 *   - Classification: applies to a specific classification group
 *   - Family: applies to a specific product family
 *
 * Also reads per-table adjustments for standard exports.
 *
 * Returns: { global: number, classif: { name: pct }, familia: { name: pct } }
 */
function lerReajustes_(wsRea) {
  var reajustes = { global: 0, classif: {}, familia: {} };
  if (!wsRea) return reajustes;

  var dados = wsRea.getDataRange().getValues();
  // Expected columns: A=tipo (GLOBAL/CLASSIF/FAMILIA), B=chave, C=pct
  for (var i = 1; i < dados.length; i++) {
    var tipo  = removerAcentos_(dados[i][0]);
    var chave = String(dados[i][1]).trim().toUpperCase();
    var pct   = parseFloat(dados[i][2]) || 0;

    if (tipo === 'GLOBAL')  reajustes.global = pct;
    if (tipo === 'CLASSIF' && chave) reajustes.classif[chave] = pct;
    if (tipo === 'FAMILIA' && chave) reajustes.familia[chave]  = pct;
  }
  return reajustes;
}


// ── [E] MAIN WORKFLOWS ────────────────────────────────────────

/**
 * carregarCliente — loads a customer (or standard table) into TABELA_NOVA.
 *
 * Steps:
 *  1. Read CONFIG B2 to get selected name
 *  2. If name matches a STD_TAB key → call carregarTabelaPadrao_()
 *  3. Else → look up the customer in EP_CLIENTES (CNPJ, UF, rep, PCT_FAT,
 *     TABELA_REF, prazo)
 *  4. Load production costs (lerCustos_) and margin parameters (lerParametros_)
 *  5. Load sales history from EP_BASE_VENDAS for this customer
 *  6. Load benchmark prices from EP_TABELAS_REF (customer's TABELA_REF)
 *  7. For each SKU in TABELA_NOVA: fill columns F–AB
 *     - Cost columns (F,G,H,I,J,K)
 *     - Margin columns (L,M) via buscarMargem_ hierarchy
 *     - Suggested prices (N,O)
 *     - Reference & last price (P,Q), history flag (R)
 *     - Base price (S): last price if history exists, else benchmark
 *     - Adjustment % (T) from REAJUSTES
 *     - New price (U), real margin (V,W), alert (X)
 *     - History stats (Y,Z,AA), source flag (AB)
 *
 * NOTE: PCT_FATURAMENTO is stored in the customer master but is NOT applied
 * to the price shown in the table. The table always shows the full invoice
 * price (preco NF cheio). PCT_FAT is only used to identify DAGDA-sourced
 * history rows (post Dec/2025 sales with PCT_FAT adjustment).
 */
function carregarCliente() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsCfg = ss.getSheetByName(ABA_CONFIG);
  var wsNova = ss.getSheetByName(ABA_NOVA);
  if (!wsCfg || !wsNova) {
    SpreadsheetApp.getUi().alert('Abas CONFIG ou TABELA_NOVA nao encontradas.');
    return;
  }

  var nomeCliente = String(wsCfg.getRange('B2').getValue()).trim();
  if (!nomeCliente) {
    SpreadsheetApp.getUi().alert('Selecione um cliente ou tabela em B2.');
    return;
  }

  // Standard table?
  if (STD_TABS[nomeCliente]) {
    carregarTabelaPadrao_(nomeCliente);
    return;
  }

  // ── Look up customer in EP_CLIENTES ──
  var dadosCliente = null;
  try {
    var ssClientes = SpreadsheetApp.openById(ID_CLIENTES);
    var wsClientes = ssClientes.getSheets()[0];
    var rows = wsClientes.getDataRange().getValues();
    var hdr  = rows[0].map(function(c) { return removerAcentos_(c); });
    var iNome  = hdr.indexOf('NOME');
    var iCnpj  = hdr.indexOf('CNPJ');
    var iUF    = hdr.indexOf('ESTADO');
    var iRep   = hdr.indexOf('REPRESENTANTE');
    var iPct   = hdr.indexOf('PCT_FATURAMENTO');
    var iTab   = hdr.indexOf('TABELA_REF');
    var iPrazo = hdr.indexOf('PRAZO');

    var nomeNorm = removerAcentos_(nomeCliente);
    for (var i = 1; i < rows.length; i++) {
      if (removerAcentos_(rows[i][iNome]) === nomeNorm) {
        dadosCliente = {
          nome:    rows[i][iNome],
          cnpj:    String(rows[i][iCnpj]).replace(/\D/g, ''),
          uf:      String(rows[i][iUF]).trim().toUpperCase(),
          rep:     String(rows[i][iRep]).trim(),
          pctFat:  parseFloat(rows[i][iPct]) || 1.0,
          tabelaRef: String(rows[i][iTab]).trim().toUpperCase() || 'PADRAO',
          prazo:   Number(rows[i][iPrazo]) || 30
        };
        break;
      }
    }
  } catch (err) {
    Logger.log('carregarCliente: erro ao ler EP_CLIENTES: ' + err.message);
  }

  if (!dadosCliente) {
    SpreadsheetApp.getUi().alert('Cliente "' + nomeCliente + '" nao encontrado em EP_CLIENTES.');
    return;
  }

  // Normalize: customers with TABELA_REF=PADRAO use LEME as benchmark source
  var abaRef = (dadosCliente.tabelaRef === 'PADRAO') ? 'LEME' : dadosCliente.tabelaRef;

  // ── Load external data ──
  var custos   = lerCustos_();
  var params   = lerParametros_();
  var classifMap = lerClassificacoes_();
  var wsRea    = ss.getSheetByName(ABA_REAJUSTES);
  var reajustes = lerReajustes_(wsRea);

  // Sales history for this customer: { sku: { ultPreco, dtUltVenda, qtdHist, nVendas, fonte } }
  var historico = {};
  try {
    var ssVendas  = SpreadsheetApp.openById(ID_BASE_VENDAS);
    var wsVendas  = ssVendas.getSheets()[0];
    var dVendas   = wsVendas.getDataRange().getValues();
    var hdrV      = dVendas[0].map(function(c) { return removerAcentos_(c); });
    var iVCnpj    = hdrV.indexOf('CNPJ');
    var iVSku     = hdrV.indexOf('SKU');
    var iVPreco   = hdrV.indexOf('PRECO_UNIT');
    var iVData    = hdrV.indexOf('DATA_VENDA');
    var iVQtd     = hdrV.indexOf('QTD');
    var iVNF      = hdrV.indexOf('NUM_NF');   // used as proxy for N_VENDAS
    var iVFonte   = hdrV.indexOf('FONTE');    // DAGDA or HISTORICO

    for (var i = 1; i < dVendas.length; i++) {
      var cnpjRow = String(dVendas[i][iVCnpj]).replace(/\D/g, '');
      if (cnpjRow !== dadosCliente.cnpj) continue;
      var sku = String(dVendas[i][iVSku]).trim();
      if (!sku) continue;

      var preco = parsePreco_(dVendas[i][iVPreco]);
      var data  = dVendas[i][iVData];
      var qtd   = Number(dVendas[i][iVQtd]) || 0;
      var fonte = String(dVendas[i][iVFonte]).trim() || 'HISTORICO';

      if (!historico[sku]) {
        historico[sku] = { ultPreco: preco, dtUltVenda: data, qtdHist: qtd,
                           nVendas: 1, fonte: fonte };
      } else {
        historico[sku].qtdHist += qtd;
        historico[sku].nVendas += 1;
        // Keep the most recent price/date
        if (data > historico[sku].dtUltVenda) {
          historico[sku].ultPreco   = preco;
          historico[sku].dtUltVenda = data;
          historico[sku].fonte      = fonte;
        }
      }
    }
  } catch (err) {
    Logger.log('carregarCliente: erro ao ler EP_BASE_VENDAS: ' + err.message);
  }

  // Benchmark prices for the customer's reference table
  var benchmarks = {};
  try {
    var ssRef  = SpreadsheetApp.openById(ID_TABELAS_REF);
    var wsRef  = ssRef.getSheetByName(abaRef);
    if (!wsRef) throw new Error('Aba ' + abaRef + ' nao encontrada em EP_TABELAS_REF');
    var dRef   = wsRef.getDataRange().getValues();
    for (var i = 1; i < dRef.length; i++) {
      var skuRef = String(dRef[i][0]).trim();
      var preco  = parsePreco_(dRef[i][4]); // col E = preco ref
      if (skuRef) benchmarks[skuRef] = preco;
    }
  } catch (err) {
    Logger.log('carregarCliente: erro ao ler EP_TABELAS_REF: ' + err.message);
  }

  // Commission % for this customer
  var pctComissao = calcComissao_(dadosCliente.cnpj, dadosCliente.uf,
                                  dadosCliente.rep, params);
  var pctFrete    = params.frete[dadosCliente.uf]   || 0;
  var pctImpostos = params.impostos[dadosCliente.uf] || 0;
  var pctCustoFin = params.custosAdicionais.custoFin || 0;

  // ── Write into TABELA_NOVA ──
  var ultimaLinha = wsNova.getLastRow();
  if (ultimaLinha < 2) return;

  var dados = wsNova.getRange(2, 1, ultimaLinha - 1, COL.SKU).getValues();

  for (var i = 0; i < dados.length; i++) {
    var sku     = String(dados[i][COL.SKU - 1]).trim();
    if (!sku) continue;

    var linhaPlan = i + 2; // 1-indexed, row 1 = header
    var classif   = classifMap[sku] || '';
    var familia   = String(wsNova.getRange(linhaPlan, COL.FAMILIA).getValue()).trim();

    var custoProd = custos[sku]      || 0;
    var margPolicy = buscarMargem_(sku, familia, classif, params);
    var custoTotalPct = pctFrete + pctComissao + pctImpostos + pctCustoFin;

    // Suggested prices: custo / (1 - custos% - margem%)
    var denominMin  = 1 - (custoTotalPct + margPolicy.min)  / 100;
    var denominAlvo = 1 - (custoTotalPct + margPolicy.alvo) / 100;
    var precoMin      = (denominMin  > 0 && custoProd > 0) ? custoProd / denominMin  : 0;
    var precoSugerido = (denominAlvo > 0 && custoProd > 0) ? custoProd / denominAlvo : 0;

    var precoRef  = benchmarks[sku] || 0;
    var hist      = historico[sku];
    var temHist   = hist ? true : false;
    var ultPreco  = temHist ? hist.ultPreco   : 0;
    var dtUlt     = temHist ? hist.dtUltVenda : '';
    var qtdHist   = temHist ? hist.qtdHist    : 0;
    var nVendas   = temHist ? hist.nVendas    : 0;
    var fonteStr  = temHist ? hist.fonte      : 'SEM HIST';

    var precoBase = temHist ? ultPreco : precoRef;

    // Aggregate adjustment: global + classif + family
    var normClassif  = classif.toUpperCase();
    var normFamilia  = familia.toUpperCase();
    var pctReajuste  = reajustes.global +
                       (reajustes.classif[normClassif] || 0) +
                       (reajustes.familia[normFamilia]  || 0);

    var novoPreco    = precoBase > 0 ? precoBase * (1 + pctReajuste / 100) : 0;

    // Real margin at novoPreco
    var margemRealRS  = novoPreco > 0 ? novoPreco - custoProd - novoPreco * custoTotalPct / 100 : 0;
    var margemRealPct = novoPreco > 0 ? (margemRealRS / novoPreco) * 100 : 0;

    // Alert
    var alerta;
    if (novoPreco === 0 || custoProd === 0) {
      alerta = 'SEM CUSTO';
    } else if (novoPreco < precoMin) {
      alerta = 'ABAIXO MINIMO';
    } else if (novoPreco < precoSugerido) {
      alerta = 'ABAIXO ALVO';
    } else if (novoPreco > precoSugerido * 1.2) {
      alerta = 'PREMIUM';
    } else {
      alerta = 'OK';
    }

    // Write all computed columns in one batch per row
    wsNova.getRange(linhaPlan, COL.CLASSIF).setValue(classif);
    wsNova.getRange(linhaPlan, COL.CUSTO_PROD).setValue(custoProd);
    wsNova.getRange(linhaPlan, COL.FRETE).setValue(pctFrete / 100);
    wsNova.getRange(linhaPlan, COL.COMISSAO).setValue(pctComissao / 100);
    wsNova.getRange(linhaPlan, COL.IMPOSTOS).setValue(pctImpostos / 100);
    wsNova.getRange(linhaPlan, COL.CUSTO_FIN).setValue(pctCustoFin / 100);
    wsNova.getRange(linhaPlan, COL.CUSTO_TOTAL).setValue(custoTotalPct / 100);
    wsNova.getRange(linhaPlan, COL.MARGEM_MIN).setValue(margPolicy.min / 100);
    wsNova.getRange(linhaPlan, COL.MARGEM_ALVO).setValue(margPolicy.alvo / 100);
    wsNova.getRange(linhaPlan, COL.PRECO_MIN).setValue(precoMin);
    wsNova.getRange(linhaPlan, COL.PRECO_SUGERIDO).setValue(precoSugerido);
    wsNova.getRange(linhaPlan, COL.PRECO_REF).setValue(precoRef);
    wsNova.getRange(linhaPlan, COL.ULT_PRECO).setValue(ultPreco);
    wsNova.getRange(linhaPlan, COL.TEM_HIST).setValue(temHist);
    wsNova.getRange(linhaPlan, COL.PRECO_BASE).setValue(precoBase);
    wsNova.getRange(linhaPlan, COL.REAJUSTE).setValue(pctReajuste / 100);
    wsNova.getRange(linhaPlan, COL.NOVO_PRECO).setValue(novoPreco);
    wsNova.getRange(linhaPlan, COL.MARGEM_REAL_PCT).setValue(margemRealPct / 100);
    wsNova.getRange(linhaPlan, COL.MARGEM_REAL_RS).setValue(margemRealRS);
    wsNova.getRange(linhaPlan, COL.ALERTA).setValue(alerta);
    wsNova.getRange(linhaPlan, COL.DT_ULT_VENDA).setValue(dtUlt);
    wsNova.getRange(linhaPlan, COL.QTD_HIST).setValue(qtdHist);
    wsNova.getRange(linhaPlan, COL.N_VENDAS).setValue(nVendas);
    wsNova.getRange(linhaPlan, COL.FONTE).setValue(fonteStr);
  }

  SpreadsheetApp.getUi().alert('Cliente carregado: ' + nomeCliente);
}

/**
 * carregarTabelaPadrao_ — loads a standard (non-customer) price table.
 *
 * Note: PADRAO GERAL uses the LEME tab as its price source, not PADRAO.
 * This ensures that the standard table reflects the commercial benchmark.
 */
function carregarTabelaPadrao_(nomeTabela) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var wsNova = ss.getSheetByName(ABA_NOVA);
  if (!wsNova) return;

  var abaSrc = STD_TABS[nomeTabela];
  // Customers with TABELA_REF=PADRAO use LEME as source (same rule applies here)
  if (abaSrc === 'PADRAO') abaSrc = 'LEME';

  try {
    var ssRef = SpreadsheetApp.openById(ID_TABELAS_REF);
    var wsRef = ssRef.getSheetByName(abaSrc);
    if (!wsRef) throw new Error('Aba ' + abaSrc + ' nao encontrada em EP_TABELAS_REF');

    var dRef = wsRef.getDataRange().getValues();
    // Clear existing data in TABELA_NOVA (keep header row 1)
    var ultimaLinha = wsNova.getLastRow();
    if (ultimaLinha > 1) {
      wsNova.getRange(2, 1, ultimaLinha - 1, COL.FONTE).clearContent();
    }

    var linhaDest = 2;
    for (var i = 1; i < dRef.length; i++) {
      var sku    = String(dRef[i][0]).trim();
      if (!sku) continue;
      var prod   = dRef[i][1];
      var familia = dRef[i][2];
      var unid   = dRef[i][3];
      var preco  = parsePreco_(dRef[i][4]);

      wsNova.getRange(linhaDest, COL.SKU).setValue(sku);
      wsNova.getRange(linhaDest, COL.PRODUTO).setValue(prod);
      wsNova.getRange(linhaDest, COL.FAMILIA).setValue(familia);
      wsNova.getRange(linhaDest, COL.UNID).setValue(unid);
      wsNova.getRange(linhaDest, COL.PRECO_BASE).setValue(preco);
      wsNova.getRange(linhaDest, COL.NOVO_PRECO).setValue(preco);
      linhaDest++;
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert('Erro ao carregar tabela padrao: ' + err.message);
  }
}

/**
 * corrigirFormulas — "Recalcular Margens"
 *
 * Re-reads costs and parameters and recalculates all margin/alert columns
 * for the currently loaded table WITHOUT reloading prices or history.
 * Use this after changing margin policy or adjustment percentages.
 *
 * Recalculates: F (cost), K (total cost%), L/M (margins), N/O (prices),
 * T (adjustment), U (new price), V/W (real margin), X (alert).
 */
function corrigirFormulas() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var wsNova = ss.getSheetByName(ABA_NOVA);
  var wsCfg  = ss.getSheetByName(ABA_CONFIG);
  if (!wsNova || !wsCfg) return;

  var custos   = lerCustos_();
  var params   = lerParametros_();
  var classifMap = lerClassificacoes_();
  var wsRea    = ss.getSheetByName(ABA_REAJUSTES);
  var reajustes = lerReajustes_(wsRea);

  // Read customer context from current table header
  var nomeCliente = String(wsCfg.getRange('B2').getValue()).trim();
  // For cost parameters we still need customer UF/rep — re-read CONFIG
  var uf  = String(wsCfg.getRange('B3').getValue()).trim().toUpperCase(); // UF in B3
  var rep = String(wsCfg.getRange('B4').getValue()).trim();               // REP in B4

  var pctComissao = calcComissao_('', uf, rep, params);
  var pctFrete    = params.frete[uf]    || 0;
  var pctImpostos = params.impostos[uf] || 0;
  var pctCustoFin = params.custosAdicionais.custoFin || 0;
  var custoTotalPct = pctFrete + pctComissao + pctImpostos + pctCustoFin;

  var ultimaLinha = wsNova.getLastRow();
  if (ultimaLinha < 2) return;

  for (var i = 2; i <= ultimaLinha; i++) {
    var sku     = String(wsNova.getRange(i, COL.SKU).getValue()).trim();
    if (!sku) continue;

    var familia  = String(wsNova.getRange(i, COL.FAMILIA).getValue()).trim();
    var classif  = classifMap[sku] || String(wsNova.getRange(i, COL.CLASSIF).getValue()).trim();
    var precoBase = parsePreco_(wsNova.getRange(i, COL.PRECO_BASE).getValue());

    var custoProd   = custos[sku] || parsePreco_(wsNova.getRange(i, COL.CUSTO_PROD).getValue());
    var margPolicy  = buscarMargem_(sku, familia, classif, params);

    var denominMin  = 1 - (custoTotalPct + margPolicy.min)  / 100;
    var denominAlvo = 1 - (custoTotalPct + margPolicy.alvo) / 100;
    var precoMin      = (denominMin  > 0 && custoProd > 0) ? custoProd / denominMin  : 0;
    var precoSugerido = (denominAlvo > 0 && custoProd > 0) ? custoProd / denominAlvo : 0;

    var normClassif = classif.toUpperCase();
    var normFamilia = familia.toUpperCase();
    var pctReajuste = reajustes.global +
                      (reajustes.classif[normClassif] || 0) +
                      (reajustes.familia[normFamilia]  || 0);

    var novoPreco    = precoBase > 0 ? precoBase * (1 + pctReajuste / 100) : 0;
    var margemRealRS  = novoPreco > 0 ? novoPreco - custoProd - novoPreco * custoTotalPct / 100 : 0;
    var margemRealPct = novoPreco > 0 ? (margemRealRS / novoPreco) * 100 : 0;

    var alerta;
    if (novoPreco === 0 || custoProd === 0) {
      alerta = 'SEM CUSTO';
    } else if (novoPreco < precoMin) {
      alerta = 'ABAIXO MINIMO';
    } else if (novoPreco < precoSugerido) {
      alerta = 'ABAIXO ALVO';
    } else if (novoPreco > precoSugerido * 1.2) {
      alerta = 'PREMIUM';
    } else {
      alerta = 'OK';
    }

    wsNova.getRange(i, COL.CUSTO_PROD).setValue(custoProd);
    wsNova.getRange(i, COL.CUSTO_TOTAL).setValue(custoTotalPct / 100);
    wsNova.getRange(i, COL.MARGEM_MIN).setValue(margPolicy.min / 100);
    wsNova.getRange(i, COL.MARGEM_ALVO).setValue(margPolicy.alvo / 100);
    wsNova.getRange(i, COL.PRECO_MIN).setValue(precoMin);
    wsNova.getRange(i, COL.PRECO_SUGERIDO).setValue(precoSugerido);
    wsNova.getRange(i, COL.REAJUSTE).setValue(pctReajuste / 100);
    wsNova.getRange(i, COL.NOVO_PRECO).setValue(novoPreco);
    wsNova.getRange(i, COL.MARGEM_REAL_PCT).setValue(margemRealPct / 100);
    wsNova.getRange(i, COL.MARGEM_REAL_RS).setValue(margemRealRS);
    wsNova.getRange(i, COL.ALERTA).setValue(alerta);
  }

  SpreadsheetApp.getUi().alert('Margens recalculadas.');
}


// ── [F] EXPORT FUNCTIONS ──────────────────────────────────────

/**
 * exportarTabelaAtual — exports the currently loaded table (single customer
 * or standard table) to a formatted Google Sheets file in a Drive folder.
 */
function exportarTabelaAtual() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var wsCfg  = ss.getSheetByName(ABA_CONFIG);
  var wsNova = ss.getSheetByName(ABA_NOVA);
  if (!wsCfg || !wsNova) return;

  var nomeCliente  = String(wsCfg.getRange('B2').getValue()).trim();
  var dataVigencia = wsCfg.getRange('B5').getValue(); // DATA_VIGENCIA cell

  var isVarejo = (nomeCliente === 'PADRAO VAREJO');
  var itens    = isVarejo
    ? montarItensVarejo_(wsNova)
    : montarItens_(wsNova);

  if (itens.length === 0) {
    SpreadsheetApp.getUi().alert('Nenhum item encontrado. Carregue o cliente primeiro.');
    return;
  }

  var logoUrl = String(wsCfg.getRange('B6').getValue()).trim(); // LOGO_URL cell
  criarArquivoTabela_(nomeCliente, itens, dataVigencia, logoUrl, isVarejo);
  SpreadsheetApp.getUi().alert('Tabela exportada: ' + nomeCliente);
}

/**
 * exportarTodasTabelas — batch export:
 *  - 4 standard tables (PADRAO GERAL, PADRAO RJ, PADRAO CONSUMO, PADRAO VAREJO)
 *  - All customers from EP_CLIENTES
 *
 * Each table is saved as a separate Sheets file inside a dedicated Drive folder.
 */
function exportarTodasTabelas() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsCfg = ss.getSheetByName(ABA_CONFIG);
  if (!wsCfg) return;

  var dataVigencia = wsCfg.getRange('B5').getValue();
  var logoUrl      = String(wsCfg.getRange('B6').getValue()).trim();

  var nomesExportar = Object.keys(STD_TABS);

  // Add all customers
  try {
    var ssClientes = SpreadsheetApp.openById(ID_CLIENTES);
    var wsClientes = ssClientes.getSheets()[0];
    var rows = wsClientes.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var nome = String(rows[i][0]).trim();
      if (nome) nomesExportar.push(nome);
    }
  } catch (err) {
    Logger.log('exportarTodasTabelas: nao foi possivel carregar EP_CLIENTES: ' + err.message);
  }

  var exportados = 0;
  nomesExportar.forEach(function(nome) {
    try {
      wsCfg.getRange('B2').setValue(nome);
      carregarCliente();
      exportarTabelaAtual();
      exportados++;
    } catch (err) {
      Logger.log('exportarTodasTabelas: erro em ' + nome + ': ' + err.message);
    }
  });

  SpreadsheetApp.getUi().alert('Exportacao concluida: ' + exportados + ' tabelas geradas.');
}

/**
 * montarItens_ — builds the item array for a standard (non-VAREJO) table.
 * Reads columns: SKU, PRODUTO, FAMILIA, UNID, NOVO_PRECO from TABELA_NOVA.
 * Returns array of { sku, produto, familia, unid, preco }.
 */
function montarItens_(wsNova) {
  var ultimaLinha = wsNova.getLastRow();
  if (ultimaLinha < 2) return [];

  var dados = wsNova.getRange(2, 1, ultimaLinha - 1, COL.FONTE).getValues();
  var itens = [];

  dados.forEach(function(row) {
    var sku    = String(row[COL.SKU - 1]).trim();
    if (!sku) return;
    var preco  = parsePreco_(row[COL.NOVO_PRECO - 1]);
    if (preco <= 0) return;

    itens.push({
      sku:     sku,
      produto: row[COL.PRODUTO - 1],
      familia: String(row[COL.FAMILIA - 1]).trim(),
      unid:    row[COL.UNID - 1],
      preco:   preco
    });
  });

  return itens;
}

/**
 * montarItensVarejo_ — builds the item array for the VAREJO table.
 *
 * VAREJO layout uses 4 price columns (base, +10%, +20%, +50% vol tiers)
 * plus FAMILIA. It does NOT show QTD or VALOR TOTAL columns.
 *
 * Returns array of { sku, produto, familia, unid, p1, p2, p3, p4 }.
 */
function montarItensVarejo_(wsNova) {
  var ultimaLinha = wsNova.getLastRow();
  if (ultimaLinha < 2) return [];

  var dados = wsNova.getRange(2, 1, ultimaLinha - 1, COL.FONTE).getValues();
  var itens = [];

  dados.forEach(function(row) {
    var sku   = String(row[COL.SKU - 1]).trim();
    if (!sku) return;
    var base  = parsePreco_(row[COL.NOVO_PRECO - 1]);
    if (base <= 0) return;

    itens.push({
      sku:     sku,
      produto: row[COL.PRODUTO - 1],
      familia: String(row[COL.FAMILIA - 1]).trim(),
      unid:    row[COL.UNID - 1],
      p1:      base,
      p2:      base * 1.10,   // Acima 10 Vol
      p3:      base * 1.20,   // Acima 20 Vol
      p4:      base * 1.50    // Acima 50 Vol
    });
  });

  return itens;
}

/**
 * criarArquivoTabela_ — creates a fully formatted Google Sheets file in Drive.
 *
 * Formatting applied:
 *  - Row 1: "Tabela de Precos [Company] Lubrificantes" (16pt bold, brand color)
 *  - Row 2: customer name in title case (13pt bold, brand color)
 *  - Row 3: validity date (centered, brand color)
 *  - Row 4: column headers (white text on brand color, bold)
 *  - Data rows: zebra striping (white / COR_LARANJA_CLARO)
 *  - Family separator rows: COR_LARANJA (SOLID_MEDIUM border), white bold text
 *  - Gridlines: hidden
 *  - Column A (SKU): always hidden
 *  - Logo: embedded from LOGO_URL in top-right corner
 *
 * @param {string}  nomeCliente  — customer or table name
 * @param {Array}   itens        — item array from montarItens_ or montarItensVarejo_
 * @param {Date}    dataVigencia — validity date
 * @param {string}  logoUrl      — URL of the company logo (Drive download link)
 * @param {boolean} isVarejo     — true = 4-tier VAREJO layout
 */
function criarArquivoTabela_(nomeCliente, itens, dataVigencia, logoUrl, isVarejo) {
  // Create file
  var nomeArquivo = 'Tabela ' + nomeCliente + ' — ' +
                    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var novoSS = SpreadsheetApp.create(nomeArquivo);
  var ws     = novoSS.getActiveSheet();
  ws.setName('TABELA');

  // Column headers
  var colunas = isVarejo
    ? ['SKU', 'PRODUTO', 'UNID', 'VAREJO', 'Acima 10 Vol', 'Acima 20 Vol', 'Acima 50 Vol', 'FAMILIA']
    : ['SKU', 'PRODUTO', 'UNID', 'PRECO', 'FAMILIA'];
  var nCols = colunas.length;

  // ── Header rows ──
  // Row 1: company title
  ws.getRange(1, 1, 1, nCols).merge()
    .setValue('Tabela de Precos — Lubrificantes')   // Sanitized: no real company name
    .setBackground(COR_LARANJA)
    .setFontColor(COR_BRANCO)
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  // Row 2: customer name
  var nomeFormatado = nomeCliente.split(' ').map(function(p) {
    return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
  }).join(' ');
  ws.getRange(2, 1, 1, nCols).merge()
    .setValue(nomeFormatado)
    .setBackground(COR_LARANJA)
    .setFontColor(COR_BRANCO)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  // Row 3: validity
  var vigStr = dataVigencia
    ? 'Vigencia: ' + Utilities.formatDate(
        new Date(dataVigencia), Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : 'Vigencia: —';
  ws.getRange(3, 1, 1, nCols).merge()
    .setValue(vigStr)
    .setBackground(COR_LARANJA)
    .setFontColor(COR_BRANCO)
    .setFontSize(10)
    .setHorizontalAlignment('center');

  // Row 4: column headers
  ws.getRange(4, 1, 1, nCols).setValues([colunas])
    .setBackground(COR_LARANJA_ESC)
    .setFontColor(COR_BRANCO)
    .setFontWeight('bold');

  // ── Data rows ──
  var linhaAtual  = 5;
  var ultimaFamilia = null;
  var zebraIndex    = 0;

  itens.forEach(function(item) {
    // Family separator
    if (item.familia !== ultimaFamilia) {
      var sepRange = ws.getRange(linhaAtual, 1, 1, nCols);
      sepRange.merge()
        .setValue(item.familia)
        .setBackground(COR_LARANJA)
        .setFontColor(COR_BRANCO)
        .setFontWeight('bold')
        .setBorder(true, true, true, true, false, false,
                   COR_LARANJA, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      linhaAtual++;
      ultimaFamilia = item.familia;
      zebraIndex    = 0;
    }

    // Data row
    var bg       = (zebraIndex % 2 === 0) ? COR_BRANCO : COR_LARANJA_CLARO;
    var rowRange = ws.getRange(linhaAtual, 1, 1, nCols);
    rowRange.setBackground(bg);

    if (isVarejo) {
      rowRange.setValues([[
        item.sku, item.produto, item.unid,
        item.p1, item.p2, item.p3, item.p4,
        item.familia
      ]]);
      ws.getRange(linhaAtual, 4, 1, 4).setNumberFormat('R$ #,##0.00');
    } else {
      rowRange.setValues([[item.sku, item.produto, item.unid, item.preco, item.familia]]);
      ws.getRange(linhaAtual, 4).setNumberFormat('R$ #,##0.00');
    }

    linhaAtual++;
    zebraIndex++;
  });

  // ── Finishing touches ──
  // Hide gridlines
  ws.setHiddenGridlines(true);

  // Hide SKU column (column A)
  ws.hideColumns(1);

  // Auto-resize visible columns
  ws.autoResizeColumns(2, nCols - 1);

  // Embed logo (top-right)
  if (logoUrl) {
    try {
      var logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
      var img = ws.insertImage(logoBlob, nCols, 1);
      img.setAnchorCell(ws.getRange(1, nCols));
    } catch (logoErr) {
      Logger.log('criarArquivoTabela_: nao foi possivel inserir logo: ' + logoErr.message);
    }
  }

  Logger.log('Arquivo criado: ' + nomeArquivo + ' (' + itens.length + ' itens)');
}
