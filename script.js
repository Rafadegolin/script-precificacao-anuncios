// ========== CONFIGURAÇÃO GLOBAL ==========
const CONFIG = {
  // Aba principal
  ABA_PRINCIPAL: 'Precificação',
  ABA_LISTA: 'Lista precificação',
  
  // Abas de frete
  ABA_FRETE_VERDE: 'Verde',
  ABA_FRETE_AMARELO: 'Amarelo',
  ABA_FRETE_VERMELHO: 'Vermelho',
  
  // Células de INPUT
  NOME_PRODUTO: 'B5',
  TIPO_ANUNCIO: 'F5',
  PRECO_VENDA: 'H5',
  PRECO_CUSTO: 'B9',
  DESPESAS_FIXAS: 'D9',
  DESPESAS_OPERACIONAIS: 'F9',
  IMPOSTO: 'H9',
  COMISSAO_ML: 'B13',
  TAXA_FIXA_ML: 'D13',
  PESO_PRODUTO: 'F13',
  REPUTACAO: 'H13',
  COOPARTICIPACAO: 'B17',
  CUSTO_FRETE: 'D17',
  CUPOM: 'F17',
  
  // Células CALCULADAS (Resumo - apenas visualização)
  CUSTO_TOTAL: 'K9',
  MARGEM_LIQUIDA: 'K13',
  MARGEM_PERCENTUAL: 'K17',
  
  // Linha inicial do resumo
  LINHA_INICIAL_RESUMO: 23
};

// ========== INICIALIZAR FÓRMULAS DO RESUMO ==========
function inicializarFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
  
  // Fórmula Custo Total em K9
  aba.getRange(CONFIG.CUSTO_TOTAL).setFormula(
    '=B9 + D9 + F9 + (H5*H9) + (H5*B13) + D13 + D17 + (H5*F17)'
  );
  
  // Fórmula Margem Líquida em K13
  aba.getRange(CONFIG.MARGEM_LIQUIDA).setFormula(
    '=H5 - K9 - B17'
  );
  
  // Fórmula Margem Percentual em K17
  aba.getRange(CONFIG.MARGEM_PERCENTUAL).setFormula(
    '=IF(H5=0; 0; K13/H5)'
  );
  
  // Aplica formatação
  aba.getRange(CONFIG.CUSTO_TOTAL).setNumberFormat('R$ #,##0.00');
  aba.getRange(CONFIG.MARGEM_LIQUIDA).setNumberFormat('R$ #,##0.00');
  aba.getRange(CONFIG.MARGEM_PERCENTUAL).setNumberFormat('0.00%');
  
  SpreadsheetApp.getUi().alert('✅ Fórmulas inicializadas com sucesso!');
}

// ========== FUNÇÃO PRINCIPAL - LISTAR PRECIFICAÇÃO ==========
function listarPrecificacao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCalculo = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
  var abaLista = ss.getSheetByName(CONFIG.ABA_LISTA);
  
  // Cria a aba Lista se não existir
  if (!abaLista) {
    abaLista = ss.insertSheet(CONFIG.ABA_LISTA);
    abaLista.getRange('B1:I1').setValues([[
      'PRODUTO', 'TIPO', 'P. VENDA', 'CUSTOS', 'CUPOM', 'COOPART.', 'MARGEM R$', 'MARGEM %'
    ]]);
    abaLista.getRange('B1:I1').setFontWeight('bold');
    abaLista.getRange('B1:I1').setBackground('#4285F4');
    abaLista.getRange('B1:I1').setFontColor('#FFFFFF');
  }
  
  // ========== CAPTURA OS DADOS ==========
  var nomeProduto = abaCalculo.getRange(CONFIG.NOME_PRODUTO).getValue();
  var tipoAnuncio = abaCalculo.getRange(CONFIG.TIPO_ANUNCIO).getValue();
  var precoVenda = abaCalculo.getRange(CONFIG.PRECO_VENDA).getValue();
  var precoCusto = abaCalculo.getRange(CONFIG.PRECO_CUSTO).getValue();
  var cupom = abaCalculo.getRange(CONFIG.CUPOM).getValue();
  var cooparticipacao = abaCalculo.getRange(CONFIG.COOPARTICIPACAO).getValue();
  
  // Lê os valores calculados das células de resumo (K9, K13, K17)
  var custoTotal = abaCalculo.getRange(CONFIG.CUSTO_TOTAL).getValue();
  var margemLiquida = abaCalculo.getRange(CONFIG.MARGEM_LIQUIDA).getValue();
  var margemPercentual = abaCalculo.getRange(CONFIG.MARGEM_PERCENTUAL).getValue();
  
  // ========== VALIDAÇÕES (apenas dos campos essenciais) ==========
  if (!nomeProduto || nomeProduto === '') {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, preencha o NOME DO PRODUTO antes de listar.');
    return;
  }
  
  if (!precoVenda || precoVenda === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, preencha o PREÇO DE VENDA antes de listar.');
    return;
  }
  
  if (!precoCusto || precoCusto === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, preencha o PREÇO DE CUSTO antes de listar.');
    return;
  }
  
  if (!tipoAnuncio || tipoAnuncio === '') {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, selecione o TIPO DE ANÚNCIO antes de listar.');
    return;
  }
  
  // ========== ADICIONA NA ABA "Lista precificação" ==========
  var ultimaLinhaLista = abaLista.getLastRow() + 1;
  
  abaLista.getRange(ultimaLinhaLista, 2, 1, 8).setValues([[
    nomeProduto,              // B - PRODUTO
    tipoAnuncio,              // C - TIPO
    precoVenda,               // D - P. VENDA
    custoTotal || 0,          // E - CUSTOS (lido de K9)
    cupom || 0,               // F - CUPOM
    cooparticipacao || 0,     // G - COOPART.
    margemLiquida || 0,       // H - MARGEM R$ (lido de K13)
    margemPercentual || 0     // I - MARGEM % (lido de K17)
  ]]);
  
  // Formata valores
  abaLista.getRange(ultimaLinhaLista, 4).setNumberFormat('R$ #,##0.00'); // D - P. VENDA
  abaLista.getRange(ultimaLinhaLista, 5).setNumberFormat('R$ #,##0.00'); // E - CUSTOS
  abaLista.getRange(ultimaLinhaLista, 6).setNumberFormat('0.00%'); // F - CUPOM
  abaLista.getRange(ultimaLinhaLista, 7).setNumberFormat('R$ #,##0.00'); // G - COOPART.
  abaLista.getRange(ultimaLinhaLista, 8).setNumberFormat('R$ #,##0.00'); // H - MARGEM R$
  abaLista.getRange(ultimaLinhaLista, 9).setNumberFormat('0.00%'); // I - MARGEM %
  
  // ========== ADICIONA NO RESUMO DA ABA PRINCIPAL ==========
  var ultimaLinhaResumo = CONFIG.LINHA_INICIAL_RESUMO;
  var rangeResumo = abaCalculo.getRange('B' + CONFIG.LINHA_INICIAL_RESUMO + ':B1000');
  var valoresResumo = rangeResumo.getValues();
  
  // Encontra primeira linha vazia
  for (var i = 0; i < valoresResumo.length; i++) {
    if (valoresResumo[i][0] === '' || valoresResumo[i][0] === null) {
      ultimaLinhaResumo = CONFIG.LINHA_INICIAL_RESUMO + i;
      break;
    }
  }
  
  // Adiciona dados no resumo (linha 22 = cabeçalho, linha 23+ = dados)
  abaCalculo.getRange(ultimaLinhaResumo, 2).setValue(nomeProduto);        // B
  abaCalculo.getRange(ultimaLinhaResumo, 6).setValue(tipoAnuncio);        // F
  abaCalculo.getRange(ultimaLinhaResumo, 7).setValue(precoVenda);         // G
  abaCalculo.getRange(ultimaLinhaResumo, 8).setValue(custoTotal || 0);    // H
  abaCalculo.getRange(ultimaLinhaResumo, 9).setValue(cupom || 0);         // I
  abaCalculo.getRange(ultimaLinhaResumo, 10).setValue(cooparticipacao || 0); // J
  abaCalculo.getRange(ultimaLinhaResumo, 11).setValue(margemLiquida || 0); // K
  abaCalculo.getRange(ultimaLinhaResumo, 12).setValue(margemPercentual || 0); // L
  
  // Formata valores no resumo
  abaCalculo.getRange(ultimaLinhaResumo, 7).setNumberFormat('R$ #,##0.00');  // G
  abaCalculo.getRange(ultimaLinhaResumo, 8).setNumberFormat('R$ #,##0.00');  // H
  abaCalculo.getRange(ultimaLinhaResumo, 9).setNumberFormat('0.00%');        // I
  abaCalculo.getRange(ultimaLinhaResumo, 10).setNumberFormat('R$ #,##0.00'); // J
  abaCalculo.getRange(ultimaLinhaResumo, 11).setNumberFormat('R$ #,##0.00'); // K
  abaCalculo.getRange(ultimaLinhaResumo, 12).setNumberFormat('0.00%');       // L
  
  // Limpa os campos (MAS NÃO LIMPA D13 QUE CONTÉM A FÓRMULA DO IF)
  limparCampos(abaCalculo);
  
  SpreadsheetApp.getUi().alert('✅ Produto "' + nomeProduto + '" adicionado com sucesso!');
}

// ========== TRIGGERS - ATUALIZA EM TEMPO REAL ==========
function onEdit(e) {
  var range = e.range;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
  // Só executa na aba Precificação
  if (sheetName !== CONFIG.ABA_PRINCIPAL) return;
  
  var cellA1 = range.getA1Notation();
  
  // ========== SE MUDAR PESO, PREÇO OU REPUTAÇÃO - CALCULA FRETE ==========
  if (cellA1 === CONFIG.PESO_PRODUTO || 
      cellA1 === CONFIG.PRECO_VENDA || 
      cellA1 === CONFIG.REPUTACAO) {
    buscarCustoFrete();
  }
  
  // ========== SE MUDAR QUALQUER CAMPO DE CÁLCULO - ATUALIZA FORMATAÇÃO ==========
  if (cellA1 === CONFIG.PRECO_VENDA ||
      cellA1 === CONFIG.PRECO_CUSTO ||
      cellA1 === CONFIG.DESPESAS_FIXAS ||
      cellA1 === CONFIG.DESPESAS_OPERACIONAIS ||
      cellA1 === CONFIG.IMPOSTO ||
      cellA1 === CONFIG.COMISSAO_ML ||
      cellA1 === CONFIG.TAXA_FIXA_ML ||
      cellA1 === CONFIG.COOPARTICIPACAO ||
      cellA1 === CONFIG.CUSTO_FRETE ||
      cellA1 === CONFIG.CUPOM) {
    
    // Aplica formatação ao resumo
    aplicarFormatacaoResumo(sheet);
    
    Logger.log('📊 Resumo atualizado em tempo real');
  }
}

// ========== APLICAR FORMATAÇÃO AO RESUMO ==========
function aplicarFormatacaoResumo(sheet) {
  var custoTotalRange = sheet.getRange(CONFIG.CUSTO_TOTAL);
  var margemLiquidaRange = sheet.getRange(CONFIG.MARGEM_LIQUIDA);
  var margemPercentualRange = sheet.getRange(CONFIG.MARGEM_PERCENTUAL);
  
  custoTotalRange.setNumberFormat('R$ #,##0.00');
  margemLiquidaRange.setNumberFormat('R$ #,##0.00');
  margemPercentualRange.setNumberFormat('0.00%');
}

// ========== BUSCAR CUSTO DE FRETE AUTOMATICAMENTE ==========
function buscarCustoFrete() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCalculo = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
  
  // Captura dados necessários
  var peso = abaCalculo.getRange(CONFIG.PESO_PRODUTO).getValue();
  var precoVenda = abaCalculo.getRange(CONFIG.PRECO_VENDA).getValue();
  var reputacao = String(abaCalculo.getRange(CONFIG.REPUTACAO).getValue()).trim();
  
  // Se algum campo está vazio, não calcula
  if (!peso || peso === 0 || !precoVenda || precoVenda === 0 || !reputacao) {
    Logger.log('⚠️ Campos incompletos para cálculo de frete. Aguardando preenchimento...');
    return;
  }
  
  // Determina qual aba de frete usar
  var abaFrete;
  if (reputacao === 'Verde' || reputacao === 'Sem reputação') {
    abaFrete = ss.getSheetByName(CONFIG.ABA_FRETE_VERDE);
  } else if (reputacao === 'Amarela' || reputacao === 'Amarelo') {
    abaFrete = ss.getSheetByName(CONFIG.ABA_FRETE_AMARELO);
  } else if (reputacao === 'Laranja' || reputacao === 'Vermelha' || reputacao === 'Vermelho') {
    abaFrete = ss.getSheetByName(CONFIG.ABA_FRETE_VERMELHO);
  } else {
    Logger.log('❌ Reputação inválida: ' + reputacao);
    return;
  }
  
  if (!abaFrete) {
    Logger.log('❌ Aba de frete não encontrada para reputação: ' + reputacao);
    return;
  }
  
  // Determina a linha do peso
  var linhaPeso = getLinhaPorPeso(peso);
  if (linhaPeso === -1) {
    Logger.log('❌ Peso não encontrado nas tabelas: ' + peso + ' kg');
    return;
  }
  
  // Determina a coluna do preço (apenas para Verde e Amarelo)
  var colunaPreco;
  if (reputacao === 'Laranja' || reputacao === 'Vermelha' || reputacao === 'Vermelho') {
    colunaPreco = 2; // Coluna B (vermelho não tem faixas de preço)
  } else {
    colunaPreco = getColunaPorPreco(precoVenda);
    if (colunaPreco === -1) {
      Logger.log('❌ Preço não encontrado nas faixas: R$ ' + precoVenda);
      return;
    }
  }
  
  // Busca o valor do frete
  var custoFrete = abaFrete.getRange(linhaPeso, colunaPreco).getValue();
  
  // Atualiza a célula de custo de frete
  abaCalculo.getRange(CONFIG.CUSTO_FRETE).setValue(custoFrete);
  abaCalculo.getRange(CONFIG.CUSTO_FRETE).setNumberFormat('R$ #,##0.00');
  
  Logger.log('✅ Custo de frete atualizado: R$ ' + custoFrete.toFixed(2));
}

// ========== FUNÇÕES AUXILIARES ==========

// Retorna a linha correspondente ao peso
function getLinhaPorPeso(peso) {
  if (peso <= 0.3) return 2;
  if (peso <= 0.5) return 3;
  if (peso <= 1) return 4;
  if (peso <= 2) return 5;
  if (peso <= 3) return 6;
  if (peso <= 4) return 7;
  if (peso <= 5) return 8;
  if (peso <= 9) return 9;
  if (peso <= 13) return 10;
  if (peso <= 17) return 11;
  if (peso <= 23) return 12;
  if (peso <= 30) return 13;
  if (peso <= 40) return 14;
  if (peso <= 50) return 15;
  if (peso <= 60) return 16;
  if (peso <= 70) return 17;
  if (peso <= 80) return 18;
  if (peso <= 90) return 19;
  if (peso <= 100) return 20;
  if (peso <= 125) return 21;
  if (peso <= 150) return 22;
  if (peso > 150) return 23;
  return -1;
}

// Retorna a coluna correspondente ao preço de venda
function getColunaPorPreco(preco) {
  if (preco < 79) return 2;        // Coluna B
  if (preco < 100) return 3;       // Coluna C (79-99.99)
  if (preco < 120) return 4;       // Coluna D (100-119.99)
  if (preco < 150) return 5;       // Coluna E (120-149.99)
  if (preco < 200) return 6;       // Coluna F (150-199.99)
  if (preco >= 200) return 7;      // Coluna G (200+)
  return -1;
}

// ========== LIMPAR CAMPOS ==========
function limparCampos(aba) {
  aba.getRange(CONFIG.NOME_PRODUTO).clearContent();
  aba.getRange(CONFIG.TIPO_ANUNCIO).clearContent();
  aba.getRange(CONFIG.PRECO_VENDA).clearContent();
  aba.getRange(CONFIG.PRECO_CUSTO).clearContent();
  aba.getRange(CONFIG.DESPESAS_FIXAS).clearContent();
  aba.getRange(CONFIG.DESPESAS_OPERACIONAIS).clearContent();
  aba.getRange(CONFIG.IMPOSTO).clearContent();
  aba.getRange(CONFIG.COMISSAO_ML).clearContent();
  // NÃO LIMPAMOS D13 - CONTÉM A FÓRMULA DO IF DO PREÇO DE VENDA
  aba.getRange(CONFIG.PESO_PRODUTO).clearContent();
  aba.getRange(CONFIG.REPUTACAO).clearContent();
  aba.getRange(CONFIG.COOPARTICIPACAO).clearContent();
  aba.getRange(CONFIG.CUSTO_FRETE).clearContent();
  aba.getRange(CONFIG.CUPOM).clearContent();
}

// ========== ABRIR SIMULADOR DE CUSTOS DO MERCADO LIVRE ==========
function abrirSimuladorML() {
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            text-align: center;
            padding: 30px;
            background: linear-gradient(135deg, #FFE600 0%, #FFC400 100%);
          }
          .container {
            background: white;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #333;
            margin-bottom: 10px;
            font-size: 24px;
          }
          p {
            color: #666;
            margin-bottom: 30px;
            line-height: 1.6;
          }
          .btn {
            background: #3483FA;
            color: white;
            padding: 16px 40px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            transition: all 0.3s ease;
            box-shadow: 0 4px 12px rgba(52, 131, 250, 0.4);
          }
          .btn:hover {
            background: #2968C8;
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(52, 131, 250, 0.5);
          }
          .icon {
            font-size: 48px;
            margin-bottom: 20px;
          }
          .info {
            background: #FFF3CD;
            border-left: 4px solid #FFE600;
            padding: 15px;
            margin: 20px 0;
            text-align: left;
            border-radius: 4px;
          }
          .info strong {
            color: #856404;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="icon">🧮</div>
          <h2>Simulador de Custos</h2>
          <p>Use o simulador oficial do Mercado Livre para calcular a comissão exata do seu produto.</p>
          
          <div class="info">
            <strong>💡 Dica:</strong> A comissão varia de acordo com a categoria do produto. Use o simulador para obter o valor correto.
          </div>
          
          <a href="https://www.mercadolivre.com.br/simulador-de-custos" class="btn" target="_blank" onclick="setTimeout(function(){google.script.host.close()}, 500)">
            📊 Abrir Simulador do ML
          </a>
        </div>
      </body>
    </html>
  `)
  .setWidth(600)
  .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, '💰 Simulador de Custos - Mercado Livre');
}


// ========== MENU PERSONALIZADO ==========
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 ML Precificação')
    .addItem('⚙️ Inicializar Fórmulas de Cálculo', 'inicializarFormulas')
    .addItem('💾 Salvar Produto na Lista', 'listarPrecificacao')
    .addItem('🚚 Calcular Frete Manualmente', 'buscarCustoFrete')
    .addSeparator()
    .addItem('🧮 Abrir Simulador de Custos ML', 'abrirSimuladorML')
    .addSeparator()
    .addItem('🧹 Limpar Campos', 'limparTudo')
    .addToUi();
}


function limparTudo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
  limparCampos(aba);
  SpreadsheetApp.getUi().alert('✅ Todos os campos foram limpos!');
}
