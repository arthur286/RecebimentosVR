/* ═══════════════════════════════════════════════════════════════════════════
   CONCILIAÇÃO · MÓDULO DE ANÁLISE DE RECEBIMENTOS — VERSÃO CORRIGIDA
   Processamento local de Excel/CSV, matching, dashboard e exportação
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

/* ── Namespace do módulo ──────────────────────────────────────────────────── */
const Analise = (() => {

  /* ── Estado interno ───────────────────────────────────────────────────── */
  let _vendas       = [];
  let _recebimentos = [];
  let _resultado    = [];
  let _filtroOp     = 'TODAS';
  let _filtroPeriod = { de: '', ate: '' };
  let _debugLogs    = [];

  const TAXA_THRESHOLD = 0.05;
  const DATE_WINDOW    = 5;

  const OP_MAP = {
    ticket: 'Ticket', tk: 'Ticket', 'ticket restaurante': 'Ticket',
    vr: 'VR', 'vale refeição': 'VR', 'vale refeicao': 'VR',
    alelo: 'Alelo',
    pluxee: 'Pluxee', sodexo: 'Pluxee',
    rede: 'Rede', getnet: 'GetNet'
  };

  function addDebugLog(category, message, data = null) {
    const log = {
      timestamp: new Date().toISOString(),
      category,
      message,
      data: data ? JSON.parse(JSON.stringify(data)) : null
    };
    _debugLogs.push(log);
    console.log(`[DEBUG][${category}]`, message, data || '');
  }

  function showDebugLogs() {
    console.group('=== DEBUG: Conciliação Financeira ===');
    _debugLogs.forEach(log => {
      console.log(`[${log.timestamp}] [${log.category}] ${log.message}`);
      if (log.data) console.log('  Dados:', log.data);
    });
    console.groupEnd();
  }

  function normalizeOp(raw) {
    if (!raw) return 'Desconhecida';
    const key = String(raw).toLowerCase().trim();
    return OP_MAP[key] || String(raw).trim();
  }

  function parseMoney(v) {
    if (v === null || v === undefined || v === '') return NaN;
    if (typeof v === 'number') return v;
    
    let s = String(v).replace(/R\$\s*/g, '').trim();
    
    if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,(?=\d{3})/g, '').replace(',', '.');
    }
    
    return parseFloat(s);
  }

  function parseDate(v) {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v) ? null : v;

    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return isNaN(d) ? null : d;
    }

    const s = String(v).trim();

    const dmY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmY) {
      let [, d, m, y] = dmY;
      if (y.length === 2) y = '20' + y;
      return new Date(Date.UTC(+y, +m - 1, +d));
    }

    const Ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (Ymd) {
      const [, y, m, d] = Ymd;
      return new Date(Date.UTC(+y, +m - 1, +d));
    }

    const dt = new Date(s);
    return isNaN(dt) ? null : dt;
  }

  function fmtDate(d) {
    if (!d) return '—';
    const dia = d.getUTCDate();
    const mes = d.getUTCMonth() + 1;
    const ano = d.getUTCFullYear();
    return `${dia.toString().padStart(2, '0')}/${mes.toString().padStart(2, '0')}/${ano}`;
  }

  function dateKey(d) {
    if (!d) return '';
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  function diffDays(a, b) {
    if (!a || !b) return null;
    return Math.abs((a - b) / 86400000);
  }

  function addDays(d, n) {
    const r = new Date(d);
    r.setUTCDate(r.getUTCDate() + n);
    return r;
  }

  function fmtMoney(v) {
    if (isNaN(v) || v === null) return '—';
    return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }

  /* ── Leitura de arquivos via SheetJS ─────────────────────────────────── */

  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb   = XLSX.read(data, { type: 'array', cellDates: false });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
          addDebugLog('FILE', `Arquivo lido: ${file.name}`, { linhas: rows.length });
          resolve(rows);
        } catch (err) {
          addDebugLog('ERROR', `Erro ao ler arquivo: ${file.name}`, err);
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Leitura ESPECÍFICA para planilha ERP (Olist)
   * Lê todas as abas (Table 1, Table 2, Table 3, Table 4)
   * Ignora linhas de cabeçalho e processa os dados corretamente
   */
  function readERPFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array', cellDates: false });
          
          let allRows = [];
          
          // Processar todas as abas do arquivo
          for (let sheetName of wb.SheetNames) {
            addDebugLog('ERP', `Processando aba: ${sheetName}`);
            
            const ws = wb.Sheets[sheetName];
            // Ler como array de arrays para ter controle total
            const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            
            // Pular linhas de cabeçalho (primeiras 2-3 linhas)
            let startRow = 0;
            for (let i = 0; i < Math.min(5, rawData.length); i++) {
              const row = rawData[i];
              if (row && row[0] && String(row[0]).includes('Cliente')) {
                startRow = i + 1;
                break;
              }
            }
            
            // Processar linhas de dados
            for (let i = startRow; i < rawData.length; i++) {
              const row = rawData[i];
              if (!row || row.length < 5) continue;
              
              const cliente = row[0];
              const historico = row[1];
              const vencimento = row[4];
              const emissao = row[5];
              const situacao = row[6];
              const valor = row[8];
              
              // Pular linhas vazias ou de total
              if (!historico || String(historico).includes('Total')) continue;
              if (!valor || valor === 0) continue;
              
              // Extrair número do pedido do histórico
              let numeroPedido = null;
              const pedidoMatch = String(historico).match(/n[º°]\s*(\d+)/i);
              if (pedidoMatch) {
                numeroPedido = pedidoMatch[1];
              }
              
              // Extrair parcela se houver
              let parcela = null;
              const parcelaMatch = String(historico).match(/parcela\s*(\d+)\/(\d+)/i);
              if (parcelaMatch) {
                parcela = `${parcelaMatch[1]}/${parcelaMatch[2]}`;
              }
              
              allRows.push({
                cliente,
                historico,
                numeroPedido,
                parcela,
                vencimento: vencimento,
                emissao: emissao,
                situacao: situacao,
                valor: valor,
                _sheet: sheetName,
                _rowIndex: i
              });
            }
          }
          
          addDebugLog('ERP', `Arquivo ERP processado`, { 
            totalLinhas: allRows.length,
            sheets: wb.SheetNames 
          });
          
          resolve(allRows);
        } catch (err) {
          addDebugLog('ERROR', `Erro ao ler arquivo ERP: ${file.name}`, err);
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  /* ── Normalização das linhas ──────────────────────────────────────────── */

  /**
   * Normaliza as linhas brutas de VENDAS (planilha Alelo)
   */
  function normalizeVendas(rows) {
    addDebugLog('NORMALIZE', 'Iniciando normalização das vendas', { totalLinhas: rows.length });
    
    // Mapeamento das colunas da planilha Alelo
    const vendasNormalizadas = rows
      .map((r, i) => {
        // Extrair dados das colunas
        const numeroAutorizacao = r['Número da Autorização'];
        const dataVendaStr = r['Data da Venda'];
        const tipoCartao = r['Tipo Cartão'];
        const valorBruto = parseMoney(r['Valor Bruto']);
        const valorLiquido = parseMoney(r['Valor Líquido']);
        const status = r['Status'];
        const dataPagamentoStr = r['Data de Pagamento'];
        const rede = r['PSR Física ou e-com'] || (r['I'] || 'Rede');
        
        // Determinar operadora baseado no tipo de cartão e rede
        let operadora = 'Desconhecida';
        if (tipoCartao) {
          if (tipoCartao.includes('Refeição')) operadora = 'VR';
          if (tipoCartao.includes('Alimentação')) operadora = 'Alelo';
        }
        if (rede && rede !== 'Rede') operadora = normalizeOp(rede);
        
        const dataVenda = parseDate(dataVendaStr);
        const dataPagamentoReal = parseDate(dataPagamentoStr);
        
        // Calcular taxa
        const taxa = !isNaN(valorBruto) && !isNaN(valorLiquido) 
          ? valorBruto - valorLiquido 
          : null;
        
        // Calcular prazo médio (30 dias padrão para VR)
        const prazo = 30;
        const dataPrevista = dataVenda ? addDays(dataVenda, prazo) : null;
        
        return {
          _idx: i,
          id: numeroAutorizacao || `V${i+1}`,
          operadora: operadora,
          dataVenda: dataVenda,
          bruto: valorBruto,
          liquido: valorLiquido,
          taxa: taxa,
          prazo: prazo,
          dataPrevista: dataPrevista,
          dataPagamentoReal: dataPagamentoReal,
          tipoCartao: tipoCartao,
          status: status,
          _raw: r,
        };
      })
      .filter(v => v.dataVenda && !isNaN(v.bruto) && v.bruto > 0);
    
    addDebugLog('NORMALIZE', 'Vendas normalizadas', { 
      total: vendasNormalizadas.length,
      amostra: vendasNormalizadas.slice(0, 3).map(v => ({
        id: v.id,
        bruto: v.bruto,
        liquido: v.liquido,
        operadora: v.operadora
      }))
    });
    
    return vendasNormalizadas;
  }

  /**
   * Normaliza as linhas do ERP (Olist)
   */
  function normalizeERP(rows) {
    addDebugLog('NORMALIZE', 'Iniciando normalização do ERP', { totalLinhas: rows.length });
    
    const erpNormalizado = rows
      .map((r, i) => {
        const dataVencimento = parseDate(r.vencimento);
        const dataEmissao = parseDate(r.emissao);
        const valor = parseMoney(r.valor);
        
        // Usar o número do pedido como ID para matching
        const id = r.numeroPedido || `${r.cliente}_${i}`;
        
        addDebugLog('ERP_ITEM', `Processando item ${i}`, {
          id: id,
          valor: valor,
          vencimento: r.vencimento,
          historico: r.historico?.substring(0, 50)
        });
        
        return {
          _idx: i,
          id: id,
          numeroPedido: r.numeroPedido,
          cliente: r.cliente,
          historico: r.historico,
          parcela: r.parcela,
          data_vencimento: dataVencimento,
          data_emissao: dataEmissao,
          situacao: r.situacao,
          valor_bruto: null,
          valor_liquido: valor,
          taxa: null,
          _raw: r,
          _used: false
        };
      })
      .filter(r => r.data_vencimento && !isNaN(r.valor_liquido) && r.valor_liquido > 0);
    
    addDebugLog('NORMALIZE', 'ERP normalizado', { 
      total: erpNormalizado.length,
      amostra: erpNormalizado.slice(0, 3).map(r => ({
        id: r.id,
        valor: r.valor_liquido,
        vencimento: fmtDate(r.data_vencimento)
      }))
    });
    
    return erpNormalizado;
  }

  /* ── Algoritmo de matching (CONCILIAÇÃO) ────────────────────────────── */

  function matchAll(vendas, recebimentos) {
    addDebugLog('MATCH', 'Iniciando conciliação', { 
      totalVendas: vendas.length, 
      totalRecebimentos: recebimentos.length 
    });
    
    const recList = recebimentos.map(r => ({ ...r, _used: false }));

    const resultado = vendas.map(venda => {
      addDebugLog('MATCH', `Conciliando venda: ${venda.id}`, {
        valor: venda.liquido,
        data: fmtDate(venda.dataVenda),
        operadora: venda.operadora
      });
      
      // Tentar match por número do pedido (se disponível)
      let matchPorId = null;
      if (venda.id) {
        matchPorId = recList.find(r => 
          !r._used && 
          (String(r.id) === String(venda.id) || 
           String(r.numeroPedido) === String(venda.id))
        );
      }
      
      if (matchPorId) {
        matchPorId._used = true;
        const taxa = venda.bruto - matchPorId.valor_liquido;
        const taxaPercentual = (taxa / venda.bruto) * 100;
        const atraso = venda.dataPrevista && matchPorId.data_vencimento 
          ? matchPorId.data_vencimento > venda.dataPrevista 
          : false;
        
        addDebugLog('MATCH_SUCCESS', `Venda ${venda.id} conciliada por ID`, {
          valor_esperado: venda.liquido,
          valor_recebido: matchPorId.valor_liquido,
          diferenca: matchPorId.valor_liquido - venda.liquido
        });
        
        return {
          id: venda.id,
          valor_bruto: venda.bruto,
          valor_liquido: venda.liquido,
          taxa: taxa,
          taxa_percentual: taxaPercentual,
          data_venda: venda.dataVenda,
          data_vencimento: venda.dataPrevista,
          data_pagamento: matchPorId.data_vencimento,
          status: Math.abs(taxa) <= (venda.bruto * TAXA_THRESHOLD) ? 'OK' : 'DIVERGENTE',
          atraso: atraso,
          operadora: venda.operadora,
          recebimento_id: matchPorId.id
        };
      }
      
      // Match por data de pagamento real da planilha de vendas
      if (venda.dataPagamentoReal) {
        const matchData = recList.find(r => 
          !r._used &&
          dateKey(r.data_vencimento) === dateKey(venda.dataPagamentoReal)
        );
        
        if (matchData) {
          matchData._used = true;
          const taxa = venda.bruto - matchData.valor_liquido;
          const taxaPercentual = (taxa / venda.bruto) * 100;
          const status = Math.abs(matchData.valor_liquido - venda.liquido) <= (venda.liquido * TAXA_THRESHOLD) 
            ? 'OK' 
            : 'DIVERGENTE';
          
          addDebugLog('MATCH_SUCCESS', `Venda ${venda.id} conciliada por data`, {
            data_pagamento_real: fmtDate(venda.dataPagamentoReal),
            data_vencimento_erp: fmtDate(matchData.data_vencimento)
          });
          
          return {
            id: venda.id,
            valor_bruto: venda.bruto,
            valor_liquido: venda.liquido,
            taxa: taxa,
            taxa_percentual: taxaPercentual,
            data_venda: venda.dataVenda,
            data_vencimento: venda.dataPrevista,
            data_pagamento: matchData.data_vencimento,
            status: status,
            atraso: false,
            operadora: venda.operadora,
            recebimento_id: matchData.id
          };
        }
      }
      
      // Match por valor aproximado + data próxima
      const candidatos = recList.filter(r => !r._used);
      
      const matchValor = candidatos.find(r => {
        const diffValor = Math.abs(r.valor_liquido - venda.liquido) / (venda.liquido || 1);
        const dentroData = venda.dataPrevista && r.data_vencimento
          ? diffDays(r.data_vencimento, venda.dataPrevista) <= DATE_WINDOW
          : true;
        return diffValor <= TAXA_THRESHOLD && dentroData;
      });
      
      if (matchValor) {
        matchValor._used = true;
        const taxa = venda.bruto - matchValor.valor_liquido;
        const taxaPercentual = (taxa / venda.bruto) * 100;
        const atraso = venda.dataPrevista && matchValor.data_vencimento 
          ? matchValor.data_vencimento > venda.dataPrevista 
          : false;
        
        addDebugLog('MATCH_SUCCESS', `Venda ${venda.id} conciliada por valor`, {
          valor_esperado: venda.liquido,
          valor_recebido: matchValor.valor_liquido,
          diferenca_percentual: ((matchValor.valor_liquido - venda.liquido) / venda.liquido * 100).toFixed(2) + '%'
        });
        
        return {
          id: venda.id,
          valor_bruto: venda.bruto,
          valor_liquido: venda.liquido,
          taxa: taxa,
          taxa_percentual: taxaPercentual,
          data_venda: venda.dataVenda,
          data_vencimento: venda.dataPrevista,
          data_pagamento: matchValor.data_vencimento,
          status: 'OK',
          atraso: atraso,
          operadora: venda.operadora,
          recebimento_id: matchValor.id
        };
      }
      
      // Match apenas por data (valor divergente)
      const matchDataOnly = candidatos.find(r => {
        const dentroData = venda.dataPrevista && r.data_vencimento
          ? diffDays(r.data_vencimento, venda.dataPrevista) <= DATE_WINDOW * 2
          : true;
        return dentroData;
      });
      
      if (matchDataOnly) {
        matchDataOnly._used = true;
        const taxa = venda.bruto - matchDataOnly.valor_liquido;
        const taxaPercentual = (taxa / venda.bruto) * 100;
        const atraso = venda.dataPrevista && matchDataOnly.data_vencimento 
          ? matchDataOnly.data_vencimento > venda.dataPrevista 
          : false;
        
        addDebugLog('MATCH_WARNING', `Venda ${venda.id} conciliada com divergência de valor`, {
          valor_esperado: venda.liquido,
          valor_recebido: matchDataOnly.valor_liquido,
          diferenca: matchDataOnly.valor_liquido - venda.liquido
        });
        
        return {
          id: venda.id,
          valor_bruto: venda.bruto,
          valor_liquido: venda.liquido,
          taxa: taxa,
          taxa_percentual: taxaPercentual,
          data_venda: venda.dataVenda,
          data_vencimento: venda.dataPrevista,
          data_pagamento: matchDataOnly.data_vencimento,
          status: 'DIVERGENTE',
          atraso: atraso,
          operadora: venda.operadora,
          recebimento_id: matchDataOnly.id
        };
      }
      
      // NENHUM MATCH ENCONTRADO
      addDebugLog('MATCH_FAIL', `Venda ${venda.id} sem correspondência no ERP`, {
        valor: venda.liquido,
        data: fmtDate(venda.dataVenda)
      });
      
      return {
        id: venda.id,
        valor_bruto: venda.bruto,
        valor_liquido: venda.liquido,
        taxa: null,
        taxa_percentual: null,
        data_venda: venda.dataVenda,
        data_vencimento: venda.dataPrevista,
        data_pagamento: venda.dataPagamentoReal || null,
        status: 'NAO_ENCONTRADO',
        atraso: null,
        operadora: venda.operadora,
        recebimento_id: null
      };
    });
    
    const naoUtilizados = recList.filter(r => !r._used);
    if (naoUtilizados.length > 0) {
      addDebugLog('MATCH_WARNING', `Recebimentos sem venda correspondente`, {
        total: naoUtilizados.length,
        valores: naoUtilizados.map(r => ({ id: r.id, valor: r.valor_liquido }))
      });
    }
    
    addDebugLog('MATCH', 'Conciliação finalizada', {
      total_ok: resultado.filter(r => r.status === 'OK').length,
      total_divergente: resultado.filter(r => r.status === 'DIVERGENTE').length,
      total_nao_encontrado: resultado.filter(r => r.status === 'NAO_ENCONTRADO').length,
      total_atrasado: resultado.filter(r => r.atraso === true).length
    });
    
    return resultado;
  }

  /* ── Agrupamento ──────────────────────────────────────────────────────── */

  function agrupar(resultado) {
    const map = {};
    resultado.forEach(v => {
      const dk = v.data_venda ? dateKey(v.data_venda) : 'sem-data';
      const key = `${dk}__${v.operadora}`;
      if (!map[key]) {
        map[key] = {
          dateKey: dk,
          dataVenda: v.data_venda,
          operadora: v.operadora,
          vendas: [],
          ok: 0,
          divergente: 0,
          nao_encontrado: 0,
          atrasado: 0,
          totalBruto: 0,
          totalLiquido: 0,
          totalRecebido: 0,
          totalTaxas: 0
        };
      }
      const g = map[key];
      g.vendas.push(v);
      if (v.status === 'OK') g.ok++;
      if (v.status === 'DIVERGENTE') g.divergente++;
      if (v.status === 'NAO_ENCONTRADO') g.nao_encontrado++;
      if (v.atraso === true) g.atrasado++;
      g.totalBruto += isNaN(v.valor_bruto) ? 0 : v.valor_bruto;
      g.totalLiquido += isNaN(v.valor_liquido) ? 0 : v.valor_liquido;
      if (v.valor_liquido && v.status !== 'NAO_ENCONTRADO') g.totalRecebido += v.valor_liquido;
      if (v.taxa) g.totalTaxas += v.taxa;
    });

    return Object.values(map).sort((a, b) => {
      if (!a.dataVenda) return 1;
      if (!b.dataVenda) return -1;
      return b.dataVenda - a.dataVenda;
    });
  }

  function applyFilters(grupos) {
    return grupos.filter(g => {
      if (_filtroOp !== 'TODAS' && g.operadora !== _filtroOp) return false;
      if (_filtroPeriod.de && g.dataVenda) {
        const de = parseDate(_filtroPeriod.de);
        if (de && g.dataVenda < de) return false;
      }
      if (_filtroPeriod.ate && g.dataVenda) {
        const ate = parseDate(_filtroPeriod.ate);
        if (ate && g.dataVenda > ate) return false;
      }
      return true;
    });
  }

  function calcIndicadores(resultado) {
    const total = resultado.length;
    const ok = resultado.filter(v => v.status === 'OK').length;
    const divergente = resultado.filter(v => v.status === 'DIVERGENTE').length;
    const naoRec = resultado.filter(v => v.status === 'NAO_ENCONTRADO').length;
    const atrasado = resultado.filter(v => v.atraso === true).length;
    const totalBruto = resultado.reduce((s, v) => s + (isNaN(v.valor_bruto) ? 0 : v.valor_bruto), 0);
    const totalRecebido = resultado.reduce((s, v) => s + (v.valor_liquido || 0), 0);
    const totalTaxas = resultado.reduce((s, v) => s + (v.taxa || 0), 0);
    const taxaMedia = totalBruto > 0 ? (totalTaxas / totalBruto) * 100 : 0;
    
    return { total, ok, divergente, naoRec, atrasado, totalBruto, totalRecebido, totalTaxas, taxaMedia };
  }

  function getOperadoras(resultado) {
    return [...new Set(resultado.map(v => v.operadora).filter(op => op && op !== 'Desconhecida'))].sort();
  }

  function exportarResultado(resultado) {
    const rows = resultado.map(v => ({
      'ID Pedido': v.id || '',
      'Operadora': v.operadora,
      'Data Venda': fmtDate(v.data_venda),
      'Data Prevista': fmtDate(v.data_vencimento),
      'Data Pagamento': fmtDate(v.data_pagamento),
      'Atraso (dias)': v.atraso && v.data_pagamento && v.data_vencimento 
        ? Math.ceil((v.data_pagamento - v.data_vencimento) / 86400000) 
        : '',
      'Valor Bruto (R$)': v.valor_bruto?.toFixed(2) || '',
      'Valor Líquido (R$)': v.valor_liquido?.toFixed(2) || '',
      'Taxa (R$)': v.taxa?.toFixed(2) || '',
      'Taxa (%)': v.taxa_percentual?.toFixed(2) || '',
      'Status': v.status === 'OK' ? 'OK' : v.status === 'DIVERGENTE' ? 'Divergente' : 'Não Encontrado',
      'Status Atraso': v.atraso === true ? 'Atrasado' : v.atraso === false ? 'No prazo' : '—',
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliação Financeira');
    XLSX.writeFile(wb, `conciliacao_financeira_${new Date().toISOString().slice(0,10)}.xlsx`);
    
    addDebugLog('EXPORT', 'Arquivo exportado', { totalLinhas: rows.length });
  }

  /* ══════════════════════════════════════════════════════════════════════
     RENDERIZAÇÃO DO DASHBOARD
  ══════════════════════════════════════════════════════════════════════ */

  function renderDashboard() {
    const section = document.getElementById('analiseDashboard');
    if (!section) return;

    const grupos = applyFilters(agrupar(_resultado));
    const ind = calcIndicadores(_resultado);

    document.getElementById('aInd_total').textContent = ind.total;
    document.getElementById('aInd_ok').textContent = ind.ok;
    document.getElementById('aInd_divergente').textContent = ind.divergente;
    document.getElementById('aInd_naoRec').textContent = ind.naoRec;
    document.getElementById('aInd_atrasado').textContent = ind.atrasado;
    document.getElementById('aInd_bruto').textContent = fmtMoney(ind.totalBruto);
    document.getElementById('aInd_recebido').textContent = fmtMoney(ind.totalRecebido);
    document.getElementById('aInd_taxas').textContent = fmtMoney(ind.totalTaxas);
    document.getElementById('aInd_taxaMedia').textContent = `${ind.taxaMedia.toFixed(2)}%`;

    const container = document.getElementById('analiseGrupos');
    container.innerHTML = '';

    if (!grupos.length) {
      container.innerHTML = '<p class="history-empty" style="display:block">Nenhum resultado para os filtros aplicados.</p>';
      return;
    }

    grupos.forEach(g => {
      container.appendChild(buildGrupoCard(g));
    });
  }

  function buildGrupoCard(g) {
    const total = g.vendas.length;
    const pctOk = total ? Math.round((g.ok / total) * 100) : 0;
    const pctDiv = total ? Math.round((g.divergente / total) * 100) : 0;
    const pctNRec = total ? Math.round((g.nao_encontrado / total) * 100) : 0;
    const pctAtraso = total ? Math.round((g.atrasado / total) * 100) : 0;

    const opClsMap = { Ticket: 'op-ticket', VR: 'op-vr', Alelo: 'op-alelo', Pluxee: 'op-pluxee', Rede: 'op-rede', GetNet: 'op-getnet' };
    const opCls = opClsMap[g.operadora] || '';

    const card = document.createElement('div');
    card.className = `analise-card ${opCls}`;

    card.innerHTML = `
      <div class="analise-card-header">
        <div class="analise-card-title">
          <span class="analise-date">📅 ${g.dataVenda ? fmtDate(g.dataVenda) : 'Data não informada'}</span>
          <span class="analise-op-badge ${opCls}">${g.operadora}</span>
        </div>
        <div class="analise-card-meta">
          <span class="analise-meta-item">${total} venda${total !== 1 ? 's' : ''}</span>
          <span class="analise-meta-item" style="color:var(--text-dim)">Bruto: ${fmtMoney(g.totalBruto)}</span>
          <span class="analise-meta-item" style="color:var(--text-dim)">Recebido: ${fmtMoney(g.totalRecebido)}</span>
          <span class="analise-meta-item" style="color:var(--text-dim)">Taxas: ${fmtMoney(g.totalTaxas)}</span>
        </div>
      </div>

      <div class="analise-bar">
        <div class="analise-bar-ok" style="width:${pctOk}%" title="OK: ${g.ok}"></div>
        <div class="analise-bar-div" style="width:${pctDiv}%" title="Divergente: ${g.divergente}"></div>
        <div class="analise-bar-nrec" style="width:${pctNRec}%" title="Não encontrado: ${g.nao_encontrado}"></div>
      </div>
      
      <div class="analise-bar" style="margin-top: 4px;">
        <div class="analise-bar-atraso" style="width:${pctAtraso}%" title="Atrasos: ${g.atrasado}"></div>
      </div>

      <div class="analise-pills">
        <span class="analise-pill pill-ok">✔ ${g.ok} OK (${pctOk}%)</span>
        <span class="analise-pill pill-warn">⚠ ${g.divergente} Divergente (${pctDiv}%)</span>
        <span class="analise-pill pill-danger">✕ ${g.nao_encontrado} Não recebido (${pctNRec}%)</span>
        <span class="analise-pill pill-atraso">⏰ ${g.atrasado} Atrasado (${pctAtraso}%)</span>
      </div>

      <button class="analise-toggle-btn" data-open="0">
        <span class="analise-toggle-icon">▾</span> Ver detalhes das vendas
      </button>
      <div class="analise-detail" style="display:none"></div>
    `;

    const btn = card.querySelector('.analise-toggle-btn');
    const detail = card.querySelector('.analise-detail');
    btn.addEventListener('click', () => {
      const isOpen = btn.dataset.open === '1';
      btn.dataset.open = isOpen ? '0' : '1';
      btn.querySelector('.analise-toggle-icon').textContent = isOpen ? '▾' : '▴';
      if (!isOpen) {
        detail.style.display = 'block';
        detail.innerHTML = buildDetailTable(g.vendas);
      } else {
        detail.style.display = 'none';
      }
    });

    return card;
  }

  function buildDetailTable(vendas) {
    const rows = vendas
      .sort((a, b) => {
        const ordem = { 'OK': 1, 'DIVERGENTE': 2, 'NAO_ENCONTRADO': 3 };
        return (ordem[a.status] || 4) - (ordem[b.status] || 4);
      })
      .map(v => {
        const stCls = v.status === 'OK' ? 'status-ok' : v.status === 'DIVERGENTE' ? 'status-warn' : 'status-danger';
        const stLbl = v.status === 'OK' ? '✔ OK' : v.status === 'DIVERGENTE' ? '⚠ Divergente' : '✕ Não recebido';
        const atrasoDias = v.atraso && v.data_pagamento && v.data_vencimento 
          ? Math.ceil((v.data_pagamento - v.data_vencimento) / 86400000)
          : null;
        
        return `
          <tr class="${v.atraso ? 'row-atrasado' : ''}">
            <td class="dt-id">${v.id || '—'}</td>
            <td>${v.operadora || '—'}</td>
            <td>${fmtMoney(v.valor_bruto)}</td>
            <td>${fmtMoney(v.valor_liquido)}</td>
            <td>${fmtMoney(v.taxa)}</td>
            <td>${v.taxa_percentual ? v.taxa_percentual.toFixed(2) + '%' : '—'}</td>
            <td>${fmtDate(v.data_vencimento)}</td>
            <td class="${stCls}">${stLbl}</td>
            <td>${fmtDate(v.data_pagamento)}</td>
            <td class="${v.atraso ? 'diff-neg' : 'diff-pos'}">${v.atraso === true ? `${atrasoDias} dias` : v.atraso === false ? 'No prazo' : '—'}</td>
          </tr>`;
      }).join('');

    return `
      <div class="analise-table-wrap">
        <table class="analise-table">
          <thead>
            <tr>
              <th>ID/NSU</th>
              <th>Operadora</th>
              <th>Valor Bruto</th>
              <th>Valor Líquido</th>
              <th>Taxa (R$)</th>
              <th>Taxa (%)</th>
              <th>Data Prevista</th>
              <th>Status</th>
              <th>Data Pagamento</th>
              <th>Atraso</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  function populateFiltroOp() {
    const sel = document.getElementById('filtroOperadora');
    if (!sel) return;
    const ops = getOperadoras(_resultado);
    sel.innerHTML = '<option value="TODAS">Todas as operadoras</option>' +
      ops.map(op => `<option value="${op}">${op}</option>`).join('');
  }

  function init() {
    const btnProcessar = document.getElementById('btnProcessarAnalise');
    if (btnProcessar) {
      btnProcessar.addEventListener('click', async () => {
        const fileVendas = document.getElementById('fileVendas').files[0];
        const fileRec = document.getElementById('fileRecebimentos').files[0];

        if (!fileVendas || !fileRec) {
          showAnaliseFeedback('⚠ Selecione os dois arquivos antes de processar.', 'warn');
          return;
        }

        setBtnLoading(btnProcessar, true);
        showAnaliseFeedback('Lendo arquivos...', 'info');
        _debugLogs = [];

        try {
          const [rawVendas, rawRec] = await Promise.all([
            readFile(fileVendas),
            readERPFile(fileRec),
          ]);

          _vendas = normalizeVendas(rawVendas);
          _recebimentos = normalizeERP(rawRec);
          _resultado = matchAll(_vendas, _recebimentos);

          showDebugLogs();

          showAnaliseFeedback(`✔ ${_vendas.length} vendas × ${_recebimentos.length} recebimentos processados.`, 'ok');
          populateFiltroOp();
          renderDashboard();

          document.getElementById('analiseDashboard').style.display = 'block';
          document.getElementById('analiseEmpty').style.display = 'none';

        } catch (err) {
          console.error(err);
          showAnaliseFeedback('❗ Erro ao ler arquivo. Verifique o formato.', 'danger');
          addDebugLog('ERROR', 'Erro no processamento', err);
        } finally {
          setBtnLoading(btnProcessar, false);
        }
      });
    }

    const selOp = document.getElementById('filtroOperadora');
    if (selOp) {
      selOp.addEventListener('change', () => {
        _filtroOp = selOp.value;
        renderDashboard();
      });
    }

    document.getElementById('filtroDe')?.addEventListener('change', (e) => {
      _filtroPeriod.de = e.target.value;
      renderDashboard();
    });

    document.getElementById('filtroAte')?.addEventListener('change', (e) => {
      _filtroPeriod.ate = e.target.value;
      renderDashboard();
    });

    document.getElementById('btnLimparFiltros')?.addEventListener('click', () => {
      _filtroOp = 'TODAS';
      _filtroPeriod = { de: '', ate: '' };
      if (selOp) selOp.value = 'TODAS';
      document.getElementById('filtroDe').value = '';
      document.getElementById('filtroAte').value = '';
      renderDashboard();
    });

    document.getElementById('btnExportar')?.addEventListener('click', () => {
      if (!_resultado.length) {
        showAnaliseFeedback('⚠ Nenhum dado para exportar.', 'warn');
        return;
      }
      exportarResultado(_resultado);
    });

    setupDropZone('dropVendas', 'fileVendas');
    setupDropZone('dropRecebimentos', 'fileRecebimentos');
  }

  function setupDropZone(zoneId, inputId) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    if (!zone || !input) return;

    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.classList.add('dz-over');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('dz-over'));
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('dz-over');
      const file = e.dataTransfer.files[0];
      if (file) {
        const dt = new DataTransfer();
        dt.items.add(file);
        input.files = dt.files;
        updateDropZoneLabel(zone, file.name);
      }
    });
    input.addEventListener('change', () => {
      if (input.files[0]) updateDropZoneLabel(zone, input.files[0].name);
    });
  }

  function updateDropZoneLabel(zone, name) {
    const lbl = zone.querySelector('.dz-label');
    if (lbl) {
      lbl.textContent = `📄 ${name}`;
      zone.classList.add('dz-loaded');
    }
  }

  function setBtnLoading(btn, loading) {
    btn.disabled = loading;
    btn.textContent = loading ? '⏳ Processando...' : '▶ Processar Análise';
  }

  function showAnaliseFeedback(msg, type) {
    const el = document.getElementById('analiseFeedback');
    if (!el) return;
    el.textContent = msg;
    el.className = `analise-feedback analise-feedback-${type}`;
  }

  return { init };

})();

document.addEventListener('DOMContentLoaded', () => Analise.init());
