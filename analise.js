/* ═══════════════════════════════════════════════════════════════════════════
   CONCILIAÇÃO · MÓDULO DE ANÁLISE DE RECEBIMENTOS — analise.js
   Processamento local de Excel/CSV, matching, dashboard e exportação
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

/* ── Namespace do módulo ──────────────────────────────────────────────────── */
const Analise = (() => {

  /* ── Estado interno ───────────────────────────────────────────────────── */
  let _vendas       = [];   // linhas normalizadas do arquivo de vendas
  let _recebimentos = [];   // linhas normalizadas do arquivo de recebimentos
  let _resultado    = [];   // vendas com status resolvido
  let _filtroOp     = 'TODAS';
  let _filtroPeriod = { de: '', ate: '' };

  /* Tolerância de valor: diferença ≤ threshold = "OK", senão "erro_taxa" */
  const TAXA_THRESHOLD = 0.05; // 5%
  /* Janela de busca de datas em dias */
  const DATE_WINDOW    = 5;

  /* Operadoras reconhecidas → normalização */
  const OP_MAP = {
    ticket: 'Ticket', tk: 'Ticket', 'ticket restaurante': 'Ticket',
    vr: 'VR', 'vale refeição': 'VR', 'vale refeicao': 'VR',
    alelo: 'Alelo',
    pluxee: 'Pluxee', sodexo: 'Pluxee',
  };

  /* ── Utilitários ──────────────────────────────────────────────────────── */

  function normalizeOp(raw) {
    if (!raw) return 'Desconhecida';
    const key = String(raw).toLowerCase().trim();
    return OP_MAP[key] || String(raw).trim();
  }

  /**
   * Converte um valor para número, aceitando
   * "R$ 1.234,56", "1234.56", "1,234.56" etc.
   */
  function parseMoney(v) {
    if (v === null || v === undefined || v === '') return NaN;
    if (typeof v === 'number') return v;
    let s = String(v).replace(/R\$\s*/g, '').trim();
    // detectar formato pt-BR: 1.234,56
    if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      // remover separadores de milhar en-US
      s = s.replace(/,(?=\d{3})/g, '').replace(',', '.');
    }
    return parseFloat(s);
  }

  /**
   * Converte qualquer representação de data para objeto Date.
   * Aceita: número serial Excel, "dd/mm/yyyy", "yyyy-mm-dd", objetos Date.
   */
  function parseDate(v) {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v) ? null : v;

    // número serial Excel (dias desde 1899-12-30)
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return isNaN(d) ? null : d;
    }

    const s = String(v).trim();

    // dd/mm/yyyy ou dd/mm/yy
    const dmY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmY) {
      let [, d, m, y] = dmY;
      if (y.length === 2) y = '20' + y;
      return new Date(+y, +m - 1, +d);
    }

    // yyyy-mm-dd
    const Ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (Ymd) {
      const [, y, m, d] = Ymd;
      return new Date(+y, +m - 1, +d);
    }

    const dt = new Date(s);
    return isNaN(dt) ? null : dt;
  }

  /** Formata Date → "dd/mm/yyyy" */
  function fmtDate(d) {
    if (!d) return '—';
    return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  /** Retorna YYYY-MM-DD para comparação */
  function dateKey(d) {
    if (!d) return '';
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  /** Diferença absoluta em dias entre dois Date */
  function diffDays(a, b) {
    return Math.abs((a - b) / 86400000);
  }

  /** Adiciona N dias a um Date */
  function addDays(d, n) {
    const r = new Date(d);
    r.setDate(r.getDate() + n);
    return r;
  }

  /** Formata moeda pt-BR */
  function fmtMoney(v) {
    if (isNaN(v) || v === null) return '—';
    return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }

  /* ── Leitura de arquivos via SheetJS ─────────────────────────────────── */

  /**
   * Lê um File (.xlsx ou .csv) e retorna array de objetos (primeira linha = cabeçalho).
   * @returns {Promise<object[]>}
   */
  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb   = XLSX.read(data, { type: 'array', cellDates: false });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  /* ── Normalização das linhas ──────────────────────────────────────────── */

  /**
   * Detecta automaticamente qual coluna corresponde a cada campo,
   * aceitando variações de nome (pt-BR, inglês, abreviações).
   */
  function detectColumns(rows, type) {
    if (!rows.length) return {};
    const keys = Object.keys(rows[0]);
    const find  = (candidates) =>
      keys.find(k => candidates.some(c => k.toLowerCase().includes(c))) || null;

    if (type === 'vendas') {
      return {
        data:      find(['data', 'emissão', 'emissao', 'date', 'venda']),
        operadora: find(['operadora', 'operator', 'bandeira', 'rede']),
        bruto:     find(['bruto', 'gross', 'valor bruto', 'total']),
        taxa:      find(['taxa', 'fee', 'desconto', 'rate']),
        liquido:   find(['líquido', 'liquido', 'net', 'valor líquido', 'valor liquido', 'líq']),
        prazo:     find(['prazo', 'd+', 'days', 'modalidade', 'tipo']),
        id:        find(['id', 'nsu', 'cod', 'código', 'codigo', 'ref']),
      };
    } else {
      return {
        data:      find(['data', 'pagamento', 'payment', 'recebimento', 'date']),
        operadora: find(['operadora', 'operator', 'bandeira', 'rede']),
        valor:     find(['valor', 'recebido', 'amount', 'pago', 'liquido', 'líquido', 'net']),
        id:        find(['id', 'nsu', 'cod', 'ref']),
      };
    }
  }

  /** Normaliza as linhas brutas de VENDAS */
  function normalizeVendas(rows) {
    const cols = detectColumns(rows, 'vendas');
    return rows
      .map((r, i) => {
        const dataVenda = parseDate(cols.data ? r[cols.data] : null);
        const bruto     = parseMoney(cols.bruto   ? r[cols.bruto]   : null);
        const taxa      = parseMoney(cols.taxa     ? r[cols.taxa]    : null);
        let   liquido   = parseMoney(cols.liquido  ? r[cols.liquido] : null);

        // calcular líquido se não informado
        if (isNaN(liquido) && !isNaN(bruto) && !isNaN(taxa)) {
          // taxa pode ser % (ex: 2.5) ou valor absoluto (ex: 25.00)
          liquido = taxa < 1
            ? bruto * (1 - taxa)
            : bruto - taxa;
        }

        // prazo D+N
        const prazoRaw = cols.prazo ? String(r[cols.prazo] || '') : '';
        const prazoNum = parseInt(prazoRaw.replace(/\D/g, '')) || 30;

        return {
          _idx:        i,
          id:          cols.id ? r[cols.id] : `V${i+1}`,
          operadora:   normalizeOp(cols.operadora ? r[cols.operadora] : null),
          dataVenda,
          bruto,
          taxa,
          liquido,
          prazo:       prazoNum,
          dataPrevista: dataVenda ? addDays(dataVenda, prazoNum) : null,
          _raw:        r,
        };
      })
      .filter(v => v.dataVenda && !isNaN(v.bruto));
  }

  /** Normaliza as linhas brutas de RECEBIMENTOS */
  function normalizeRecebimentos(rows) {
    const cols = detectColumns(rows, 'recebimentos');
    return rows
      .map((r, i) => {
        const dataPagamento = parseDate(cols.data ? r[cols.data] : null);
        const valor         = parseMoney(cols.valor ? r[cols.valor] : null);
        return {
          _idx:        i,
          id:          cols.id ? r[cols.id] : null,
          operadora:   normalizeOp(cols.operadora ? r[cols.operadora] : null),
          dataPagamento,
          valor,
          _used:       false,   // marca se já foi vinculada a uma venda
          _raw:        r,
        };
      })
      .filter(r => r.dataPagamento && !isNaN(r.valor));
  }

  /* ── Algoritmo de matching ────────────────────────────────────────────── */

  /**
   * Para cada venda, tenta encontrar um recebimento correspondente.
   * Prioriza: mesma operadora + valor dentro do threshold + data dentro da janela.
   * Se não achar por valor exato, marca como "erro_taxa".
   * Se não achar nada, marca como "nao_recebido".
   */
  function matchAll(vendas, recebimentos) {
    // Cópia para não mutar os originais
    const recList = recebimentos.map(r => ({ ...r, _used: false }));

    return vendas.map(venda => {
      // Candidatos: mesma operadora (ou desconhecida nos dois lados)
      const candidatos = recList.filter(r =>
        !r._used &&
        (r.operadora === venda.operadora || r.operadora === 'Desconhecida' || venda.operadora === 'Desconhecida')
      );

      // Tenta match perfeito: valor dentro do threshold + data na janela
      const matchExato = candidatos.find(r => {
        const diffValor = Math.abs(r.valor - venda.liquido) / (venda.liquido || 1);
        const dentroData = venda.dataPrevista
          ? diffDays(r.dataPagamento, venda.dataPrevista) <= DATE_WINDOW
          : true;
        return diffValor <= TAXA_THRESHOLD && dentroData;
      });

      if (matchExato) {
        matchExato._used = true;
        const diffValor = Math.abs(matchExato.valor - venda.liquido);
        return {
          ...venda,
          status:        'ok',
          recebimento:   matchExato,
          dataPagamento: matchExato.dataPagamento,
          valorRecebido: matchExato.valor,
          diferenca:     matchExato.valor - venda.liquido,
          diffDias:      venda.dataPrevista ? diffDays(matchExato.dataPagamento, venda.dataPrevista) : null,
        };
      }

      // Tenta match por operadora + data (valor diferente = erro de taxa)
      const matchTaxa = candidatos.find(r => {
        const dentroData = venda.dataPrevista
          ? diffDays(r.dataPagamento, venda.dataPrevista) <= DATE_WINDOW * 2
          : true;
        return dentroData;
      });

      if (matchTaxa) {
        matchTaxa._used = true;
        return {
          ...venda,
          status:        'erro_taxa',
          recebimento:   matchTaxa,
          dataPagamento: matchTaxa.dataPagamento,
          valorRecebido: matchTaxa.valor,
          diferenca:     matchTaxa.valor - venda.liquido,
          diffDias:      venda.dataPrevista ? diffDays(matchTaxa.dataPagamento, venda.dataPrevista) : null,
        };
      }

      // Não encontrado
      return {
        ...venda,
        status:        'nao_recebido',
        recebimento:   null,
        dataPagamento: null,
        valorRecebido: null,
        diferenca:     null,
        diffDias:      null,
      };
    });
  }

  /* ── Agrupamento ──────────────────────────────────────────────────────── */

  /**
   * Agrupa _resultado por (dateKey(dataVenda), operadora).
   * Retorna array ordenado por data desc.
   */
  function agrupar(resultado) {
    const map = {};
    resultado.forEach(v => {
      const dk  = dateKey(v.dataVenda);
      const key = `${dk}__${v.operadora}`;
      if (!map[key]) {
        map[key] = {
          dateKey:   dk,
          dataVenda: v.dataVenda,
          operadora: v.operadora,
          vendas:    [],
          ok:        0,
          erro_taxa: 0,
          nao_recebido: 0,
          totalBruto:  0,
          totalLiquido: 0,
          totalRecebido: 0,
        };
      }
      const g = map[key];
      g.vendas.push(v);
      g[v.status]++;
      g.totalBruto   += isNaN(v.bruto)   ? 0 : v.bruto;
      g.totalLiquido += isNaN(v.liquido) ? 0 : v.liquido;
      if (v.valorRecebido) g.totalRecebido += v.valorRecebido;
    });

    return Object.values(map).sort((a, b) => b.dataVenda - a.dataVenda);
  }

  /* ── Filtros ──────────────────────────────────────────────────────────── */

  function applyFilters(grupos) {
    return grupos.filter(g => {
      if (_filtroOp !== 'TODAS' && g.operadora !== _filtroOp) return false;
      if (_filtroPeriod.de) {
        const de = parseDate(_filtroPeriod.de);
        if (de && g.dataVenda < de) return false;
      }
      if (_filtroPeriod.ate) {
        const ate = parseDate(_filtroPeriod.ate);
        if (ate && g.dataVenda > ate) return false;
      }
      return true;
    });
  }

  /* ── Indicadores globais ──────────────────────────────────────────────── */

  function calcIndicadores(resultado) {
    const total    = resultado.length;
    const ok       = resultado.filter(v => v.status === 'ok').length;
    const erroTaxa = resultado.filter(v => v.status === 'erro_taxa').length;
    const naoRec   = resultado.filter(v => v.status === 'nao_recebido').length;
    const totalBruto    = resultado.reduce((s, v) => s + (isNaN(v.bruto)   ? 0 : v.bruto), 0);
    const totalRecebido = resultado.reduce((s, v) => s + (v.valorRecebido || 0), 0);
    return { total, ok, erroTaxa, naoRec, totalBruto, totalRecebido };
  }

  /* ── Operadoras únicas do resultado ──────────────────────────────────── */

  function getOperadoras(resultado) {
    return [...new Set(resultado.map(v => v.operadora))].sort();
  }

  /* ── Exportação CSV/Excel ─────────────────────────────────────────────── */

  function exportarResultado(resultado) {
    const rows = resultado.map(v => ({
      ID:           v.id || '',
      Operadora:    v.operadora,
      'Data Venda': fmtDate(v.dataVenda),
      'Prazo (D+)': v.prazo,
      'Data Prevista': fmtDate(v.dataPrevista),
      'Valor Bruto':  isNaN(v.bruto)    ? '' : v.bruto,
      'Valor Líquido Esperado': isNaN(v.liquido) ? '' : v.liquido,
      Status:       v.status === 'ok' ? 'OK' : v.status === 'erro_taxa' ? 'Erro de Taxa' : 'Não Recebido',
      'Data Pagamento': fmtDate(v.dataPagamento),
      'Valor Recebido': v.valorRecebido ?? '',
      'Diferença':      v.diferenca    ?? '',
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Análise');
    XLSX.writeFile(wb, `conciliacao_analise_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  /* ══════════════════════════════════════════════════════════════════════
     RENDERIZAÇÃO DO DASHBOARD
  ══════════════════════════════════════════════════════════════════════ */

  function renderDashboard() {
    const section = document.getElementById('analiseDashboard');
    if (!section) return;

    const grupos = applyFilters(agrupar(_resultado));
    const ind    = calcIndicadores(_resultado);

    /* ── Indicadores ── */
    document.getElementById('aInd_total').textContent    = ind.total;
    document.getElementById('aInd_ok').textContent       = ind.ok;
    document.getElementById('aInd_erro').textContent     = ind.erroTaxa;
    document.getElementById('aInd_naoRec').textContent   = ind.naoRec;
    document.getElementById('aInd_bruto').textContent    = fmtMoney(ind.totalBruto);
    document.getElementById('aInd_recebido').textContent = fmtMoney(ind.totalRecebido);

    /* ── Cards de grupos ── */
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

  /** Card visual de um grupo (data × operadora) */
  function buildGrupoCard(g) {
    const total    = g.vendas.length;
    const pctOk    = total ? Math.round((g.ok / total) * 100) : 0;
    const pctErro  = total ? Math.round((g.erro_taxa / total) * 100) : 0;
    const pctNRec  = total ? Math.round((g.nao_recebido / total) * 100) : 0;

    // Classe de borda por operadora
    const opClsMap = { Ticket: 'op-ticket', VR: 'op-vr', Alelo: 'op-alelo', Pluxee: 'op-pluxee' };
    const opCls    = opClsMap[g.operadora] || '';

    const card = document.createElement('div');
    card.className = `analise-card ${opCls}`;

    /* Cabeçalho */
    card.innerHTML = `
      <div class="analise-card-header">
        <div class="analise-card-title">
          <span class="analise-date">📅 ${fmtDate(g.dataVenda)}</span>
          <span class="analise-op-badge ${opCls}">${g.operadora}</span>
        </div>
        <div class="analise-card-meta">
          <span class="analise-meta-item">${total} venda${total !== 1 ? 's' : ''}</span>
          <span class="analise-meta-item" style="color:var(--text-dim)">Bruto: ${fmtMoney(g.totalBruto)}</span>
          <span class="analise-meta-item" style="color:var(--text-dim)">Recebido: ${fmtMoney(g.totalRecebido)}</span>
        </div>
      </div>

      <!-- Barra de progresso tricolor -->
      <div class="analise-bar">
        <div class="analise-bar-ok"    style="width:${pctOk}%"   title="OK: ${g.ok}"></div>
        <div class="analise-bar-erro"  style="width:${pctErro}%"  title="Erro: ${g.erro_taxa}"></div>
        <div class="analise-bar-nrec"  style="width:${pctNRec}%"  title="Não rec: ${g.nao_recebido}"></div>
      </div>

      <!-- Pills de resumo -->
      <div class="analise-pills">
        <span class="analise-pill pill-ok">    ✔ ${g.ok} OK (${pctOk}%)</span>
        <span class="analise-pill pill-warn">  ⚠ ${g.erro_taxa} Erro de taxa (${pctErro}%)</span>
        <span class="analise-pill pill-danger">✕ ${g.nao_recebido} Não recebido (${pctNRec}%)</span>
      </div>

      <!-- Detalhe expansível -->
      <button class="analise-toggle-btn" data-open="0">
        <span class="analise-toggle-icon">▾</span> Ver detalhes das vendas
      </button>
      <div class="analise-detail" style="display:none"></div>
    `;

    /* Toggle de detalhes */
    const btn     = card.querySelector('.analise-toggle-btn');
    const detail  = card.querySelector('.analise-detail');
    btn.addEventListener('click', () => {
      const isOpen = btn.dataset.open === '1';
      btn.dataset.open = isOpen ? '0' : '1';
      btn.querySelector('.analise-toggle-icon').textContent = isOpen ? '▾' : '▴';
      btn.innerHTML = btn.innerHTML.replace(
        isOpen ? 'Fechar' : 'Ver detalhes das vendas',
        isOpen ? 'Ver detalhes das vendas' : 'Fechar'
      );
      if (!isOpen) {
        detail.style.display = 'block';
        detail.innerHTML     = buildDetailTable(g.vendas);
      } else {
        detail.style.display = 'none';
      }
    });

    return card;
  }

  /** Tabela detalhada de vendas de um grupo */
  function buildDetailTable(vendas) {
    const rows = vendas
      .sort((a, b) => (a.status === 'ok' ? 1 : -1))
      .map(v => {
        const stCls  = v.status === 'ok' ? 'status-ok' : v.status === 'erro_taxa' ? 'status-warn' : 'status-danger';
        const stLbl  = v.status === 'ok' ? '✔ OK' : v.status === 'erro_taxa' ? '⚠ Erro de taxa' : '✕ Não recebido';
        const diff   = v.diferenca !== null ? fmtMoney(Math.abs(v.diferenca)) : '—';
        const diffCls = v.diferenca > 0 ? 'diff-pos' : v.diferenca < 0 ? 'diff-neg' : '';

        return `
          <tr>
            <td class="dt-id">${v.id || '—'}</td>
            <td>${fmtMoney(v.bruto)}</td>
            <td>${fmtMoney(v.liquido)}</td>
            <td>${fmtDate(v.dataPrevista)}</td>
            <td class="${stCls}">${stLbl}</td>
            <td>${fmtDate(v.dataPagamento)}</td>
            <td>${fmtMoney(v.valorRecebido)}</td>
            <td class="${diffCls}">${v.diferenca !== null ? (v.diferenca >= 0 ? '+' : '-') + diff : '—'}</td>
          </tr>`;
      }).join('');

    return `
      <div class="analise-table-wrap">
        <table class="analise-table">
          <thead>
            <tr>
              <th>ID/NSU</th>
              <th>Bruto</th>
              <th>Líquido Esperado</th>
              <th>Prev. Receb.</th>
              <th>Status</th>
              <th>Data Pagto.</th>
              <th>Valor Recebido</th>
              <th>Diferença</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  /* ── Filtros UI ───────────────────────────────────────────────────────── */

  function populateFiltroOp() {
    const sel = document.getElementById('filtroOperadora');
    if (!sel) return;
    const ops = getOperadoras(_resultado);
    sel.innerHTML = '<option value="TODAS">Todas as operadoras</option>' +
      ops.map(op => `<option value="${op}">${op}</option>`).join('');
  }

  /* ── Inicialização da UI ──────────────────────────────────────────────── */

  function init() {
    /* ── Botão de processamento ── */
    const btnProcessar = document.getElementById('btnProcessarAnalise');
    if (btnProcessar) {
      btnProcessar.addEventListener('click', async () => {
        const fileVendas = document.getElementById('fileVendas').files[0];
        const fileRec    = document.getElementById('fileRecebimentos').files[0];

        if (!fileVendas || !fileRec) {
          showAnaliseFeedback('⚠ Selecione os dois arquivos antes de processar.', 'warn');
          return;
        }

        setBtnLoading(btnProcessar, true);
        showAnaliseFeedback('Lendo arquivos...', 'info');

        try {
          const [rawVendas, rawRec] = await Promise.all([
            readFile(fileVendas),
            readFile(fileRec),
          ]);

          _vendas       = normalizeVendas(rawVendas);
          _recebimentos = normalizeRecebimentos(rawRec);
          _resultado    = matchAll(_vendas, _recebimentos);

          showAnaliseFeedback(`✔ ${_vendas.length} vendas × ${_recebimentos.length} recebimentos processados.`, 'ok');
          populateFiltroOp();
          renderDashboard();

          // Mostrar seção do dashboard
          document.getElementById('analiseDashboard').style.display = 'block';
          document.getElementById('analiseEmpty').style.display     = 'none';

        } catch (err) {
          console.error(err);
          showAnaliseFeedback('❗ Erro ao ler arquivo. Verifique o formato.', 'danger');
        } finally {
          setBtnLoading(btnProcessar, false);
        }
      });
    }

    /* ── Filtro operadora ── */
    const selOp = document.getElementById('filtroOperadora');
    if (selOp) {
      selOp.addEventListener('change', () => {
        _filtroOp = selOp.value;
        renderDashboard();
      });
    }

    /* ── Filtros de período ── */
    document.getElementById('filtroDe')?.addEventListener('change', (e) => {
      _filtroPeriod.de = e.target.value;
      renderDashboard();
    });

    document.getElementById('filtroAte')?.addEventListener('change', (e) => {
      _filtroPeriod.ate = e.target.value;
      renderDashboard();
    });

    /* ── Limpar filtros ── */
    document.getElementById('btnLimparFiltros')?.addEventListener('click', () => {
      _filtroOp = 'TODAS';
      _filtroPeriod = { de: '', ate: '' };
      if (selOp) selOp.value = 'TODAS';
      document.getElementById('filtroDe').value  = '';
      document.getElementById('filtroAte').value = '';
      renderDashboard();
    });

    /* ── Exportar ── */
    document.getElementById('btnExportar')?.addEventListener('click', () => {
      if (!_resultado.length) {
        showAnaliseFeedback('⚠ Nenhum dado para exportar.', 'warn');
        return;
      }
      exportarResultado(_resultado);
    });

    /* ── Drag & Drop nos inputs de arquivo ── */
    setupDropZone('dropVendas',       'fileVendas');
    setupDropZone('dropRecebimentos', 'fileRecebimentos');
  }

  function setupDropZone(zoneId, inputId) {
    const zone  = document.getElementById(zoneId);
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
    el.className   = `analise-feedback analise-feedback-${type}`;
  }

  /* ── API pública ──────────────────────────────────────────────────────── */
  return { init };

})();

/* ── Bootstrap ─────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => Analise.init());
