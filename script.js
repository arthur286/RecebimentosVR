/* ═══════════════════════════════════════════════════════════════════════════
   CONCILIAÇÃO · OPERADORAS — script.js
   Toda a lógica do sistema: estado, renderização, localStorage, indicadores
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

/* ── Constantes ─────────────────────────────────────────────────────────── */

const STORAGE_KEY_CURRENT  = 'conciliacao_current';
const STORAGE_KEY_HISTORY  = 'conciliacao_history';

/** Definição das operadoras: id, nome, classe CSS e iniciais do logo */
const OPERATORS = [
  { id: 'ticket', name: 'Ticket',  cls: 'op-ticket', logo: 'TK', histCls: 'hw-op-ticket' },
  { id: 'vr',     name: 'VR',      cls: 'op-vr',     logo: 'VR', histCls: 'hw-op-vr'     },
  { id: 'alelo',  name: 'Alelo',   cls: 'op-alelo',  logo: 'AL', histCls: 'hw-op-alelo'  },
  { id: 'pluxee', name: 'Pluxee',  cls: 'op-pluxee', logo: 'PL', histCls: 'hw-op-pluxee' },
];

/** Configuração dos grupos de botões de status de cada card */
const STATUS_GROUPS = [
  {
    key: 'conciliacao',
    label: 'Conciliação',
    buttons: [
      { value: 'ok',         label: '✔ OK',         cls: 'active-ok'     },
      { value: 'incompleto', label: '⚠ Incompleto',  cls: 'active-warn'   },
      { value: 'nao_feito',  label: '✕ Não feito',   cls: 'active-danger' },
    ],
  },
  {
    key: 'transferencia',
    label: 'Transferência',
    buttons: [
      { value: 'feito',          label: '✔ Feito',           cls: 'active-ok'     },
      { value: 'pendente',       label: '⏳ Pendente',         cls: 'active-warn'   },
      { value: 'nao_se_aplica',  label: '— N/A',              cls: 'active-warn'   },
    ],
  },
  {
    key: 'pendencias',
    label: 'Pendências',
    buttons: [
      { value: 'sem_pendencia', label: '✔ Sem pendência',  cls: 'active-ok'     },
      { value: 'com_pendencia', label: '❗ Com pendência', cls: 'active-danger' },
    ],
  },
];

/* ── Estado global ──────────────────────────────────────────────────────── */

/**
 * Estado atual (semana em aberto).
 * Estrutura:
 * {
 *   period: string,
 *   operators: {
 *     [id]: {
 *       unidade: 'filial' | 'matriz',
 *       period: string,
 *       conciliacao: string,
 *       transferencia: string,
 *       pendencias: string,
 *       obs: string,
 *     }
 *   }
 * }
 */
let state = buildEmptyState();

/** Histórico: array de { savedAt, period, operators } */
let history = [];

/* ── Inicialização ──────────────────────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  loadFromStorage();
  renderCards();
  renderWeekBadge();
  renderHistory();
  updateIndicators();
  bindGlobalEvents();
});

/* ── Helpers de estado ──────────────────────────────────────────────────── */

/** Retorna um objeto de estado vazio para uma nova semana */
function buildEmptyState() {
  const ops = {};
  OPERATORS.forEach(op => {
    ops[op.id] = {
      unidade:      'filial',
      period:       '',
      conciliacao:  '',
      transferencia:'',
      pendencias:   '',
      obs:          '',
    };
  });
  return { period: '', operators: ops };
}

/* ── localStorage ───────────────────────────────────────────────────────── */

function saveToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY_CURRENT, JSON.stringify(state));
    localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(history));
  } catch (e) {
    console.warn('Erro ao salvar no localStorage:', e);
  }
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_CURRENT);
    if (raw) state = JSON.parse(raw);

    const rawH = localStorage.getItem(STORAGE_KEY_HISTORY);
    if (rawH) history = JSON.parse(rawH);
  } catch (e) {
    console.warn('Erro ao carregar localStorage:', e);
  }
}

/* ── Renderização dos Cards ─────────────────────────────────────────────── */

function renderCards() {
  const grid = document.getElementById('cardsGrid');
  grid.innerHTML = '';

  OPERATORS.forEach(op => {
    const opState = state.operators[op.id];
    const card = buildCard(op, opState);
    grid.appendChild(card);
  });
}

/**
 * Constrói o elemento DOM de um card de operadora.
 * @param {object} op        - Definição da operadora (OPERATORS[i])
 * @param {object} opState   - Estado atual desta operadora
 * @returns {HTMLElement}
 */
function buildCard(op, opState) {
  const card = document.createElement('div');
  card.className = `op-card ${op.cls}`;
  card.dataset.opId = op.id;

  /* ── Cabeçalho ── */
  const header = document.createElement('div');
  header.className = 'op-card-header';

  const nameEl = document.createElement('div');
  nameEl.className = 'op-name';
  nameEl.textContent = op.name;

  const logoEl = document.createElement('div');
  logoEl.className = 'op-logo';
  logoEl.textContent = op.logo;

  header.appendChild(nameEl);
  header.appendChild(logoEl);
  card.appendChild(header);

  /* ── Corpo ── */
  const body = document.createElement('div');
  body.className = 'op-card-body';

  /* Linha: Unidade + Período */
  const row1 = document.createElement('div');
  row1.className = 'field-row';

  row1.appendChild(buildSelectField(op.id, 'unidade', 'Unidade', [
    { value: 'filial', label: 'Filial' },
    { value: 'matriz', label: 'Matriz' },
  ], opState.unidade));

  row1.appendChild(buildInputField(op.id, 'period', 'Período', 'ex: 22/04 a 29/04', opState.period));

  body.appendChild(row1);

  /* Grupos de status */
  STATUS_GROUPS.forEach(group => {
    body.appendChild(buildStatusGroup(op.id, group, opState[group.key]));
  });

  /* Observação */
  body.appendChild(buildTextareaField(op.id, 'obs', 'Observação', 'Notas adicionais...', opState.obs));

  /* Pílula de resumo do card */
  const pill = buildStatusPill(opState);
  body.appendChild(pill);

  card.appendChild(body);
  return card;
}

/** Campo <select> */
function buildSelectField(opId, fieldKey, label, options, currentVal) {
  const group = document.createElement('div');
  group.className = 'field-group';

  const lbl = document.createElement('label');
  lbl.className = 'field-label';
  lbl.textContent = label;

  const sel = document.createElement('select');
  sel.className = 'field-select';
  sel.dataset.opId = opId;
  sel.dataset.field = fieldKey;

  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt.value;
    o.textContent = opt.label;
    if (opt.value === currentVal) o.selected = true;
    sel.appendChild(o);
  });

  sel.addEventListener('change', () => {
    state.operators[opId][fieldKey] = sel.value;
    updateIndicators();
  });

  group.appendChild(lbl);
  group.appendChild(sel);
  return group;
}

/** Campo <input type="text"> */
function buildInputField(opId, fieldKey, label, placeholder, currentVal) {
  const group = document.createElement('div');
  group.className = 'field-group';

  const lbl = document.createElement('label');
  lbl.className = 'field-label';
  lbl.textContent = label;

  const inp = document.createElement('input');
  inp.type = 'text';
  inp.className = 'field-input';
  inp.placeholder = placeholder;
  inp.value = currentVal || '';
  inp.dataset.opId = opId;
  inp.dataset.field = fieldKey;

  inp.addEventListener('input', () => {
    state.operators[opId][fieldKey] = inp.value;
    updateIndicators();
  });

  group.appendChild(lbl);
  group.appendChild(inp);
  return group;
}

/** Campo <textarea> */
function buildTextareaField(opId, fieldKey, label, placeholder, currentVal) {
  const group = document.createElement('div');
  group.className = 'field-group';

  const lbl = document.createElement('label');
  lbl.className = 'field-label';
  lbl.textContent = label;

  const ta = document.createElement('textarea');
  ta.className = 'field-textarea';
  ta.placeholder = placeholder;
  ta.value = currentVal || '';
  ta.dataset.opId = opId;
  ta.dataset.field = fieldKey;

  ta.addEventListener('input', () => {
    state.operators[opId][fieldKey] = ta.value;
  });

  group.appendChild(lbl);
  group.appendChild(ta);
  return group;
}

/** Grupo de botões de status (conciliação / transferência / pendências) */
function buildStatusGroup(opId, group, currentVal) {
  const wrapper = document.createElement('div');
  wrapper.className = 'status-group';

  const lbl = document.createElement('div');
  lbl.className = 'status-label';
  lbl.textContent = group.label;

  const btnsRow = document.createElement('div');
  btnsRow.className = 'status-buttons';

  group.buttons.forEach(btnDef => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'status-btn';
    btn.textContent = btnDef.label;
    btn.dataset.value = btnDef.value;
    btn.dataset.activeCls = btnDef.cls;

    /* Aplicar estado salvo */
    if (currentVal === btnDef.value) {
      btn.classList.add(btnDef.cls);
    }

    btn.addEventListener('click', () => {
      /* Desselecionar todos os botões do grupo */
      btnsRow.querySelectorAll('.status-btn').forEach(b => {
        b.classList.remove('active-ok', 'active-warn', 'active-danger');
      });
      /* Selecionar este */
      btn.classList.add(btnDef.cls);
      /* Atualizar estado */
      state.operators[opId][group.key] = btnDef.value;
      /* Atualizar pílula do card */
      updateCardPill(opId);
      updateIndicators();
    });

    btnsRow.appendChild(btn);
  });

  wrapper.appendChild(lbl);
  wrapper.appendChild(btnsRow);
  return wrapper;
}

/** Pílula de resumo no rodapé do card */
function buildStatusPill(opState) {
  const pill = document.createElement('span');
  pill.className = 'card-status-pill';
  setStatusPillContent(pill, opState);
  return pill;
}

/** Atualiza a pílula de resumo de um card já renderizado */
function updateCardPill(opId) {
  const card = document.querySelector(`.op-card[data-op-id="${opId}"]`);
  if (!card) return;
  const pill = card.querySelector('.card-status-pill');
  if (!pill) return;
  setStatusPillContent(pill, state.operators[opId]);
}

function setStatusPillContent(pill, opState) {
  pill.classList.remove('pill-ok', 'pill-warn', 'pill-danger');
  const { conciliacao, transferencia, pendencias } = opState;

  const hasData = conciliacao || transferencia || pendencias;
  if (!hasData) {
    pill.textContent = '— Sem dados';
    return;
  }

  const isOk =
    conciliacao === 'ok' &&
    (transferencia === 'feito' || transferencia === 'nao_se_aplica') &&
    pendencias === 'sem_pendencia';

  const isDanger =
    conciliacao === 'nao_feito' || pendencias === 'com_pendencia';

  if (isOk) {
    pill.classList.add('pill-ok');
    pill.textContent = '✔ Tudo OK';
  } else if (isDanger) {
    pill.classList.add('pill-danger');
    pill.textContent = '❗ Com problema';
  } else {
    pill.classList.add('pill-warn');
    pill.textContent = '⚠ Incompleto / Pendente';
  }
}

/* ── Indicadores no topo ─────────────────────────────────────────────────── */

function updateIndicators() {
  let ok = 0, pending = 0, problem = 0, incomplete = 0;

  OPERATORS.forEach(op => {
    const s = state.operators[op.id];
    const { conciliacao, transferencia, pendencias } = s;

    const hasAny = conciliacao || transferencia || pendencias;
    if (!hasAny) return; /* Ignora operadoras sem nenhum dado */

    const isFullOk =
      conciliacao === 'ok' &&
      (transferencia === 'feito' || transferencia === 'nao_se_aplica') &&
      pendencias === 'sem_pendencia';

    const hasDanger =
      conciliacao === 'nao_feito' || pendencias === 'com_pendencia';

    const hasPending =
      transferencia === 'pendente';

    if (isFullOk)       ok++;
    else if (hasDanger) problem++;
    else if (hasPending) pending++;
    else                incomplete++;
  });

  document.getElementById('indOk').textContent         = ok;
  document.getElementById('indPending').textContent    = pending;
  document.getElementById('indProblem').textContent    = problem;
  document.getElementById('indIncomplete').textContent = incomplete;
}

/* ── Badge da semana no header ──────────────────────────────────────────── */

function renderWeekBadge() {
  const badge = document.getElementById('currentWeekBadge');
  const period = state.period;
  if (period && period.trim()) {
    badge.textContent = `Semana: ${period}`;
  } else {
    const now = new Date();
    badge.textContent = `Sem. ${getWeekNumber(now)} · ${now.getFullYear()}`;
  }
}

/** Retorna o número da semana ISO de uma data */
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/* ── Período global ─────────────────────────────────────────────────────── */

function bindGlobalEvents() {
  /* Aplicar período a todos os cards */
  document.getElementById('btnApplyPeriod').addEventListener('click', () => {
    const val = document.getElementById('globalPeriod').value.trim();
    if (!val) return;

    state.period = val;

    OPERATORS.forEach(op => {
      state.operators[op.id].period = val;
    });

    /* Atualizar todos os inputs de período nos cards */
    document.querySelectorAll('[data-field="period"]').forEach(inp => {
      inp.value = val;
    });

    renderWeekBadge();
    flashFeedback('Período aplicado a todos ✔');
  });

  /* Salvar semana */
  document.getElementById('btnSave').addEventListener('click', () => {
    saveToStorage();
    flashFeedback('Semana salva com sucesso! 💾');
  });

  /* Limpar histórico */
  document.getElementById('btnClearHistory').addEventListener('click', () => {
    if (!confirm('Tem certeza que deseja limpar todo o histórico?')) return;
    history = [];
    saveToStorage();
    renderHistory();
  });

  /* Modal: Nova Semana */
  document.getElementById('btnNovaSemana').addEventListener('click', () => {
    openModal();
  });

  document.getElementById('modalCancel').addEventListener('click', closeModal);

  document.getElementById('modalConfirm').addEventListener('click', () => {
    commitNewWeek();
    closeModal();
  });

  /* Fechar modal clicando fora */
  document.getElementById('modalOverlay').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeModal();
  });
}

/* ── Nova Semana ────────────────────────────────────────────────────────── */

/**
 * Salva o estado atual no histórico, limpa os campos
 * e re-renderiza tudo.
 */
function commitNewWeek() {
  /* Só salva no histórico se houver algum dado preenchido */
  const hasAnyData = OPERATORS.some(op => {
    const s = state.operators[op.id];
    return s.conciliacao || s.transferencia || s.pendencias || s.period || s.obs;
  });

  if (hasAnyData) {
    const entry = {
      savedAt:   new Date().toISOString(),
      period:    state.period,
      operators: JSON.parse(JSON.stringify(state.operators)),
    };
    history.unshift(entry); /* mais recente primeiro */
  }

  /* Limpar estado atual */
  state = buildEmptyState();

  /* Limpar campo de período global */
  document.getElementById('globalPeriod').value = '';

  saveToStorage();
  renderCards();
  renderWeekBadge();
  renderHistory();
  updateIndicators();
}

/* ── Modal ─────────────────────────────────────────────────────────────── */

function openModal() {
  document.getElementById('modalOverlay').classList.add('visible');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('visible');
}

/* ── Feedback visual ────────────────────────────────────────────────────── */

let feedbackTimer = null;

function flashFeedback(msg) {
  const el = document.getElementById('saveFeedback');
  el.textContent = msg;
  el.classList.add('visible');
  clearTimeout(feedbackTimer);
  feedbackTimer = setTimeout(() => el.classList.remove('visible'), 3000);
}

/* ── Histórico ──────────────────────────────────────────────────────────── */

function renderHistory() {
  const list    = document.getElementById('historyList');
  const empty   = document.getElementById('historyEmpty');
  list.innerHTML = '';

  if (!history.length) {
    empty.style.display = 'block';
    return;
  }

  empty.style.display = 'none';

  history.forEach((entry, idx) => {
    list.appendChild(buildHistoryWeekBlock(entry, idx));
  });
}

/**
 * Constrói o bloco de histórico de uma semana (acordeão).
 */
function buildHistoryWeekBlock(entry, idx) {
  const week = document.createElement('div');
  week.className = 'history-week';

  /* ── Cabeçalho do bloco ── */
  const header = document.createElement('div');
  header.className = 'history-week-header';

  const meta = document.createElement('div');
  meta.className = 'history-week-meta';

  const periodEl = document.createElement('span');
  periodEl.className = 'hw-period';
  periodEl.textContent = entry.period ? `📅 ${entry.period}` : '— Sem período';

  const dateEl = document.createElement('span');
  dateEl.className = 'hw-date';
  dateEl.textContent = formatSavedAt(entry.savedAt);

  /* Pílulas de resumo por operadora */
  const pillsEl = document.createElement('div');
  pillsEl.className = 'hw-pills';

  OPERATORS.forEach(op => {
    const opData = entry.operators[op.id];
    if (!opData) return;
    const pillClass = getHistoryPillClass(opData);
    if (!pillClass) return;
    const p = document.createElement('span');
    p.className = `hw-pill ${pillClass}`;
    p.textContent = op.name;
    pillsEl.appendChild(p);
  });

  meta.appendChild(periodEl);
  meta.appendChild(dateEl);
  meta.appendChild(pillsEl);

  const toggle = document.createElement('span');
  toggle.className = 'hw-toggle';
  toggle.textContent = '▾';

  header.appendChild(meta);
  header.appendChild(toggle);

  /* ── Corpo expansível ── */
  const body = document.createElement('div');
  body.className = 'history-week-body';

  OPERATORS.forEach(op => {
    const opData = entry.operators[op.id];
    if (!opData) return;
    body.appendChild(buildHistoryRow(op, opData));
  });

  /* Botão de restaurar */
  const restoreRow = document.createElement('div');
  restoreRow.style.cssText = 'display:flex;gap:8px;margin-top:4px;';

  const btnRestore = document.createElement('button');
  btnRestore.className = 'btn-clear-history';
  btnRestore.style.cssText = 'color:var(--accent);border-color:var(--accent-dim);';
  btnRestore.textContent = '↩ Restaurar esta semana';
  btnRestore.addEventListener('click', () => {
    if (!confirm('Restaurar esta semana irá substituir os dados atuais. Continuar?')) return;
    state = {
      period:    entry.period || '',
      operators: JSON.parse(JSON.stringify(entry.operators)),
    };
    document.getElementById('globalPeriod').value = state.period || '';
    saveToStorage();
    renderCards();
    renderWeekBadge();
    updateIndicators();
  });

  const btnDelete = document.createElement('button');
  btnDelete.className = 'btn-clear-history';
  btnDelete.textContent = '🗑 Excluir';
  btnDelete.addEventListener('click', () => {
    if (!confirm('Excluir este registro do histórico?')) return;
    history.splice(idx, 1);
    saveToStorage();
    renderHistory();
  });

  restoreRow.appendChild(btnRestore);
  restoreRow.appendChild(btnDelete);
  body.appendChild(restoreRow);

  /* Toggle acordeão */
  header.addEventListener('click', () => {
    const isOpen = body.classList.toggle('open');
    header.classList.toggle('open', isOpen);
    toggle.classList.toggle('open', isOpen);
  });

  week.appendChild(header);
  week.appendChild(body);
  return week;
}

/**
 * Linha de uma operadora no histórico.
 */
function buildHistoryRow(op, opData) {
  const row = document.createElement('div');
  row.className = 'hw-row';

  const nameEl = document.createElement('span');
  nameEl.className = `hw-op ${op.histCls}`;
  nameEl.textContent = op.name;

  const unitEl = document.createElement('span');
  unitEl.className = 'hw-unit';
  unitEl.textContent = opData.unidade ? opData.unidade.charAt(0).toUpperCase() + opData.unidade.slice(1) : '—';

  const concEl = buildHwStatusBadge(opData.conciliacao, {
    ok: '✔ OK', incompleto: '⚠ Incompleto', nao_feito: '✕ Não feito',
  });
  const transEl = buildHwStatusBadge(opData.transferencia, {
    feito: '✔ Feito', pendente: '⏳ Pendente', nao_se_aplica: '— N/A',
  });
  const pendEl = buildHwStatusBadge(opData.pendencias, {
    sem_pendencia: '✔ Sem pend.', com_pendencia: '❗ Com pend.',
  });

  const obsEl = document.createElement('span');
  obsEl.className = 'hw-obs';
  obsEl.textContent = opData.obs || '—';
  obsEl.title = opData.obs || '';

  row.appendChild(nameEl);
  row.appendChild(unitEl);
  row.appendChild(concEl);
  row.appendChild(transEl);
  row.appendChild(pendEl);
  row.appendChild(obsEl);

  return row;
}

/** Mini badge de status no histórico */
function buildHwStatusBadge(value, labelMap) {
  const el = document.createElement('span');
  el.style.cssText = 'font-size:10px;font-family:var(--font-mono);';

  if (!value || !labelMap[value]) {
    el.textContent = '—';
    el.style.color = 'var(--text-dim)';
    return el;
  }

  el.textContent = labelMap[value];

  /* Cor */
  if (value === 'ok' || value === 'feito' || value === 'sem_pendencia') {
    el.style.color = 'var(--ok)';
  } else if (value === 'com_pendencia' || value === 'nao_feito') {
    el.style.color = 'var(--danger)';
  } else {
    el.style.color = 'var(--warn)';
  }

  return el;
}

/** Retorna a classe CSS da pílula de resumo no histórico */
function getHistoryPillClass(opData) {
  const { conciliacao, transferencia, pendencias } = opData;
  const hasAny = conciliacao || transferencia || pendencias;
  if (!hasAny) return null;

  const isOk =
    conciliacao === 'ok' &&
    (transferencia === 'feito' || transferencia === 'nao_se_aplica') &&
    pendencias === 'sem_pendencia';

  const isDanger =
    conciliacao === 'nao_feito' || pendencias === 'com_pendencia';

  if (isOk)      return 'hw-pill-ok';
  if (isDanger)  return 'hw-pill-danger';
  return 'hw-pill-warn';
}

/** Formata o timestamp de salvamento */
function formatSavedAt(isoStr) {
  try {
    const d = new Date(isoStr);
    return d.toLocaleString('pt-BR', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit',
    });
  } catch {
    return isoStr;
  }
}