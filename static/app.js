/* Escala Supervisor — lógica da interface (3 blocos) */
'use strict';

const $ = (s) => document.querySelector(s);

const POSTOS = [
  { grupo: 'Oficiais Subalternos', itens: [
    ['Asp', 'Aspirante / Guarda-Marinha'],
    ['2º Ten', 'Segundo-Tenente'],
    ['1º Ten', 'Primeiro-Tenente'],
  ]},
  { grupo: 'Oficiais Intermediários', itens: [
    ['Cap', 'Capitão (EB/FAB)'],
    ['Cap-Ten', 'Capitão-Tenente (MB)'],
    ['CT', 'Capitão-Tenente (MB)'],
  ]},
  { grupo: 'Oficiais Superiores', itens: [
    ['Maj', 'Major (EB/FAB)'],
    ['CC', 'Capitão de Corveta (MB)'],
    ['TC', 'Tenente-Coronel (EB/FAB)'],
    ['CF', 'Capitão de Fragata (MB)'],
    ['Cel', 'Coronel (EB/FAB)'],
    ['CMG', 'Capitão de Mar e Guerra (MB)'],
  ]},
  { grupo: 'Oficiais-Generais', itens: [
    ['Gen Bda', 'General de Brigada (EB)'],
    ['Brig', 'Brigadeiro (FAB)'],
    ['CA', 'Contra-Almirante (MB)'],
    ['Gen Div', 'General de Divisão (EB)'],
    ['Maj Brig', 'Major-Brigadeiro (FAB)'],
    ['VA', 'Vice-Almirante (MB)'],
    ['Gen Ex', 'General de Exército (EB)'],
    ['Ten Brig', 'Tenente-Brigadeiro (FAB)'],
    ['AE', 'Almirante de Esquadra (MB)'],
  ]},
];

/* Concorrentes de exemplo, fictícios — [nome, posto, força, especialidade,
   não escalar fixo], ordenado por hierarquia (mais antigo -> mais moderno).
   São apenas demonstração: apague e coloque o seu quadro no Bloco 3. */
const CONCORRENTES_PADRAO = [
  ['ALFA', 'TC', 'EB', 'Med', ''],
  ['BRAVO', 'Maj', 'EB', 'Dent', ''],
  ['CHARLIE', 'Cap', 'Aer', 'Med', ''],
  ['DELTA', 'CT', 'MB', 'Farm', ''],
  ['ECHO', 'Cap', 'EB', 'ENF', ''],
];

let seqQ = 0;
let disp = {};    // id do quadro -> { disp: 'sim'|'nao', ind: '' }
let heranca = {}; // id do quadro -> { folga_util, folga_fds, ult_escala, posto_grade }

/* ======================= PERSISTÊNCIA (memória do navegador) ======================= */
const LS_ESTADO = 'escala-supervisor-estado-v1';
const LS_GERADO = 'escala-supervisor-gerado-v1';

function salvarEstado() {
  try {
    const quadro = lerQuadro();
    localStorage.setItem(LS_ESTADO, JSON.stringify({
      mes: $('#mes').value,
      feriados: [...document.querySelectorAll('#corpo-feriados .item-linha')].map(l => ({
        data: l.querySelector('.data').value, desc: l.querySelector('.desc').value })),
      quadro: quadro.map(m => ({ nome: m.nome, posto: m.posto, forca: m.forca,
        espec: m.espec, desempate: m.desempate, obs: m.obs, novo: m.novo })),
      disp: quadro.map(m => disp[m.id] || { ind: '', ne: '' }),
      heranca: quadro.map(m => heranca[m.id] || {}),
    }));
  } catch (e) { /* storage indisponível — segue sem memória */ }
}

let _timerSave = null;
function agendarSave() {
  clearTimeout(_timerSave);
  _timerSave = setTimeout(salvarEstado, 350);
}

function restaurarEstado(st) {
  try {
    $('#mes').value = st.mes || '';
    $('#corpo-feriados').innerHTML = '';
    (st.feriados || []).forEach(f => addLinhaFeriado(f));
    $('#corpo-quadro').innerHTML = '';
    disp = {};
    heranca = {};
    (st.quadro || []).forEach(m => addMilitar(m, m.novo === '1' || m.novo === true));
    const quadro = lerQuadro();
    (st.disp || []).forEach((d, i) => {
      if (!quadro[i]) return;
      const a = d || {};
      // migração de formatos antigos (com estado/rádio) -> colunas abertas {ind, ne}
      const e = a.estado || (a.disp === 'nao' ? 'ind' : 'disp');
      const texto = (a.texto ?? '') || '';
      let ind = (a.ind ?? '') || '';
      let ne = (a.ne ?? '') || '';
      if (e === 'ind' && !ind) ind = texto;
      if (e === 'ne' && !ne) ne = texto;
      disp[quadro[i].id] = { ind, ne };
    });
    (st.heranca || []).forEach((h, i) => { if (quadro[i]) heranca[quadro[i].id] = h; });
    ordenarQuadro();
    renderConcorrentes();
    renderHeranca();
    return true;
  } catch (e) {
    return false;
  }
}

function carregarEstadoSalvo() {
  try {
    return JSON.parse(localStorage.getItem(LS_ESTADO) || 'null');
  } catch (e) {
    return null;
  }
}

/* Backup da memória do navegador em JSON. Não altera o estado; só lê e grava. */
function exportarEstado() {
  salvarEstado(); // garante que a última edição já entrou no localStorage
  const raw = localStorage.getItem(LS_ESTADO);
  if (!raw) return setStatus('Nada para exportar: a memória está vazia.', 'erro');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([raw], { type: 'application/json' }));
  const hoje = new Date();
  const data = [hoje.getFullYear(), String(hoje.getMonth() + 1).padStart(2, '0'),
                String(hoje.getDate()).padStart(2, '0')].join('-');
  a.download = 'escala-estado-' + data + '.json';
  a.click();
  URL.revokeObjectURL(a.href);
  setStatus('Estado exportado em JSON.', 'ok');
}

function importarEstado(arquivo) {
  const leitor = new FileReader();
  leitor.onload = () => {
    try {
      const st = JSON.parse(leitor.result);
      if (!st || !Array.isArray(st.quadro) || !st.quadro.length) {
        throw new Error('O arquivo não tem um quadro válido.');
      }
      if (!confirm(`Importar ${st.quadro.length} militares e substituir a memória atual?`)) {
        return;
      }
      localStorage.setItem(LS_ESTADO, JSON.stringify(st));
      location.reload();
    } catch (e) {
      mostraErros([e.message || 'Falha ao importar o estado.']);
    }
  };
  leitor.onerror = () => mostraErros(['Falha ao ler o arquivo.']);
  leitor.readAsText(arquivo);
}

function salvarGerado(nome) {
  try {
    localStorage.setItem(LS_GERADO, JSON.stringify({
      nome, quando: new Date().toLocaleString('pt-BR') }));
  } catch (e) { /* ignora */ }
}

function geradoInfo() {
  try {
    return JSON.parse(localStorage.getItem(LS_GERADO) || 'null');
  } catch (e) {
    return null;
  }
}

function esc(s) {
  return String(s ?? '').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

function optionsPosto() {
  let html = '<option value="">— escolha —</option>';
  for (const g of POSTOS) {
    html += `<optgroup label="${g.grupo}">`;
    for (const [v, t] of g.itens) html += `<option value="${v}">${v} — ${t}</option>`;
    html += '</optgroup>';
  }
  return html;
}

function optionsForca() {
  return '<option value="">—</option>' +
    '<option value="EB">EB — Exército</option>' +
    '<option value="MB">MB — Marinha</option>' +
    '<option value="Aer">Aer — Aeronáutica</option>';
}

/* ======================= CALENDÁRIO (flatpickr, locale pt-BR) ======================= */
/* O campo continua livre para digitação; o botão 📅 abre o calendário ancorado.
   Clique em duas datas = período (ou na mesma data = dia único). */

function fmtData(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}`;
}

function mesDaConfig() {
  const v = $('#mes').value; // aaaa-mm
  if (v) {
    const [aa, mm] = v.split('-').map(Number);
    return new Date(aa, mm - 1, 1);
  }
  const h = new Date();
  return new Date(h.getFullYear(), h.getMonth(), 1);
}

function calendarioAnexar(inputEl) {
  if (!inputEl || inputEl.dataset.cal) return;
  inputEl.dataset.cal = '1';
  const wrap = document.createElement('span');
  wrap.className = 'cal-wrap';
  inputEl.parentNode.insertBefore(wrap, inputEl);
  wrap.appendChild(inputEl);
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'cal-btn';
  btn.title = 'Calendário: escolha o início e o fim do período (ou a mesma data para um dia)';
  btn.textContent = '📅';
  wrap.appendChild(btn);
  let fp = null;

  function inserir(txt) {
    const atual = inputEl.value.trim();
    inputEl.value = atual ? atual + '; ' + txt : txt;
    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
    salvarEstado();
  }

  function abrir() {
    const alvo = mesDaConfig();
    if (fp) { try { fp.destroy(); } catch (e) { /* ignora */ } fp = null; }
    const ghost = document.createElement('input');
    ghost.type = 'text';
    ghost.className = 'cal-ghost';
    ghost.setAttribute('aria-hidden', 'true');
    document.body.appendChild(ghost);
    try {
      fp = flatpickr(ghost, {
        locale: flatpickr.l10ns.pt,
        dateFormat: 'd/m/Y',
        defaultDate: alvo,
        mode: 'range',
        positionElement: btn,
        onChange: (sel) => {
          if (sel.length === 2) {
            const a = sel[0];
            const b = sel[1];
            inserir(fmtData(a) === fmtData(b)
              ? fmtData(a)
              : `${fmtData(a)}-${fmtData(b)}`);
            fp.close();
          }
        },
      });
    } catch (e) {
      // flatpickr indisponível: mantém só a digitação livre
      fp = null;
      return;
    }
    fp.clear();
    fp.open();
    fp.jumpToDate(alvo, false);
  }

  btn.addEventListener('pointerdown', (ev) => ev.stopPropagation());
  btn.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    if (fp && fp.isOpen) fp.close();
    else abrir();
  });
}

/* ========================= BLOCO 3 — QUADRO ========================= */

function addMilitar(dados, novo) {
  const tbody = $('#corpo-quadro');
  const id = ++seqQ;
  const tr = document.createElement('tr');
  tr.dataset.id = id;
  tr.dataset.novo = novo ? '1' : '0';
  tr.innerHTML = `
    <td><input data-campo="nome" value="${esc(dados?.nome || '')}" placeholder="Nome de guerra"></td>
    <td><select data-campo="posto">${optionsPosto()}</select></td>
    <td><select data-campo="forca">${optionsForca()}</select></td>
    <td><input data-campo="espec" value="${esc(dados?.espec || '')}" placeholder="ex.: Med"></td>
    <td class="fixo"><input data-campo="nao_escalar_fixo" value="${esc(dados?.nao_escalar_fixo || '')}" placeholder="ex.: ter; qua; qui"></td>
    <td><input data-campo="desempate" type="number" min="1" value="${esc(dados?.desempate ?? '')}" placeholder="1"></td>
    <td><input data-campo="obs" value="${esc(dados?.obs || '')}"></td>
    <td class="ck-novo-td"><input type="checkbox" class="ck-novo" ${novo ? 'checked' : ''} title="Novo na escala: entra com folga 0 e é escalado logo (não herda folgas da planilha anterior)"></td>
    <td><button type="button" class="remover" data-quadro="${id}" title="Remover">✕</button></td>`;
  tbody.appendChild(tr);
  if (dados?.posto) tr.querySelector('[data-campo="posto"]').value = dados.posto;
  if (dados?.forca) tr.querySelector('[data-campo="forca"]').value = dados.forca;
  calendarioAnexar(tr.querySelector('[data-campo="nao_escalar_fixo"]'));
  return id;
}

function lerQuadro() {
  const out = [];
  for (const tr of document.querySelectorAll('#corpo-quadro tr')) {
    const m = {};
    for (const el of tr.querySelectorAll('[data-campo]')) m[el.dataset.campo] = el.value.trim();
    if (!m.nome && !m.posto && !m.forca) continue; // linha vazia
    m.id = tr.dataset.id;
    m.novo = tr.dataset.novo === '1';
    out.push(m);
  }
  return out;
}

/* Ordem visual: posto mais antigo primeiro e nome alfabético dentro do posto.
   O motor não usa esta ordem: a escolha do dia continua pela regra do mais folgado. */
const RANK = {
  'Asp': 1, 'GM': 1,
  '2º Ten': 2,
  '1º Ten': 3,
  'Cap': 4, 'Cap-Ten': 4, 'CT': 4,
  'Maj': 5, 'CC': 5,
  'TC': 6, 'CF': 6,
  'Cel': 7, 'CMG': 7,
  'Gen Bda': 8, 'Brig': 8, 'CA': 8,
  'Gen Div': 9, 'Maj Brig': 9, 'VA': 9,
  'Gen Ex': 10, 'Ten Brig': 10, 'AE': 10,
};

function ordenarQuadro() {
  const tbody = $('#corpo-quadro');
  const linhas = Array.from(tbody.querySelectorAll('tr'));
  const valor = (tr, campo) =>
    tr.querySelector(`[data-campo="${campo}"]`)?.value || '';
  linhas.sort((a, b) => {
    const pa = RANK[valor(a, 'posto')] || 0;
    const pb = RANK[valor(b, 'posto')] || 0;
    if (pa !== pb) return pb - pa; // mais antigo primeiro
    return valor(a, 'nome').localeCompare(valor(b, 'nome'), 'pt-BR',
                                          { sensitivity: 'base' });
  });
  for (const tr of linhas) tbody.appendChild(tr);
}

function carregarEfetivo() {
  $('#corpo-quadro').innerHTML = '';
  disp = {};
  heranca = {};
  for (const [nome, posto, forca, espec, fixo] of CONCORRENTES_PADRAO) {
    addMilitar({ nome, posto, forca, espec, nao_escalar_fixo: fixo });
  }
  ordenarQuadro();
  renderConcorrentes();
  esconderHeranca();
  $('#corpo-heranca').innerHTML = '';
  salvarEstado();
}

/* ========================= BLOCO 2 — CONCORRENTES ========================= */

function renderConcorrentes() {
  const tbody = $('#corpo-concorrentes');
  tbody.innerHTML = '';
  const quadro = lerQuadro();
  quadro.forEach((m) => {
    if (!disp[m.id]) disp[m.id] = { ind: '', ne: '' };
    const d = disp[m.id];
    const posto = m.posto ? m.posto + (m.forca ? ' (' + m.forca + ')' : '') : (m.forca || '—');
    const tr = document.createElement('tr');
    tr.dataset.id = m.id;
    tr.innerHTML = `
      <td class="posto">${esc(posto)}</td>
      <td class="nome">${esc(m.nome)}${m.novo ? ' <span class="selo-novo" title="Novo na escala: entra com folga 0 e é escalado logo">novo</span>' : ''}</td>
      <td class="campo-ind">
        <input data-campo="ind" value="${esc(d.ind)}" placeholder="Período de afastamento (ex.: 10-20; 10out-15out)">
      </td>
      <td class="campo-ne">
        <input data-campo="ne" value="${esc(d.ne)}" placeholder="ex.: ter; qua; qui (opcional)">
      </td>`;
    tbody.appendChild(tr);
    tr.querySelector('.campo-ind input').addEventListener('input', (ev) => {
      disp[m.id].ind = ev.target.value;
    });
    tr.querySelector('.campo-ne input').addEventListener('input', (ev) => {
      disp[m.id].ne = ev.target.value;
    });
    calendarioAnexar(tr.querySelector('.campo-ind input'));
    calendarioAnexar(tr.querySelector('.campo-ne input'));
  });
  $('#contagem').textContent = quadro.length ? `— ${quadro.length} concorrente(s)` : '';
}

/* ======================== BLOCO 1 — FERIADOS ======================== */

function addLinhaFeriado(dados) {
  const div = document.createElement('div');
  div.className = 'item-linha';
  div.innerHTML = `<input class="data" placeholder="dd/mm ou 10out" value="${esc(dados?.data || '')}">
    <input class="desc" placeholder="Descrição (ex.: Independência do Brasil)" value="${esc(dados?.desc || '')}">
    <button type="button" class="remover" title="Remover">✕</button>`;
    $('#corpo-feriados').appendChild(div);
  calendarioAnexar(div.querySelector('.data'));
}

/* ======================== BLOCO 1b — HERANÇA DO MÊS ANTERIOR ======================== */

function mesAnterior(metade) {
  const mesVal = $('#mes').value; // aaaa-mm
  if (!mesVal) return '';
  const [aa, mm] = mesVal.split('-').map(Number);
  const d = new Date(Date.UTC(aa, mm - 1, 1));
  d.setUTCMonth(d.getUTCMonth() - 1);
  return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
}

function arquivoEmBase64(arquivo) {
  return new Promise((ok, falha) => {
    const fr = new FileReader();
    fr.onload = () => {
      const s = String(fr.result);
      ok(s.slice(s.indexOf(',') + 1));   // data:...;base64,XXXX
    };
    fr.onerror = () => falha(fr.error || new Error('falha ao ler o arquivo'));
    fr.readAsDataURL(arquivo);
  });
}

async function importarFolgas() {
  const arquivo = $('#arquivo-anterior').files[0];
  const ref = mesAnterior();
  if (!arquivo) return mostraErros(['Escolha a planilha do mês anterior (.xlsx).']);
  if (!ref) return mostraErros(['Informe o mês em edição (para saber qual é o anterior).']);
  const quadro = lerQuadro();
  if (!quadro.length) return mostraErros(['O quadro (bloco 3) está vazio.']);
  mostraErros([]);
  mostraAvisos([]);
  setStatus('Lendo folgas do mês anterior…');
  try {
    const arquivo_b64 = await arquivoEmBase64(arquivo);
    // posto e força vão junto: são o desempate entre homônimos na grade anterior
    const militares = quadro.map(m => ({ nome: m.nome, posto: m.posto, forca: m.forca }));
    const r = await requisicao('/api/importar', { arquivo_b64, mes_ref: ref, militares });
    // novo mês: a herança anterior é substituída pela leitura desta planilha
    heranca = {};
    let aplicadas = 0;
    const semRegistro = [];
    const naoEncontrados = [];
    // resultados vêm na ordem do quadro enviado: casa por posição, sem adivinhar nome
    r.resultados.forEach((res, i) => {
      const alvo = quadro[i];
      if (!alvo) return;
      if (!res.encontrado) {
        naoEncontrados.push(res.nome + (res.motivo ? ` (${res.motivo})` : ''));
        return;
      }
      // "novo" (inserido agora) não herda folgas da planilha anterior — entra com 0
      if (alvo.novo) return;
      heranca[alvo.id] = {
        folga_util: res.folga_util ?? '',
        folga_fds: res.folga_fds ?? '',
        ult_escala: res.ult_escala || '',
        posto_grade: res.posto_grade || '',
        sem_registro: !!res.sem_registro,
      };
      if (res.sem_registro) semRegistro.push(res.nome); else aplicadas += 1;
    });
    // militares que estão na planilha e não estavam no quadro entram agora
    const vindosDaPlanilha = [];
    for (const ex of (r.extras || [])) {
      const id = addMilitar({ nome: ex.nome, posto: ex.posto, forca: ex.forca,
                              espec: ex.espec });
      heranca[id] = {
        folga_util: ex.folga_util ?? '',
        folga_fds: ex.folga_fds ?? '',
        ult_escala: ex.ult_escala || '',
        posto_grade: ex.posto_grade || '',
        sem_registro: !!ex.sem_registro,
      };
      vindosDaPlanilha.push(`${ex.posto_grade} ${ex.nome}`.trim());
      if (ex.sem_registro) semRegistro.push(ex.nome); else aplicadas += 1;
    }
    ordenarQuadro();
    renderConcorrentes();
    renderHeranca();
    $('#secao-heranca').hidden = false;
    salvarEstado();
    const partes = [`Folgas carregadas para ${aplicadas} militar(es) (mês anterior: ${ref}).`];
    if (vindosDaPlanilha.length) partes.push(`Entraram no quadro vindos da planilha: ${vindosDaPlanilha.join(', ')}.`);
    if (semRegistro.length) partes.push(`Sem folga registrada no mês anterior: ${semRegistro.join(', ')}.`);
    if (naoEncontrados.length) partes.push(`Fora da planilha: ${naoEncontrados.join(', ')}.`);
    setStatus(partes.join(' '), 'ok');
    const avisos = [];
    if (vindosDaPlanilha.length) {
      avisos.push(`Estavam na planilha anterior e não no quadro; incluí agora (confira posto, força e especialidade): ${vindosDaPlanilha.join(', ')}.`);
    }
    if (semRegistro.length) {
      avisos.push(`Ficaram indisponíveis o mês inteiro, sem folga para herdar; a linha está na tabela para você preencher à mão: ${semRegistro.join(', ')}.`);
    }
    if (naoEncontrados.length) {
      avisos.push(`Não encontrados na planilha anterior (conferir): ${naoEncontrados.join(', ')}.`);
    }
    mostraAvisos(avisos);
  } catch (e) {
    setStatus('Falha ao carregar folgas.', 'erro');
    mostraErros([e.message]);
  }
}

function renderHeranca() {
  const tbody = $('#corpo-heranca');
  tbody.innerHTML = '';
  for (const m of lerQuadro()) {
    const h = heranca[m.id];
    // {} é militar sem importação; quem veio da planilha tem os campos definidos,
    // ainda que vazios (ficou indisponível o mês inteiro) — esse não some da tela
    if (!h || (h.folga_util === undefined && h.folga_fds === undefined
               && !h.ult_escala && !h.sem_registro)) continue;
    const tr = document.createElement('tr');
    tr.dataset.id = m.id;
    const marca = h.sem_registro
      ? ' <span class="selo-novo" title="Ficou indisponível o mês inteiro na planilha anterior: não há folga para herdar">sem registro</span>'
      : '';
    tr.innerHTML = `
      <td class="pg">${esc(h.posto_grade || m.posto + (m.forca ? ' (' + m.forca + ')' : ''))}</td>
      <td class="nm">${esc(m.nome)}${marca}</td>
      <td><input class="fu" type="number" min="0" value="${esc(h.folga_util)}"></td>
      <td><input class="ff" type="number" min="0" value="${esc(h.folga_fds)}"></td>
      <td><input class="ue" placeholder="dd/mm/aaaa" value="${esc(h.ult_escala)}"></td>`;
    tbody.appendChild(tr);
    tr.querySelector('.fu').addEventListener('input', ev => { heranca[m.id].folga_util = ev.target.value; });
    tr.querySelector('.ff').addEventListener('input', ev => { heranca[m.id].folga_fds = ev.target.value; });
    tr.querySelector('.ue').addEventListener('input', ev => { heranca[m.id].ult_escala = ev.target.value; });
  }
  $('#secao-heranca').hidden = !tbody.children.length;
}

function esconderHeranca() {
  $('#secao-heranca').hidden = true;
}

/* ========================= COLETA ========================= */

function coleta() {
  const mesVal = $('#mes').value; // aaaa-mm
  let mes = '';
  if (mesVal) {
    const [aa, mm] = mesVal.split('-');
    mes = `${mm}/${aa}`;
  }
  const feriados = [];
  for (const linha of document.querySelectorAll('#corpo-feriados .item-linha')) {
    const data = linha.querySelector('.data').value.trim();
    const desc = linha.querySelector('.desc').value.trim();
    if (!data) continue;
    feriados.push({ data, desc });
  }
  const militares = [];
  for (const m of lerQuadro()) {
    const d = disp[m.id] || { ind: '', ne: '' };
    const h = heranca[m.id] || {};
    militares.push({
      nome: m.nome, posto: m.posto, forca: m.forca, espec: m.espec,
      desempate: m.desempate, obs: m.obs,
      nao_escalar: (d.ind || '').trim(),
      nao_escalar_mensal: (d.ne || '').trim(),
      nao_escalar_fixo: m.nao_escalar_fixo || '',
      folga_util: h.folga_util ?? '', folga_fds: h.folga_fds ?? '',
      ult_escala: h.ult_escala ?? '',
    });
  }
  return { config: { mes, feriados }, militares, erros: [] };
}

/* ========================= REQUISIÇÕES ========================= */

function setStatus(texto, classe) {
  const st = $('#status');
  st.textContent = texto;
  st.className = 'status' + (classe ? ' ' + classe : '');
}

function mostraErros(erros) {
  const box = $('#erros');
  box.hidden = !erros || !erros.length;
  box.textContent = (erros || []).join('\n');
  // a caixa fica no fim da página: sem isto o dono só vê "Falha ao carregar folgas."
  // no topo e nunca lê o motivo
  if (!box.hidden) box.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function mostraAvisos(avisos) {
  const box = $('#avisos');
  box.hidden = !avisos || !avisos.length;
  box.innerHTML = (avisos && avisos.length)
    ? '<strong>Avisos:</strong><ul>' + avisos.map(a => `<li>${esc(a)}</li>`).join('') + '</ul>'
    : '';
}

async function requisicao(caminho, payload) {
  let res;
  try {
    res = await fetch(caminho, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
  } catch (e) {
    throw new Error('Não foi possível falar com o servidor da escala (' + caminho + '). ' +
      'Ele roda no servidor local: confira http://127.0.0.1:8770/ e o serviço.');
  }
  if (caminho === '/api/gerar' && res.ok) return res;
  const dados = await res.json().catch(() => null);
  if (!res.ok) throw new Error((dados?.erros || ['Erro inesperado (HTTP ' + res.status + ').']).join('\n'));
  return dados;
}

/* ========================= RESULTADO ========================= */

function coresPorEscala(escala) {
  if (escala === 'VERMELHA') return { th: 've', cel: 'cel-ve' };
  return { th: 'pr', cel: 'cel-pr' };
}

function renderGrade(visao) {
  const dias = visao.dias;
  const head1 = ['<th>MILITAR</th>'];
  const head2 = ['<th></th>'];
  for (const d of dias) {
    const c = coresPorEscala(d.escala);
    head1.push(`<th class="${c.th}">${d.escala}</th>`);
    head2.push(`<th class="${c.th}">${d.data.slice(0, 5)}</th>`);
  }
  let html = `<table class="grade-t"><thead><tr>${head1.join('')}</tr><tr>${head2.join('')}</tr></thead><tbody>`;
  for (const m of visao.grade) {
    const cels = m.celulas.map((v, i) => {
      const c = coresPorEscala(dias[i].escala);
      let conteudo = String(v);
      if (v === 'D') conteudo = '<span class="d">D</span>';
      else if (v === 'IND') conteudo = '<span class="ind">IND</span>';
      return `<td class="${c.cel}">${conteudo}</td>`;
    }).join('');
    html += `<tr><th>${esc(m.militar)}</th>${cels}</tr>`;
  }
  $('#grade').innerHTML = html + '</tbody></table>';
}

const DIA_SEMANA_ABR = { 'segunda-feira': 'seg', 'terça-feira': 'ter', 'quarta-feira': 'qua',
  'quinta-feira': 'qui', 'sexta-feira': 'sex', 'sábado': 'sáb', 'domingo': 'dom' };

function renderDiario(visao) {
  const linhas = visao.dias.map(d => {
    const cls = d.escala === 'VERMELHA' && d.semana === 'sábado' ? 'linha-sab' : '';
    return `<tr class="${cls}"><td>${d.data}</td><td>${d.semana} (${DIA_SEMANA_ABR[d.semana]})</td>` +
      `<td>${esc(d.militar)}</td>` +
      `<td class="nota-cell">${esc(d.nota || '')}</td></tr>`;
  }).join('');
  $('#diario').innerHTML =
    '<thead><tr><th>DATA</th><th>DIA DA SEMANA</th><th>MILITAR DE SERVIÇO</th><th>OBSERVAÇÃO</th></tr></thead>' +
    '<tbody>' + linhas + '</tbody>';
}

async function preview() {
  mostraErros([]);
  mostraAvisos([]);
  const { config, militares, erros } = coleta();
  if (erros.length) {
    setStatus('Corrija os erros antes de pré-visualizar.', 'erro');
    mostraErros(erros);
    return;
  }
  if (!config.mes) {
    mostraErros(['Informe o mês.']);
    return;
  }
  if (!militares.length) {
    mostraErros(['Adicione ao menos um militar no bloco 3 (Quadro).']);
    return;
  }
  const btn = $('#btn-preview');
  btn.disabled = true;
  setStatus('Calculando…');
  try {
    const v = await requisicao('/api/preview', { config, militares });
    renderGrade(v);
    renderDiario(v);
    mostraAvisos(v.avisos);
    const porSerie = {};
    for (const d of v.dias) porSerie[d.escala] = (porSerie[d.escala] || 0) + 1;
    $('#resumo').textContent =
      `Mês ${v.mes} — ${v.dias.length} dias ` +
      Object.entries(porSerie).map(([k, n]) => `${k}: ${n}`).join(' · ') +
      ` — ${v.grade.length} militar(es)`;
    $('#secao-resultado').hidden = false;
    setStatus('Pré-visualização gerada.', 'ok');
    $('#secao-resultado').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (e) {
    setStatus('Falha na pré-visualização.', 'erro');
    mostraErros([e.message]);
  } finally {
    btn.disabled = false;
  }
}

async function gerar() {
  const { config, militares, erros } = coleta();
  if (erros.length) {
    setStatus('Corrija os erros antes de gerar.', 'erro');
    mostraErros(erros);
    return;
  }
  if (!config.mes) return mostraErros(['Informe o mês.']);
  if (!militares.length) return mostraErros(['Adicione ao menos um militar no bloco 3.']);
  if (!confirm('Confirmar a geração da planilha .xlsx com os dados atuais?')) return;
  setStatus('Gerando planilha…');
  try {
    const res = await requisicao('/api/gerar', { config, militares });
    const blob = await res.blob();
    const nome = 'escala_' + config.mes.replace('/', '_') + '.xlsx';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nome;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    salvarGerado(nome);
    setStatus(`Planilha gerada: ${nome}`, 'ok');
  } catch (e) {
    setStatus('Falha ao gerar.', 'erro');
    mostraErros([e.message]);
  }
}

/* ========================= INICIALIZAÇÃO ========================= */

$('#btn-preview').addEventListener('click', preview);
$('#btn-gerar').addEventListener('click', gerar);
$('#btn-rolar-gerar').addEventListener('click', gerar);
$('#btn-add-quadro').addEventListener('click', () => {
  addMilitar({}, true);
  ordenarQuadro();
});
$('#btn-add-feriado').addEventListener('click', () => addLinhaFeriado());
$('#btn-importar').addEventListener('click', importarFolgas);
$('#btn-exportar').addEventListener('click', exportarEstado);
$('#btn-importar-estado').addEventListener('click', () => $('#arquivo-estado').click());
$('#arquivo-estado').addEventListener('change', (ev) => {
  const f = ev.target.files[0];
  if (f) importarEstado(f);
  ev.target.value = '';
});
$('#arquivo-anterior').addEventListener('change', ev => {
  const f = ev.target.files[0];
  $('#arquivo-anterior-nome').textContent = f ? f.name : '';
});
$('#btn-limpar-heranca').addEventListener('click', () => {
  heranca = {};
  renderHeranca();
  salvarEstado();
  setStatus('Herança do mês anterior removida.', 'ok');
});
$('#btn-efetivo').addEventListener('click', () => {
  if (confirm('Substituir o quadro pelos 5 concorrentes de exemplo? ' +
              'As disponibilidades e folgas herdadas serão perdidas.')) {
    carregarEfetivo();
    $('#secao-resultado').hidden = true;
  }
});

document.addEventListener('click', (ev) => {
  if (!ev.target.classList.contains('remover')) return;
  if (ev.target.dataset.quadro) {
    const id = ev.target.dataset.quadro;
    delete disp[id];
    delete heranca[id];
    ev.target.closest('tr').remove();
    renderConcorrentes();
    renderHeranca();
    salvarEstado(); // remove efetivamente: não ressuscita ao recarregar
    return;
  }
  const linha = ev.target.closest('.item-linha');
  if (linha) {
    linha.remove();
    salvarEstado();
  }
});

document.addEventListener('change', (ev) => {
  if (ev.target.closest('#corpo-quadro')) {
    ordenarQuadro();
    const tr = ev.target.closest('tr');
    if (ev.target.classList.contains('ck-novo')) {
      tr.dataset.novo = ev.target.checked ? '1' : '0';
      salvarEstado();
    }
    renderConcorrentes();
    renderHeranca();
  }
});

document.addEventListener('keydown', (ev) => {
  if (ev.key === 'Enter' && ev.target.matches('#corpo-quadro input, #corpo-quadro select')) {
    ev.preventDefault();
    preview();
  }
});

// inicialização: restaura o que estava na tela (memória do navegador) ou o padrão
const estadoSalvo = carregarEstadoSalvo();
if (estadoSalvo && Array.isArray(estadoSalvo.quadro) && estadoSalvo.quadro.length) {
  if (restaurarEstado(estadoSalvo)) {
    const g = geradoInfo();
    setStatus(g
      ? `Tudo restaurado — última geração: ${g.nome} (${g.quando}).`
      : 'Estado restaurado da última sessão.', 'ok');
  } else {
    carregarEfetivo();
    addLinhaFeriado({ data: '12/10', desc: 'Nossa Senhora Aparecida' });
  }
} else {
  carregarEfetivo();
  addLinhaFeriado({ data: '12/10', desc: 'Nossa Senhora Aparecida' });
}

// memória: qualquer edição é salva automaticamente
document.addEventListener('input', agendarSave);
document.addEventListener('change', agendarSave);

// recomeçar do zero (limpa a memória local)
$('#btn-recomecar').addEventListener('click', () => {
  if (confirm('Apagar a memória do navegador e recomeçar com os 5 concorrentes de exemplo?\n'
    + 'Disponibilidades, folgas herdadas e feriados digitados serão perdidos.')) {
    localStorage.removeItem(LS_ESTADO);
    carregarEfetivo();
    $('#corpo-feriados').innerHTML = '';
    addLinhaFeriado({});
    $('#secao-resultado').hidden = true;
    mostraErros([]);
    mostraAvisos([]);
    setStatus('Recomeçado — memória local limpa.', 'ok');
  }
});