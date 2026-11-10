#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Escala de Serviço por Folga — organização militar (Exército, Marinha e Aeronáutica)
===========================================================
Gera a escala mensal no padrão "mais folgado", com 2 planilhas de saída
(Diário e Grade), a partir de uma planilha de entrada (aba "Entrada")
ou de dados JSON (interface web — app.py).

REGRAS IMPLEMENTADAS (conforme combinado)
-------------------------------------------
1. Duas séries independentes de escala:
   - PRETA    : dia útil (segunda a sexta, fora feriado)
   - VERMELHA : sábado, domingo e feriado
2. A cada dia, escala-se o militar ELEGÍVEL com MAIS dias de folga na série
   do dia. Elegível = disponível na data E com diferença >= 3 dias desde a
   última escala (2 dias inteiros de folga).
3. Desempate: menor hierarquia (mais moderno) -> DESEMPATE (menor = mais
   moderno) -> nome.
4. Contador de folga por série e independente: incrementa 1 a cada dia da
   própria série e zera no dia em que o militar escala na série.
   Célula do escalado = "D"; célula de indisponibilidade = "IND".
5. Início do mês: contadores semeados pelas últimas escalas (PRETA e
   VERMELHA) informadas; a regra do intervalo mínimo também vale para a
   última escala do mês anterior.
6. Sem candidato elegível, a regra é relaxada com AVISO registrado na
   planilha (nunca fica dia sem escalado).

USO
---
    python3 escala.py --template entrada.xlsx   # cria/recria o modelo de entrada
    python3 escala.py entrada.xlsx              # lê a entrada e (re)gera Diário e Grade
    python3 app.py                              # interface web (ver app.py)
"""

import datetime as dt
import io
import re
import sys
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Normalização de texto (acentos e caixa)
# ---------------------------------------------------------------------------

_ACENTOS = str.maketrans({
    'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
    'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
    'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
    'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
    'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c', 'Ç': 'C', 'ª': '', 'º': '',
})


def norm(texto):
    return str(texto or '').lower().translate(_ACENTOS).strip()


def tostr(valor):
    if isinstance(valor, dt.datetime):
        return valor.strftime('%d/%m/%Y')
    if isinstance(valor, dt.date):
        return valor.strftime('%d/%m/%Y')
    return str(valor or '')


# ---------------------------------------------------------------------------
# Postos de oficiais (Lei 6.880/1980) — índice = senioridade (menor = mais moderno)
# ---------------------------------------------------------------------------

POSTO_RULES = [
    (r'^(guarda.?marinha|marinheiro|aspirante.?a.?oficial|aspirante|gm|asp)$', 1),
    (r'^(2[ºo]?|segundo).?ten', 2),
    (r'^(1[ºo]?|primeiro).?ten', 3),
    (r'^(capitao.?tenente|cap.?ten|ct)$', 4),
    (r'^(capitao|cap)$', 4),
    (r'^(capitao.?de.?corveta|cc)$', 5),
    (r'^(major|maj)$', 5),
    (r'^(capitao.?de.?fragata|cf)$', 6),
    (r'^(tenente.?coronel|ten.?cel|tc)$', 6),
    (r'^(capitao.?de.?mar.?e.?guerra|cmg)$', 7),
    (r'^(coronel|cel|col)$', 7),
    (r'^(contra.?almirante|ca)$', 8),
    (r'^(general.?de.?brigada|gen.?bda|brigadeiro|brig)$', 8),
    (r'^(vice.?almirante|va)$', 9),
    (r'^(general.?de.?divisao|gen.?div|major.?brigadeiro|maj.?brig)$', 9),
    (r'^(almirante.?de.?esquadra|ae)$', 10),
    (r'^(general.?de.?exercito|gen.?ex|tenente.?brigadeiro|ten.?brig)$', 10),
    (r'^(almirante|marechal.?do.?ar|marechal)$', 11),
]

ROT_EB_FA = {1: 'Asp', 2: '2º Ten', 3: '1º Ten', 4: 'Cap', 5: 'Maj', 6: 'TC',
             7: 'Cel', 8: 'Gen Bda', 9: 'Gen Div', 10: 'Gen Ex', 11: 'Marechal'}
ROT_MB = {1: 'GM', 2: '2º Ten', 3: '1º Ten', 4: 'CT', 5: 'CC', 6: 'CF',
          7: 'CMG', 8: 'CA', 9: 'VA', 10: 'AE', 11: 'Almte'}

POSTOS_VALIDOS = ('Aspirante/Guarda-Marinha; 2º Ten; 1º Ten; Capitão/Cap-Ten; '
                  'Major/CC; TC/CF; Cel/CMG; Gen Bda/Brig/CA; Gen Div/Maj Brig/VA; '
                  'Gen Ex/Ten Brig/AE')


def parse_posto(texto):
    t = norm(texto)
    for rx, idx in POSTO_RULES:
        if re.match(rx, t):
            return idx
    raise ValueError(f'Posto não reconhecido: {texto!r}. Aceitos: {POSTOS_VALIDOS}')


def parse_forca(texto):
    t = norm(texto)
    if not t:
        return ''
    if 'exercito' in t or t.startswith('eb'):
        return 'EB'
    if t.startswith('mb') or 'marinha' in t:
        return 'MB'
    if 'aeronautica' in t or 'fab' in t or t.startswith('aer'):
        return 'Aer'
    raise ValueError(f'Força não reconhecida: {texto!r}. Use EB, MB ou Aer')


def rotulo_posto(idx, forca):
    return ROT_MB[idx] if forca == 'MB' else ROT_EB_FA[idx]


# ---------------------------------------------------------------------------
# Datas: "dd/mm", "dd/mm/aaaa", intervalos "dd/mm-DD/MM", sementes
# ---------------------------------------------------------------------------

_MES_POR_PREFIXO = {p: n for n, nomes in {
    1: ('jan',), 2: ('fev',), 3: ('mar',), 4: ('abr',), 5: ('mai',),
    6: ('jun',), 7: ('jul',), 8: ('ago',), 9: ('set',), 10: ('out',),
    11: ('nov',), 12: ('dez',),
}.items() for p in nomes}


def _dia_mes(tok):
    """Decompõe um token de data flexível em (dia, mês|None, ano|0).
    Aceita: '10', '10/10', '10/10/2026', '10out', '10 de outubro', '10/out'."""
    m = re.fullmatch(r'(\d{1,2})', tok)
    if m:
        return int(m.group(1)), None, 0
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?', tok)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    m = re.fullmatch(r'(\d{1,2})\s*(?:/|de\s+)?([a-zç]{3,})', tok)
    if m:
        mes = _MES_POR_PREFIXO.get(m.group(2)[:3])
        if mes:
            return int(m.group(1)), mes, 0
    return None


def _data_flex(tok, ano, mes, allow_prev_year=False):
    dm = _dia_mes(tok)
    if not dm:
        return None
    dia, mm, aa = dm
    if mm is None:
        mm = mes
    if not aa:
        aa = ano
        if allow_prev_year and mm > mes:
            aa -= 1
    try:
        return dt.date(aa, mm, dia)
    except ValueError:
        return None


def parse_datas(tok, ano, mes, allow_prev_year=False):
    """Lista de datas de um token: '07/09', '10/09-12/09', '10', '10out',
    '10 de outubro', '10-15', '10out-15out', '07/09 (Independência)',
    '20/12/2026-04/01/2027'."""
    tok = norm(tok).replace('–', '-').replace('—', '-')
    tok = re.sub(r'\(.*?\)', '', tok).strip()
    if '-' in tok:
        a, b = tok.split('-', 1)
        da = _data_flex(a.strip(), ano, mes, allow_prev_year)
        db = _data_flex(b.strip(), ano, mes, allow_prev_year)
        if da and db and db >= da:
            out, d = [], da
            while d <= db:
                out.append(d)
                d += dt.timedelta(days=1)
            return out
    d = _data_flex(tok, ano, mes, allow_prev_year)
    return [d] if d else []


def parse_semente(s):
    """Data da última escala: dd/mm/aaaa, dd/mm/aa ou aaaa-mm-dd."""
    s = str(s or '').strip()
    if not s:
        return None
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y'):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return 'invalido'


SEMANA = {'seg': 0, 'segunda': 0, 'segunda-feira': 0,
          'ter': 1, 'terca': 1, 'terca-feira': 1,
          'qua': 2, 'quarta': 2, 'quarta-feira': 2,
          'qui': 3, 'quinta': 3, 'quinta-feira': 3,
          'sex': 4, 'sexta': 4, 'sexta-feira': 4,
          'sab': 5, 'sabado': 5,
          'dom': 6, 'domingo': 6}
# Nota: numeral solto (ex.: '10') é dia do mês corrente; dias da semana usam
# palavras (seg, ter, qua, qui, sex, sab, dom).


def parse_indisponibilidade(texto, ano, mes):
    """Retorna (datas, dias_da_semana). Ex.: '10/09-15/09; dom; sex',
    '10out-15out', '10-15; 20' (dia do mês), '10 de outubro'."""
    datas, semana = set(), set()
    for tok in re.split(r'[;,\n]+| e ', norm(texto)):
        tok = tok.strip()
        if not tok:
            continue
        if tok in SEMANA:
            semana.add(SEMANA[tok])
            continue
        antes = len(datas)
        for d in parse_datas(tok, ano, mes):
            datas.add(d)
        if len(datas) == antes:
            raise ValueError(f'Indisponibilidade não reconhecida: {tok!r}. '
                             f'Use datas (10/09, 10out, 10 de outubro, 10), '
                             f'intervalos (10-15, 10/09-15/09) ou dias da semana '
                             f'(seg, ter, qua, qui, sex, sab, dom)')
    return datas, semana


# ---------------------------------------------------------------------------
# Tipo do dia
# ---------------------------------------------------------------------------

NAMES_SERIE = {'pr': 'PRETA', 've': 'VERMELHA'}


def tipo_dia(d, feriados):
    if d in feriados or d.weekday() >= 5:
        return 've'
    return 'pr'


# ---------------------------------------------------------------------------
# Configuração / entrada
# ---------------------------------------------------------------------------

CHAVES_CFG = ('titulo 1', 'titulo 2', 'mes (mm/aaaa)',
              'feriados (dd/mm; dd/mm... ou intervalos)')

CAB = ['NOME', 'POSTO', 'FORÇA', 'ESPECIALIDADE',
       'INDISPONIBILIDADE (afastamento: período; ex.: 10-15/09, 10out)',
       'ÚLT. ESCALA ÚTIL (dd/mm/aaaa)', 'ÚLT. ESCALA FDS/FERIADO (dd/mm/aaaa)',
       'FOLGA ÚTIL INICIAL', 'FOLGA FDS/FERIADO INICIAL',
       'ÚLT. ESCALA (dd/mm/aaaa)', 'DESEMPATE (menor = mais moderno)', 'OBSERVAÇÕES',
       'NÃO ESCALAR FIXO (dias/datas; vale todo mês; ex.: ter; qua; qui)']


def validar_mes(cfg, erros):
    mes_txt = cfg.get('mes (mm/aaaa)', '')
    if not mes_txt or '/' not in mes_txt:
        erros.append('Configuração "Mês (mm/aaaa)" ausente ou inválida.')
        return None
    try:
        mm, aa = mes_txt.strip().split('/')
        cfg['mes'] = (int(aa), int(mm))
        return cfg['mes']
    except ValueError:
        erros.append(f'Configuração "Mês (mm/aaaa)" inválida: {mes_txt!r}')
        return None


def montar_militar(linha, posto_idx, posto_txt, forca, nome, mid=0):
    """Constrói o dict padrão de militar a partir de campos validados.
    Eixos: 'nao_escalar' = indisponibilidade (afastamento, período — gera retorno);
    'nao_escalar_fixo' e 'nao_escalar_mensal' = dias em que não pode escalar
    (não geram retorno)."""
    return {'id': mid, 'linha': linha.get('linha', mid + 1), 'nome': nome,
            'posto': posto_idx, 'posto_txt': posto_txt, 'forca': forca,
            'espec': str(linha.get('espec') or '').strip(),
            'nao_escalar': str(linha.get('nao_escalar') or '').strip(),
            'nao_escalar_fixo': str(linha.get('nao_escalar_fixo') or '').strip(),
            'nao_escalar_mensal': str(linha.get('nao_escalar_mensal') or '').strip(),
            'ult_util': str(linha.get('ult_util') or '').strip(),
            'ult_verm': str(linha.get('ult_verm') or '').strip(),
            'folga_util': str(linha.get('folga_util') or '').strip(),
            'folga_fds': str(linha.get('folga_fds') or '').strip(),
            'ult_escala': str(linha.get('ult_escala') or '').strip(),
            'desempate': None, 'obs': str(linha.get('obs') or '').strip()}


def _itens_para_texto(item_list):
    """Feriados [{data, desc}] da UI -> texto legível para a aba Entrada:
    '07/09 (Independência); 08/09'."""
    partes = []
    for it in item_list or []:
        data = str(it.get('data') or '').strip()
        desc = str(it.get('desc') or '').strip()
        if data:
            partes.append(f'{data} ({desc})' if desc else data)
    return '; '.join(partes)


def entrada_json(cfg_raw, mil_raw):
    """Valida dados vindos da interface web (JSON) e devolve
    (cfg, militares, erros). cfg_raw: {titulo1, titulo2, mes, feriados}
    Feriados podem ser texto (formato antigo) ou listas com descrição."""
    erros = []
    cfg_raw = cfg_raw or {}
    fer_txt = cfg_raw.get('feriados')
    if not isinstance(fer_txt, str):
        fer_txt = _itens_para_texto(fer_txt)
    cfg = {
        'titulo 1': str(cfg_raw.get('titulo1') or 'ORGANIZAÇÃO MILITAR'),
        'titulo 2': str(cfg_raw.get('titulo2') or 'ESCALA DE SERVIÇO POR FOLGA'),
        'mes (mm/aaaa)': str(cfg_raw.get('mes') or '').strip(),
        'feriados (dd/mm; dd/mm... ou intervalos)': str(fer_txt or ''),
    }
    validar_mes(cfg, erros)
    militares = []
    for i, row in enumerate(mil_raw or [], 1):
        nome = str(row.get('nome') or '').strip()
        posto_txt = str(row.get('posto') or '').strip()
        forca_txt = str(row.get('forca') or '').strip()
        if not nome and not posto_txt and not forca_txt:
            continue  # linha vazia
        if not nome:
            erros.append(f'Militar {i}: nome obrigatório.')
            continue
        try:
            posto = parse_posto(posto_txt)
        except ValueError as e:
            erros.append(f'Militar {i} ({nome}): {e}')
            continue
        try:
            forca = parse_forca(forca_txt)
        except ValueError as e:
            erros.append(f'Militar {i} ({nome}): {e}')
            continue
        m = montar_militar(row, posto, posto_txt, forca, nome, mid=len(militares))
        desempate_txt = str(row.get('desempate') or '').strip()
        if desempate_txt:
            try:
                m['desempate'] = int(float(desempate_txt.replace(',', '.')))
            except ValueError:
                erros.append(f'Militar {i} ({nome}): DESEMPATE deve ser número '
                             f'(menor = mais moderno).')
                continue
        militares.append(m)
    if not militares and not erros:
        erros.append('Informe ao menos um militar.')
    return cfg, militares, erros


def ler_entrada(ws):
    """Lê a aba Entrada da planilha (fluxo CLI). Retorna (cfg, militares, aviso)."""
    cfg = {}
    militares = []
    erro = []
    aviso = []
    secao = 'cfg'
    cab_pos = None
    for r in range(1, ws.max_row + 1):
        a = norm(ws.cell(r, 1).value)
        if not a:
            if secao == 'mil':
                break  # fim da tabela MILITARES (mantenha-a sem linhas em branco no meio)
            continue
        if a in ('configuracao', 'militares'):
            secao = 'cfg' if a == 'configuracao' else 'mil'
            continue
        if secao == 'cfg':
            cfg[a] = tostr(ws.cell(r, 2).value)
            continue
        # secao 'mil' (militares)
        if cab_pos is None:
            cab_pos = r  # linha de cabeçalho
            continue
        nome = tostr(ws.cell(r, 1).value).strip()
        if not nome:
            break
        posto_txt = tostr(ws.cell(r, 2).value)
        try:
            posto = parse_posto(posto_txt)
        except ValueError as e:
            erro.append(f'Linha {r}: {e}')
            continue
        try:
            forca = parse_forca(tostr(ws.cell(r, 3).value))
        except ValueError as e:
            erro.append(f'Linha {r}: {e}')
            continue
        desempate_txt = tostr(ws.cell(r, 11).value).strip()
        desempate = None
        if desempate_txt:
            try:
                desempate = int(float(desempate_txt.replace(',', '.')))
            except ValueError:
                erro.append(f'Linha {r}: DESEMPATE deve ser número '
                            f'(menor = mais moderno). Recebido: {desempate_txt!r}')
                continue
        m = montar_militar({'espec': ws.cell(r, 4).value,
                            'nao_escalar': ws.cell(r, 5).value,
                            'nao_escalar_fixo': ws.cell(r, 13).value,
                            'ult_util': ws.cell(r, 6).value,
                            'ult_verm': ws.cell(r, 7).value,
                            'folga_util': ws.cell(r, 8).value,
                            'folga_fds': ws.cell(r, 9).value,
                            'ult_escala': ws.cell(r, 10).value,
                            'obs': ws.cell(r, 12).value},
                           posto, posto_txt, forca, nome, mid=len(militares))
        m['desempate'] = desempate
        militares.append(m)
    validar_mes(cfg, erro)
    if erro:
        raise ValueError('\n'.join(erro))
    return cfg, militares, aviso


def config_datas(cfg):
    """Feriados a partir da config (chaves normed), guardando as descrições
    por data em cfg['_feriados_desc']."""
    feriados = set()
    desc_fer = {}
    label = 'feriados (dd/mm; dd/mm... ou intervalos)'
    if label in cfg and cfg[label]:
        ano, mes = cfg['mes']
        for tok in re.split(r'[;,\n]+', str(cfg[label])):
            tok = tok.strip()
            if not tok:
                continue
            mdesc = re.search(r'\(\s*([^)]+?)\s*\)', tok)
            desc = mdesc.group(1).strip() if mdesc else ''
            for d in parse_datas(tok, ano, mes):
                if d.year == ano and d.month == mes and d not in feriados:
                    feriados.add(d)
                    if desc:
                        desc_fer[d] = desc
    cfg['_feriados'] = feriados
    cfg['_feriados_desc'] = desc_fer
    return feriados


# ---------------------------------------------------------------------------
# Cômputo da escala
# ---------------------------------------------------------------------------

def computar(cfg, militares, feriados):
    aviso = []
    ano, mes = cfg['mes']
    primeiro = dt.date(ano, mes, 1)
    ultimo = (dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)) if mes < 12 \
        else dt.date(ano, 12, 31)
    dias = []
    d = primeiro
    while d <= ultimo:
        dias.append(d)
        d += dt.timedelta(days=1)

    def conta_serie(serie, de, ate):
        n, x = 0, de
        while x <= ate:
            if tipo_dia(x, feriados) == serie:
                n += 1
            x += dt.timedelta(days=1)
        return n

    for m in militares:
        def parse_restricao(texto, rotulo):
            try:
                return parse_indisponibilidade(texto, ano, mes)
            except ValueError as e:
                aviso.append(f'{m["nome"]}: {e} — {rotulo} ignorada')
                return (set(), set())

        m['indis'] = parse_restricao(m['nao_escalar'], 'indisponibilidade')
        ne_fixo = parse_restricao(m['nao_escalar_fixo'], 'não escalar fixo')
        ne_mes = parse_restricao(m['nao_escalar_mensal'], 'não escalar')
        m['ne'] = (ne_fixo[0] | ne_mes[0], ne_fixo[1] | ne_mes[1])
        # Regra do retorno: findo um período de indisponibilidade (datas) em E,
        # a apresentação ocorre no 1º dia útil após E que NÃO seja dia de "não
        # escalar"; nesse dia o militar não escala — liberado a partir do seguinte.
        datas_ind, semana_ind = m['indis']
        ne_datas, ne_semana = m['ne']
        retornos = set()
        ord_datas = sorted(datas_ind)
        for i, e in enumerate(ord_datas):
            if i + 1 < len(ord_datas) and ord_datas[i + 1] == e + dt.timedelta(days=1):
                continue  # não é o fim do bloco
            a = e + dt.timedelta(days=1)
            while a <= ultimo and (tipo_dia(a, feriados) != 'pr'
                                   or a in ne_datas or a.weekday() in ne_semana):
                a += dt.timedelta(days=1)
            if a <= ultimo:
                retornos.add(a)
        m['retornos'] = retornos
        cont = {'pr': 0, 've': 0}
        m['ult'] = None  # última escala (qualquer série), p/ regra do intervalo
        if m.get('ult_escala'):
            sd = parse_semente(m['ult_escala'])
            if sd == 'invalido':
                aviso.append(f'{m["nome"]}: última escala ignorada '
                             f'(formato inválido: {m["ult_escala"]!r})')
            elif sd is not None:
                m['ult'] = sd
        # sementes legadas por SÉRIE (data da última escala da série)
        for chave, serie in (('ult_util', 'pr'), ('ult_verm', 've')):
            t = m[chave]
            if not t:
                continue
            sd = parse_semente(t)
            if sd == 'invalido':
                aviso.append(f'{m["nome"]}: última escala ignorada '
                             f'(formato inválido: {t!r})')
                continue
            if sd is None:
                continue
            if sd >= primeiro or sd < dt.date(ano - 1, mes, 1):
                aviso.append(f'{m["nome"]}: última escala {t} fora do período '
                             f'válido — ignorada')
                continue
            cont[serie] = conta_serie(serie, sd + dt.timedelta(days=1),
                                      primeiro - dt.timedelta(days=1))
            if m['ult'] is None or sd > m['ult']:
                m['ult'] = sd
        # folgas iniciais NUMÉRICAS (herança do mês anterior — prevalecem)
        for serie, campo in (('pr', 'folga_util'), ('ve', 'folga_fds')):
            f = str(m.get(campo) or '').strip()
            if not f:
                continue
            try:
                cont[serie] = int(float(f))
            except ValueError:
                aviso.append(f'{m["nome"]}: folga inicial {campo}={f!r} '
                             f'inválida — ignorada')
        m['cont'] = cont
        # Novato: entrou sem folga herdada e sem última escala (nenhuma semente).
        # Folga 0 = "nunca escalou": tem estreia garantida na PRETA E na VERMELHA
        # (1ª oportunidade elegível de cada série, d+0, sem retorno), respeitando o
        # intervalo mínimo entre escalas. Depois segue todas as regras.
        m['estreia'] = (not m['folga_util'] and not m['folga_fds']
                        and not m['ult_escala'] and not m['ult_util']
                        and not m['ult_verm'])
        m['estreou'] = {'pr': False, 've': False}

    def indisponivel(m, dia):
        datas, semana = m['indis']
        ne_datas, ne_semana = m['ne']
        return (dia in datas or dia.weekday() in semana
                or dia in ne_datas or dia.weekday() in ne_semana
                or dia in m.get('retornos', ()))

    def bloq_escala(m, dia):
        """Dias em que o militar não pode escalar (indisponibilidade, não escalar
        ou dia da semana) — exibidos como IND na grade."""
        datas, semana = m['indis']
        ne_datas, ne_semana = m['ne']
        return (dia in datas or dia.weekday() in semana
                or dia in ne_datas or dia.weekday() in ne_semana)

    def elegiveis(dia):
        out = []
        for m in militares:
            if indisponivel(m, dia):
                continue
            if m['ult'] is not None and (dia - m['ult']).days < 3:
                continue
            out.append(m)
        return out

    grade = {}  # valor exibido por (militar, dia): int | 'D' | 'IND'
    escalados = {}
    for dia in dias:
        serie = tipo_dia(dia, feriados)
        relaxado = None
        pend_estreia = []
        if serie in ('pr', 've'):
            pend_estreia = [m for m in militares
                            if m['estreia'] and not m['estreou'][serie]
                            and not indisponivel(m, dia)
                            and (m['ult'] is None or (dia - m['ult']).days >= 3)]
        if pend_estreia:
            escolhido = min(pend_estreia, key=lambda m: (
                m['posto'],
                m['desempate'] if m['desempate'] is not None else 10 ** 9,
                norm(m['nome'])))
            escolhido['estreou'][serie] = True
        else:
            cand = elegiveis(dia)
            if not cand:
                cand = [m for m in militares if not indisponivel(m, dia)]
                relaxado = 'intervalo'
            if not cand:
                cand = list(militares)  # ninguém disponível — desconsidera indisponibilidade
                relaxado = 'indisponibilidade'
            escolhido = min(cand, key=lambda m: (
                -m['cont'][serie], m['posto'],
                m['desempate'] if m['desempate'] is not None else 10 ** 9,
                norm(m['nome'])))
        escolhido['ult'] = dia
        for m in militares:
            if m is escolhido:
                m['cont'][serie] = 0
            else:
                m['cont'][serie] += 1
        escalados[dia] = escolhido
        if relaxado:
            motivo = ('sem candidatos elegíveis — regra do intervalo (2 dias) relaxada'
                      if relaxado == 'intervalo'
                      else 'NENHUM militar disponível — indisponibilidade desconsiderada')
            aviso.append(f'{dia:%d/%m/%Y}: {motivo} — escalado {escolhido["nome"]}')
        for m in militares:
            if m is escolhido:
                grade[(m['id'], dia)] = 'D'
            elif bloq_escala(m, dia):
                grade[(m['id'], dia)] = 'IND'
            else:
                grade[(m['id'], dia)] = m['cont'][serie]
    return dias, escalados, grade, aviso


# ---------------------------------------------------------------------------
# Visão JSON para a interface web (pré-visualização)
# ---------------------------------------------------------------------------

def nota_dia(d, cfg):
    if d in cfg.get('_feriados_desc', {}):
        return 'Feriado: ' + cfg['_feriados_desc'][d]
    return ''


def visao(cfg, dias, escalados, militares, grade, aviso):
    feriados = cfg.get('_feriados', set())
    ano, mes = cfg['mes']
    return {
        'mes': f'{mes:02d}/{ano}',
        'dias': [{'data': d.strftime('%d/%m/%Y'),
                  'semana': SEMANAS[d.weekday()],
                  'escala': NAMES_SERIE[tipo_dia(d, feriados)],
                  'militar': nome_exibicao(escalados[d]),
                  'nota': nota_dia(d, cfg)} for d in dias],
        'grade': [{'militar': nome_exibicao(m),
                   'celulas': [grade[(m['id'], d)] for d in dias]} for m in militares],
        'avisos': aviso,
    }


def visao_para(cfg, militares):
    feriados = config_datas(cfg)
    dias, escalados, grade, aviso = computar(cfg, militares, feriados)
    return visao(cfg, dias, escalados, ordenar_militares(militares), grade, aviso)


# ---------------------------------------------------------------------------
# Renderização
# ---------------------------------------------------------------------------

THIN = Side(style='thin', color='BFBFBF')
BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TIT1 = Font(bold=True, size=14, color='1F3864')
F_TIT2 = Font(bold=True, size=12, color='1F3864')
F_MES = Font(bold=True, size=11, color='1F3864')
F_HDR = Font(bold=True, size=10)
F_HDRV = Font(bold=True, size=10, color='9C0006')
F_CEL = Font(size=9)
F_D = Font(bold=True, size=9, color='FFFFFF')
F_IND = Font(size=9, color='7F6000')
F_NOME = Font(bold=True, size=8, color='375623')
F_OBS = Font(size=9, italic=True, color='595959')
F_AVISO = Font(size=9, color='9C0006')
CEN = Alignment(horizontal='center', vertical='center')
ESQ = Alignment(horizontal='left', vertical='center')
FILL_VERM = PatternFill(fill_type='solid', fgColor='FFC7CE')
FILL_CINZA = PatternFill(fill_type='solid', fgColor='E7E6E6')
FILL_D = PatternFill(fill_type='solid', fgColor='2F5597')
FILL_IND = PatternFill(fill_type='solid', fgColor='FFEB9C')
FILL_NOME = PatternFill(fill_type='solid', fgColor='C6EFCE')
FILL_SAB = PatternFill(fill_type='solid', fgColor='DDEBF7')
MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
SEMANAS = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira',
           'sexta-feira', 'sábado', 'domingo']


def nome_exibicao(m):
    partes = [rotulo_posto(m['posto'], m['forca'])]
    if m['espec']:
        partes.append(m['espec'])
    if m['forca']:
        partes.append(m['forca'])
    partes.append(m['nome'])
    return ' '.join(partes)


def chave_ordem(m):
    """Ordem de exibição: posto mais antigo primeiro e nome alfabético dentro do
    posto. Vale só para a lista (Quadro, Entrada, Grade); não entra no cômputo."""
    return (-m['posto'], norm(m['nome']))


def ordenar_militares(militares):
    return sorted(militares, key=chave_ordem)


def render_entrada(wb, cfg, militares):
    """Aba Entrada preenchida com a configuração e os concorrentes (reprodutível via CLI)."""
    ws = wb.active
    ws.title = 'Entrada'
    ws.column_dimensions['A'].width = 46
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 26
    ws.column_dimensions['H'].width = 26
    ws.column_dimensions['I'].width = 30
    ws['A1'] = 'ENTRADA — ESCALA DE SERVIÇO POR FOLGA'
    ws['A1'].font = F_TIT2
    ws['A3'] = 'CONFIGURAÇÃO'
    ws['A3'].font = F_HDR
    cfg_rows = [
        ('Título 1', cfg.get('titulo 1', '')),
        ('Título 2', cfg.get('titulo 2', '')),
        ('Mês (mm/aaaa)', f'{cfg["mes"][1]:02d}/{cfg["mes"][0]}'),
        ('Feriados (dd/mm; dd/mm... ou intervalos)',
         cfg.get('feriados (dd/mm; dd/mm... ou intervalos)', '')),
    ]
    for i, (rot, val) in enumerate(cfg_rows, 4):
        ws.cell(i, 1, rot).font = F_CEL
        ws.cell(i, 2, val).font = F_CEL
        ws.cell(i, 2).border = BORD
    ws['A10'] = 'MILITARES'
    ws['A10'].font = F_HDR
    for c, h in enumerate(CAB, 1):
        cell = ws.cell(11, c, h)
        cell.font = F_HDR
        cell.fill = FILL_CINZA
        cell.border = BORD
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    r = 12
    for m in ordenar_militares(militares):
        valores = [m['nome'], m['posto_txt'], m['forca'], m['espec'], m['nao_escalar'],
                   m['ult_util'], m['ult_verm'],
                   m['folga_util'], m['folga_fds'], m['ult_escala'],
                   str(m['desempate']) if m['desempate'] is not None else '',
                   m['obs'], m['nao_escalar_fixo']]
        for c, v in enumerate(valores, 1):
            cell = ws.cell(r, c, v)
            cell.font = F_CEL
            cell.border = BORD
        r += 1
    ws.cell(r + 2, 1, 'Gerada pela interface web. Edite e regere no app, ou rode: '
                     'python3 escala.py entrada.xlsx').font = F_OBS


def render_diario(wb, cfg, dias, escalados):
    """Diária no padrão do arquivo antigo: títulos TNR 16 negrito mesclados,
    cabeçalho TNR 18 negrito, dados Arial 10, faixa azul no sábado."""
    ano, mes = cfg['mes']
    feriados = cfg.get('_feriados', set())
    ws = wb.create_sheet('Diário')
    f_tit = Font(name='Times New Roman', size=16, bold=True)
    f_hdr = Font(name='Times New Roman', size=18, bold=True)
    f_cel = Font(name='Arial', size=10)
    f_obs = Font(name='Arial', size=9, italic=True, color='595959')
    fill_sab = PatternFill(fill_type='solid', fgColor='A3DBFF')
    cena = Alignment(horizontal='center', vertical='center')
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 17
    ws.column_dimensions['C'].width = 54
    ws.column_dimensions['D'].width = 40
    titulos = [cfg.get('titulo 1', 'ORGANIZAÇÃO MILITAR'),
               cfg.get('titulo 2', 'ESCALA DE SERVIÇO POR FOLGA'),
               f'MÊS DE {MESES[mes - 1]} DE {ano}']
    for i, t in enumerate(titulos, 1):
        cell = ws.cell(i, 1, t)
        cell.font = f_tit
        cell.alignment = cena
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
    for c, h in enumerate(['DATA', 'DIA DA SEMANA', 'MILITAR DE SERVIÇO', 'OBSERVAÇÃO'], 1):
        cell = ws.cell(5, c, h)
        cell.font = f_hdr
        cell.alignment = cena
    r = 6
    for dia in dias:
        serie = tipo_dia(dia, feriados)
        dcell = ws.cell(r, 1, dia)
        dcell.number_format = 'dd/mm/yyyy'
        ws.cell(r, 2, SEMANAS[dia.weekday()])
        ws.cell(r, 3, nome_exibicao(escalados[dia]))
        ws.cell(r, 4, nota_dia(dia, cfg))
        for c in range(1, 5):
            cell = ws.cell(r, c)
            cell.font = f_obs if c == 4 else f_cel
            cell.alignment = cena
            if serie == 've' and dia.weekday() == 5:
                cell.fill = fill_sab
        r += 1
    ws.cell(r + 1, 1, f'Gerada em {dt.datetime.now():%d/%m/%Y %H:%M} — escala.py').font = f_obs
    ws.freeze_panes = 'A6'


def render_grade(wb, cfg, dias, militares, escalados, grade, aviso):
    ano, mes = cfg['mes']
    feriados = cfg.get('_feriados', set())
    ws = wb.create_sheet('Grade')
    ws.column_dimensions['A'].width = 46
    for c in range(2, 2 + len(dias)):
        ws.column_dimensions[get_column_letter(c)].width = 7.2
    a1 = ws.cell(1, 1, 'MILITAR')
    a1.font = Font(bold=True, size=9)
    a1.fill = FILL_NOME
    a1.alignment = CEN
    a1.border = BORD
    for i, dia in enumerate(dias, 2):
        serie = tipo_dia(dia, feriados)
        h1 = ws.cell(1, i, NAMES_SERIE[serie])
        h2 = ws.cell(2, i, dia.strftime('%d/%m'))
        for h in (h1, h2):
            h.border = BORD
            h.alignment = CEN
            if serie == 've':
                h.fill, h.font = FILL_VERM, F_HDRV
            else:
                h.fill, h.font = FILL_CINZA, F_HDR
    r = 3
    for m in ordenar_militares(militares):
        n = ws.cell(r, 1, nome_exibicao(m))
        n.font = F_NOME
        n.fill = FILL_NOME
        n.border = BORD
        n.alignment = ESQ
        for i, dia in enumerate(dias, 2):
            v = grade[(m['id'], dia)]
            serie = tipo_dia(dia, feriados)
            cell = ws.cell(r, i, v)
            cell.border = BORD
            cell.alignment = CEN
            if v == 'D':
                cell.fill, cell.font = FILL_D, F_D
            elif v == 'IND':
                cell.fill, cell.font = FILL_IND, F_IND
            else:
                cell.font = F_CEL
                if serie == 've':
                    cell.fill = FILL_VERM
        r += 1
    r += 1
    o = ws.cell(r, 1, 'OBS:')
    o.font = F_HDR
    o.alignment = ESQ
    linhas_obs = []
    if aviso:
        linhas_obs.append('AVISOS: ' + ' | '.join(aviso))
    legenda = ('LEGENDA: D = escalado | IND = indisponível | '
               'vermelho = fim de semana/feriado | '
               'amarelo = indisponibilidade | verde = nomes')
    linhas_obs.append(legenda)
    for m in militares:
        if m['obs']:
            linhas_obs.append(f'{m["nome"]}: {m["obs"]}')
    r += 1
    for texto in linhas_obs:
        ws.cell(r, 1, texto).font = F_AVISO if texto.startswith('AVISOS') else F_OBS
        ws.cell(r, 1).alignment = ESQ
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2 + len(dias))
        r += 1
    ws.freeze_panes = 'C3'


def preencher_diario_grade(wb, cfg, militares):
    """Computa e (re)escreve as abas Diário e Grade no workbook dado.
    Retorna (aviso, dias, escalados, grade)."""
    feriados = config_datas(cfg)
    dias, escalados, grade, aviso = computar(cfg, militares, feriados)
    for aba in ('Diário', 'Grade'):
        if aba in wb.sheetnames:
            del wb[aba]
    render_diario(wb, cfg, dias, escalados)
    render_grade(wb, cfg, dias, militares, escalados, grade, aviso)
    return aviso, dias, escalados, grade


def montar_workbook(cfg, militares):
    """Workbook completo (Entrada + Diário + Grade) para a interface web."""
    wb = Workbook()
    render_entrada(wb, cfg, militares)
    preencher_diario_grade(wb, cfg, militares)
    return wb


def salvar_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    # openpyxl 3.1.2 do Ubuntu grava "+00:00Z" no core.xml, que o Excel recusa
    # como conteúdo ilegível. O save ainda reescreve "modified" com fuso, então
    # a correção é feita no ZIP: só essa parte vira "...Z", o formato válido.
    origem = io.BytesIO(buf.getvalue())
    saida = io.BytesIO()
    with zipfile.ZipFile(origem) as zin, \
            zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                dados = dados.replace(b"+00:00Z", b"Z")
            zout.writestr(item, dados)
    return saida.getvalue()


# ---------------------------------------------------------------------------
# Importação de folgas do mês anterior (planilha do modelo antigo, aba "Escala")
# ---------------------------------------------------------------------------

def importar_folgas(caminho, ano, mes, militares):
    """Lê a grade do mês anterior e extrai, para cada militar do quadro, a folga do
    último dia útil, a do último fds/feriado e a última escala ('D'/'d'/0).

    `militares` aceita nomes soltos ou dicts {'nome', 'posto', 'forca'}. O posto é o
    desempate entre homônimos (Maj Farm EB FERNANDA x Cap Farm EB FERNANDA): sem ele
    os dois recebiam sempre a primeira linha da grade.

    Aceita dois formatos:
    1. Contínua do modelo antigo (aba 'Escala'): datas datetime na linha 1, posto/nome nas
       colunas B/C;
    2. Mensal gerada pelo app (aba 'Grade'): linha 2 com datas 'dd/mm', nomes
       completos (posto espec força NOME) na coluna A.

    Fim de semana é identificado por sáb/dom; feriado em dia útil no fim do mês
    anterior pode defasar a herança em 1 dia — conferir no painel.
    Retorna (resultados, extras, erro); `extras` são os militares que estão na
    planilha e não estão no quadro."""
    try:
        wb = load_workbook(caminho, data_only=True)
    except Exception as e:
        return [], [], f'Não consegui abrir a planilha: {e}'

    def abas_candidatas():
        pref = [a for a in ('Escala', 'Grade') if a in wb.sheetnames]
        if pref:
            return pref
        exc = {'Entrada', 'ORIENTAÇÃO', 'CONTROLE DE DISPENSAS',
               'DISP NATAL ANO NOVO', 'Planilha1'}
        return [a for a in wb.sheetnames if a not in exc] or list(wb.sheetnames)

    pedidos = [p for p in (_pedido_militar(m) for m in (militares or []))
               if p['nome']]
    for aba in abas_candidatas():
        ws = wb[aba]
        layout = _detectar_layout(ws, ano, mes)
        if layout:
            resultados, extras = _extrair_folgas(ws, layout, ano, mes, pedidos,
                                                 _catalogo_entrada(wb))
            return resultados, extras, None
    return [], [], _erro_mes_ausente(wb, ano, mes)


def _detectar_layout(ws, ano, mes):
    """Reconhece a grade e devolve {'cols': [(col, date)], 'series': {col: serie},
    'mensal': bool} ou None. 'series' só existe no formato mensal do app, onde a
    linha 1 diz PRETA/VERMELHA; sem ela o leitor cai no dia da semana."""
    # formato contínuo do modelo antigo: datas datetime na linha 1
    cols = []
    for c in range(1, min(ws.max_column, 2000) + 1):
        v = ws.cell(1, c).value
        if isinstance(v, dt.datetime) and v.year == ano and v.month == mes:
            cols.append((c, v.date()))
    if len(cols) >= 3:
        return {'cols': sorted(cols, key=lambda t: t[1]), 'mensal': False}
    # formato mensal do app: datas 'dd/mm' na linha 2 (cabeçalho: linha 1 = série)
    cols2 = []
    series = {}
    rotulos = {'PRETA': 'pr', 'VERMELHA': 've'}
    for c in range(1, min(ws.max_column, 60) + 1):
        v = str(ws.cell(2, c).value or '').strip()
        m = re.fullmatch(r'(\d{1,2})/(\d{1,2})', v)
        if not m:
            continue
        dia, mm = int(m.group(1)), int(m.group(2))
        if mm == mes:
            try:
                cols2.append((c, dt.date(ano, mm, dia)))
            except ValueError:
                continue
            serie = rotulos.get(str(ws.cell(1, c).value or '').strip().upper())
            if serie:
                series[c] = serie
    if len(cols2) >= 3:
        return {'cols': sorted(cols2, key=lambda t: t[1]), 'series': series,
                'mensal': True}
    return None


def _pedido_militar(m):
    """Item do quadro: aceita só o nome (uso antigo) ou dict com posto/força."""
    if isinstance(m, dict):
        return {'nome': str(m.get('nome') or '').strip(),
                'posto': str(m.get('posto') or '').strip(),
                'forca': str(m.get('forca') or '').strip()}
    return {'nome': str(m or '').strip(), 'posto': '', 'forca': ''}


_FORCA_POR_TOKEN = {'eb': 'EB', 'mb': 'MB', 'aer': 'Aer', 'fab': 'Aer'}


def _partes_rotulo(texto):
    """Quebra 'Cap Farm EB FERNANDA' em posto/espec/força/nome (rótulo da Grade)."""
    toks = str(texto or '').split()
    if not toks:
        return None
    for i in range(1, min(len(toks) - 2, 3) + 1):
        forca = _FORCA_POR_TOKEN.get(norm(toks[i]))
        if forca:
            return {'posto': toks[0], 'espec': ' '.join(toks[1:i]),
                    'forca': forca, 'nome': ' '.join(toks[i + 1:])}
    return {'posto': toks[0], 'espec': '', 'forca': '', 'nome': ' '.join(toks[1:])}


def _rotulos_aceitos(posto_txt, forca_txt):
    """Rótulos que o posto do quadro pode ter na grade ('Maj' -> {'maj'}; sem força
    declarada, {'maj', 'cc'})."""
    try:
        idx = parse_posto(posto_txt)
    except ValueError:
        return set()
    try:
        forca = parse_forca(forca_txt)
    except ValueError:
        forca = ''
    if forca:
        return {norm(rotulo_posto(idx, forca))}
    return {norm(ROT_EB_FA[idx]), norm(ROT_MB[idx])}


def _casa_nome(linha, chave):
    c = norm(chave)
    if not c:
        return False
    for alvo in (norm(linha['nome']), norm(linha['rotulo'])):
        if alvo == c or alvo.endswith(' ' + c):
            return True
    return False


def _pontos(linha, pedido):
    """Quanto a linha da grade combina com o militar do quadro."""
    pts = 0
    rots = _rotulos_aceitos(pedido['posto'], pedido['forca'])
    if rots:
        pts += 6 if norm(linha['posto']) in rots else -6
    if pedido['forca'] and linha['forca']:
        try:
            pts += 3 if parse_forca(pedido['forca']) == linha['forca'] else -3
        except ValueError:
            pass
    if norm(linha['nome']) == norm(pedido['nome']):
        pts += 1
    return pts


def _linhas_grade(ws, layout):
    """Militares da grade, um dict por linha, com posto/força já separados."""
    linhas = []
    for r in range(3 if layout['mensal'] else 2, min(ws.max_row, 300) + 1):
        if layout['mensal']:
            rotulo = str(ws.cell(r, 1).value or '').strip()
            n = norm(rotulo)
            if not rotulo or n in ('militar', 'obs', 'obs:') or n.startswith('legenda'):
                continue
            p = _partes_rotulo(rotulo)
            if not p or not p['nome']:
                continue
        else:
            nome = str(ws.cell(r, 3).value or '').strip()
            col_a = str(ws.cell(r, 1).value or '').strip()
            if nome:
                posto = str(ws.cell(r, 2).value or '').strip()
                p = {'posto': posto, 'espec': '', 'forca': '', 'nome': nome}
                rotulo = col_a or f'{posto} {nome}'.strip()
            elif col_a:
                p = _partes_rotulo(col_a)
                rotulo = col_a
                if not p or not p['nome']:
                    continue
            else:
                continue
        p.update({'r': r, 'rotulo': rotulo})
        linhas.append(p)
    return linhas


def _folgas_da_linha(ws, r, cols, series=None):
    """(folga da PRETA, folga da VERMELHA, data da última escala).

    A série de cada coluna vem do cabeçalho da Grade ('series'); só quando ele não
    existe (planilha antiga, sem cabeçalho de série) é que o sábado/domingo
    serve de critério."""
    folga = {'pr': None, 've': None}
    ult_escala = None
    for c, dd in cols:
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.strip().upper() in ('D', 'DD'):
            v = 'D'
        if v in (0, '0', 'D', 'd'):
            ult_escala = dd
            valor = 0
        elif isinstance(v, str) and v.strip().upper() == 'IND':
            continue
        else:
            try:
                valor = int(float(str(v).replace(',', '.')))
            except (TypeError, ValueError):
                continue
        serie = (series or {}).get(c) or ('ve' if dd.weekday() >= 5 else 'pr')
        if serie not in folga:
            continue
        folga[serie] = valor
    return folga['pr'], folga['ve'], ult_escala


def _prefixo(linha):
    """Posto como aparece na grade, sem o nome ('Cap Farm EB')."""
    rotulo, nome = linha['rotulo'].strip(), linha['nome'].strip()
    if nome and rotulo.upper().endswith(' ' + nome.upper()):
        return rotulo[:-(len(nome) + 1)].strip()
    return ' '.join(x for x in (linha['posto'], linha['espec'], linha['forca']) if x)


def _catalogo_entrada(wb):
    """Nome/posto/força/especialidade exatos da aba Entrada (planilha do app)."""
    if 'Entrada' not in wb.sheetnames:
        return []
    ws = wb['Entrada']
    cab = None
    itens = []
    for r in range(1, min(ws.max_row, 300) + 1):
        if cab is None:
            if norm(ws.cell(r, 1).value) == 'nome':
                cab = r
            continue
        nome = str(ws.cell(r, 1).value or '').strip()
        if not nome:
            break
        itens.append({'nome': nome,
                      'posto': str(ws.cell(r, 2).value or '').strip(),
                      'forca': str(ws.cell(r, 3).value or '').strip(),
                      'espec': str(ws.cell(r, 4).value or '').strip()})
    return itens


def _do_catalogo(catalogo, linha):
    """Dados do militar da grade conforme a aba Entrada; na falta, o que der para ler
    do próprio rótulo."""
    for it in catalogo:
        if norm(it['nome']) != norm(linha['nome']):
            continue
        rots = _rotulos_aceitos(it['posto'], it['forca'])
        if rots and norm(linha['posto']) not in rots:
            continue
        return it
    return {'nome': linha['nome'], 'posto': linha['posto'],
            'forca': linha['forca'], 'espec': linha['espec']}


def _mes_da_entrada(wb):
    """(ano, mês) declarado na aba Entrada; (0, 0) quando não houver."""
    if 'Entrada' not in wb.sheetnames:
        return 0, 0
    ws = wb['Entrada']
    for r in range(1, min(ws.max_row, 40) + 1):
        if norm(ws.cell(r, 1).value).startswith('mes'):
            m = re.fullmatch(r'(\d{1,2})/(\d{4})', tostr(ws.cell(r, 2).value).strip())
            if m:
                return int(m.group(2)), int(m.group(1))
    return 0, 0


def _meses_na_planilha(wb):
    """Meses que a planilha enviada contém, para dizer o que ela é de fato."""
    achados = set()
    ano_ent, mes_ent = _mes_da_entrada(wb)
    if mes_ent:
        achados.add((ano_ent, mes_ent))
    for aba in wb.sheetnames:
        ws = wb[aba]
        for c in range(1, min(ws.max_column, 2000) + 1):
            v = ws.cell(1, c).value
            if isinstance(v, dt.datetime):
                achados.add((v.year, v.month))
            if c > 60 or not ano_ent:
                continue
            m = re.fullmatch(r'(\d{1,2})/(\d{1,2})',
                             str(ws.cell(2, c).value or '').strip())
            if m:
                achados.add((ano_ent, int(m.group(2))))
    return sorted(achados)


def _erro_mes_ausente(wb, ano, mes):
    pedido = f'{mes:02d}/{ano}'
    achados = _meses_na_planilha(wb)
    if not achados:
        return (f'Não encontrei grade de datas nesta planilha (procurei {pedido} nas '
                f'abas Escala/Grade).')
    a1, m1 = achados[-1]
    prox_m = (m1 % 12) + 1
    prox_a = a1 + (1 if prox_m == 1 else 0)
    lista = ', '.join(f'{m:02d}/{a}' for a, m in achados)
    return (f'A planilha enviada é de {lista} e eu procurei {pedido}, o mês anterior '
            f'ao que está em edição. Para herdar as folgas dela, o mês em edição '
            f'precisa ser {prox_m:02d}/{prox_a}.')


def _extrair_folgas(ws, layout, ano, mes, pedidos, catalogo):
    cols = layout['cols']
    series = layout.get('series')
    linhas = _linhas_grade(ws, layout)
    usadas = {}
    resultados = []
    for p in pedidos:
        cands = [l for l in linhas if _casa_nome(l, p['nome'])]
        livres = [l for l in cands if l['r'] not in usadas]
        if not livres:
            motivo = ''
            if cands:
                motivo = ('a única linha com esse nome ("%s") já foi usada por %s'
                          % (cands[0]['rotulo'], usadas[cands[0]['r']]))
            resultados.append({'nome': p['nome'], 'encontrado': False,
                               'motivo': motivo})
            continue
        alvo = max(livres, key=lambda l: _pontos(l, p))
        usadas[alvo['r']] = p['nome']
        fu, ff, ue = _folgas_da_linha(ws, alvo['r'], cols, series)
        resultados.append({'nome': p['nome'], 'encontrado': True,
                           'posto_grade': _prefixo(alvo),
                           'folga_util': fu, 'folga_fds': ff,
                           'ult_escala': ue.strftime('%d/%m/%Y') if ue else '',
                           'sem_registro': fu is None and ff is None
                           and ue is None})
    extras = []
    for l in linhas:
        if l['r'] in usadas:
            continue
        fu, ff, ue = _folgas_da_linha(ws, l['r'], cols, series)
        cat = _do_catalogo(catalogo, l)
        extras.append({'nome': cat['nome'], 'posto': cat['posto'],
                       'forca': cat['forca'], 'espec': cat['espec'],
                       'posto_grade': _prefixo(l), 'folga_util': fu,
                       'folga_fds': ff,
                       'ult_escala': ue.strftime('%d/%m/%Y') if ue else '',
                       'sem_registro': fu is None and ff is None
                       and ue is None})
    return resultados, extras


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def processar(arquivo):
    wb = load_workbook(arquivo)
    if 'Entrada' not in wb.sheetnames:
        raise ValueError(f'Planilha {arquivo} não tem a aba "Entrada". '
                         f'Gere o modelo com: escala.py --template {arquivo}')
    ws = wb['Entrada']
    cfg, militares, _ = ler_entrada(ws)
    aviso, dias, escalados, _ = preencher_diario_grade(wb, cfg, militares)
    wb.save(arquivo)

    feriados = cfg['_feriados']
    n_pr = sum(1 for d in dias if tipo_dia(d, feriados) == 'pr')
    n_ve = sum(1 for d in dias if tipo_dia(d, feriados) == 've')
    freq = {}
    for m in militares:
        freq[m['nome']] = sum(1 for d in dias if escalados[d] is m)
    print(f'Mês: {cfg["mes"][1]:02d}/{cfg["mes"][0]} — {len(dias)} dias '
          f'(PRETA {n_pr}, VERMELHA {n_ve})')
    print(f'Militares: {len(militares)}')
    print('Escalações por militar: ' + ', '.join(
        f'{nome} {q}' for nome, q in sorted(freq.items())))
    if aviso:
        print('AVISOS:')
        for a in aviso:
            print(f'  - {a}')
    print(f'OK — {arquivo} atualizado (abas Diário e Grade).')
    return aviso


def criar_template(arquivo):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Entrada'
    ws.column_dimensions['A'].width = 46
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 24
    ws.column_dimensions['G'].width = 28
    ws.column_dimensions['H'].width = 26
    ws.column_dimensions['I'].width = 30
    ws['A1'] = 'ENTRADA — ESCALA DE SERVIÇO POR FOLGA'
    ws['A1'].font = F_TIT2
    ws['A3'] = 'CONFIGURAÇÃO'
    ws['A3'].font = F_HDR
    cfg_rows = [
        ('Título 1', 'ORGANIZAÇÃO MILITAR'),
        ('Título 2', 'ESCALA DE SERVIÇO POR FOLGA'),
        ('Mês (mm/aaaa)', '09/2026'),
        ('Feriados (dd/mm; dd/mm... ou intervalos)', '07/09'),
    ]
    for i, (rot, val) in enumerate(cfg_rows, 4):
        ws.cell(i, 1, rot).font = F_CEL
        ws.cell(i, 2, val).font = F_CEL
        ws.cell(i, 2).border = BORD
    ws['A10'] = 'MILITARES'
    ws['A10'].font = F_HDR
    for c, h in enumerate(CAB, 1):
        cell = ws.cell(11, c, h)
        cell.font = F_HDR
        cell.fill = FILL_CINZA
        cell.border = BORD
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    exemplo = ['FULANO DE TAL (EXEMPLO — APAGAR)', 'Cap', 'EB', 'Med',
               '10-15/09', '20/08/2026', '16/08/2026', '', '', '',
               '2', 'exemplo — apagar antes de usar', 'ter; qua']
    for c, v in enumerate(exemplo, 1):
        cell = ws.cell(12, c, v)
        cell.font = F_CEL
        cell.border = BORD
    nota = (
        'COMO USAR: edite a CONFIGURAÇÃO (título e mês; feriados do mês, '
        'separados por ; ou ,) e a tabela MILITARES (mantenha as linhas contíguas, '
        'sem linhas em branco no meio). POSTO: Asp, 2º Ten, 1º Ten, Cap/Cap-Ten, '
        'Maj/CC, TC/CF, Cel/CMG, Gen Bda/Brig/CA, Gen Div/Maj Brig/VA, '
        'Gen Ex/Ten Brig/AE. NÃO ESCALAR: datas (10/09), intervalos (10/09-15/09) e '
        'dias da semana (dom, sex...). ÚLT. ESCALA ÚTIL/FDS-FERIADO: data da última '
        'escala do mês anterior (semeia a folga e respeita o intervalo mínimo). '
        'DESEMPATE: número menor = mais moderno (mesmo posto). '
        'Depois rode: python3 escala.py entrada.xlsx — o script (re)cria as abas '
        'Diário e Grade nesta mesma planilha. Se um dia não tiver ninguém '
        'elegível, o script relaxa a regra e registra AVISO na Grade — '
        'nunca deixa dia sem escalado.'
    )
    ws.cell(14, 1, nota).font = F_OBS
    ws.merge_cells('A14:I20')
    ws['A14'].alignment = Alignment(wrap_text=True, vertical='top')
    wb.save(arquivo)
    print(f'OK — modelo {arquivo} criado (aba Entrada).')


if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == '--template':
        criar_template(sys.argv[2] if len(sys.argv) > 2 else 'entrada.xlsx')
    else:
        arquivo = sys.argv[1] if len(sys.argv) > 1 else 'entrada.xlsx'
        try:
            processar(arquivo)
        except ValueError as e:
            print('ERRO:', e, file=sys.stderr)
            sys.exit(1)