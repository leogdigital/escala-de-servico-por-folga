#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Teste de regressão da escala: constrói entrada de exemplo, processa e
valida as regras contra uma referência independente (reimplementação).
Uso: .venv/bin/python teste.py   (cria/remove entrada_teste.xlsx sozinho)"""

import datetime as dt
import os
import sys

from openpyxl import Workbook, load_workbook

sys.path.insert(0, '.')
from escala import (NAMES_SERIE, SEMANAS, ler_entrada, nome_exibicao,  # noqa: E402
                    parse_indisponibilidade, processar, tipo_dia)

ANO, MES = 2026, 9
FERIADOS = {dt.date(ANO, MES, 7)}                               # 07/09 — Independência

LINHAS = [
    # nome, posto, força, espec, nao escalar, ult util, ult verm, desempate, obs
    ('ALFA', 'TC', 'EB', 'Med', '', '20/08/2026', '16/08/2026', '', ''),
    ('BRAVO', 'CF', 'MB', '', '', '03/08/2026', '10/08/2026', '', ''),
    ('CHARLIE', 'Maj', 'Aer', 'Farm', 'dom; sex', '30/08/2026', '', '', ''),
    ('DELTA', 'Cap', 'EB', 'QCO', '10/09-15/09', '', '15/08/2026', '2', ''),
    ('ECHO', 'Cap-Ten', 'MB', 'Med', '', '20/08/2026', '25/08/2026', '5', ''),
    ('FOXTROT', 'Cel', 'EB', 'Dent', '', '05/08/2026', '02/08/2026', '', ''),
]


def constroi_entrada(path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Entrada'
    cfg = [
        ('Título 1', 'ORGANIZAÇÃO MILITAR'),
        ('Título 2', 'ESCALA DE SERVIÇO POR FOLGA'),
        ('Mês (mm/aaaa)', '09/2026'),
        ('Feriados (dd/mm; dd/mm... ou intervalos)', '07/09'),
    ]
    r = 1
    for rot, val in cfg:
        ws.cell(r, 1, rot)
        ws.cell(r, 2, val)
        r += 1
    ws.cell(r, 1, 'MILITARES')
    r += 1
    cab = ['NOME', 'POSTO', 'FORÇA', 'ESPECIALIDADE', 'INDISPONIBILIDADE',
           'ÚLT. ESCALA ÚTIL', 'ÚLT. ESCALA FDS/FERIADO',
           'FOLGA ÚTIL INICIAL', 'FOLGA FDS/FERIADO INICIAL', 'ÚLT. ESCALA',
           'DESEMPATE', 'OBSERVAÇÕES', 'NÃO ESCALAR FIXO']
    for c, h in enumerate(cab, 1):
        ws.cell(r, c, h)
    r += 1
    for (nome, posto, forca, espec, nao, uu, uv, emp, obs) in LINHAS:
        valores = [nome, posto, forca, espec, nao, uu, uv, '', '', '', emp, obs, '']
        for c, v in enumerate(valores, 1):
            ws.cell(r, c, v)
        r += 1
    wb.save(path)


def dias_do_mes():
    out = []
    d = dt.date(ANO, MES, 1)
    while d.month == MES:
        out.append(d)
        d += dt.timedelta(days=1)
    return out


def nome_simples(nome_coluna):
    for nome in (linha[0] for linha in LINHAS):
        if nome_coluna == nome or nome_coluna.endswith(' ' + nome):
            return nome
    w = nome_coluna.split()
    return ' '.join(w[-2:]) if len(w) >= 2 else nome_coluna


def ref_contadores(m, dias, escalado_nome, dia_atual):
    """Contador de referência no INÍCIO do dia (pré-seleção):
    semeia pela última escala, reseta no dia da escala e incrementa 1 a cada
    dia ANTERIOR da própria série."""
    cont = {'pr': 0, 've': 0}
    primero = dias[0]
    for chave, serie in (('ult_util', 'pr'), ('ult_verm', 've')):
        if not m[chave]:
            continue
        sd = dt.datetime.strptime(m[chave], '%d/%m/%Y').date()
        d = sd + dt.timedelta(days=1)
        while d < primero:
            if tipo_dia(d, FERIADOS) == serie:
                cont[serie] += 1
            d += dt.timedelta(days=1)
    for d in dias:
        if d >= dia_atual:
            break
        s = tipo_dia(d, FERIADOS)
        if escalado_nome[d] == m['nome']:
            cont[s] = 0
        else:
            cont[s] += 1
    return cont


def main():
    path = 'entrada_teste.xlsx'
    constroi_entrada(path)
    avisos = processar(path)
    wb = load_workbook(path)
    dias = dias_do_mes()

    # ---- Diário (novo formato: DATA datetime, sem coluna ESCALA) ----
    di = wb['Diário']
    rows = []
    r = 6
    while di.cell(r, 1).value:
        rows.append(tuple(di.cell(r, c).value for c in range(1, 5)))
        r += 1
    assert len(rows) == len(dias), f'Diário com {len(rows)} dias, esperado {len(dias)}'
    for i, (data, dsem, milit, obs) in enumerate(rows):
        d = dias[i]
        assert isinstance(data, dt.datetime) and data.date() == d, \
            f'{data!r} != {d:%d/%m/%Y}'
        assert dsem == SEMANAS[d.weekday()], f'{dsem} != {SEMANAS[d.weekday()]}'
        assert milit, 'militar vazio'
    # fontes do padrão antigo
    assert di.cell(1, 1).font.name == 'Times New Roman' and di.cell(1, 1).font.size == 16
    assert di.cell(5, 1).font.name == 'Times New Roman' and di.cell(5, 1).font.size == 18
    assert di.cell(6, 3).font.name == 'Arial' and di.cell(6, 3).font.size == 10
    assert di.cell(1, 1).alignment.horizontal == 'center'

    # ---- Grade: cabeçalho, D único por dia ----
    g = wb['Grade']
    escalado_nome = {}
    for i, d in enumerate(dias, 2):
        col = g.cell(1, i).value
        assert col == NAMES_SERIE[tipo_dia(d, FERIADOS)], f'{col} em {d:%d/%m}'
        assert g.cell(2, i).value == d.strftime('%d/%m')
        quem = []
        for rl in range(3, 3 + len(LINHAS)):
            if g.cell(rl, i).value == 'D':
                quem.append(nome_simples(g.cell(rl, 1).value))
        assert len(quem) == 1, f'{d:%d/%m}: {len(quem)} escalados → {quem}'
        escalado_nome[d] = quem[0]

    # gap >= 3 dias para cada militar
    for rl in range(3, 3 + len(LINHAS)):
        nome = nome_simples(g.cell(rl, 1).value)
        datas_d = [d for i, d in enumerate(dias, 2) if g.cell(rl, i).value == 'D']
        assert len(datas_d) >= 1, f'{nome} nunca escalado'
        for a, b in zip(datas_d, datas_d[1:]):
            assert (b - a).days >= 3, \
                f'{nome}: escalado {a:%d/%m} e {b:%d/%m} (gap {(b - a).days} < 3)'

    # indisponibilidades: CHARLIE (dom; sex), DELTA (10-15/09)
    for rl in range(3, 3 + len(LINHAS)):
        nome = nome_simples(g.cell(rl, 1).value)
        for i, d in enumerate(dias, 2):
            v = g.cell(rl, i).value
            if nome == 'CHARLIE' and d.weekday() in (4, 6):
                assert v in ('IND', 'D'), f'CHARLIE {d:%d/%m} = {v}'
                if v == 'D':
                    raise AssertionError(f'CHARLIE escalada em dia restrito {d:%d/%m}')
            if nome == 'DELTA' and dt.date(ANO, MES, 10) <= d <= dt.date(ANO, MES, 15):
                assert v in ('IND', 'D'), f'DELTA {d:%d/%m} = {v}'
                if v == 'D':
                    raise AssertionError(f'DELTA escalado em 10-15/09 ({d:%d/%m})')

    entrada = ler_entrada(wb['Entrada'])[1]
    por_coluna = {nome_exibicao(m): m for m in entrada}
    indis = {m['nome']: parse_indisponibilidade(m['nao_escalar'], ANO, MES)
             for m in entrada}

    def retornos_de(datas):
        ords = sorted(datas)
        fins = [e for i, e in enumerate(ords)
                if i + 1 >= len(ords) or ords[i + 1] != e + dt.timedelta(days=1)]
        rets = set()
        for e in fins:
            a = e + dt.timedelta(days=1)
            while a <= dias[-1] and tipo_dia(a, FERIADOS) != 'pr':
                a += dt.timedelta(days=1)
            if a <= dias[-1]:
                rets.add(a)
        return rets

    retornos = {m['nome']: retornos_de(indis[m['nome']][0]) for m in entrada}

    # contadores exibidos == referência (células numéricas)
    erros_cont = 0
    for rl in range(3, 3 + len(LINHAS)):
        m = por_coluna[g.cell(rl, 1).value]
        for i, d in enumerate(dias, 2):
            v = g.cell(rl, i).value
            if isinstance(v, int):
                serie = tipo_dia(d, FERIADOS)
                ref = ref_contadores(m, dias, escalado_nome, d)[serie] + 1
                if v != ref:
                    erros_cont += 1
                    print(f'  cont divergiu: {m["nome"]} {d:%d/%m} = {v}, ref {ref}')
    assert erros_cont == 0, f'{erros_cont} células com contador divergente'

    # escolha de cada dia == argmax folga entre elegíveis (referência)
    erros_escolha = 0
    for d in dias:
        serie = tipo_dia(d, FERIADOS)
        melhor, melhor_chave = None, None
        for m in entrada:
            datas, semana = indis[m['nome']]
            if d in datas or d.weekday() in semana or d in retornos[m['nome']]:
                continue
            ult = None
            for d2 in dias:
                if d2 >= d:
                    break
                if escalado_nome[d2] == m['nome']:
                    ult = d2
            if ult is not None and (d - ult).days < 3:
                continue
            ref = ref_contadores(m, dias, escalado_nome, d)[serie]
            chave = (-ref, m['posto'],
                     m['desempate'] if m['desempate'] is not None else 10 ** 9,
                     m['nome'])
            if melhor_chave is None or chave < melhor_chave:
                melhor, melhor_chave = m, chave
        if melhor is None:
            continue  # sem elegíveis → regra relaxada (coberto pelo aviso)
        if escalado_nome[d] != melhor['nome']:
            erros_escolha += 1
            print(f'  escolha divergiu em {d:%d/%m}: grade={escalado_nome[d]}, '
                  f'ref={melhor["nome"]}')
    assert erros_escolha == 0, f'{erros_escolha} dias com escolha divergente'

    # regra relaxada não deve ter sido disparada (sempre houve elegível)
    assert not avisos, f'avisos inesperados: {avisos}'

    print(f'PASS — {len(dias)} dias, {len(entrada)} militares.')
    for d in dias:
        print(f'  {d:%d/%m} {NAMES_SERIE[tipo_dia(d, FERIADOS)]:9s} '
              f'→ {escalado_nome[d]}')

    # ---- caminho do template: CONFIGURAÇÃO + linha em branco + nota ----
    from escala import criar_template

    tpl = 'entrada_template_teste.xlsx'
    criar_template(tpl)
    processar(tpl)
    wb2 = load_workbook(tpl)
    assert wb2.sheetnames == ['Entrada', 'Diário', 'Grade'], wb2.sheetnames
    assert len(ler_entrada(wb2['Entrada'])[1]) == 1  # só a linha de exemplo
    os.remove(tpl)
    print('PASS template: CONFIGURAÇÃO/nota/linha em branco lidas; 1 militar (exemplo).')

    # ---- formato estruturado da UI: feriados com descrição ----
    from escala import entrada_json, parse_datas, visao_para

    cfg, mils, erros = entrada_json(
        {'mes': '09/2026',
         'feriados': [{'data': '07/09', 'desc': 'Independência do Brasil'},
                      {'data': '08/09', 'desc': ''}]},
        [{'nome': 'X Y', 'posto': 'Cap', 'forca': 'EB'}])
    assert not erros, erros
    v = visao_para(cfg, mils)
    notas = {d['data']: d['nota'] for d in v['dias']}
    assert notas['07/09/2026'] == 'Feriado: Independência do Brasil', notas
    assert notas['08/09/2026'] == '', notas
    assert v['dias'][6]['escala'] == 'VERMELHA'
    assert parse_datas('07/09 (Independência)', 2026, 9) == [dt.date(2026, 9, 7)]
    # formatos flexíveis
    assert parse_datas('10', 2026, 10) == [dt.date(2026, 10, 10)]
    assert parse_datas('10out', 2026, 10) == [dt.date(2026, 10, 10)]
    assert parse_datas('10 de outubro', 2026, 10) == [dt.date(2026, 10, 10)]
    assert parse_datas('3 de setembro', 2026, 10) == [dt.date(2026, 9, 3)]
    assert parse_datas('10-15', 2026, 10) == [dt.date(2026, 10, d) for d in range(10, 16)]
    assert parse_datas('10out-15out', 2026, 10) == \
        [dt.date(2026, 10, d) for d in range(10, 16)]
    assert parse_datas('25/12', 2026, 1, allow_prev_year=True) == [dt.date(2025, 12, 25)]
    # indisponibilidade com formatos novos e palavra fora de ordem
    ds, sm = parse_indisponibilidade('10out-15out; dom; 3', 2026, 10)
    assert ds == {dt.date(2026, 10, d) for d in range(10, 16)} | {dt.date(2026, 10, 3)}
    assert sm == {6}
    try:
        parse_indisponibilidade('xyz', 2026, 10)
        raise AssertionError('deveria rejeitar token inválido')
    except ValueError:
        pass
    print('PASS formatos flexíveis de data.')

    # ---- folgas iniciais (herança do mês anterior) ----
    from escala import entrada_json, visao_para

    # A..F com última escala em 25/09 (gap ok); G escalado em 30/09 (gap bloqueia 01-02/10)
    dados_folga = [
        ('A', '21', '8', '25/09/2026'), ('B', '100', '0', '25/09/2026'),
        ('C', '30', '1', '25/09/2026'), ('D', '29', '2', '25/09/2026'),
        ('E', '28', '3', '25/09/2026'), ('F', '27', '4', '25/09/2026'),
        ('G', '200', '200', '30/09/2026'),
    ]
    cfg, mils, erros = entrada_json(
        {'mes': '10/2026', 'feriados': []},
        [{'nome': n, 'posto': 'Cap', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': pu, 'folga_fds': pv, 'ult_escala': ue}
         for (n, pu, pv, ue) in dados_folga])
    assert not erros, erros
    v = visao_para(cfg, mils)
    g = {m['militar']: m['celulas'] for m in v['grade']}
    fu = g['Cap EB A']
    gx = g['Cap EB G']
    assert v['dias'][0]['data'] == '01/10/2026' and v['dias'][0]['escala'] == 'PRETA'
    assert fu[0] == 22 and fu[1] == 23, fu[:4]           # 21 herdada + dias úteis
    assert v['dias'][2]['escala'] == 'VERMELHA' and fu[2] == 9, fu[:4]  # 8 herdada + sáb
    assert v['dias'][0]['militar'] == 'Cap EB B'         # G (200) não vence: gap
    assert v['dias'][1]['militar'] == 'Cap EB C'
    assert v['dias'][2]['militar'] == 'Cap EB G'         # 03/10: gap liberado (3 dias)
    assert gx[0] == 201 and gx[1] == 202 and gx[2] == 'D', gx[:4]
    print('PASS folgas iniciais (herança + gap).')

    # ---- importar_folgas (workbook fake no padrão da aba "Escala") ----
    from escala import importar_folgas
    from openpyxl import Workbook

    fake = 'fake_grade.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Escala'
    ini = dt.datetime(2026, 9, 1)
    for i in range(30):
        ws.cell(1, 1 + i, ini + dt.timedelta(days=i))
    ws['A2'] = 'prefixo X'
    ws['B2'] = 'Cap Med EB'
    ws['C2'] = 'FULANO'
    ws.cell(2, 26, 8)    # 26/09 sáb
    ws.cell(2, 27, 'd')  # 27/09 dom — escalado
    ws.cell(2, 29, 20)   # 29/09 ter
    ws.cell(2, 30, 21)   # 30/09 qua
    wb.save(fake)
    res, extras, err = importar_folgas(fake, 2026, 9, ['FULANO', 'ZÉ NINGUÉM'])
    assert not err, err
    r0 = res[0]
    assert r0['encontrado'] and r0['folga_util'] == 21 and r0['folga_fds'] == 0 \
        and r0['ult_escala'] == '27/09/2026', r0
    assert not res[1]['encontrado']
    assert extras == [], extras
    os.remove(fake)
    print('PASS importação do mês anterior.')

# ---- regra do retorno (períodos/datas; 1º dia útil seguinte bloqueado) ----
    # A afastado 08-11/10 (fim 11/10, dom). 12/10 é feriado → 13/10 é o 1º dia útil
    # (apresentação, inelegível); a partir de 14/10 A pode escalar.
    cfg, mils, erros = entrada_json(
        {'mes': '10/2026', 'feriados': [{'data': '12/10', 'desc': 'Padroeira'}]},
        [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB', 'nao_escalar': '08-11',
          'folga_util': '200', 'folga_fds': '200'},
         {'nome': 'B', 'posto': 'Maj', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': '1', 'folga_fds': '1'}])
    assert not erros, erros
    v = visao_para(cfg, mils)
    dm = {d['data']: d['militar'] for d in v['dias']}
    assert dm['13/10/2026'] != 'Cap EB A', dm['13/10/2026']  # inelegível no retorno
    assert dm['14/10/2026'] == 'Cap EB A', dm['14/10/2026']  # liberado no d+1 útil
    print('PASS regra do retorno.')

    # ---- roundtrip: arquivo gerado pelo app (Grade mensal) vira fonte ----
    from escala import montar_workbook

    cfg, mils, erros = entrada_json(
        {'mes': '10/2026', 'feriados': [{'data': '12/10', 'desc': 'Padroeira'}]},
        [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': '21', 'folga_fds': '8', 'ult_escala': '25/09/2026'},
         {'nome': 'A', 'posto': 'Maj', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': '30', 'folga_fds': '30'},
         {'nome': 'C', 'posto': 'Cap', 'forca': 'MB', 'nao_escalar': '',
          'folga_util': '5', 'folga_fds': '5'}])
    assert not erros, erros
    wb = montar_workbook(cfg, mils)
    arquivo = 'roundtrip.xlsx'
    wb.save(arquivo)
    quadro = [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB'},
              {'nome': 'A', 'posto': 'Maj', 'forca': 'EB'}]
    res, extras, err = importar_folgas(arquivo, 2026, 10, quadro)
    assert not err, err
    g = load_workbook(arquivo)['Grade']

    def achcol(dia):
        for c in range(1, g.max_column + 1):
            if str(g.cell(2, c).value or '').strip() == f'{dia:02d}/10':
                return c
        return None

    def achlinha(rotulo):
        return next(r for r in range(3, g.max_row + 1)
                    if str(g.cell(r, 1).value or '').strip() == rotulo)

    c30, c31 = achcol(30), achcol(31)
    rA = achlinha('Cap EB A')
    esp = lambda v: 0 if str(v).upper() in ('D', 'DD') else int(v)
    assert res[0]['encontrado'] and res[0]['folga_util'] == esp(g.cell(rA, c30).value) \
        and res[0]['folga_fds'] == esp(g.cell(rA, c31).value), (res, g.cell(rA, c30).value,
                                                                 g.cell(rA, c31).value)
    # homônimos: cada um na sua linha, desempatados pelo posto (bug do Cap x Maj)
    rM = achlinha('Maj EB A')
    assert res[0]['posto_grade'] == 'Cap EB' and res[1]['posto_grade'] == 'Maj EB', res
    assert res[1]['folga_util'] == esp(g.cell(rM, c30).value), (res[1],
                                                                g.cell(rM, c30).value)
    # quem está na planilha e não no quadro volta como extra, com posto e força
    assert [e['nome'] for e in extras] == ['C'], extras
    assert extras[0]['posto'] == 'Cap' and extras[0]['forca'] == 'MB', extras[0]
    assert extras[0]['folga_util'] == esp(g.cell(achlinha('CT MB C'), c30).value), extras[0]
    # mês ausente: o erro diz o que a planilha é e qual mês em edição a usa
    _, _, err = importar_folgas(arquivo, 2026, 9, quadro)
    assert err and '10/2026' in err and '11/2026' in err, err
    os.remove(arquivo)
    print('PASS roundtrip: Grade do app como fonte do mês seguinte.')

    # ---- três eixos: indisponível + não escalar fixo (exemplo "quinta") ----
    # A: férias (indisponível) 05-08/11 (fim dom 08/11) + NÃO ESCALAR fixo seg/ter.
    # Apresentação = 1º dia útil fora de não-escalar = qua 11/11 (inelegível);
    # escala a partir de qui 12/11.
    cfg, mils, erros = entrada_json(
        {'mes': '11/2026', 'feriados': [{'data': '15/11', 'desc': 'Proclamação'},
                                        {'data': '20/11', 'desc': 'Consciência Negra'}]},
        [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB', 'nao_escalar': '05-08',
          'nao_escalar_fixo': 'seg; ter', 'folga_util': '200', 'folga_fds': '200'},
         {'nome': 'B', 'posto': 'Maj', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': '1', 'folga_fds': '1'}])
    assert not erros, erros
    v = visao_para(cfg, mils)
    dm = {d['data']: d['militar'] for d in v['dias']}
    assert dm['09/11/2026'] != 'Cap EB A'      # seg: não escalar fixo
    assert dm['10/11/2026'] != 'Cap EB A'      # ter: não escalar fixo
    assert dm['11/11/2026'] != 'Cap EB A'      # qua: apresentação (inelegível)
    assert dm['12/11/2026'] == 'Cap EB A'      # qui: primeiro dia escalável = QUINTA
    # sem indisponibilidade: o não escalar fixo bloqueia sem gerar retorno (qua escala)
    cfg2, mils2, erros = entrada_json(
        {'mes': '11/2026', 'feriados': []},
        [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB', 'nao_escalar_fixo': 'seg; ter',
          'folga_util': '200', 'folga_fds': '200'},
         {'nome': 'B', 'posto': 'Maj', 'forca': 'EB', 'nao_escalar': '',
          'folga_util': '1', 'folga_fds': '1'}])
    v2 = visao_para(cfg2, mils2)
    dm2 = {d['data']: d['militar'] for d in v2['dias']}
    assert dm2['09/11/2026'] != 'Cap EB A' and dm2['10/11/2026'] != 'Cap EB A'
    assert dm2['11/11/2026'] == 'Cap EB A'     # qua: disponível, sem retorno
    print('PASS três eixos (indisponível / não escalar fixo) e retorno composto.')

    # ---- estreia de novatos (entram sem herança; folga 0 = escalar já) ----
    # Nov/2026 começa num domingo (vermelha). Seis veteranos com folga alta herdada
    # e dois novatos. Cada novato estreia em CADA série (vermelha e preta), o mais
    # moderno primeiro, respeitando o intervalo; depois a regra normal retoma.
    vet = [{'nome': f'VET{i}', 'posto': p, 'forca': 'EB',
            'folga_util': '25', 'folga_fds': '25'}
           for i, p in enumerate(['TC', 'CF', 'Maj', 'CC', 'TC', 'Maj'])]
    cfg, mils, erros = entrada_json(
        {'mes': '11/2026', 'feriados': []},
        vet + [{'nome': 'RENATA TC', 'posto': 'TC', 'forca': 'EB'},
               {'nome': 'CUNHA CAP', 'posto': 'Cap', 'forca': 'EB'}])
    assert not erros, erros
    v = visao_para(cfg, mils)
    dm = {d['data']: d['militar'] for d in v['dias']}
    g = {m['militar']: m['celulas'] for m in v['grade']}
    eh = (lambda x: x == 'D')
    cunha = g['Cap EB CUNHA CAP']
    renata = g['TC EB RENATA TC']
    dC = [v['dias'][i]['data'] for i, x in enumerate(cunha) if eh(x)]
    dR = [v['dias'][i]['data'] for i, x in enumerate(renata) if eh(x)]
    # estreia na VERMELHA (01/11 dom) e na PRETA (04/11 qua) da mais moderna
    assert dC[0] == '01/11/2026' and dC[1] == '04/11/2026', dC
    # a outra novata: PRETA em 02/11 e VERMELHA no 1º fds liberado (07/11)
    assert dR[0] == '02/11/2026' and dR[1] == '07/11/2026', dR
    # regra normal retoma para os veteranos nos demais dias iniciais
    assert 'VET' in dm['03/11/2026'] or 'VET' in dm['05/11/2026'] or \
        'VET' in dm['06/11/2026'], dm['03/11/2026']
    print('PASS estreia de novatos na PRETA e na VERMELHA (d+0, mais moderno primeiro).')

    # ---- ordem de exibição: posto mais antigo primeiro, nome alfabético ----
    from escala import computar, entrada_json, ordenar_militares

    linhas_ordem = [
        {'nome': 'ZECA', 'posto': 'Cap', 'forca': 'EB'},
        {'nome': 'ANA', 'posto': 'TC', 'forca': 'EB'},
        {'nome': 'BIA', 'posto': 'Maj', 'forca': 'EB'},
        {'nome': 'ALBERTO', 'posto': 'Cap', 'forca': 'EB'},
    ]
    cfg, mils, erros = entrada_json(
        {'mes': '11/2026', 'feriados': []}, linhas_ordem)
    assert not erros, erros
    ordem = [m['nome'] for m in ordenar_militares(mils)]
    assert ordem == ['ANA', 'BIA', 'ALBERTO', 'ZECA'], ordem
    # a ordem visual não pode mudar a escala
    cfg2, mils2, erros = entrada_json(
        {'mes': '11/2026', 'feriados': []}, linhas_ordem)
    d1, e1, _, _ = computar(cfg, list(mils), set())
    d2, e2, _, _ = computar(cfg2, ordenar_militares(mils2), set())
    assert [e1[d]['nome'] for d in d1] == [e2[d]['nome'] for d in d2]
    print('PASS ordem de exibição: posto e alfabética, sem alterar a escala.')

    os.remove(path)
    return 0


if __name__ == '__main__':
    sys.exit(main())