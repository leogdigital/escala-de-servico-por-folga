#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testes do servidor HTTP (app.py): portão de IP e upload da planilha.

Rode:  .venv/bin/python teste_servidor.py
"""

import base64
import http.client
import io
import json
import os
import shutil
import tempfile
import threading
import unittest

from http.server import ThreadingHTTPServer
from openpyxl import load_workbook

import app as servidor
from escala import entrada_json, montar_workbook, salvar_bytes

CFG_SEP = {'mes': '09/2026', 'feriados': []}
MIL = [
    {'nome': 'A', 'posto': 'Cap', 'forca': 'EB',
     'folga_util': '21', 'folga_fds': '8', 'ult_escala': '25/09/2026'},
    {'nome': 'B', 'posto': 'Maj', 'forca': 'EB',
     'folga_util': '3', 'folga_fds': '4', 'ult_escala': '28/09/2026'},
]


def planilha_setembro():
    """Gera a planilha do mês anterior (setembro) no formato do próprio app."""
    cfg, mils, erros = entrada_json(CFG_SEP, MIL)
    assert not erros, erros
    return salvar_bytes(montar_workbook(cfg, mils))


class ServidorTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        servidor.Handler.log_message = lambda *a, **k: None   # silencia o log de acesso
        cls.srv = ThreadingHTTPServer(('127.0.0.1', 0), servidor.Handler)
        cls.porta = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        os.environ['PERMITIR'] = '127.0.0.1'

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        os.environ.pop('PERMITIR', None)

    def pedir(self, metodo, caminho, corpo=None):
        c = http.client.HTTPConnection('127.0.0.1', self.porta, timeout=15)
        cabecalhos = {}
        dados = None
        if corpo is not None:
            dados = json.dumps(corpo).encode('utf-8')
            cabecalhos['Content-Type'] = 'application/json'
        c.request(metodo, caminho, body=dados, headers=cabecalhos)
        r = c.getresponse()
        texto = r.read()
        c.close()
        return r.status, r.getheader('Content-Type') or '', texto

    # ---------------- portão de IP ----------------

    def test_bloqueia_ip_fora_da_lista(self):
        os.environ['PERMITIR'] = '203.0.113.10'
        try:
            status, ctype, corpo = self.pedir('GET', '/')
            self.assertEqual(status, 403)
            self.assertIn('text/html', ctype)
            self.assertIn('Acesso restrito', corpo.decode('utf-8'))
            self.assertIn('127.0.0.1', corpo.decode('utf-8'))   # mostra o IP de quem tentou
            # a API fica atrás do mesmo portão
            status, _, _ = self.pedir('POST', '/api/gerar',
                                      {'config': CFG_SEP, 'militares': MIL})
            self.assertEqual(status, 403)
            status, _, _ = self.pedir('GET', '/static/app.js')
            self.assertEqual(status, 403)
        finally:
            os.environ['PERMITIR'] = '127.0.0.1'

    def test_permite_ip_da_lista(self):
        status, ctype, corpo = self.pedir('GET', '/')
        self.assertEqual(status, 200)
        self.assertIn('text/html', ctype)
        self.assertIn('Escala de Serviço por Folga'.encode('utf-8'), corpo)
        status, ctype, _ = self.pedir('GET', '/static/app.js')
        self.assertEqual(status, 200)
        self.assertIn('javascript', ctype)

    def test_recusa_content_type_texto(self):
        c = http.client.HTTPConnection('127.0.0.1', self.porta, timeout=15)
        c.request('POST', '/api/preview', body=b'{}',
                  headers={'Content-Type': 'text/plain'})
        r = c.getresponse()
        texto = r.read()
        c.close()
        self.assertEqual(r.status, 415)
        self.assertIn('application/json', texto.decode('utf-8'))

    def test_recusa_host_desconhecido(self):
        c = http.client.HTTPConnection('127.0.0.1', self.porta, timeout=15)
        c.request('GET', '/', headers={'Host': 'attacker.example'})
        r = c.getresponse()
        texto = r.read()
        c.close()
        self.assertEqual(r.status, 403)
        self.assertIn('Acesso restrito', texto.decode('utf-8'))

    def test_content_length_negativo(self):
        import socket
        s = socket.create_connection(('127.0.0.1', self.porta), timeout=15)
        s.sendall(b'POST /api/preview HTTP/1.1\r\nHost: 127.0.0.1\r\n'
                  b'Content-Type: application/json\r\n'
                  b'Content-Length: -1\r\n\r\n{}')
        data = s.recv(4096)
        s.close()
        self.assertIn(b'413', data)

    def test_corpo_json_nao_objeto(self):
        status, _, _ = self.pedir('POST', '/api/preview', [])
        self.assertEqual(status, 400)

    def test_config_nao_objeto(self):
        status, _, _ = self.pedir('POST', '/api/preview',
                                  {'config': 'x', 'militares': MIL})
        self.assertEqual(status, 400)

    def test_mes_invalido(self):
        status, _, _ = self.pedir('POST', '/api/preview', {
            'config': {'mes': '13/2026', 'feriados': []},
            'militares': MIL})
        self.assertEqual(status, 400)

    def test_intervalo_de_datas_gigante(self):
        status, _, _ = self.pedir('POST', '/api/preview', {
            'config': {'mes': '09/2026',
                       'feriados': [{'data': '01/01/1900-31/12/2026',
                                     'desc': ''}]},
            'militares': MIL})
        self.assertEqual(status, 400)

    def test_muitos_militares(self):
        mils = [{'nome': f'M{i}', 'posto': 'Cap', 'forca': 'EB'}
                for i in range(1001)]
        status, _, _ = self.pedir('POST', '/api/preview',
                                  {'config': CFG_SEP, 'militares': mils})
        self.assertEqual(status, 400)

    def test_desempate_infinito(self):
        status, _, _ = self.pedir('POST', '/api/preview', {
            'config': CFG_SEP,
            'militares': [{'nome': 'A', 'posto': 'Cap', 'forca': 'EB',
                           'desempate': float('inf')}]})
        self.assertEqual(status, 400)

    # ---------------- geração ----------------

    def test_gerar_devolve_xlsx_valido(self):
        status, ctype, corpo = self.pedir('POST', '/api/gerar',
                                          {'config': CFG_SEP, 'militares': MIL})
        self.assertEqual(status, 200)
        self.assertIn('spreadsheetml', ctype)
        self.assertTrue(corpo.startswith(b'PK'))
        wb = load_workbook(io.BytesIO(corpo))
        for aba in ('Entrada', 'Diário', 'Grade'):
            self.assertIn(aba, wb.sheetnames)

    # ---------------- upload da planilha do mês anterior ----------------

    def test_importar_por_upload(self):
        dados = planilha_setembro()
        nomes = [m['nome'] for m in MIL]
        corpo = {'mes_ref': '09/2026', 'militares': nomes}

        status, _, resp = self.pedir('POST', '/api/importar', dict(
            corpo, arquivo_b64=base64.b64encode(dados).decode('ascii')))
        self.assertEqual(status, 200, resp[:200])
        resultados = json.loads(resp)['resultados']
        self.assertTrue(all(r['encontrado'] for r in resultados), resultados)
        for r in resultados:
            self.assertIsInstance(r['folga_util'], int)

    def test_importar_sem_upload_devolve_400(self):
        status, _, corpo = self.pedir('POST', '/api/importar',
                                      {'mes_ref': '09/2026',
                                       'militares': ['A'],
                                       'caminho': '/etc/passwd'})
        self.assertEqual(status, 400)
        self.assertIn('Escolha a planilha', json.loads(corpo)['erros'][0])

    def test_upload_ilegivel_devolve_400(self):
        status, _, corpo = self.pedir('POST', '/api/importar',
                                      {'mes_ref': '09/2026', 'militares': ['A'],
                                       'arquivo_b64': 'isto nao e base64###'})
        self.assertEqual(status, 400)
        self.assertIn('ilegível', json.loads(corpo)['erros'][0])

    def test_arquivo_que_nao_e_planilha_devolve_400(self):
        lixo = base64.b64encode(b'nao sou uma planilha').decode('ascii')
        status, _, corpo = self.pedir('POST', '/api/importar',
                                      {'mes_ref': '09/2026', 'militares': ['A'],
                                       'arquivo_b64': lixo})
        self.assertEqual(status, 400)
        self.assertTrue(json.loads(corpo)['erros'][0])

    def test_sem_arquivo_devolve_400(self):
        status, _, corpo = self.pedir('POST', '/api/importar',
                                      {'mes_ref': '09/2026', 'militares': ['A']})
        self.assertEqual(status, 400)
        self.assertIn('Escolha a planilha', json.loads(corpo)['erros'][0])


if __name__ == '__main__':
    unittest.main(verbosity=2)
