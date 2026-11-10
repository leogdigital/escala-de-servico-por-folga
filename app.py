#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor — Escala de Supervisor de Dia (interface web)

Uso:
  python3 app.py [--no-browser] [--host 127.0.0.1] [--port 8770]
                 [--permitir 127.0.0.1]

Abre a página em http://127.0.0.1:8770/ — insira a configuração e o quadro de
concorrentes, pré-visualize a escala e baixe a planilha .xlsx (abas Entrada +
Diário + Grade).

Acesso restrito por IP de origem: só os endereços em PERMITIR (variável de
ambiente ou --permitir) abrem a página. Qualquer outro cliente recebe 403 e
fica registrado no log.
"""

import base64
import json
import os
import sys
import tempfile
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from escala import (entrada_json, importar_folgas, montar_workbook,
                    salvar_bytes, visao_para)

BASE = os.path.dirname(os.path.abspath(__file__))
HOST = os.environ.get('HOST', '127.0.0.1')
PORT = int(os.environ.get('PORT', '8770'))
MAX_BODY = 5 * 1024 * 1024
PERMITIR_PADRAO = '127.0.0.1'


def permitidos():
    """IPs autorizados (lidos a cada requisição: facilita teste e ajuste)."""
    bruto = os.environ.get('PERMITIR', PERMITIR_PADRAO)
    return {p.strip() for p in bruto.split(',') if p.strip()}


PAGINA_403 = """<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="utf-8"><title>Acesso restrito</title>
<style>body{font:16px/1.5 -apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;
max-width:34em;margin:12vh auto;padding:0 1.2em;color:#23272e}
h1{font-size:1.25em}code{background:#f2f3f6;padding:2px 5px;border-radius:4px}</style></head>
<body>
<h1>Acesso restrito</h1>
<p>Esta página só abre a partir de um computador autorizado (o Mac mini da casa).</p>
<p>Se você está no Mac mini e leu isto, o endereço dele na rede pode ter mudado:
avise o Leo informando o número abaixo.</p>
<p>Endereço deste computador na rede: <code>{{IP}}</code></p>
</body></html>
"""


def pagina_403(ip):
    return PAGINA_403.replace('{{IP}}', ip).encode('utf-8')

STATIC = {
    '/': ('static/index.html', 'text/html; charset=utf-8'),
    '/static/app.css': ('static/app.css', 'text/css; charset=utf-8'),
    '/static/app.js': ('static/app.js', 'text/javascript; charset=utf-8'),
    '/static/vendor/flatpickr/flatpickr.min.js':
        ('static/vendor/flatpickr/flatpickr.min.js', 'text/javascript; charset=utf-8'),
    '/static/vendor/flatpickr/flatpickr.min.css':
        ('static/vendor/flatpickr/flatpickr.min.css', 'text/css; charset=utf-8'),
    '/static/vendor/flatpickr/pt.js':
        ('static/vendor/flatpickr/pt.js', 'text/javascript; charset=utf-8'),
}

XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


class Handler(BaseHTTPRequestHandler):
    def _begin(self, code, ctype, length):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(length))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()

    def _send(self, code, ctype, body):
        self._begin(code, ctype, len(body))
        self.wfile.write(body)

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self._send(code, 'application/json; charset=utf-8', body)

    def _404(self):
        self._send(404, 'text/plain; charset=utf-8', b'nao encontrado')

    def _bloqueado(self):
        """403 para quem não está na lista de IPs autorizados (fail-closed)."""
        ip = self.client_address[0]
        if ip in permitidos():
            return False
        sys.stderr.write('[app] BLOQUEADO %s -> %s\n' % (ip, self.path))
        self._send(403, 'text/html; charset=utf-8', pagina_403(ip))
        return True

    def do_GET(self):
        if self._bloqueado():
            return
        caminho = self.path.split('?', 1)[0]
        if caminho in STATIC:
            rel, ctype = STATIC[caminho]
            try:
                with open(os.path.join(BASE, rel), 'rb') as f:
                    self._send(200, ctype, f.read())
            except OSError:
                self._404()
            return
        self._404()

    def do_HEAD(self):
        if self._bloqueado():
            return
        caminho = self.path.split('?', 1)[0]
        if caminho in STATIC:
            rel, ctype = STATIC[caminho]
            try:
                with open(os.path.join(BASE, rel), 'rb') as f:
                    data = f.read()
            except OSError:
                return self._404()
            self._begin(200, ctype, len(data))
            return
        self._404()

    def do_POST(self):
        if self._bloqueado():
            return
        ctype = (self.headers.get('Content-Type') or '').split(';', 1)[0].strip().lower()
        if ctype != 'application/json':
            return self._json(415, {'erros': ['Content-Type deve ser application/json.']})
        try:
            n = int(self.headers.get('Content-Length') or 0)
            if n > MAX_BODY:
                return self._json(413, {'erros': ['Requisição muito grande.']})
            raw = self.rfile.read(n) or b'{}'
            payload = json.loads(raw.decode('utf-8'))
        except Exception as e:
            return self._json(400, {'erros': [f'JSON inválido: {e}']})

        if self.path == '/api/importar':
            b64 = str(payload.get('arquivo_b64') or '')
            ref = str(payload.get('mes_ref') or '')
            militares = []
            for n in (payload.get('militares') or []):
                if not isinstance(n, dict):
                    n = {'nome': n}
                nome = str(n.get('nome') or '').strip()
                if nome:
                    militares.append({'nome': nome,
                                      'posto': str(n.get('posto') or '').strip(),
                                      'forca': str(n.get('forca') or '').strip()})
            try:
                mm, aa = ref.split('/')
                ano, mes = int(aa), int(mm)
            except Exception:
                return self._json(400, {'erros': ['Referência "mes_ref" inválida '
                                                  '(use mm/aaaa).']})
            # A planilha do mês anterior chega por upload: o servidor não enxerga
            # o disco de quem usa o app. Não aceite caminho de arquivo do servidor.
            if not b64:
                return self._json(400, {'erros': ['Escolha a planilha do mês '
                                                  'anterior (.xlsx).']})
            try:
                dados = base64.b64decode(b64, validate=True)
            except Exception:
                return self._json(400, {'erros': ['Planilha enviada ilegível.']})
            if not dados:
                return self._json(400, {'erros': ['Planilha enviada está vazia.']})
            fd, temporario = tempfile.mkstemp(prefix='escala-anterior-',
                                              suffix='.xlsx')
            with os.fdopen(fd, 'wb') as f:
                f.write(dados)
            try:
                resultados, extras, erro = importar_folgas(temporario, ano, mes,
                                                           militares)
            finally:
                try:
                    os.unlink(temporario)
                except OSError:
                    pass
            if erro:
                return self._json(400, {'erros': [erro]})
            return self._json(200, {'resultados': resultados, 'extras': extras})

        if self.path in ('/api/preview', '/api/gerar'):
            mil_raw = payload.get('militares')
            if not isinstance(mil_raw, list) or any(
                    not isinstance(x, dict) for x in mil_raw):
                return self._json(400, {'erros': ['Lista "militares" inválida.']})
            cfg, militares, erros = entrada_json(payload.get('config'), mil_raw)
            if erros:
                return self._json(400, {'erros': erros})
            if self.path == '/api/preview':
                return self._json(200, visao_para(cfg, militares))
            wb = montar_workbook(cfg, militares)
            data = salvar_bytes(wb)
            ano, mes = cfg['mes']
            nome = f'escala_{mes:02d}_{ano}.xlsx'
            self.send_response(200)
            self.send_header('Content-Type', XLSX_TYPE)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Content-Disposition',
                             f'attachment; filename="{nome}"')
            self.end_headers()
            self.wfile.write(data)
            return

        self._404()

    def log_message(self, fmt, *args):
        sys.stderr.write('[app %s] %s\n' % (self.address_string(), fmt % args))


def ip_local():
    try:
        import ipaddress
        import socket
        host = socket.gethostname()
        for ip in socket.gethostbyname_ex(host)[2]:
            try:
                if ipaddress.ip_address(ip).is_private:
                    return ip
            except ValueError:
                continue
    except Exception:
        pass
    return None


def main():
    global HOST, PORT
    args = list(sys.argv[1:])
    abrir = '--no-browser' not in args
    for i, a in enumerate(args):
        if a == '--port' and i + 1 < len(args):
            PORT = int(args[i + 1])
        elif a == '--host' and i + 1 < len(args):
            HOST = args[i + 1]
        elif a == '--permitir' and i + 1 < len(args):
            os.environ['PERMITIR'] = args[i + 1]
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as e:
        if getattr(e, 'errno', None) == 48 or 'address already in use' in str(e).lower():
            print(f'ERRO: a porta {PORT} já está em uso — provavelmente outra '
                  f'instância do app.py está rodando. Encerre-a ou use '
                  f'python3 app.py --port {PORT + 1}', file=sys.stderr)
            sys.exit(2)
        raise
    url = f'http://127.0.0.1:{PORT}/'
    print(f'Escala de Serviço por Folga — servidor em {HOST}:{PORT}')
    print(f'Acesso local: {url}')
    ip = ip_local()
    if ip:
        print(f'Acesso pela rede (Wi-Fi): http://{ip}:{PORT}/')
    else:
        print('Acesso pela rede local: use o IP desta máquina na mesma rede Wi-Fi.')
    print('IPs autorizados: ' + ', '.join(sorted(permitidos())))
    print('Para encerrar: Ctrl+C')
    if abrir:
        webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\nEncerrado.')


if __name__ == '__main__':
    main()