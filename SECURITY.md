# Segurança

Este app não tem login. O controle de acesso é a lista de IPs `PERMITIR`,
mais a validação do `Host` (loopback, IP local e `HOSTS`). Por padrão, ele
escuta apenas em `127.0.0.1`, no próprio computador.

## Uso recomendado

- Rode no seu computador, em `127.0.0.1`, para uso individual.
- Para uso em equipe, rode em um servidor da rede local e libere apenas os IPs
  autorizados em `PERMITIR`.
- Não exponha o app diretamente na internet sem autenticação ou um proxy com
  login na frente.
- `POST` exige `Content-Type: application/json`; outro tipo recebe 415.
- Para um proxy reverso com outro nome, defina `HOSTS=nome.exemplo` (lista
  separada por vírgula). Sem isso, só loopback e o IP local são aceitos.
- Os dados ficam no `localStorage` do navegador. Não há banco de dados no
  servidor.
- Não coloque `.env`, tokens, senhas ou dados pessoais no repositório.

## Reportar um problema

Abra uma issue no repositório descrevendo o problema e como reproduzir. Não
inclua dados pessoais, nomes reais ou capturas com informação sensível.
