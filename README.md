# Escala de Serviço por Folga

Português | [English](README.en.md)

Uma aplicação local para montar a escala diária de serviço por folga. Roda em
Python, guarda os dados no navegador e gera uma planilha Excel. Serve para
organizações militares e para qualquer equipe que precise de uma pessoa por dia.

## Índice

1. O que é
2. Para quem serve
3. Como funciona
4. Requisitos
5. Início rápido
6. Rodar na rede com systemd
7. Usando o app
8. Dados e backup
9. Importar folgas do mês anterior
10. Planilhas geradas
11. Modo CLI
12. Possibilidades e melhorias
13. Testes
14. Docker (opcional)
15. Autoria e transparência
16. Segurança
17. Licença

## 1. O que é

É um servidor web local em Python. Você abre a página no navegador, cadastra o
quadro de concorrentes, informa as indisponibilidades e gera a escala do mês.

O app não tem banco de dados. O estado fica no `localStorage` do navegador, e a
planilha é baixada para o seu computador. Nada é enviado para a internet.

O quadro de demonstração vem com 5 nomes fictícios:

| Nome | Posto | Força | Especialidade |
|---|---|---|---|
| ALFA | TC | EB | Med |
| BRAVO | Maj | EB | Dent |
| CHARLIE | Cap | Aer | Med |
| DELTA | CT | MB | Farm |
| ECHO | Cap | EB | ENF |

Eles são apenas exemplo. Apague e coloque o seu quadro.

## 2. Para quem serve

Serve para uma escala de **uma pessoa por dia**, com hierarquia, folgas e
indisponibilidades. Exemplos:

- plantão de enfermaria, farmácia ou clínica;
- abertura e fechamento de comércio;
- serviço de dia em organização militar;
- plantão de portaria, zeladoria ou operação.

O app **não** faz vários turnos por dia, várias posições no mesmo dia ou escala
por hora. Para esses casos, veja a seção 12, Possibilidades e melhorias.

## 3. Como funciona

O motor trabalha com duas séries independentes:

- **PRETA**: dia útil.
- **VERMELHA**: sábado, domingo e feriado.

As regras principais:

1. Cada série tem o seu contador de dias de folga.
2. A cada dia, escolhe-se o concorrente **elegível com mais dias de folga** na
   série daquele dia.
3. É elegível quem não está indisponível, não está em dia de "não escalar", não
   está no dia de apresentação e está há pelo menos 3 dias sem servir.
4. No empate, decide a menor hierarquia (mais moderno), depois o campo
   **Desempate** (menor = mais moderno) e, por último, o nome.
5. No começo do mês, cada série herda o contador do último dia dela no mês
   anterior. O painel de herança permite ajustar.
6. Quem entra sem folga herdada é **novo**: tem estreia garantida na PRETA e na
   VERMELHA.
7. Se não houver ninguém elegível, o app relaxa a regra e registra um aviso.
   Nunca deixa um dia sem escalado.

## 4. Requisitos

- Python 3.10 ou superior.
- `openpyxl` e `defusedxml`, instalados pelo `requirements.txt`.
- Um navegador atualizado.

Não precisa de banco de dados, servidor web separado ou internet depois da
instalação.

## 5. Início rápido

### 1. Veja se o Python já está instalado

No Linux ou no macOS, copie e cole no terminal:

```
python3 --version
```

No Windows, copie e cole:

```
python --version
```

Se aparecer `Python 3.10` ou maior, vá para o passo 2.

Se aparecer uma versão menor, ou o comando não for encontrado, instale o
Python 3.10 ou superior em https://www.python.org/downloads/.

No Windows, marque a opção **Add Python to PATH** durante a instalação.
No Ubuntu ou Debian, também serve:

```
sudo apt install python3 python3-venv
```

No macOS com Homebrew:

```
brew install python
```

### 2. Clone o repositório

Troque `SEU_USUARIO` pelo seu usuário do GitHub. Copie e cole:

```
git clone https://github.com/SEU_USUARIO/escala-de-servico-por-folga.git
cd escala-de-servico-por-folga
```

### 3. Crie o ambiente virtual e instale a dependência

Copie e cole:

```
python3 -m venv .venv
```

Ative o ambiente.

No Linux ou no macOS:

```
source .venv/bin/activate
```

No Windows:

```
.venv\Scripts\activate
```

Depois instale:

```
pip install -r requirements.txt
```

### 4. Rode o app

No Linux ou no macOS:

```
python3 app.py
```

No Windows:

```
python app.py
```

### 5. Abra no navegador

Copie e cole no navegador:

```
http://127.0.0.1:8770/
```

Se o navegador não abrir sozinho, use o endereço acima. Para parar o app, volte
ao terminal e pressione Ctrl+C.

## 6. Rodar na rede com systemd

Se várias pessoas precisam acessar, rode o app em um servidor Linux.

1. Copie a pasta do projeto para o servidor, por exemplo em `/caminho/do/app`.
2. Edite `servico/escala.service.example` e troque:
   - `SEU_USUARIO` pelo usuário do servidor;
   - `/caminho/do/app` pelo caminho real;
   - `PERMITIR` pelos IPs autorizados.
3. Copie e cole no servidor:

```
sudo cp servico/escala.service.example /etc/systemd/system/escala.service
sudo systemctl daemon-reload
sudo systemctl enable --now escala
```

4. Confira:

```
systemctl status escala
```

Para acessar de outro computador da rede, troque no `ExecStart`:

```
--host 127.0.0.1
```

por:

```
--host 0.0.0.0
```

e coloque os IPs dos clientes em `PERMITIR`. Sem isso, o app só responde no
próprio servidor.

## 7. Usando o app

### Bloco 1: configuração do mês

| Campo ou botão | O que faz |
|---|---|
| Mês | Mês e ano que serão gerados |
| Feriados | Datas e descrição dos feriados do mês |
| Arquivo do mês anterior | Planilha do mês anterior para herdar as folgas |
| Carregar folgas | Lê o último dia de cada série e preenche a herança |
| Limpar | Limpa a herança carregada |
| Exportar estado | Baixa um JSON com toda a memória do navegador |
| Importar estado | Restaura a memória de um JSON |

### Bloco 2: quem concorre à escala

| Campo | O que faz |
|---|---|
| Indisponibilidade (período) | Afastamento por datas ou dias da semana. Gera retorno |
| Não escalar (mês) | Dias em que a pessoa não pode servir, mesmo presente. Não gera retorno |

A tabela mostra o posto, o nome, a disponibilidade e o "não escalar". Nada
preenchido significa disponível.

### Bloco 3: quadro de concorrentes

| Campo ou botão | O que faz |
|---|---|
| + Inserir militar | Adiciona uma linha nova |
| Restaurar concorrentes à escala | Recarrega os 5 nomes de exemplo |
| Nome | Nome de guerra ou nome curto |
| Posto | Hierarquia. Menor índice é mais moderno |
| Força | EB, MB, Aer ou vazio |
| Especialidade | Texto livre, como Med, Dent ou Farm |
| Não escalar (fixo) | Dias que valem todo mês, como `ter; qua; qui` |
| Desempate | Número. Menor é mais moderno dentro do mesmo posto |
| Observações | Texto que aparece no rodapé da Grade |
| Novo | Marca quem entrou sem folga herdada. Tem estreia garantida |
| Remover | Tira a pessoa do quadro e da memória |

### Botões do topo

| Botão | O que faz |
|---|---|
| Pré-visualizar escala | Calcula e mostra a grade e o diário na tela |
| Gerar planilha .xlsx | Baixa a planilha com Entrada, Diário e Grade |
| Recomeçar | Apaga a memória do navegador e volta aos 5 exemplos |

## 8. Dados e backup

O estado fica no `localStorage` do navegador. Isso significa:

- não fica no servidor;
- não é compartilhado entre navegadores ou computadores;
- limpar os dados do site apaga o estado;
- trocar de navegador não leva o estado.

Use **Exportar estado** para gerar um JSON de backup e **Importar estado** para
restaurar. O JSON tem o mês, feriados, quadro, indisponibilidades e heranças.

## 9. Importar folgas do mês anterior

O botão **Carregar folgas** espera uma planilha do mês anterior. Ela pode ser:

- uma planilha gerada por este app, com as abas `Entrada` e `Grade`;
- a planilha antiga do modelo, com a aba `Escala`.

Ele lê o último dia de cada série e preenche o painel de herança. Uma planilha
aleatória, sem essas abas, não é importada.

Quem ficou indisponível o mês inteiro aparece como **sem registro**, com os
campos vazios. Preencha à mão antes de gerar, senão a pessoa entra com folga 0.

O mês do arquivo precisa ser o **anterior** ao mês em edição. Se não for, o app
diz de que mês é a planilha e qual mês em edição a usaria.

## 10. Planilhas geradas

O botão **Gerar planilha .xlsx** baixa um arquivo com três abas:

- **Entrada**: configuração, quadro e folgas iniciais. Serve para o modo CLI.
- **Diário**: um dia por linha, com data, dia da semana, pessoa escalada e
  observação.
- **Grade**: uma linha por pessoa e uma coluna por dia, com `D` para escalado,
  `IND` para indisponível e o contador de folga nos demais dias.

## 11. Modo CLI

O `entrada.xlsx` é o modelo do modo de linha de comando. O web app **não** usa
esse arquivo; ele é opcional.

Para usar o CLI:

```
python3 escala.py entrada.xlsx
```

O comando lê a aba `Entrada` e grava as abas `Diário` e `Grade` no mesmo
arquivo. O modelo vem com os 5 nomes fictícios e o mês `01/2027`. Edite a
configuração e o quadro antes de rodar.

## 12. Possibilidades e melhorias

O app foi feito para uma pessoa por dia, com a regra do mais folgado. Se você
quiser mudar o modelo, estes são os pontos de entrada.

### Mais de uma pessoa por dia

Hoje o `computar` escolhe um único militar por dia. Para escalar dois ou três:

- `escala.py`, função `computar`: troque `escolhido` por uma lista de escolhidos.
  Marque um `D` para cada um no `grade`.
- `escala.py`, funções `visao`, `render_diario` e `render_grade`: aceitar uma
  lista de nomes por dia.
- `static/app.js`, funções `renderDiario` e `renderGrade`: mostrar a lista.
- `teste.py`: criar um caso com dois escalados no mesmo dia.

Cuidado: o `escalados[data]` deixa de ser um militar e passa a ser uma lista.
Todos os pontos que leem esse valor precisam acompanhar.

### Trocar "mais folgado" por "corrida"

Hoje a escolha é:

```
min(cand, key=lambda m: (-m['cont'][serie], m['posto'], m['desempate'], m['nome']))
```

O `-m['cont'][serie]` é o mais folgado. Para uma fila por ordem de serviço,
troque a chave por quem está há mais tempo sem servir:

```
min(cand, key=lambda m: (m['ult'] or data_minima, m['posto'], m['desempate'], m['nome']))
```

O campo `ult` já guarda a última escala. A semente do mês anterior continua
valendo. Atualize o README e o `teste.py` junto.

### Outros pontos

- Indisponibilidade e retorno: `parse_indisponibilidade`, `indisponivel` e
  `retornos`.
- Séries e tipos de dia: `tipo_dia`, `config_datas` e `NAMES_SERIE`.
- Títulos e nome da organização: `index.html`, `cfg` e as funções `render_*`.
- Hierarquia e patentes: `POSTOS`, `parse_posto`, `ROT_EB_FA` e `ROT_MB`.
- Campos do quadro: `app.js` (`salvarEstado` e `restaurarEstado`) e `escala.py`
  (`montar_militar` e `CAB`).

## 13. Testes

> **Rode os testes sempre que mudar qualquer regra, campo, botão ou função.**
> Eles são a garantia de que a escala continua igual. Se um teste falhar, pare
> e veja o que mudou antes de gerar uma planilha.

### Motor

Copie e cole:

```
python3 teste.py
```

Resultado esperado: uma linha `PASS` para cada caso e `OK` no final. Hoje são
11 casos.

### Servidor

Copie e cole:

```
python3 teste_servidor.py
```

Resultado esperado: 17 testes e `OK`. Ele confere o portão de IP, o Host, o upload e a
geração do xlsx.

### O que cada teste cobre

| Teste | O que confere |
|---|---|
| `teste.py` | Regras do motor, importação, retorno, estreia e ordem |
| `teste_servidor.py` | 403 fora da lista, 200 dentro, upload e geração do xlsx |

Se você mexeu no motor, rode os dois. Se mexeu só no HTML ou no CSS, rode o do
servidor. Se mexeu no `app.js`, rode os dois e abra a tela para conferir.

## 14. Docker (opcional)

O modo recomendado em servidor Linux é o systemd. O Docker é uma alternativa
para quem não quer instalar Python. **Não precisa de Docker Compose**: é um
único container.

1. Clone o repositório e entre na pasta, como no passo 2 da seção 5. Para Docker,
   não precisa do passo 1 nem do passo 3, ou seja, não precisa instalar Python.

2. Confirme que você está na pasta do projeto, onde está o `Dockerfile`.

3. Construa a imagem. Copie e cole no terminal. O ponto no final é o contexto
   de build, a pasta atual, não é erro de digitação:

```
docker build -t escala-de-servico-por-folga .
```

4. Rode o container. Copie e cole:

```
docker run --rm -p 127.0.0.1:8770:8770 --read-only --tmpfs /tmp --cap-drop=ALL --security-opt no-new-privileges escala-de-servico-por-folga
```

5. Abra:

```
http://127.0.0.1:8770/
```

O `docker build` só cria a imagem. O `docker run` é que sobe o app. A imagem
não vem pronta de um registro; o `Dockerfile` copia o código da pasta para
dentro dela, por isso o clone é obrigatório.

Os IPs `172.17.0.1` e `127.0.0.1` no `Dockerfile` são padrões do Docker e do
próprio computador, não são IPs seus. O `172.17.0.1` é o gateway padrão da rede
bridge do Docker no Linux. Se a sua rede Docker for diferente, veja com:

```
docker network inspect bridge
```

Para acessar de outros computadores com o IP real, use `--network host` no
Linux e ajuste o `PERMITIR`. Com `-p`, o app vê o gateway do Docker, não o IP
do cliente, então a lista `PERMITIR` não distingue quem está na rede. No
Windows e no macOS, o caminho mais simples para uso local é rodar com Python.

## 15. Autoria e transparência

Este projeto foi idealizado por LeoG.Digital. A implementação contou com
assistência de IA. As decisões de produto, a revisão e a aprovação são de
LeoG.Digital.

## 16. Segurança

O app não tem login. O controle de acesso é a lista de IPs `PERMITIR`. Por
padrão, ele escuta apenas em `127.0.0.1`, no próprio computador.

- Rode em rede local ou atrás de um proxy com autenticação.
- Não exponha o app diretamente na internet.
- Os dados ficam no `localStorage` do navegador.
- Não coloque `.env`, tokens, senhas ou dados pessoais no repositório.

Mais detalhes em [SECURITY.md](SECURITY.md).

## 17. Licença

MIT. Copyright (c) 2026 LeoG.Digital. Veja [LICENSE](LICENSE).
