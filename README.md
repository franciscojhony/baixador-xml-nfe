# Baixador de XML de NF-e - Python

Aplicação web **local** que baixa os XMLs de notas fiscais eletrônicas (NF-e) a partir das chaves de acesso de uma planilha. Roda no seu computador e é usada pelo navegador.

## Como funciona

O programa sobe um servidor local (Flask) em `http://127.0.0.1:8765` e abre uma janela real do Google Chrome controlada por Selenium. Ele repete o que uma pessoa faria à mão, sem depender de API:

**Etapa 1 — Baixar XMLs** (consultadanfe.com)

1. Você escolhe a planilha (`.xlsx`, `.xlsm` ou `.csv`) e a coluna com a chave de acesso (44 dígitos). A coluna é sugerida automaticamente.
2. Você informa a linha inicial e final (ex.: 1753 a 1802) e a pasta onde salvar.
3. Para cada chave, o Chrome cola a chave no site, clica em **Imprimir DANFE** e depois em **Baixar XML**. O arquivo é salvo como `<chave>.xml`.
4. Chaves inválidas são puladas e chaves já baixadas não são baixadas de novo. Basta rodar o mesmo intervalo outra vez para refazer só as que falharam.
5. Há pausas aleatórias entre as notas (e uma pausa maior a cada lote) para reduzir o risco de bloqueio. Se o site exibir a verificação anti-robô, você a resolve na própria janela e o programa continua.

A tela mostra o andamento em tempo real (barra, contadores, log e lista de falhas), tem botão **Parar** e um modo **Apenas simular**, que mostra o plano sem abrir o Chrome nem baixar ou enviar nada.

## Requisitos

- Python 3.9 ou superior
- Google Chrome instalado
- Pacotes: `flask`, `pandas`, `openpyxl`, `selenium`

```bash
pip install -r requirements.txt
```

## Como executar
- iniciar.bat       atalho de execução no Windows
- iniciar.sh        atalho de execução no Mac/Linux

O navegador abre sozinho em `http://127.0.0.1:8765`. No Windows também é possível dar dois cliques em `iniciar.bat` (ele procura o Anaconda automaticamente); no Mac/Linux, use `./iniciar.sh`.

## Estrutura

```
app.py            servidor Flask, automação com Selenium e lógica de download/importação
static/index.html interface (HTML, CSS e JavaScript em um único arquivo)
requirements.txt  dependências
iniciar.bat       atalho de execução no Windows
iniciar.sh        atalho de execução no Mac/Linux
```

## Observações

- O servidor escuta apenas em `127.0.0.1`: só o próprio computador acessa a aplicação.
- O modo invisível (headless) não é usado porque o site de consulta costuma bloqueá-lo.
- Rode em lotes pequenos (30 a 50 notas) e respeite os termos de uso do site consultado.
- Os fluxos dependem dos botões e IDs das páginas externas (consultadanfe.com e SISGI). Se elas mudarem, os seletores em `app.py` precisam ser ajustados.
