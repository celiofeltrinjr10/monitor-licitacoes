# Monitor Semanal de Licitações de Obras — PNCP

Script que consulta a API pública do PNCP (Portal Nacional de Contratações
Públicas), filtra licitações de **obras/construção civil** e mantém uma
planilha (`controle_licitacoes_obras.xlsx`) como banco de dados acumulativo —
cada execução só adiciona linhas novas, nunca duplica ou sobrescreve.

## 1. Instalação (fazer uma única vez)

Pré-requisito: Python 3.10 ou superior instalado.

```bash
cd pncp_monitor
pip install -r requirements.txt
```

## 2. Uso manual

```bash
# Busca os últimos 9 dias (padrão) e atualiza a planilha
python pncp_scraper.py

# Busca um período maior (ex: primeira execução, para "carga inicial")
python pncp_scraper.py --dias 30

# Testa sem alterar a planilha (mostra o que encontraria)
python pncp_scraper.py --dry-run
```

A planilha `controle_licitacoes_obras.xlsx` é criada automaticamente na
primeira execução, na mesma pasta do script.

## 3. Por que o padrão é "9 dias" e não "7 dias"?

Para uma rotina semanal, usar uma janela um pouco maior que 7 dias (com
2 dias de folga) evita perder licitações publicadas perto da virada da
semana, por causa de pequenas diferenças de fuso horário ou atraso na
publicação pelos órgãos. Como o script já elimina duplicatas automaticamente
(pela coluna `numeroControlePNCP`), não há risco de duplicar registros — a
folga só garante que nada passe despercebido.

## 4. Agendamento automático (rodar 1x por semana sem precisar lembrar)

### Windows — Agendador de Tarefas
1. Abra o **Agendador de Tarefas** (Task Scheduler).
2. "Criar Tarefa Básica" → defina um nome (ex: "Monitor PNCP Obras").
3. Disparador: **Semanalmente**, escolha o dia e horário (ex: toda segunda, 08h).
4. Ação: **Iniciar um programa**.
   - Programa/script: caminho do `python.exe` (ex: `C:\Python312\python.exe`)
   - Argumentos: `pncp_scraper.py`
   - Iniciar em: a pasta `pncp_monitor` (caminho completo da pasta no seu PC)
5. Finalize. A planilha será atualizada automaticamente todo início de semana.

### macOS / Linux — cron
```bash
crontab -e
```
Adicione a linha (exemplo: toda segunda-feira às 08h):
```
0 8 * * 1 cd /caminho/completo/pncp_monitor && /usr/bin/python3 pncp_scraper.py >> log_cron.txt 2>&1
```

## 5. Personalizando o filtro

Abra `pncp_scraper.py` e edite:

- **`PALAVRAS_CHAVE_OBRAS`**: lista de termos buscados no objeto da licitação.
  Adicione ou remova termos livremente (sem necessidade de acento — o script
  já ignora acentuação e maiúsculas/minúsculas).
- **`MODALIDADES`**: códigos de modalidade consultados. Por padrão:
  Concorrência (Eletrônica e Presencial) e Pregão (Eletrônico e Presencial).
  Para incluir Dispensa de Licitação (código 8), descomente a linha
  correspondente — atenção: isso aumenta bastante o volume de resultados.

## 6. Colunas da planilha

| Coluna | Descrição |
|---|---|
| numeroControlePNCP | Chave única da licitação (usada para evitar duplicatas) |
| dataColeta | Quando o robô capturou esse registro |
| statusInterno | Você controla manualmente (ex: "analisando", "descartado", "proposta enviada") |
| orgao / cnpjOrgao | Órgão público responsável |
| uf / municipio | Localização |
| modalidade | Concorrência, Pregão, etc. |
| objetoCompra | Descrição do que está sendo contratado |
| valorEstimado | Valor estimado da contratação |
| dataPublicacaoPNCP | Data de publicação |
| dataAberturaProposta / dataEncerramentoProposta | Janela de propostas |
| situacao | Status no PNCP |
| palavraChaveEncontrada | Qual termo do filtro disparou essa linha (útil para ajustar a lista de palavras) |
| linkPNCP | Link direto para o edital |
| processo | Número do processo no órgão |

A coluna `statusInterno` é sua área de trabalho manual — o robô só escreve
"novo - a analisar" nela; toda a triagem e decisão (próxima etapa do nosso
workflow de segunda a sexta) acontece editando essa coluna diretamente na
planilha.

## 7. Importante: teste a primeira execução

Como o ambiente onde este script foi gerado não tem acesso de rede ao
`pncp.gov.br`, ele foi validado quanto à sintaxe e lógica, mas a primeira
chamada real à API deve ser testada no seu computador. Rode primeiro com
`--dry-run` para confirmar que está tudo funcionando antes de deixar
agendado.
