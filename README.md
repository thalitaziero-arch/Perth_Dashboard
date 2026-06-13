# Painel Perth Azzurri — Guia de Atualização Semanal

## Arquivos do pacote

- **perth_azzurri_painel_v2.html** — o painel completo, já com os dados da NPL Comparison embutidos dentro do arquivo (objeto `REPORT_DATA` no `<script>`). **É este arquivo que você edita/substitui toda semana.**
- **npl_comparison_data.json** — cópia de referência dos mesmos dados em JSON puro, útil para conferir valores ou para rodar o `extract_wyscout.py`.
- **extract_wyscout.py** — script Python que lê o PDF "NPL Comparison" e gera/atualiza esses dados automaticamente.
- **README.md** — este guia.

## Como funciona

Os dados da aba "NPL Comparison" estão dentro do próprio `perth_azzurri_painel_v2.html`, no objeto `const REPORT_DATA = { ... }`. Não há mais arquivo externo — o painel funciona sozinho, inclusive abrindo localmente (`file://`).

Toda semana você atualiza esse objeto `REPORT_DATA` dentro do HTML com os novos valores do PDF "NPL Comparison".

## Atualização semanal (passo a passo)

Toda semana você recebe o PDF **"NPL Comparison"** (Season Report) do Hudl Wyscout. Para atualizar:

1. Abra `perth_azzurri_painel_v2.html` em um editor de texto/código (VS Code recomendado).
2. Procure por `const REPORT_DATA = {` — é o bloco que contém todos os dados da NPL Comparison.
3. Abra o novo PDF e localize cada seção abaixo. Os comentários originais indicavam as páginas do PDF — use como referência (os números de página podem variar levemente entre rodadas):

| Campo no JSON | Conteúdo | Página típica do PDF |
|---|---|---|
| `lastUpdated` | Texto tipo "Round 11 · Junho 2026" — atualize a cada rodada | — |
| `standings` | Tabela de classificação (pos, pontos, jogos, V/E/D, gols pró/contra, xG/xGA) | p.2 |
| `goalsScoredTypology` | Tipologia dos gols marcados (chutes de longe, contra-ataque, bola parada, falta direta) | p.3 |
| `goalsScoredDynamics` | Gols marcados por período (1º/2º tempo, sextos de jogo) + xG | p.3 |
| `goalsConcededTypology` | Tipologia dos gols sofridos | p.4 |
| `goalsConcededDynamics` | Gols sofridos por período + xGA | p.4 |
| `formations` | Formações mais usadas por time (% de uso) | p.5 |
| `attackFlanks` | % de ataques por lado (esquerda/centro/direita) | p.5 |
| `teamStats.shots` até `teamStats.shotsXG` | Chutes por 90min, chutes sofridos, xG por chute | p.6 |
| `teamStats.possession`, `teamStats.passes` | Posse de bola, passes por 90min | p.8 |
| `teamStats.progressive`, `teamStats.deepComp` | Passes progressivos, recepções profundas | p.9 |
| `teamStats.crosses`, `teamStats.dribbles`, `teamStats.touchesPA` | Cruzamentos, dribles, toques na área | p.10 |
| `teamStats.offDuels`, `teamStats.defDuels`, `teamStats.aerialDuels` | Duelos ofensivos/defensivos/aéreos | p.11 |
| `teamStats.ppda`, `teamStats.interceptions` | PPDA (intensidade de pressão), interceptações | p.12 |
| `teamStats.recoveries`, `teamStats.losses` | Recuperações e perdas de bola por 90min | p.7 |

3. Para cada seção, substitua os valores pelos números do novo PDF. **Mantenha a estrutura exata** (chaves `{ }`, vírgulas, aspas) — só troque os números e nomes.

   Exemplo — uma linha da tabela `standings`:
   ```json
   { "pos": 1, "team": "Perth", "pts": 24, "p": 10, "w": 7, "d": 3, "l": 0, "gf": 23, "ga": 2, "gd": 21, "xg": 17.8, "xga": 4.3 }
   ```
   Troque apenas os números (`pts`, `p`, `w`, `d`, `l`, `gf`, `ga`, `gd`, `xg`, `xga`) pelos da nova rodada. Se a ordem dos times na classificação mudar, é só reordenar as linhas e ajustar `pos`.

4. Para as listas em `teamStats` (ex: `shots`, `possession`), cada time tem um objeto `{"team": "Nome", "val": número}`. Atualize o `val` de cada time e, se a ordem do ranking mudar, reordene as linhas (a ordem na lista define a posição no gráfico/tabela).

5. Salve o arquivo `perth_azzurri_painel_v2.html`.

6. Verifique se o JSON dentro do `REPORT_DATA` continua válido (sem vírgulas sobrando, todas as chaves e valores entre aspas quando texto). O VS Code já avisa se houver erro de sintaxe dentro do bloco.

7. Suba o `perth_azzurri_painel_v2.html` atualizado para o servidor.

8. Abra o painel no navegador, vá até a aba "NPL Comparison" e confirme que os dados aparecem corretos e que o texto "Last updated" no topo da aba mostra a rodada certa.

## Atualização automática com extract_wyscout.py

Em vez de editar manualmente, você pode usar o script Python para gerar os dados automaticamente a partir do PDF:

```bash
pip install pdfplumber --break-system-packages
python extract_wyscout.py caminho/para/o_novo_relatorio.pdf --round "Round 11 · Junho 2026"
```

O script gera/atualiza `npl_comparison_data.json` e mostra um resumo das diferenças em relação à rodada anterior (pontos na classificação, etc.), além de avisos se alguma seção não tiver os 8 times esperados — confira esses avisos antes de prosseguir.

Depois, copie o conteúdo de `npl_comparison_data.json` e cole no lugar do objeto `REPORT_DATA` dentro do `perth_azzurri_painel_v2.html` (substituindo apenas o que está entre `const REPORT_DATA = ` e o `;` final do bloco).

Se quiser, peça para o Claude fazer essa parte: basta enviar o novo PDF e o HTML atual, e o Claude roda o script, confere os avisos e atualiza o HTML para você.

## Solução de problemas

- **A aba "NPL Comparison" aparece vazia, com dados antigos ou trava o carregamento**: provavelmente há um erro de sintaxe no objeto `REPORT_DATA` (vírgula faltando ou sobrando, chave sem fechar). Abra o Console do navegador (F12) para ver a mensagem de erro exata, ou abra o arquivo no VS Code, que sinaliza erros de JSON/JS automaticamente.
- **Quer confirmar que o `REPORT_DATA` está correto antes de subir**: copie o conteúdo entre `const REPORT_DATA = ` e o `;` final, cole em [jsonlint.com](https://jsonlint.com) e valide.
