# Painel Perth Azzurri — Guia de Atualização Semanal

## Arquivos do pacote

- **perth_azzurri_painel_v2.html** — o painel completo (não precisa editar este arquivo nas atualizações semanais).
- **npl_comparison_data.json** — os dados da aba "NPL Comparison". **É este arquivo que você atualiza toda semana.**
- **README.md** — este guia.

## Como funciona

O painel carrega o arquivo `npl_comparison_data.json` automaticamente quando a aba "NPL Comparison" é aberta. Por isso:

1. Os dois arquivos (`perth_azzurri_painel_v2.html` e `npl_comparison_data.json`) **precisam estar na mesma pasta** no servidor (GitHub Pages, Netlify, etc.).
2. Toda semana, basta substituir o `npl_comparison_data.json` pelo arquivo atualizado — **não precisa tocar no HTML**.

## Atualização semanal (passo a passo)

Toda semana você recebe o PDF **"NPL Comparison"** (Season Report) do Hudl Wyscout. Para atualizar:

1. Abra `npl_comparison_data.json` em um editor de texto (VS Code, Notepad++, ou até o Bloco de Notas).
2. Abra o novo PDF e localize cada seção abaixo. Os comentários originais do HTML indicavam as páginas do PDF — use como referência (os números de página podem variar levemente entre rodadas):

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

5. Salve o arquivo `npl_comparison_data.json`.

6. Verifique se o JSON é válido (sem vírgulas sobrando, todas as chaves e valores entre aspas quando texto). Você pode colar o conteúdo em [jsonlint.com](https://jsonlint.com) para checar rapidamente.

7. Suba o `npl_comparison_data.json` atualizado para o servidor, na mesma pasta do `perth_azzurri_painel_v2.html`.

8. Abra o painel no navegador, vá até a aba "NPL Comparison" e confirme que os dados aparecem corretos e que o texto "Last updated" no topo da aba mostra a rodada certa.

## Dica

Se quiser, peça para o Claude processar o PDF da semana e gerar o `npl_comparison_data.json` atualizado automaticamente — basta enviar o PDF junto com este arquivo JSON atual.

## Solução de problemas

- **A aba "NPL Comparison" aparece vazia ou com mensagem de erro vermelha**: verifique se `npl_comparison_data.json` está na mesma pasta do HTML no servidor e se o nome do arquivo está exatamente correto (sensível a maiúsculas/minúsculas).
- **JSON inválido**: um erro de sintaxe (vírgula faltando ou sobrando, aspas erradas) impede o carregamento. Use o jsonlint.com para encontrar o erro exato.
