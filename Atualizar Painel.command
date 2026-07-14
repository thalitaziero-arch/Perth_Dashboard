#!/bin/bash
# ─────────────────────────────────────────────────────────────
#  Atualizar Painel Perth Azzurri — dê DOIS CLIQUES neste arquivo
#  (antes, troque na pasta: Perth_SC.pdf, NPL_Comparison.pdf,
#   team_stats_perth.xlsx  — mesmos nomes, conteúdo novo)
# ─────────────────────────────────────────────────────────────

# vai para a pasta onde este arquivo está (funciona de qualquer lugar)
cd "$(dirname "$0")" || exit 1

echo "======================================================"
echo "   ATUALIZANDO O PAINEL PERTH AZZURRI"
echo "   Pasta: $(pwd)"
echo "======================================================"
echo ""

# usa o python3 disponível (anaconda ou do sistema)
if command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=/usr/bin/python3
fi

"$PY" atualizar_tudo.py
STATUS=$?

echo ""
if [ $STATUS -eq 0 ]; then
    echo "======================================================"
    echo "   ✅ PRONTO! O site atualiza em 1-2 minutos."
    echo "======================================================"
else
    echo "======================================================"
    echo "   ⚠️  Algo deu errado (código $STATUS). Veja a mensagem acima."
    echo "======================================================"
fi

echo ""
echo "Pode fechar esta janela."
# mantém a janela aberta pra você ler o resultado
read -n 1 -s -r -p "Aperte qualquer tecla para fechar..."
echo ""
