#!/usr/bin/env bash
# Exemplo: invocar a skill V2 com a planilha no layout de duas colunas
# paralelas (padrão BSO a partir do ciclo 14/2026).
#
# Uso: edite PLANILHA/SAIDA/CICLO e rode.

set -euo pipefail

PLANILHA="${1:-/caminho/para/DRAFT_PDVS_SEM_XXXXX.xlsx}"
SAIDA="${2:-./catalogo_ciclo_15_2026.html}"
CICLO="${3:-15/2026}"

# 1) Substituir o template pelo patcheado (badge usa valor da coluna)
cp template_patched.html template.html

# 2) Rodar o generate.py com as flags do layout de duas colunas
python3 generate.py "$PLANILHA" "$SAIDA" \
  --cycle "$CICLO" \
  --focus-column "Promoção Próximo Ciclo + 1" \
  --focus-discount-column "Desconto Promoção Próximo Ciclo + 1" \
  --projection-column "Projeção Próximo Ciclo + 2"

echo
echo "Gerado: $SAIDA"
echo "Abra no browser: file://$(realpath "$SAIDA")"
