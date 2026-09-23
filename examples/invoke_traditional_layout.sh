#!/usr/bin/env bash
# Exemplo: invocar a skill V2 com planilha no formato V1 (uma coluna só,
# cada linha no formato "YYYYMM - NN% - descrição").
#
# Uso: edite PLANILHA/SAIDA/CICLO e rode.

set -euo pipefail

PLANILHA="${1:-/caminho/para/planilha_v1.xlsx}"
SAIDA="${2:-./catalogo_v1.html}"
CICLO="${3:-12/2026}"

python3 generate.py "$PLANILHA" "$SAIDA" \
  --cycle "$CICLO" \
  --focus-column "Promoção Próximo Ciclo + 1"

echo
echo "Gerado: $SAIDA"
echo "Abra no browser: file://$(realpath "$SAIDA")"
