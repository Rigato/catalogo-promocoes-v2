# Exemplos de invocação

## Layout 1 — Duas colunas paralelas (layout novo, pós ciclo 14/2026)

A planilha traz duas colunas:
- `Desconto Promoção Próximo Ciclo + 1` (col 36)
- `Promoção Próximo Ciclo + 1` (col 37)

Cada linha de uma casa pareia com a linha equivalente da outra (ex: linha 3
do desconto = linha 3 da promoção).

### Comando

```bash
# 1) Com template patcheado (badge usa valor da coluna)
python3 run.py
# (edite o topo do run.py pra apontar pro seu xlsx)

# OU equivalente manual:
cp template_patched.html template.html
python3 generate.py planilha.xlsx saida.html \
  --cycle 15/2026 \
  --focus-column "Promoção Próximo Ciclo + 1" \
  --focus-discount-column "Desconto Promoção Próximo Ciclo + 1"
```

### Quando usar esse layout

Quando o dado vem direto da BSO mensal e tem 4 colunas de promoção
(Ciclo Atual, Próximo Ciclo, Próximo Ciclo +1, Próximo Ciclo +2) cada uma
com sua coluna de desconto pareada.

## Layout 2 — Formato V1 (uma coluna só, linhas `YYYYMM - NN% - desc`)

A coluna foco traz linhas no formato:

```
202612 - 10% - Itens selecionados com 10% de desconto direto.
202612 - 5% - Lucro Extra progressivo de 5% na compra de 50 ou mais itens.
```

### Comando

```bash
python3 generate.py planilha.xlsx saida.html \
  --cycle 12/2026 \
  --focus-column "Promoção Próximo Ciclo + 1"
```

### Quando usar

Quando a planilha foi exportada manualmente ou o script de extração não
quebrou em duas colunas. O V2 detecta automaticamente o formato e dispensa
a coluna de desconto.

## Layout 3 — Ciclo Atual (ao invés de Próximo Ciclo + 1)

Pra gerar catálogo do ciclo corrente em vez do próximo:

```bash
python3 generate.py planilha.xlsx saida.html \
  --cycle 14/2026 \
  --focus-column "Promoção Ciclo Atual" \
  --focus-discount-column "Desconto Promoção Ciclo Atual"
```

## Auto-detecção de ciclo

Sem `--cycle`, o script tenta detectar pelo header da coluna "Histórico de
Vendas ... (atual)". Se essa coluna for `Histórico de Vendas do Ciclo
202614 (atual)`, o script assume que a coluna foco é o ciclo **+1** →
`202615`. Para planos de ciclo atual ou horizonte diferente, use `--cycle`
explicitamente.

## Atualizando a lista de linhas canônicas

A primeira execução puxa de `linhas_canonicas.json` (carregado em runtime).
Se aparecer uma linha nova (ex: "BOTI SPORT"), atualize passando uma
planilha da portfolio:

```bash
python3 generate.py --refresh-linhas portfolio.xlsx
```

(Isso regenera `linhas_canonicas.json` e imprime um resumo; não gera
HTML.)
