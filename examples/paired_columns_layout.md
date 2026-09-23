# Layout de planilha — Duas colunas paralelas (pós ciclo 14/2026)

## Visão geral

A partir do ciclo 14/2026 a BSO passou a exportar a planilha de PDVs com
**duas colunas pareadas** por ciclo: uma de desconto e uma de descrição.

Para cada ciclo, a planilha traz 2 colunas adjacentes:

```
... | Desconto Promoção Ciclo Atual | Promoção Ciclo Atual | Desconto Promoção Próximo Ciclo | ...
... | Desconto Promoção Próximo Ciclo + 1 | Promoção Próximo Ciclo + 1 | Desconto Promoção Próximo Ciclo + 2 | ...
```

Cada linha da coluna desconto pareia com a linha equivalente da coluna
descrição (mesmo índice na quebra por `\n`).

## Exemplo real (linha do Arbo Atlantica, SKU 55366, ciclo 15/2026)

```
Coluna "Desconto Promoção Próximo Ciclo + 1" (col 36):
    25.0

Coluna "Promoção Próximo Ciclo + 1" (col 37):
    Franqueado: Ressarcimento via BSO. Consumidor/Revendedor:
    Itens selecionados com até 50% de desconto.
```

→ 1 linha em cada coluna → 1 promoção:
- `discount = 25` (da coluna)
- `description = "Franqueado: ... com até 50% de desconto."`

## Exemplo com múltiplas ações (linha do ACCORDES, SKU 52882)

```
Coluna desconto:
    11.0
    10.0

Coluna descrição:
    Progressiva de itens selecionados, a partir de 2 itens 15% (desconto
    adicional 5,56%), 3 itens 20% (desconto adicional 11,11%).
    Itens selecionados com 10% de desconto direto.
```

→ 2 linhas em cada coluna → 2 promoções:
- Promoção 1: `discount = 11`, descrição da progressiva
- Promoção 2: `discount = 10`, descrição do desconto direto

## Como o V2 lê

O script detecta automaticamente esse layout: se as linhas não começam
com `YYYYMM - NN% -`, ele assume o formato de duas colunas e pareia
linha-a-linha. Se a coluna desconto estiver vazia e a descrição tiver um
`%`, faz fallback pro valor extraído do texto (mas isso é raro e o badge
vai mostrar `?%`).

## Por que parear linha-a-linha em vez de procurar % no texto

O desconto na coluna é a **fonte da verdade do desconto real do PDV**.
O texto pode mencionar:

- "**até X%**" → o X é o teto, não o desconto real (Arbo Atlantica: 50% no
  texto, 25% na coluna)
- "**de X% a Y%**" → progressivo, mas a coluna traz um valor fixo por tier
- "Compre X pague Y" → sem % no texto, mas a coluna traz o desconto
- Casos em que a coluna está vazia → fallback pro texto é a única opção

A regra é: **se a coluna tem valor positivo, usar ele. Se for null ou 0,
fazer fallback pro texto.**
