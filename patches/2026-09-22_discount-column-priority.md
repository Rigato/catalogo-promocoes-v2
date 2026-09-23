# Patch 2026-09-22 — discount-column-priority

## Problema

O `template.html` original da skill V2 tem uma função `renderActionChips`
que priorizava `promo.allDiscounts` (regex extraindo % do texto da
descrição) sobre `promo.discount` (valor da coluna `Desconto Promoção ...`)
na hora de renderizar o badge inline do produto:

```js
// ANTES (errado):
if (promo.allDiscounts && promo.allDiscounts.length) {
  text = fmtPctList(promo.allDiscounts);    // ← usava texto
} else if (promo.discount != null) {
  text = `${promo.discount}%`;              // ← fallback pra coluna
}
```

Isso é incorreto quando a planilha traz **duas colunas paralelas**: a coluna
"Desconto Promoção ..." (fonte da verdade do desconto real do PDV) e a
coluna "Promoção ..." (texto descritivo). Em muitos casos o texto menciona
"até X%" como teto promocional, enquanto a coluna traz o valor real.

### Caso concreto (Arbo Atlantica, SKU 55366, ciclo 15/2026)

| Fonte | Valor |
|---|---|
| Coluna `Desconto Promoção Próximo Ciclo + 1` | `25.0` ← real |
| Texto `Promoção Próximo Ciclo + 1` | "...com até **50%** de desconto" |

| Badge antes do patch | Badge depois |
|---|---|
| `DD 50.0%` ❌ | `DD 25%` ✅ |

### Escala do impacto (ciclo 15/2026)

- 1.676 participações com promoção
- **743 badges estavam com valor errado** (coluna > 0 e divergia do texto)
- 35 promoções com `coluna = 0` (ex: "Compre 3 pague 2, até 33%") → fazem
  fallback pro texto intencionalmente

## Solução

```js
// DEPOIS (correto):
if (promo.discount != null && promo.discount > 0) {
  text = `${promo.discount}%`;              // ← usa coluna primeiro
} else if (promo.allDiscounts && promo.allDiscounts.length) {
  text = fmtPctList(promo.allDiscounts);    // ← fallback pro texto
}
```

### Por que `> 0` e não só `!= null`?

Algumas promoções têm `discount = 0` na coluna, o que indica **"sem
desconto base; só dá desconto na compra por volume"**. Nesses casos o
desconto real vem do texto:

- Texto: "Compre 3 itens selecionados EUDORA ROXO, IMPRESSION, DIVA e VOLPE
  e pague 2, até 33% de desconto em cada unidade."
- Coluna: `0`
- allDiscounts extraído: `[{pct: 33.0, ...}]`

Com `discount > 0`, o badge cai pro fallback `DD 33%` corretamente.

## Como aplicar

O patch já está aplicado em `template_patched.html`. Se quiser reaplicar em
outro `template.html`:

```bash
patch template.html < patches/2026-09-22_discount-column-priority.patch
```

Ou simplesmente copie `template_patched.html` por cima do original.

## Compatibilidade com versões anteriores da skill

O patch é **retrocompatível**: planilhas no formato V1 (uma coluna só,
formato `YYYYMM - NN% - desc`) continuam funcionando igual ao template
puro, porque nesse formato a coluna desconto paralela não existe e
`promo.discount` é setado a partir do `NN%` do próprio texto — ou seja, é
igual ao `allDiscounts` extraído.

## Outras partes do template que não foram tocadas

- `effMax = a.maxDiscount` (badge grande no action header): já usava
  coluna corretamente via `Math.max(...p.discount)`. Não precisou patch.
- Tabela de produtos: já renderizava `${p.discount}%` (coluna). Não
  precisou patch.

Só o badge inline `renderActionChips` tinha a inversão.
