---
name: ciclo-promocoes-v2
description: |
  Generate a self-contained interactive HTML catalog of promotional actions
  from a Boticário / O.U.i / Eudora / Quem Disse product spreadsheet,
  organised **by Tipo de Promoção** (Desconto Direto, Lucro Extra, Combo,
  Progressivo, Volumetria, etc.) instead of by individual action card.
  Reads the .xlsx, classifies each promotion structurally via regex,
  parses combo deals with implied discount, derives product linha from
  the description, and renders a single index.html with collapsible
  type sections + collapsible subgroups, ready to open locally.
---

# ciclo-promocoes-v2 (V2)

V2 is a parallel skill to `ciclo-promocoes` (V1). The output page has the same
data envelope (`DATA = { products: [...] }`) and the same `.xlsx` input, but
the page groups promotions **by tipo de promoção** (which actionType each one
falls into) rather than one card per action. That makes it much easier to
plan supply decisions across many products and tier variants at once.

## When to use

Use this skill — **not** `ciclo-promocoes` (V1) — when the user sends a new
`Promoção Próximo Ciclo + 1` `.xlsx` and wants to look at the same volume of
promotions from the angle of *what kind of promotion each one is*, not *one
card per reseller tier*. The user phrases this as:

- "gera o catálogo v2 dessa planilha"
- "roda a v2 com esses dados"
- "cria a página nova do ciclo 12"

If the user is okay with the V1 layout (action card per row), stay on
`ciclo-promocoes`. The skills don't share data; each one is read-only against
its own template.

## What V2 produces that V1 doesn't

| Aspect | V1 | V2 |
|---|---|---|
| Grouping inside Tipo | Per-action card (one per normalised description) | One section per `actionType`, products listed in subgroups |
| Promotion classification | None — every action displayed as-is | **actionType** field per promotion via regex (desconto_direto, lucro_extra, combo, combo_preco, progressivo, volumetria, primeiro_pedido, brinde, outra) |
| Subgroup axis | Categoria only | Categoria / **Linha** / **Profundidade de desconto** (toggle on toolbar) |
| Tooltip on hover | Action description in a custom dark box above the row | Action description via the browser's native `title=` (small delay, near cursor — no custom DOM) |
| Subgroup collapse | n/a | Every subgroup is collapsible, default closed, click to expand |
| Header tone | Cool gray | Warm beige (`#f3efe6` / `#efeae0` on hover) on the type rows, matching V1's original action card look |
| Linha derivation | Static first-word extraction in Python | JS-side `deriveLinha()` with brand-aware rules (BOTI BABY/HOME/SUN are linhas; BOTI NECESS is not; QDB/OUI fall back to subcategoria) |

## Expected input format

Identical to V1: a `.xlsx` with **3 sheets**, in this order:

1. `QUEM_DISSE_BERENICE` — products of the QDB / Quem Disse brand
2. `EUDORA` — products of the Eudora brand
3. `BOTICARIO` — products of Boticário **and** OUI (OUI products are
   re-classified by their description prefix — see below)

Column layout (Excel column letters, 1-indexed):

| Col | Field | Used for |
|---|---|---|
| B (2) | Classe (A/B/C/…) | Curva ABC color coding |
| C (3) | Classe Segmentada (BA/BB/BC) | Sub-class display |
| D (4) | SKU | Product identifier |
| E (5) | Descrição | Product name + used to detect OUI products and derive Linha |
| F (6) | Categoria | Grouping, filters, table column |
| G (7) | Subcategoria | Fallback Linha for OUI/QDB products |
| H (8) | Lançamento | Cycle when product was launched (C07 / C08 / …) |
| I (9) | Desativação | Cycle when product is deactivated |
| J..AA (10–27) | Histórico de Vendas do Ciclo 202510..202609 | Used to compute `max_hist` |
| AF (32) | Projeção ciclo + 2 | Forward projection (used in Top 10 + table) |
| AH (34) | **Promoção Próximo Ciclo + 1** | **THE FOCUS COLUMN** — each newline is one action |

### Action parsing

Each line of the focus column is parsed as:

```
YYYYMM - NN% - <description text>
```

or, less often, without the discount percent:

```
YYYYMM - <description text>
```

### Action-type classification (V2-specific)

Each parsed promotion gets an `actionType` field via this precedence of
regexes (the **first** match wins — order matters):

| Order | Type | Pattern (lowercase) |
|---|---|---|
| 1 | `combo` | `(ao )?comprar? \d+ (unidades?|itens?)…ganhe \d+ (unidade|item)` |
| 2 | `combo_preco` | `pre(c|ç)o de:?\s*r\$ \| combo \w+\s+por\s+r\$` |
| 3 | `primeiro_pedido` | `primeiro pedido \| primeira compra` |
| 4 | `volumetria` | `\d+ a \d+ itens? \| compre \d+ a \d+ \| \d+ ou mais (unidades?|itens?)` |
| 5 | `progressivo` | `\d+ itens?:?\s*\d+% \| progressiv` |
| 6 | `lucro_extra` | `lucro extra \| \ble\b` |
| 7 | `brinde` | `junte.*embalagens \| botirecicla \| junte \d+\s*embal` |
| 8 | `desconto_direto` | `\d+%\s+de\s+desconto \| até \d+%` |
| 9 | `combo` (fallback) | `pague? \d+…leve \d+` |
| 10 | `outra` | nothing matched |

### Combo detection

Same as V1 — patterns like "comprar 5 unidades, ganhe 1 unidade" or
"pague 1, leve 2" produce an implicit discount % (`y/(x+y)`) shown in the
combo badge (only used as a hint inside the tooltip — the page itself
visualises the actionType classification, not the combo badge).

### Action normalisation

To avoid exploding the page with one section per reseller tier
(Cobre/Bronze/Prata/Ouro/Platina/Rubi/Esmeralda/Diamante), V1 strips tier
names from descriptions before grouping. V2 inherits the same logic so the
same tier variants of one promotion land in the same Tipo de promoção
section, not in duplicates.

### Linha derivation (JS, runtime)

`p.linha` is left empty in the data.json (Python leaves `""`). At runtime the
JS `deriveLinha()` function fills it in, applying brand-aware rules:

| Brand | Rule |
|---|---|
| BOTICARIO | First word of the description IS the linha — MALBEC, LILY, FLORATTA, MAKE, etc. |
| BOTICARIO with first="BOTI" or "BOT" | If second word is in `{BABY, HOME, SUN}` → `BOTI BABY`/`BOTI HOME`/`BOTI SUN`. Otherwise (e.g. "BOTI NECESS", "BOTI NECESSAIRE") → fall back to subcategoria (those are product types, not families). |
| EUDORA with first in `{EUD, REF, EUDORA}` | Falls back to subcategoria — these are brand prefixes, no real linha. |
| EUDORA other | First word IS the linha (NIINA, SIAGE, GLAM, INSTANCE, CHIC, …). |
| OUI / QDB | Description doesn't encode a linha → falls back to subcategoria. |

## Files in this skill

| File | Purpose |
|---|---|
| `SKILL.md` | This file |
| `template.html` | The HTML template with `__DATA_PLACEHOLDER__` for the data payload. All CSS/JS is inlined (no network). |
| `generate.py` | Reads the `.xlsx`, classifies each promotion, detects combos, emits the final `index.html`. |

## How to run

```bash
python3 .skills/ciclo-promocoes-v2/generate.py <input.xlsx> [output.html]
```

If `output.html` is omitted, the script writes to `./index.html` next to the
input file. Defaults to looking for `template.html` in the same directory
as `generate.py`.

## Using V2 side-by-side with V1

The two skills read the **same** `.xlsx` shape and write **independent** HTML
files. To produce both pages from one input:

```bash
python3 .skills/ciclo-promocoes/generate.py   ciclo12.xlsx ciclo12_v1.html
python3 .skills/ciclo-promocoes-v2/generate.py ciclo12.xlsx ciclo12_v2.html
```

Open them in two tabs and compare. They share header style, brand color
coding, filters and Top 10 — V2 just reorganises what V1 calls the
"action cards" into structural buckets (`actionType`) plus a sub-axis
toggle (Categoria / Linha / Profundidade).

## What the output page does

- **Header**: title + cycle badge + 5 metric cards (totals per brand).
- **Sticky toolbar** with filters:
  - Brand chips (Todas / Bot / Eud / OUI / QDB)
  - Curva chips (A / B / C / outros)
  - Lançamentos do ciclo
  - Top 10 toggle
  - Categoria multi-select dropdown
  - Fundo (gray ↔ cream) toggle
  - Search box (SKU / descrição / promoção)
  - **Ordenar ações por** (desconto / volume / nprods / ciclo)
  - **Ordenar produtos por** (default / desconto / vol hist / projeção) — V2 only
  - **Agrupar produtos por** (Categoria / **Linha** / **Profundidade de desconto**) — V2 only
  - Expandir tudo / Recolher tudo / Limpar filtros
- **Top 10 view** (toggle ON): grid per brand with rank, SKU, descrição,
  categoria, curva, vol, projeção, desconto.
- **Actions view** (toggle OFF, default): per Marca → per `actionType` →
  per subgrupo → lista de produtos. Every level is collapsible:
  - Type sections: dense beige header, default closed on page load.
  - Subgroups: gray, default closed; click to open and see the products.
- **Hover tooltip**: native browser tooltip on each product row showing the
  parent promotion's full description (cycle + tipo + texto). Appears after
  the browser's standard delay (close to the cursor).
- **Inline linha pill**: on each row, a secondary info chip with either
  `Linha: X` or `Categoria: X` depending on the active subgroup axis.

## Notes & assumptions

- The skill assumes the focus column is **column 34 (AH)**. If the user has
  a different spreadsheet layout, update the constants in `generate.py`
  (look for `row[32]` etc.).
- The skill uses **column 32 (AF)** for "Projeção ciclo + 2" and columns
  10–27 (J–AA) for historical sales to compute `max_hist`. Same comment as
  V1.
- The first-load state is **all closed** by design — both type sections and
  subgroups. After the user opens anything, that choice persists (the
  `state.typesOpen` / `state.subgroupsOpen` Sets survive re-renders). A
  final "Limpar filtros" reset returns both to fully-closed.
- If a new spreadsheet introduces a brand prefix or product family that
  isn't in the derivation rules (e.g. "BOTI SPORT"), add it to
  `BOTI_SUB_LINHAS` in `template.html` JS (search for `BOTI_SUB_LINHAS`).
  Or update the corresponding brand rule if it's a bigger change.

## Quick sanity check after regenerating

After running `generate.py`, verify in the output HTML:

1. **Counts** in the header metric cards add up to the number of products in
   the input that had at least one promotion on the focus column.
2. The **current cycle** shown in the header (`Ciclo 11/2026`) matches the
   most common cycle in the focus column.
3. The **OUI** count matches what you expect (sanity check by description
   prefix).
4. The **Tipo de promoção** sections appear in the canonical order
   (Desconto Direto, Lucro Extra, Combo, Combo Preço Fixo, Desconto
   Progressivo, Desconto por Volume, Primeiro Pedido, Brinde/Recompra,
   Outra) — types with no products are simply not rendered.
5. Each promotion runs through the `classify_action` regex and gets a
   sensible `actionType`. Logs at the end of `generate.py` print the
   distribution: `Action types: {...}`.

## Relationship to V1

`ciclo-promocoes` (V1) and `ciclo-promocoes-v2` are **independent** skills:

- They have separate `generate.py` files with separate logic.
- They have separate `template.html` files.
- They produce different HTML pages from the **same** `.xlsx`.
- They do **not** share data, so changing one does not affect the other.

The user runs whichever best fits the planning angle they want for each
cycle.
