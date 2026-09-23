# ciclo-promocoes-v2 — pacote standalone

Skill para gerar catálogo HTML de promoções dos ciclos do Boticário / O.U.i
/ Eudora / Quem Disse, agrupando por **tipo de promoção** (Desconto Direto,
Lucro Extra, Combo, etc.) em vez de um card por ação.

Este pacote é self-contained: dá pra extrair em qualquer máquina com Python
3.11+ e `openpyxl` instalado, e gerar catálogos a partir de uma planilha
`.xlsx` de PDVs.

## Conteúdo do pacote

| Arquivo | O que é |
|---|---|
| `SKILL.md` | Descrição original da skill V2 (formato, regex de classificação, etc.) |
| `generate.py` | Script principal — lê xlsx, classifica promoções, emite HTML |
| `template.html` | Template V2 **puro** (preserva o comportamento original) |
| `template_patched.html` | Template V2 **com o patch de 2026-09-22** aplicado |
| `linhas_canonicas.json` | Cache de linhas canônicas por marca (BOTI, EUD, QDB, OUI) |
| `run.py` | Wrapper de exemplo que usa o template patcheado |
| `patches/` | Diff e changelog de patches aplicados |
| `examples/` | Comandos de invocação e exemplos de layout de entrada |

## Quando usar `template.html` vs `template_patched.html`

- **`template.html`** (puro): use se a planilha de entrada usa o formato V1
  antigo, com cada linha da coluna de promoção no formato
  `YYYYMM - NN% - descrição`. O script detecta esse formato automaticamente.

- **`template_patched.html`**: use se a planilha de entrada usa o formato de
  **duas colunas paralelas** ("Desconto Promoção ..." + "Promoção ..."), que
  é o que vem da BSO desde o ciclo 14/2026. O script já sabe ler esse
  formato, mas o template puro renderiza o badge com o `allDiscounts`
  extraído do texto (que pode divergir do valor real). O patch força o
  badge a usar o valor da coluna quando ela tem valor positivo.

  Detalhes em `patches/2026-09-22_discount-column-priority.md`.

## Instalação

```bash
pip install openpyxl
```

Sem outras dependências — o HTML gerado é standalone (CSS/JS inline, sem
chamadas externas).

## Como gerar um catálogo

### Modo 1 — formato V1 (uma coluna, formato `YYYYMM - NN% - desc`)

```bash
python3 generate.py planilha.xlsx saida.html \
  --cycle 14/2026 \
  --focus-column "Promoção Próximo Ciclo + 1" \
  --focus-discount-column "Desconto Promoção Próximo Ciclo + 1"
```

### Modo 2 — formato duas colunas paralelas (recomendado)

Quando a planilha tem colunas separadas para desconto e descrição da
promoção (cada linha de uma casa com a linha equivalente da outra), use o
**template patcheado** via `run.py`:

```bash
python3 run.py
```

Edite o topo do `run.py` para apontar pra sua planilha e seu diretório de
saída. Por padrão ele está configurado pra ler `Desconto Promoção Próximo
Ciclo + 1` + `Promoção Próximo Ciclo + 1` e gravar no ciclo 15/2026.

Se preferir chamar via CLI sem `run.py`, é só instalar um symlink:

```bash
# substitui o template padrão pelo patcheado
cp template_patched.html template.html
python3 generate.py planilha.xlsx saida.html \
  --cycle 15/2026 \
  --focus-column "Promoção Próximo Ciclo + 1" \
  --focus-discount-column "Desconto Promoção Próximo Ciclo + 1"
```

## Argumentos do `generate.py`

```
python3 generate.py <input.xlsx> <output.html>
                    [--cycle MM/AAAA]                              (override do ciclo)
                    [--focus-column NOME]                          (header da coluna descrição)
                    [--focus-discount-column NOME]                 (header da coluna desconto)
                    [--projection-column NOME]                     (header da projeção C+2)
                    [--refresh-linhas <marcas.xlsx>]               (atualiza linhas_canonicas.json)
```

Sem flag de ciclo, o script auto-detecta pela coluna "Histórico de Vendas
... (atual)" — assume que a coluna foco é o ciclo **+1** daquele.

## Layout esperado da planilha de entrada

A planilha precisa ter (pelo menos) uma das sheets reconhecidas:

- `BOTICARIO`
- `EUDORA`
- `QUEMDISSEBERENICE` ou `QUEM_DISSE_BERENICE`
- `O.U.I` ou `OUI` (O.U.i como sheet separada é layout novo; como linhas
  dentro de `BOTICARIO` é layout antigo — o script detecta pelos dois)

Colunas relevantes (1-indexed):

| Coluna | Campo |
|---|---|
| B (2) | Classe (Curva ABC) |
| C (3) | Classe Segmentada (BA/BB/BC) |
| D (4) | SKU |
| E (5) | Descrição |
| F (6) | Categoria |
| G (7) | Subcategoria |
| H (8) | Lançamento |
| I (9) | Desativação |
| J..AA (10–27) | Histórico de vendas |
| AF (32) | Projeção Próximo Ciclo + 2 |
| AH (34) ou similar | **Promoção do foco** (varia — use `--focus-column` pra indicar) |
| Coluna paralela | **Desconto da promoção** (use `--focus-discount-column` se existir) |

O script descobre as colunas por nome de header quando possível — mais
robusto contra drift entre ciclos.

## Tipos de promoção reconhecidos (`actionType`)

V2 classifica cada promoção via regex na ordem de precedência:

1. `combo` (compre X unidades ganhe Y)
2. `combo_preco` (preço de: R$ X / combo X por R$ Y)
3. `primeiro_pedido` (primeiro pedido / primeira compra)
4. `progressivo` (X itens Y% / progressiv*)
5. `lucro_extra` (lucro extra / X% LE / LE Boti / LE Eudora)
6. `volumetria` (X ou mais / X a Y itens)
7. `brinde` (junte embalagens / botirecicla)
8. `desconto_direto` (X% de desconto / até X%)
9. `outra` (fallback)

O HTML agrupa por Marca → `actionType` → subgrupo (Categoria / Linha /
Profundidade de desconto, alternável na toolbar).

## Estrutura do HTML gerado

- Self-contained (~1 MB), zero network dependency
- Hero com 5 cards de métrica (Visão geral + 4 marcas)
- Toolbar sticky: marca · curva · lançamentos · top 10 · categoria · fundo ·
  busca · ordenação (ações e produtos) · agrupar produtos por (Cat/Linha/Prof) ·
  expandir/recolher/limpar filtros
- Type sections colapsáveis (default fechado)
- Subgrupos colapsáveis (default fechado)
- Hover nativo do browser pra tooltip da descrição da promoção
- Linha pill inline por linha de produto
- Badge de profundidade do desconto: ≤15% verde · ≤30% amarelo · ≤40%
  vermelho · >40% magenta (cor derivada do **valor real da coluna desconto**)

## Sanity check rápido após gerar

1. Os counts nos cards do header batem com o número de produtos que tinham
   promoção na coluna foco.
2. O ciclo mostrado no hero é o esperado.
3. O count de OUI bate com o que você espera (cheque pelas sheets).
4. As seções de tipo aparecem na ordem canônica (DD, LE, Combo, Combo Preço,
   Progressivo, Volumetria, Primeiro Pedido, Brinde, Outra).
5. Cada promoção tem um `actionType` sensato (veja a tabela acima).
6. **Os badges mostram o valor da coluna desconto, não o "até X%" do texto**
   (só vale quando `template_patched.html` está em uso).

## Histórico de patches

- **2026-09-22 — discount-column-priority**: o template puro priorizava o
  `allDiscounts` (regex do texto) sobre o `discount` da coluna no badge
  inline. Corrigido: coluna > 0 vence; null/0 cai pro texto. Diff em
  `patches/2026-09-22_discount-column-priority.patch`. Aplicado em
  `template_patched.html`. Impacto: 743 promoções do ciclo 15/2026 estavam
  com badge errado (ex: Arbo Atlantica mostrava "DD 50%" em vez de "DD 25%").

## Onde a outra IA aprende mais

- `SKILL.md` — doc canônica da skill (mesmo conteúdo do upstream).
- `patches/2026-09-22_discount-column-priority.md` — changelog do fix.
- `examples/` — comandos de invocação reais + comparação de layouts de
  planilha.
- `run.py` — exemplo de wrapper Python que carrega o módulo V2 e chama
  `generate_html()` diretamente com template patcheado.
