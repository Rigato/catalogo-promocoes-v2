#!/usr/bin/env python3
"""
Generate a self-contained interactive HTML catalog of promotional actions
from a Boticário / O.U.i / Eudora / Quem Disse product spreadsheet.

This is the **V2** flavor. It produces the same page structure as v1 but with
structural-action classification (actionType) per promotion, which lets the
page group products by *tipo de promoção* (Desconto Direto, Lucro Extra,
Combo, Progressivo, Volumetria, etc.) instead of one card per action.

Usage:
    python3 generate.py <input.xlsx> [output.html]
    python3 generate.py <input.xlsx> --cycle 12/2026
    python3 generate.py <input.xlsx> --focus-column "Promoção Próximo Ciclo" \
                          --focus-discount-column "Desconto Promoção Próximo Ciclo"
    python3 generate.py --refresh-linhas <marcas.xlsx>

The output filename default is
    Catálogo de Promoções Ciclo CC/AAAA.html
where CC/AAAA is the focus cycle derived from the data (or the value
supplied via --cycle).

The first run on a new machine needs --refresh-linhas to seed the canonical
list of linhas (called MARCAS in the user's catalog spreadsheet). After
that, the list is read from the cache file
`.skills/ciclo-promocoes-v2/linhas_canonicas.json`. Re-run with
--refresh-linhas whenever the user provides a refreshed MARCAS xlsx.
"""

import json
import os
import re
import sys
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:
    sys.stderr.write(
        "This script needs openpyxl. Install with:\n"
        "  pip install openpyxl\n"
    )
    sys.exit(1)


# --------------------------------------------------------------------------------------
# 1. Constants
# --------------------------------------------------------------------------------------
# Column layout: we discover columns by header name (more robust against
# layout drift between cycles). The fallbacks below are used when the
# expected header isn't found.
FALLBACK_HIST_RANGE = (9, 29)  # V2 layout: columns I..AB

# Sheet-name → brand. V2 supports the older V1 names (with underscores and
# OUI nested in BOTICARIO) as well as the newer layout (O.U.I as its own
# sheet, QUEMDISSEBERENICE without underscores).
SHEET_TO_BRAND = {
    # Older layout (V1)
    "QUEM_DISSE_BERENICE": "QDB",
    "QUEMDISSEBERENICE":   "QDB",  # newer spelling (no underscores)
    "EUDORA":              "EUDORA",
    "BOTICARIO":           "BOTICARIO",
    "O.U.I":               "OUI",     # newer layout (separate sheet)
    "OUI":                 "OUI",     # alt name
}

# Header name candidates. The first match wins.
HDR_FOCUS_PROMO  = ("Promoção Próximo Ciclo + 1",)
HDR_FOCUS_DISC   = ("Desconto Promoção Próximo Ciclo + 1",)
# Projection column candidates. The default is "Projeção Próximo Ciclo + 2"
# (matches the V2 layout when the focus is "Próximo Ciclo + 1"). When the
# focus is the same slot as the projection (e.g. user wants the projection
# of the cycle being analysed), pass --projection-column to override.
HDR_PROJ_C2      = ("Projeção Próximo Ciclo + 1", "Projeção Próximo Ciclo + 2")
HDR_SKU          = ("SKU", "Cód.", "Codigo", "Código")
HDR_DESC         = ("Descrição", "Descricao", "Description")
HDR_CLASSE       = ("Classe",)
HDR_CATEGORIA    = ("Categoria",)
HDR_SUBCATEGORIA = ("Subcategoria",)
HDR_LANCAMENTO   = ("Lançamento", "Lancamento")
HDR_DESATIVACAO  = ("Desativação", "Desativacao")


# --------------------------------------------------------------------------------------
# 2. Action-type regex classifier (V2-specific feature)
# --------------------------------------------------------------------------------------
# Order matters! `combo_re` must be tested BEFORE `brinde_reciclagem` because
# combo descriptions often mention "junte" / "cupom" too.
#
# Precedence rationale (cycle 12/2026 lessons learned):
#   - "Lucro Extra Progressivo" / "Progressiva de X: 2 itens 15% | 3 ou mais 25%"
#     should match PROGRESSIVO, not LE / volumetria → PROGRESSIVO before both.
#   - "Compre 100 a 199... tenha 20% LE Boti" should match LUCRO_EXTRA, not
#     VOLUMETRIA (the "X a Y" / "X ou mais" patterns fire too greedily) →
#     LUCRO_EXTRA before VOLUMETRIA.
#   - "compre X unidades" alone (without a following LE/ganhe/desconto) → volumetria.
COMBO_RE = r"(?:ao\s+)?comprar?\s+(\d+)\s+(?:unidades?|itens?)[^.]*?ganhe\s+(\d+)\s+(?:unidade|item)"
COMPRE_PAGUE_LEVE = r"pague?\s+(\d+)[^.]*?leve\s+(\d+)"
PRECO_DE_POR = r"pre[çc]o\s+de[\s:]+\s*r\$|combo\s+\w+\s+por\s+r\$"
PRIMEIRO_PEDIDO = r"primeiro pedido|primeira compra|no primeiro pedido"
# Progressivo: explicit keyword "progressiv(a/o)" OR the classic "X itens Y%"
# pattern. Must come before lucro_extra / volumetria.
PROGRESSIVO = r"progressiv|\d+\s*itens?:?\s*\d+%"
# Lucro Extra: literal phrase OR "X% LE" (the corporate shorthand).
# The "LE" shorthand is a strong signal — when present, classify as LE even
# if the text also has "X a Y itens" (e.g. "Compre 100 a 199 itens e tenha
# 20% LE Boti" → LE, not volumetria).
LUCRO_EXTRA = r"lucro extra|\b\d+\s*%\s*le\b|\ble\s+boti\b|\ble\s+eudora\b|\ble\s+franqueado"
VOLUMETRIA_OU_MAIS = r"\d+\s*ou\s+mais\s*(?:unidades?|itens?)"
VOLUMETRIA = r"\d+\s*a\s*\d+\s*itens?|compre\s+\d+\s*a\s+\d+"
DESCONTO_DIRETO = r"\d+%\s+de\s+desconto|at[ée]\s+\d+%|desconto direto"
BRINDE_RECICLAGEM = r"junte.*embalagens|botirecicla|junte\s*\d+\s*embal"

ACTION_RE = [
    ("combo",        COMBO_RE),
    ("combo_preco",  PRECO_DE_POR),
    ("primeiro_pedido", PRIMEIRO_PEDIDO),
    ("progressivo",  PROGRESSIVO),
    ("lucro_extra",  LUCRO_EXTRA),
    ("volumetria",   VOLUMETRIA_OU_MAIS + "|" + VOLUMETRIA),
    ("brinde",       BRINDE_RECICLAGEM),
    ("desconto_direto", DESCONTO_DIRETO),
    ("combo",        COMPRE_PAGUE_LEVE),  # fallback X+Y pattern (pague X leve Y)
]

# Short codes per action type — used as inline badges on the product row
# (e.g. "DD 35%" for Desconto Direto, "LE 10%" for Lucro Extra).
ACTION_SHORT_CODES = {
    "desconto_direto":  "DD",
    "lucro_extra":      "LE",
    "combo":            "CB",
    "combo_preco":      "CP",
    "progressivo":      "PG",
    "volumetria":       "VL",
    "primeiro_pedido":  "PP",
    "brinde":           "BR",
    "outra":            "??",
}


def classify_action(desc: str) -> str:
    """Return one of: desconto_direto, lucro_extra, combo, combo_preco,
    progressivo, volumetria, primeiro_pedido, brinde, outra."""
    if not desc:
        return "outra"
    d = desc.lower()
    for name, pattern in ACTION_RE:
        if re.search(pattern, d):
            return name
    return "outra"


# --------------------------------------------------------------------------------------
# 3. Combo detection (for "X+Y ≈ NN% implícito" badge)
# --------------------------------------------------------------------------------------
COMBO_PATTERNS = [
    r"(?:ao\s+)?comprar?\s+(\d+)\s+(?:unidades?|itens?)[^.]*?ganhe\s+(\d+)\s+(?:unidade|item)",
    r"compre?\s+(\d+)\s+(?:unidades?|itens?)[^.]*?ganhe\s+(\d+)",
    r"compre?\s+(\d+)[^.]*?leve\s+(\d+)",
    r"pague?\s+(\d+)[^.]*?leve\s+(\d+)",
]


def detect_combo(desc):
    """Look for 'compre X, ganhe Y' / 'pague X leve Y' patterns and compute
    the implied discount %. Returns None if no combo found."""
    for pat in COMBO_PATTERNS:
        m = re.search(pat, (desc or "").lower())
        if not m:
            continue
        try:
            x, y = int(m.group(1)), int(m.group(2))
        except (ValueError, IndexError):
            continue
        if y >= 1 and x >= 2 and y < x:
            disc = round(y / (x + y) * 100, 1)
            return {
                "is_combo": True,
                "X": x,
                "Y": y,
                "pattern": f"{x}+{y}",
                "disc_pct": disc,
                "display": f"{x}+{y}",
            }
    return None


# --------------------------------------------------------------------------------------
# 3.5 Extract all discount percentages mentioned in a description
# --------------------------------------------------------------------------------------
# A single action description can mention multiple %s (e.g. "Lucro Extra
# Progressivo: 8% em 2 a 3 unidades, 10% em 4 ou mais"). For inline badges
# we want to surface ALL of them with the surrounding context, so the user
# can see the ladder of tiers at a glance.
_PCT_FIND_RE = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*%",
    flags=re.IGNORECASE,
)
# Words that should NOT be counted as a discount tier (e.g. "100% de lucro"
# in a brinde is a context indicator, not a discount).
_PCT_STOPWORDS = {"100", "100.0"}


def _parse_pct(s: str):
    """Parse '12', '12,5', '12.5' as float. Returns None on failure."""
    try:
        return float(s.replace(",", "."))
    except (ValueError, AttributeError):
        return None


def extract_discounts(desc):
    """Return a list of discount objects mentioned in `desc`, each shaped:
        {"pct": <int>, "raw": "12,5%", "ctx": "8% de Lucro Extra em 2 a 3 unid"}
    Duplicates are preserved (so the user can see "10% / 10%" if mentioned
    twice), and percentages > 60 are filtered out (those are usually
    conditions like "100% de lucro" on a brinde, not the discount itself).
    Percentages inside parentheses like "(desconto adicional 5,56%)" are
    skipped — those are explanatory decimals, not the headline tier.
    """
    if not desc:
        return []
    # Build a list of (start, end) ranges that are inside parens; we'll
    # skip any % matches that fall inside one of those ranges.
    paren_ranges = []
    depth = 0
    open_idx = None
    for i, ch in enumerate(desc):
        if ch == "(":
            if depth == 0:
                open_idx = i
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0 and open_idx is not None:
                paren_ranges.append((open_idx, i))
                open_idx = None

    def inside_paren(pos):
        for (s, e) in paren_ranges:
            if s < pos <= e:
                return True
        return False

    out = []
    seen_positions = set()
    for m in _PCT_FIND_RE.finditer(desc):
        if m.start() in seen_positions:
            continue
        seen_positions.add(m.start())
        pct = _parse_pct(m.group(1))
        if pct is None:
            continue
        raw = m.group(0)
        if m.group(1) in _PCT_STOPWORDS:
            continue
        if inside_paren(m.start()):
            continue
        # Capture a short context window around the match (40 chars each side)
        s = max(0, m.start() - 40)
        e = min(len(desc), m.end() + 40)
        ctx = desc[s:e].strip()
        if s > 0:
            ctx = "…" + ctx
        if e < len(desc):
            ctx = ctx + "…"
        out.append({
            "pct": pct,
            "raw": raw,
            "ctx": ctx,
        })
    return out


# --------------------------------------------------------------------------------------
# 4. OUI detection + brand assignment
# --------------------------------------------------------------------------------------
def is_oui(desc):
    if not desc:
        return False
    d = desc.strip().upper()
    return (
        d.startswith("OUI ")
        or d.startswith("OUI\t")
        or d.startswith("ESTJ OUI ")
        or d.startswith("EST OUI ")
        or d.startswith("REF OUI ")
    )


def get_brand(sheet_name, desc):
    """Map an xlsx sheet name to a brand label. For the older V1 layout
    where OUI products live inside BOTICARIO, we re-classify by description
    prefix; newer layouts use O.U.I as its own sheet."""
    brand = SHEET_TO_BRAND.get(sheet_name)
    if brand is not None:
        # Older V1 layout had OUI products inside BOTICARIO. If we're using
        # the newer dedicated O.U.I sheet, brand is already "OUI" — no extra
        # work needed.
        if sheet_name == "BOTICARIO" and is_oui(desc):
            return "OUI"
        return brand
    return sheet_name


# --------------------------------------------------------------------------------------
# 5. Promotion normalisation (strip reseller tiers and prefixes)
# --------------------------------------------------------------------------------------
TIER_REGEX = r"revendedor(?:es)?\s+(cobre|bronze|prata|ouro|platina|rubi|esmeralda|diamante)[\s:,]+"
CATEGORIA_REGEX = r"\bcategoria(?:\s+de\s+cadastro)?(?:\s+(?:da|do))?\s*"
CADASTRO_REGEX = r"\bcadastro(?:\s+da\s+categoria)?\s*"


def normalize_action_desc(desc: str) -> str:
    """Strip tier prefixes and normalization noise so similar actions merge."""
    if not desc:
        return ""
    d = desc
    # Remove reseller-tier prefixes like "Revendedor Cobre/Bronze/..." when
    # followed by a colon or comma in the description.
    d = re.sub(TIER_REGEX, "", d, flags=re.IGNORECASE)
    # Strip "Categoria [de cadastro] X" / "Cadastro [da categoria] X" patterns.
    d = re.sub(CATEGORIA_REGEX, "", d, flags=re.IGNORECASE)
    d = re.sub(CADASTRO_REGEX, "", d, flags=re.IGNORECASE)
    # Collapse whitespace and trim trailing punctuation.
    d = re.sub(r"\s+", " ", d).strip()
    d = re.sub(r"[.,;:]+$", "", d)
    return d


# --------------------------------------------------------------------------------------
# 6. Helpers
# --------------------------------------------------------------------------------------
def is_oui_product(marca, desc):
    return marca == "OUI"


def parse_cycle_discount(line):
    """Parse a single line from the focus column. Returns dict or None."""
    line = line.strip()
    if not line:
        return None
    m = re.match(r"^(20\d{4})\s*-\s*(\d+)%\s*-\s*(.+)$", line, re.DOTALL)
    if m:
        cycle = m.group(1)
        discount = int(m.group(2))
        desc = m.group(3).strip()
    else:
        # Try without leading discount % (e.g. "202612 - Junte embalagens...")
        m = re.match(r"^(20\d{4})\s*-\s*(.+)$", line, re.DOTALL)
        if not m:
            return None
        cycle = m.group(1)
        desc = m.group(2).strip()
        discount = None

    return {
        "cycle": cycle,
        "cycle_label": f"{cycle[-2:]}/{cycle[:4]}",
        "discount": discount,
        "description": desc,
        "raw": line,
    }


# --------------------------------------------------------------------------------------
# 7. Main: read workbook + emit data.json
# --------------------------------------------------------------------------------------
def _norm_header(s):
    return (s or "").strip().lower()


def _find_col(headers, candidates, fallback=None):
    """Find the 0-based column index matching one of the candidate header
    names (case-insensitive, whitespace-tolerant). Returns None if not found."""
    norm = [_norm_header(h) for h in headers]
    for cand in candidates:
        c = _norm_header(cand)
        if c in norm:
            return norm.index(c)
    return fallback


def _find_hist_cols(headers, fallback_range):
    """Return all 0-based column indexes whose header starts with the
    historical-cycle pattern. If none found, fall back to a fixed range."""
    cols = []
    for i, h in enumerate(headers):
        if _norm_header(h).startswith("histórico de vendas do ciclo") or \
           _norm_header(h).startswith("historico de vendas do ciclo"):
            cols.append(i)
    if cols:
        return cols
    lo, hi = fallback_range
    return list(range(lo, hi))


def load_products_from_xlsx(xlsx_path, override_cycle=None, cycle_source='override',
                             focus_column=None, focus_discount_column=None,
                             projection_column=None):
    """cycle_source: 'override' (user passed --cycle) or 'auto' (auto-detected).
    focus_column / focus_discount_column: optional header names that override
    the default HDR_FOCUS_PROMO / HDR_FOCUS_DISC. Used to point the skill at a
    different cycle's column (e.g. "Promoção Próximo Ciclo" instead of
    "Promoção Próximo Ciclo + 1").
    projection_column: optional header name to override the default projection
    column (HDR_PROJ_C2). The default tries "Projeção Próximo Ciclo + 1"
    first, then falls back to "Projeção Próximo Ciclo + 2".
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    products = []
    if override_cycle:
        cycle_label = f"{override_cycle[-2:]}/{override_cycle[:4]}"
        msg = (
            f"  (using --cycle override: {cycle_label})"
            if cycle_source == 'override'
            else f"  (cycle auto-detected: {cycle_label})"
        )
        print(msg)
    if focus_column or focus_discount_column:
        print(
            f"  (using custom focus columns: desc={focus_column or HDR_FOCUS_PROMO[0]!r}, "
            f"disc={focus_discount_column or HDR_FOCUS_DISC[0]!r})"
        )
    # Sheets to process, in display order. Only the ones present in the
    # workbook are used.
    sheet_candidates = [
        "QUEM_DISSE_BERENICE", "QUEMDISSEBERENICE",
        "EUDORA",
        "BOTICARIO",
        "O.U.I", "OUI",
    ]
    for sheet_name in sheet_candidates:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))

        # Discover columns by header (fall back to the V1 indexes if absent)
        idx_sku          = _find_col(headers, HDR_SKU,          fallback=3)
        idx_desc         = _find_col(headers, HDR_DESC,         fallback=4)
        idx_categoria    = _find_col(headers, HDR_CATEGORIA,    fallback=5)
        idx_subcategoria = _find_col(headers, HDR_SUBCATEGORIA, fallback=6)
        idx_lancamento   = _find_col(headers, HDR_LANCAMENTO,   fallback=7)
        idx_desativacao  = _find_col(headers, HDR_DESATIVACAO,  fallback=8)
        idx_classe       = _find_col(headers, HDR_CLASSE,       fallback=1)
        idx_focus        = _find_col(headers, (focus_column,) if focus_column else HDR_FOCUS_PROMO, fallback=32)
        idx_focus_disc   = _find_col(headers, (focus_discount_column,) if focus_discount_column else HDR_FOCUS_DISC,  fallback=None)
        # Projection column: prefer user override, then try the candidates in
        # HDR_PROJ_C2 order. Fall back to the +2 index for older layouts.
        if projection_column:
            proj_candidates = (projection_column,)
        else:
            proj_candidates = HDR_PROJ_C2 if isinstance(HDR_PROJ_C2, tuple) else (HDR_PROJ_C2,)
        idx_proj_c2      = _find_col(headers, proj_candidates, fallback=30)
        hist_cols        = _find_hist_cols(headers, FALLBACK_HIST_RANGE)

        # Derive the current cycle from the "(atual)" historical column.
        current_cycle = "202610"
        for h in headers:
            if _norm_header(h).endswith("(atual)"):
                m = re.search(r"(\d{6})", h)
                if m:
                    current_cycle = m.group(1)
                    break
        # The focus column in the V2 layout ("Promoção Próximo Ciclo + 1")
        # refers to the cycle that is two steps ahead of the historical
        # "(atual)" reference cycle. The convention is:
        #   "(atual)"        = the last fully-closed cycle
        #   "Próximo Ciclo"  = the cycle being planned right now
        #   "Próximo Ciclo + 1" = the cycle after that — this is what the
        #                          focus column is for.
        cc_year = int(current_cycle[:4])
        cc_mon  = int(current_cycle[4:6])
        next_mon = cc_mon + 2
        next_year = cc_year
        while next_mon > 12:
            next_mon -= 12
            next_year += 1
        default_focus_cycle = f"{next_year}{next_mon:02d}"
        if override_cycle:
            default_focus_cycle = override_cycle

        # classe_seg is optional in newer layouts; default to empty
        idx_classe_seg = _find_col(headers, ("Classe Segmentada", "ClasseSeg"), fallback=2)

        # Normalize to brand via sheet name; OUI is detected inside BOTICARIO
        # in the V1 layout (older xlsx) but is its own sheet in newer layouts.
        brand = get_brand(sheet_name, "")  # desc is irrelevant for newer sheet-based detection

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or idx_sku is None or idx_sku >= len(row):
                continue
            sku_cell = row[idx_sku]
            if sku_cell is None or sku_cell == "":
                continue
            sku = str(int(sku_cell)) if isinstance(sku_cell, (int, float)) else str(sku_cell)
            desc = (row[idx_desc] or "").strip() if idx_desc is not None and idx_desc < len(row) else ""
            if not desc:
                continue
            # Re-apply OUI detection for the older V1 layout (OUI nested
            # inside BOTICARIO) so description-prefix re-classification still
            # works there.
            if sheet_name == "BOTICARIO" and is_oui(desc):
                brand = "OUI"
            categoria = (row[idx_categoria] or "").strip() if idx_categoria is not None and idx_categoria < len(row) else ""
            if not categoria:
                categoria = "(SEM CATEGORIA)"
            subcategoria = (row[idx_subcategoria] or "").strip() if idx_subcategoria is not None and idx_subcategoria < len(row) else ""
            if not subcategoria:
                subcategoria = categoria
            classe = (row[idx_classe] or "").strip() if idx_classe is not None and idx_classe < len(row) else "-"
            classe_seg = (row[idx_classe_seg] or "").strip() if idx_classe_seg is not None and idx_classe_seg < len(row) else ""

            lancamento = row[idx_lancamento] if idx_lancamento is not None and idx_lancamento < len(row) else None
            if hasattr(lancamento, "strftime"):
                lancamento = lancamento.strftime("%Y-%m-%d")
            desativacao = row[idx_desativacao] if idx_desativacao is not None and idx_desativacao < len(row) else None
            if hasattr(desativacao, "strftime"):
                desativacao = desativacao.strftime("%Y-%m-%d")

            proj_c2 = row[idx_proj_c2] if idx_proj_c2 is not None and idx_proj_c2 < len(row) else 0
            if proj_c2 is None or proj_c2 == "":
                proj_c2 = 0

            hist_vals = []
            for c in hist_cols:
                if c < len(row):
                    v = row[c]
                    if isinstance(v, (int, float)) and v > 0:
                        hist_vals.append(v)
            max_hist = max(hist_vals) if hist_vals else 0

            prom_cell = row[idx_focus] if idx_focus is not None and idx_focus < len(row) else None
            disc_cell = row[idx_focus_disc] if idx_focus_disc is not None and idx_focus_disc < len(row) else None
            promotions = []
            if prom_cell and isinstance(prom_cell, str):
                # Normalize line endings (newer xlsx uses \r\n)
                lines = [ln.rstrip("\r").strip() for ln in prom_cell.split("\n")]
                lines = [ln for ln in lines if ln]

                # V1 layout: each line encodes cycle+discount+description.
                if any(" - " in ln and re.match(r"^\d{6}\s*-", ln) for ln in lines):
                    for line in lines:
                        parsed = parse_cycle_discount(line)
                        if parsed is None:
                            continue
                        atype = classify_action(parsed["description"])
                        parsed["actionType"] = atype
                        parsed["combo"] = detect_combo(parsed["description"])
                        parsed["shortCode"] = ACTION_SHORT_CODES.get(atype, "??")
                        parsed["allDiscounts"] = extract_discounts(parsed["description"])
                        promotions.append(parsed)
                else:
                    # V2 layout: discount is in a parallel column (col AK),
                    # description is in the focus column (col AL).
                    disc_lines = []
                    if disc_cell and isinstance(disc_cell, str):
                        disc_lines = [v.rstrip("\r").strip() for v in disc_cell.split("\n")]
                        disc_lines = [v for v in disc_lines if v]
                    elif isinstance(disc_cell, (int, float)):
                        disc_lines = [str(disc_cell)]
                    if not disc_lines:
                        disc_lines = ["0"] * len(lines)
                    if len(disc_lines) < len(lines):
                        disc_lines += ["0"] * (len(lines) - len(disc_lines))
                    disc_lines = disc_lines[:len(lines)]

                    for line, disc_str in zip(lines, disc_lines):
                        if not line:
                            continue
                        try:
                            disc_val = int(float(disc_str))
                        except (ValueError, TypeError):
                            m = re.search(r"(\d+)\s*%", line)
                            disc_val = int(m.group(1)) if m else None
                        # Default cycle = +1 from current. The current cycle
                        # is taken from the "atual" histórico header (e.g. 202610),
                        # so +1 = 202611. Override if the description carries
                        # a "ciclo MM/YYYY" reference.
                        cycle = default_focus_cycle
                        m_cycle = re.search(r"ciclo\s+(\d{1,2})/(\d{2,4})", line.lower())
                        if m_cycle:
                            mm, yy = m_cycle.groups()
                            if len(yy) == 2:
                                cycle = f"20{yy}{int(mm):02d}"
                            else:
                                cycle = f"{yy}{int(mm):02d}"
                        atype = classify_action(line)
                        promotions.append({
                            "cycle": cycle,
                            "cycle_label": f"{cycle[-2:]}/{cycle[:4]}",
                            "discount": disc_val,
                            "description": line,
                            "raw": line,
                            "actionType": atype,
                            "shortCode": ACTION_SHORT_CODES.get(atype, "??"),
                            "combo": detect_combo(line),
                            "allDiscounts": extract_discounts(line),
                        })
            if not promotions:
                continue  # skip products without any promotion

            products.append(
                {
                    "sku": sku,
                    "descricao": desc,
                    "marca": brand,
                    "categoria": categoria,
                    "subcategoria": subcategoria,
                    "linha": "",  # V2 derives linha from description in JS at runtime
                    "classe": classe,
                    "classe_seg": classe_seg,
                    "lancamento": str(lancamento) if lancamento else "",
                    "desativacao": str(desativacao) if desativacao else "",
                    "max_hist": max_hist,
                    "proj_c2": proj_c2,
                    "promotions": promotions,
                }
            )
    if not products:
        print("  WARN: no products parsed. Check the sheet names and focus column header.", file=sys.stderr)
    return products


def build_data_json(products):
    return {"products": products}


# --------------------------------------------------------------------------------------
# 7.5 Canonical list of linhas (loaded from a cache file)
# --------------------------------------------------------------------------------------
LINHAS_CACHE_FILENAME = "linhas_canonicas.json"

# Built-in abbreviation map for the common cases where the description's
# first word is a known short form (e.g. "CBEM" → "CUIDE-SE BEM"). The
# canonical list still wins when its first word matches exactly; this map
# only kicks in when there's no direct first-word match.
ABBREVIATION_MAP = {
    "CBEM":    "CUIDE-SE BEM",
    "NSPA":    "NATIVA SPA",
    "DR":      "DR BOTICA",
    "HER":     "HER CODE",
    "EL":      "ELYSEE",
    "ELY":     "ELYSEE",
    "MAKE":    "MAKE B.",
    "LYZ":     "LIZ",
    "LIZ":     "LIZ",
    "BOTI":    None,            # BOTI alone isn't a linha; needs BABY/HOME/SUN
}


def _linhas_cache_path():
    skill_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(skill_dir, LINHAS_CACHE_FILENAME)


def refresh_linhas(marcas_xlsx_path):
    """Read the user's catalog spreadsheet (sheet 'portfolio', col 'Marca')
    and persist the unique list of linha names to a local cache file.

    Only the names are stored; SKU / price / sortimento data is discarded
    because it changes often and would bloat the cache. Run this whenever
    the user gives a refreshed MARCAS xlsx.
    """
    if not os.path.exists(marcas_xlsx_path):
        sys.stderr.write(f"MARCAS xlsx not found: {marcas_xlsx_path}\n")
        sys.exit(1)
    wb = openpyxl.load_workbook(marcas_xlsx_path, data_only=True, read_only=True)
    if "portfolio" not in wb.sheetnames:
        sys.stderr.write(
            f"Sheet 'portfolio' not found in {marcas_xlsx_path} "
            f"(got {wb.sheetnames})\n"
        )
        sys.exit(1)
    ws = wb["portfolio"]
    headers = None
    for r in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        headers = list(r)
        break
    if not headers or "Marca" not in headers:
        sys.stderr.write("Column 'Marca' not found in portfolio sheet\n")
        sys.exit(1)
    marca_idx = headers.index("Marca")
    seen = []
    seen_set = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if marca_idx < len(row):
            v = row[marca_idx]
            if v and v not in seen_set:
                seen.append(v)
                seen_set.add(v)
    out_path = _linhas_cache_path()
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(seen, f, ensure_ascii=False, indent=2)
    print(f"Refreshed canonical linhas cache → {out_path}")
    print(f"  {len(seen)} unique names: {', '.join(seen[:8])}{'…' if len(seen) > 8 else ''}")
    return seen


def load_linhas_cache():
    """Return the canonical list of linhas, refreshing from a sibling
    portfolio.xlsx if the cache is missing and the file is present."""
    p = _linhas_cache_path()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                cached = json.load(f)
            if isinstance(cached, list) and cached:
                return cached
        except (OSError, ValueError):
            pass
    # Auto-bootstrap: look for a portfolio.xlsx in the same directory as the
    # generate.py file.
    skill_dir = os.path.dirname(os.path.abspath(__file__))
    sibling = os.path.join(skill_dir, "portfolio.xlsx")
    if os.path.exists(sibling):
        print(f"  (bootstrapping linhas cache from {sibling})")
        return refresh_linhas(sibling)
    # Last resort: empty list (skill still works, derivation just doesn't get
    # the canonical-name normalization).
    return []


# --------------------------------------------------------------------------------------
# 8. Emit HTML by injecting data into template.html
# --------------------------------------------------------------------------------------
# Friendly label for a focus-column pair, used in the page <title>.
_FOCUS_SLOT_LABELS = {
    ("Promoção Ciclo Atual", "Desconto Promoção Ciclo Atual"):       "Ciclo Atual",
    ("Promoção Próximo Ciclo", "Desconto Promoção Próximo Ciclo"):     "Próximo Ciclo",
    ("Promoção Próximo Ciclo + 1", "Desconto Promoção Próximo Ciclo + 1"): "Próximo Ciclo + 1",
    ("Promoção Próximo Ciclo + 2", "Desconto Promoção Próximo Ciclo + 2"): "Próximo Ciclo + 2",
}


def _cycle_slot_label(focus_column, focus_discount_column):
    """Return a human-friendly slot name like 'Ciclo Atual' / 'Próximo Ciclo + 1'.
    Returns None if the column pair doesn't match a known slot (i.e. user
    passed a custom pair)."""
    if not focus_column or not focus_discount_column:
        return None
    return _FOCUS_SLOT_LABELS.get((focus_column, focus_discount_column))


def generate_html(xlsx_path, output_path, template_path=None, override_cycle=None,
                  cycle_source='override', canonical_linhas=None,
                  focus_column=None, focus_discount_column=None,
                  projection_column=None):
    if template_path is None:
        skill_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(skill_dir, "template.html")
    if not os.path.exists(template_path):
        sys.stderr.write(f"Template not found: {template_path}\n")
        sys.exit(1)
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()
    if "__DATA_PLACEHOLDER__" not in template:
        sys.stderr.write(
            "Template is missing the __DATA_PLACEHOLDER__ token. "
            "Add it where the DATA constant lives.\n"
        )
        sys.exit(1)

    products = load_products_from_xlsx(xlsx_path, override_cycle=override_cycle,
                                       cycle_source=cycle_source,
                                       focus_column=focus_column,
                                       focus_discount_column=focus_discount_column,
                                       projection_column=projection_column)
    data = build_data_json(products)
    # Embed the canonical linhas list as a JS const so the page can use
    # it for derivation without making any network calls.
    if canonical_linhas is None:
        canonical_linhas = load_linhas_cache()
    data["_meta"] = {
        "canonical_linhas": canonical_linhas,
        "abbreviation_map": ABBREVIATION_MAP,
    }

    json_str = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    out_html = template.replace("__DATA_PLACEHOLDER__", json_str, 1)

    # Replace the <title> tag with a cycle-aware one. Default template has
    # "Catálogo de Promoções — Próximo Ciclo + 1"; we substitute with the
    # focus cycle label (e.g. "Ciclo 12/2026") and the slot name when known.
    title_cycle_label = f"{override_cycle[-2:]}/{override_cycle[:4]}" if override_cycle else "?"
    slot_label = _cycle_slot_label(focus_column, focus_discount_column)
    if slot_label:
        new_title = f"Catálogo de Promoções — {slot_label} ({title_cycle_label})"
    else:
        new_title = f"Catálogo de Promoções — {title_cycle_label}"
    out_html = re.sub(
        r"<title>.*?</title>",
        f"<title>{new_title}</title>",
        out_html,
        count=1,
        flags=re.DOTALL,
    )

    # Create any missing parent directories (e.g. 'Catálogo de Promoções
    # Ciclo 12/' when the output name has a "/" in the cycle).
    parent = os.path.dirname(output_path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(out_html)

    # Print summary
    brands = Counter(p["marca"] for p in products)
    cycles = Counter()
    types = Counter()
    for p in products:
        for pr in p["promotions"]:
            cycles[pr["cycle"]] += 1
            types[pr["actionType"]] += 1
    print(f"Wrote {output_path}")
    print(f"  Products with promotions: {len(products)}")
    for b in ["BOTICARIO", "EUDORA", "OUI", "QDB"]:
        print(f"    {b}: {brands.get(b, 0)}")
    print(f"  Cycle participations: {dict(cycles)}")
    print(f"  Action types: {dict(types)}")
    print(f"  Canonical linhas loaded: {len(canonical_linhas)}")


# --------------------------------------------------------------------------------------
# 9. CLI
# --------------------------------------------------------------------------------------
def _parse_cycle_arg(s):
    """Accepts '12/2026' or '12/26' and returns '202612'."""
    m = re.match(r"^\s*(\d{1,2})/(\d{2,4})\s*$", s)
    if not m:
        return None
    mm, yy = m.groups()
    if len(yy) == 2:
        yy = "20" + yy
    return f"{yy}{int(mm):02d}"


def _default_output_path(xlsx_path, focus_cycle):
    """Returns 'Catálogo de Promoções Ciclo CC-AAAA.html' next to the input
    xlsx, where CC is the cycle number (e.g. 12) and AAAA the year (e.g. 2026).

    We use '-' (not '/') between the cycle and the year so the result is a
    single file on disk; '/' would be interpreted as a path separator by the
    OS and split the name into a directory + file (the user then sees only the
    year as the filename).
    """
    cycle_num = focus_cycle[-2:]
    year = focus_cycle[:4]
    base = os.path.dirname(os.path.abspath(xlsx_path))
    return os.path.join(base, f"Catálogo de Promoções Ciclo {cycle_num}-{year}.html")


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__, file=sys.stderr)
        sys.exit(1)

    # Subcommand: --refresh-linhas <xlsx>
    if args[0] == "--refresh-linhas":
        if len(args) < 2:
            sys.stderr.write("Usage: --refresh-linhas <marcas.xlsx>\n")
            sys.exit(1)
        refresh_linhas(os.path.abspath(args[1]))
        return

    # Parse flags
    xlsx_arg = None
    override_cycle = None
    output_path = None
    focus_column = None
    focus_discount_column = None
    projection_column = None
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--cycle" and i + 1 < len(args):
            override_cycle = _parse_cycle_arg(args[i + 1])
            if not override_cycle:
                sys.stderr.write(f"Invalid --cycle value: {args[i+1]} (expected MM/AAAA or MM/AA)\n")
                sys.exit(1)
            i += 2
        elif a == "--focus-column" and i + 1 < len(args):
            focus_column = args[i + 1]
            i += 2
        elif a == "--focus-discount-column" and i + 1 < len(args):
            focus_discount_column = args[i + 1]
            i += 2
        elif a == "--projection-column" and i + 1 < len(args):
            projection_column = args[i + 1]
            i += 2
        elif a == "--refresh-linhas":
            sys.stderr.write("--refresh-linhas must be the first argument (alone)\n")
            sys.exit(1)
        elif a.startswith("--"):
            sys.stderr.write(f"Unknown flag: {a}\n")
            sys.exit(1)
        else:
            if xlsx_arg is None:
                xlsx_arg = a
            else:
                output_path = a
            i += 1

    if xlsx_arg is None:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    xlsx_path = os.path.abspath(xlsx_arg)
    if not os.path.exists(xlsx_path):
        sys.stderr.write(f"Input not found: {xlsx_path}\n")
        sys.exit(1)

    # Auto-derive focus cycle from the data if the user didn't override it
    canonical_linhas = load_linhas_cache()
    cycle_source = 'override'
    if override_cycle is None:
        override_cycle = _detect_focus_cycle(xlsx_path)
        cycle_source = 'auto'

    if output_path is None:
        output_path = _default_output_path(xlsx_path, override_cycle)

    print(f"Reading {xlsx_path} ...")
    generate_html(xlsx_path, output_path,
                  override_cycle=override_cycle,
                  cycle_source=cycle_source,
                  canonical_linhas=canonical_linhas,
                  focus_column=focus_column,
                  focus_discount_column=focus_discount_column,
                  projection_column=projection_column)


def _detect_focus_cycle(xlsx_path):
    """Detect the focus cycle from the headers of the input xlsx.
    Returns YYYYMM (e.g. '202612')."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    try:
        for sn in wb.sheetnames:
            ws = wb[sn]
            for r in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                headers = list(r)
                break
            if not headers:
                continue
            # Find the "(atual)" historical cycle
            current_cycle = None
            for h in headers:
                if h and _norm_header(h).endswith("(atual)"):
                    m = re.search(r"(\d{6})", h)
                    if m:
                        current_cycle = m.group(1)
                        break
            if not current_cycle:
                continue
            # The focus column "Promoção Próximo Ciclo + 1" refers to
            # current + 2.
            cc_year = int(current_cycle[:4])
            cc_mon  = int(current_cycle[4:6])
            mon = cc_mon + 2
            yr  = cc_year
            while mon > 12:
                mon -= 12
                yr  += 1
            return f"{yr}{mon:02d}"
    finally:
        wb.close()
    return None


if __name__ == "__main__":
    main()
