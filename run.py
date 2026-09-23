#!/usr/bin/env python3
"""Wrapper that runs the V2 generate.py with a custom template path."""
import sys
sys.path.insert(0, "/workspace/.skills/ciclo-promocoes-v2")

# Import the V2 module
import generate as v2

# Override generate_html call with our patched template
xlsx = "/workspace/attachments/d5bb5c83e14100fe/DRAFT_PDVS_SEM_13259_20260922081036.xlsx"
output = "/workspace/ciclo15_2026_v2_patched/index.html"
template = "/workspace/ciclo15_2026_v2_patched/template.html"

# Auto-detect cycle from xlsx (same as default behavior)
cycle = v2._detect_focus_cycle(xlsx)
print(f"Auto-detected cycle from xlsx: {cycle}")
# Override to 15/2026
cycle = "202615"
print(f"Using override cycle: {cycle} (15/2026)")

# Load canonical linhas
canonical_linhas = v2.load_linhas_cache()

print(f"Reading {xlsx}...")
v2.generate_html(
    xlsx_path=xlsx,
    output_path=output,
    template_path=template,
    override_cycle=cycle,
    cycle_source="override",
    canonical_linhas=canonical_linhas,
    focus_column="Promoção Próximo Ciclo + 1",
    focus_discount_column="Desconto Promoção Próximo Ciclo + 1",
    projection_column="Projeção Próximo Ciclo + 2",
)
