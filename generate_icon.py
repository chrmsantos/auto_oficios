"""
generate_icon.py — Gera o ícone icon.ico para o Auto Ofícios.

Conceito visual:
  - Fundo quadrado arredondado (azul-escuro do tema da app)
  - Documento/papel com canto dobrado (representando o ofício)
  - Raio dourado em destaque (representando automação)

Execute:  python generate_icon.py
Requer:   pip install pillow
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


# ── Cores da paleta do app (ui.py) ─────────────────────────────────────────
BG      = (15,  17,  26, 255)   # #0f111a  fundo principal
CARD    = (26,  29,  46, 255)   # #1a1d2e  cartão / fundo do ícone
BORDER  = (46,  49,  80, 200)   # #2e3150  borda suave
PAPER   = (210, 218, 240, 255)  # papel levemente azulado
PAPER_F = (158, 172, 210, 255)  # dobra do papel (mais escura)
LINE    = (120, 138, 185, 140)  # linhas de texto (semitransparente)
BOLT    = (255, 180,  84, 255)  # #ffb454 dourado — raio
BOLT_S  = (200, 120,  30, 180)  # sombra do raio


def _rr(draw: ImageDraw.ImageDraw, xy, radius: int, fill, outline=None, width=1):
    """Rounded rectangle helper (PIL <9.2 compat)."""
    try:
        draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)
    except AttributeError:
        x0, y0, x1, y1 = xy
        draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
        draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
        draw.ellipse([x0, y0, x0 + 2*radius, y0 + 2*radius], fill=fill)
        draw.ellipse([x1 - 2*radius, y0, x1, y0 + 2*radius], fill=fill)
        draw.ellipse([x0, y1 - 2*radius, x0 + 2*radius, y1], fill=fill)
        draw.ellipse([x1 - 2*radius, y1 - 2*radius, x1, y1], fill=fill)


def _scale(pts, s):
    """Scale a list of (x, y) tuples by factor s."""
    return [(int(x * s), int(y * s)) for x, y in pts]


def draw_frame(size: int) -> Image.Image:
    """Render one icon frame at the given pixel size."""
    s = size / 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # ── Background rounded square ──────────────────────────────────────────
    pad = max(1, int(4 * s))
    _rr(d, [pad, pad, size - pad, size - pad],
        radius=int(44 * s), fill=CARD, outline=BORDER, width=max(1, int(2 * s)))

    # ── Document (papel) ──────────────────────────────────────────────────
    # Rectangle with a dog-eared top-right corner
    dx1, dy1 = int(56 * s), int(34 * s)
    dx2, dy2 = int(188 * s), int(218 * s)
    fold = int(34 * s)

    doc_body = [
        (dx1,          dy1),
        (dx2 - fold,   dy1),
        (dx2,          dy1 + fold),
        (dx2,          dy2),
        (dx1,          dy2),
    ]
    d.polygon(doc_body, fill=PAPER)

    # Fold triangle (darker corner)
    fold_tri = [
        (dx2 - fold,   dy1),
        (dx2,          dy1 + fold),
        (dx2 - fold,   dy1 + fold),
    ]
    d.polygon(fold_tri, fill=PAPER_F)

    # ── Text lines on document ────────────────────────────────────────────
    lx1 = dx1 + int(13 * s)
    lx2 = dx2 - int(9 * s)
    lx2_short = dx2 - fold - int(4 * s)   # short line to avoid fold area
    lw = max(1, int(4 * s))
    lines = [
        (int(82  * s), lx1, lx2_short),
        (int(108 * s), lx1, lx2),
        (int(134 * s), lx1, lx2),
        (int(160 * s), lx1, int(lx1 + (lx2 - lx1) * 0.60)),
    ]
    for ly, lxA, lxB in lines:
        if lxB > lxA and size >= 32:
            d.line([(lxA, ly), (lxB, ly)], fill=LINE, width=lw)

    # ── Lightning bolt ────────────────────────────────────────────────────
    # Classic zigzag bolt: two parallelogram halves meeting at a kink.
    # Designed on 256 grid; scaled by s.
    #
    #        a──b
    #       /    \
    #      /      c──d   ← kink (overhang right)
    #     h──g    |  |
    #     |   \   /  |
    #     |    e──f
    #     |
    #  (h to a closes the polygon)
    #
    # Upper half: top-left, top-right, kink-right-top, kink-right-bottom
    # Lower half: kink-left-bottom, kink-left-top, bottom-right, bottom-left

    bolt_raw = [
        (124, 22),   # a  top-left
        (198, 22),   # b  top-right
        (148, 128),  # c  kink — right-upper
        (206, 128),  # d  kink — far-right overhang
        (126, 242),  # e  bottom-right
        (60,  242),  # f  bottom-left
        (104, 138),  # g  kink — left-lower
        (52,  138),  # h  kink — far-left overhang
    ]
    bolt_pts = _scale(bolt_raw, s)

    # Shadow (slightly offset, for depth)
    if size >= 32:
        off = max(2, int(4 * s))
        bolt_shadow = [(x + off, y + off) for x, y in bolt_pts]
        d.polygon(bolt_shadow, fill=BOLT_S)

    d.polygon(bolt_pts, fill=BOLT)

    # Thin bright highlight on left edge of upper bolt
    if size >= 48:
        hi_raw = [(126, 22), (136, 22), (112, 115), (104, 115)]
        hi_pts = _scale(hi_raw, s)
        d.polygon(hi_pts, fill=(255, 220, 140, 180))

    return img


def build_ico(out_path: Path):
    sizes = [256, 128, 64, 48, 32, 16]
    frames = [draw_frame(sz) for sz in sizes]

    frames[0].save(
        out_path,
        format="ICO",
        sizes=[(sz, sz) for sz in sizes],
        append_images=frames[1:],
    )
    print(f"✓  {out_path}  ({out_path.stat().st_size // 1024} KB, {len(sizes)} sizes)")


if __name__ == "__main__":
    out = Path(__file__).parent / "icon.ico"
    build_ico(out)
