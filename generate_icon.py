"""
generate_icon.py — Gera o ícone icon.ico para o Auto Ofícios.

Conceito visual:
  - Fundo quadrado arredondado (azul-escuro do tema da app)
  - Folha de papel com canto dobrado (representando o ofício)
  - Caneta tinteiro em diagonal sobre o papel

Execute:  python generate_icon.py
Requer:   pip install pillow
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw


# ── Cores da paleta do app (ui.py) ─────────────────────────────────────────
CARD    = (26,  29,  46, 255)   # #1a1d2e  fundo do ícone
BORDER  = (46,  49,  80, 200)   # #2e3150  borda suave
PAPER   = (210, 218, 240, 255)  # papel levemente azulado
PAPER_F = (158, 172, 210, 255)  # dobra do papel (mais escura)
LINE    = (120, 138, 185, 140)  # linhas de texto (semitransparente)
SHADOW  = (  0,   0,   0,  90)  # sombra da caneta

# Caneta tinteiro
PEN_NIB    = (232, 196,  72, 255)  # bico dourado
PEN_NIB_D  = (155, 118,  28, 210)  # face escura do bico
PEN_GRIP   = ( 52,  58,  82, 255)  # grip cinza-metálico
PEN_BODY   = ( 22,  44, 118, 255)  # corpo azul (accent do app)
PEN_BODY_H = ( 80, 120, 200, 130)  # reflexo lateral no corpo
PEN_CAP    = ( 16,  30,  82, 255)  # capuchão azul-marinho
PEN_RING   = (200, 172,  58, 255)  # anel dourado
PEN_CLIP   = (200, 172,  58, 255)  # clip dourado


def _rr(draw: ImageDraw.ImageDraw, xy, radius: int, fill, outline=None, width=1):
    """Rounded rectangle helper (PIL <9.2 compat)."""
    try:
        draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)
    except AttributeError:
        x0, y0, x1, y1 = xy
        r = radius
        draw.rectangle([x0 + r, y0, x1 - r, y1], fill=fill)
        draw.rectangle([x0, y0 + r, x1, y1 - r], fill=fill)
        for ex, ey in [(x0, y0), (x1-2*r, y0), (x0, y1-2*r), (x1-2*r, y1-2*r)]:
            draw.ellipse([ex, ey, ex + 2*r, ey + 2*r], fill=fill)


def _sp(pts, s: float):
    """Scale float (x, y) tuples to int pixel coordinates."""
    return [(int(x * s + 0.5), int(y * s + 0.5)) for x, y in pts]


def draw_frame(size: int) -> Image.Image:
    """Render one icon frame at the given pixel size."""
    s = size / 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # ── Background rounded square ──────────────────────────────────────────
    pad = max(1, int(4 * s))
    _rr(d, [pad, pad, size - pad, size - pad],
        radius=int(44 * s), fill=CARD, outline=BORDER, width=max(1, int(2 * s)))

    # ── Folha de papel com canto dobrado ──────────────────────────────────
    x1, y1 = 48.0, 30.0
    x2, y2 = 180.0, 220.0
    fold = 36.0

    paper_body = [
        (x1,        y1),
        (x2 - fold, y1),
        (x2,        y1 + fold),
        (x2,        y2),
        (x1,        y2),
    ]
    d.polygon(_sp(paper_body, s), fill=PAPER)

    fold_tri = [
        (x2 - fold, y1),
        (x2,        y1 + fold),
        (x2 - fold, y1 + fold),
    ]
    d.polygon(_sp(fold_tri, s), fill=PAPER_F)

    # Linhas de texto no papel
    lx1 = x1 + 13
    lx2 = x2 - 8
    lx2s = x2 - fold - 3
    lw = max(1, int(3 * s))
    if size >= 32:
        for ly, lxA, lxB in [
            ( 76, lx1, lx2s),
            (100, lx1, lx2),
            (124, lx1, lx2),
            (148, lx1, lx2),
            (172, lx1, lx1 + (lx2 - lx1) * 0.55),
        ]:
            if lxB > lxA:
                d.line(_sp([(lxA, ly), (lxB, ly)], s), fill=LINE, width=lw)

    # ── Caneta tinteiro em diagonal ────────────────────────────────────────
    # Eixo: ponta do bico (inferior-esquerda) → topo do capuchão (superior-direita)
    # Coordenadas na grade 256×256.
    NIB_TIP = (62.0, 202.0)
    CAP_END = (193.0,  48.0)

    dx = CAP_END[0] - NIB_TIP[0]   # 131.0
    dy = CAP_END[1] - NIB_TIP[1]   # -154.0
    L  = math.hypot(dx, dy)         # ≈ 201.9

    # Vetor unitário do eixo (bico→cap) e perpendicular
    ux, uy = dx / L, dy / L         # ≈ ( 0.649, -0.763)
    px, py = -uy, ux                # ≈ ( 0.763,  0.649) — lado direito da caneta

    def P(t: float, w: float = 0.0):
        """Ponto a distância t do bico, deslocado w na perpendicular."""
        return (NIB_TIP[0] + t * ux + w * px,
                NIB_TIP[1] + t * uy + w * py)

    def quad(t0: float, t1: float, w0: float, w1: float):
        """Quadrilátero entre duas posições axiais com meias-larguras w0 e w1."""
        return [P(t0, +w0), P(t1, +w1), P(t1, -w1), P(t0, -w0)]

    # Posições axiais (px na grade 256)
    T_NIB_END  =  26.0   # base do bico
    T_GRIP_END =  56.0   # fim da seção de preensão
    T_R1_S     =  59.0   # início do anel 1 (grip/corpo)
    T_R1_E     =  65.0   # fim do anel 1
    T_BODY_END = 152.0   # fim do corpo
    T_R2_S     = 155.0   # início do anel 2 (corpo/capuchão)
    T_R2_E     = 163.0   # fim do anel 2
    T_CAP_END  =   L     # topo do capuchão (≈ 202)

    # Meias-larguras (px na grade 256)
    W_NIB  =  7.0
    W_GRIP =  8.0
    W_BODY = 11.0
    W_CAP  = 12.0
    W_RING = 13.0

    # Sombra suave da caneta (offset uniforme)
    if size >= 32:
        off = max(2, int(3 * s))
        for poly in [
            quad(0,         T_NIB_END,  0,      W_NIB),
            quad(T_NIB_END, T_GRIP_END, W_NIB,  W_GRIP),
            quad(T_R1_E,    T_BODY_END, W_BODY, W_BODY),
            quad(T_R2_E,    T_CAP_END,  W_CAP,  W_CAP),
        ]:
            d.polygon([(int(x * s + off + 0.5), int(y * s + off + 0.5))
                       for x, y in poly], fill=SHADOW)

    # Bico (triângulo dourado com face clara e escura)
    nb_r = P(T_NIB_END, +W_NIB)
    nb_l = P(T_NIB_END, -W_NIB)
    nb_c = P(T_NIB_END,  0.0)
    d.polygon(_sp([NIB_TIP, nb_r, nb_l], s), fill=PEN_NIB_D)   # face escura
    d.polygon(_sp([NIB_TIP, nb_r, nb_c], s), fill=PEN_NIB)     # face clara

    # Ponto de tinta na ponta do bico
    if size >= 48:
        tx, ty = int(NIB_TIP[0] * s + 0.5), int(NIB_TIP[1] * s + 0.5)
        rd = max(2, int(3 * s))
        d.ellipse([tx - rd, ty - rd, tx + rd, ty + rd], fill=PEN_NIB)

    # Seção de preensão (grip)
    d.polygon(_sp(quad(T_NIB_END, T_GRIP_END, W_GRIP, W_GRIP), s), fill=PEN_GRIP)

    # Anel 1 (grip / corpo)
    if size >= 32:
        d.polygon(_sp(quad(T_R1_S, T_R1_E, W_RING, W_RING), s), fill=PEN_RING)

    # Corpo (barrel)
    d.polygon(_sp(quad(T_R1_E, T_BODY_END, W_BODY, W_BODY), s), fill=PEN_BODY)

    # Reflexo lateral no corpo
    if size >= 48:
        hl_s = P(T_R1_E + 8, -(W_BODY - 3))
        hl_e = P(T_BODY_END - 8, -(W_BODY - 3))
        d.line(_sp([hl_s, hl_e], s), fill=PEN_BODY_H, width=max(1, int(2 * s)))

    # Anel 2 (corpo / capuchão)
    if size >= 32:
        d.polygon(_sp(quad(T_R2_S, T_R2_E, W_RING, W_RING), s), fill=PEN_RING)

    # Capuchão
    d.polygon(_sp(quad(T_R2_E, T_CAP_END, W_CAP, W_CAP), s), fill=PEN_CAP)

    # Topo arredondado do capuchão (círculo)
    cap_center = P(T_CAP_END, 0.0)
    r_cap = int(W_CAP * s + 0.5)
    cx, cy = int(cap_center[0] * s + 0.5), int(cap_center[1] * s + 0.5)
    d.ellipse([cx - r_cap, cy - r_cap, cx + r_cap, cy + r_cap], fill=PEN_CAP)

    # Clip dourado (faixa lateral no capuchão, lado superior)
    if size >= 48:
        cl_s = P(T_R2_E + 5, -(W_CAP + 1))
        cl_e = P(T_CAP_END - 5, -(W_CAP + 1))
        d.line(_sp([cl_s, cl_e], s), fill=PEN_CLIP, width=max(1, int(2.5 * s)))

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
