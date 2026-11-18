"""Генератор SVG-пейзажа «долина с холмами» для start_page.html (детерминированный)."""
import random
import sys

R = random.Random(7)
W, H = 1600, 900
out = []
a = out.append

a(f'<svg viewBox="0 0 {W} {H}" preserveAspectRatio="xMidYMax slice" xmlns="http://www.w3.org/2000/svg">')
a('''<defs>
  <linearGradient id="vsky" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#9fd0f3"/><stop offset=".45" stop-color="#cfe7f7"/><stop offset=".7" stop-color="#eef6f6"/>
  </linearGradient>
  <linearGradient id="vmtn" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7f9fc4"/><stop offset="1" stop-color="#a9c1d8"/></linearGradient>
  <linearGradient id="vh1" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#5e9e35"/><stop offset="1" stop-color="#4f8a2b"/></linearGradient>
  <linearGradient id="vh2" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#9ed65a"/><stop offset=".6" stop-color="#7fbd3c"/><stop offset="1" stop-color="#6aa82f"/></linearGradient>
  <linearGradient id="vh3" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#a8dc63"/><stop offset="1" stop-color="#86c445"/></linearGradient>
  <linearGradient id="vh4" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#7cbf3d"/><stop offset="1" stop-color="#95d152"/></linearGradient>
  <linearGradient id="vh5" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#9ad457"/><stop offset="1" stop-color="#78b83a"/></linearGradient>
  <linearGradient id="vh6" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8fcd4d"/><stop offset="1" stop-color="#6fb032"/></linearGradient>
  <radialGradient id="vshade" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#2f5d1b" stop-opacity=".35"/><stop offset="1" stop-color="#2f5d1b" stop-opacity="0"/></radialGradient>
  <linearGradient id="vhaze" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffffff" stop-opacity=".55"/><stop offset="1" stop-color="#ffffff" stop-opacity="0"/></linearGradient>
</defs>''')
# небо
a(f'<rect width="{W}" height="{H}" fill="url(#vsky)"/>')
# далёкая синяя гора
a('<path d="M690 345 C 760 300, 820 285, 880 300 C 930 312, 980 300, 1040 322 C 1080 335, 1110 345, 1140 352 L 690 352 Z" fill="url(#vmtn)" opacity=".85"/>')

# тёмный лес на гребне слева
trees = []
x = -20
while x < 760:
    y = 312 + (x - 380) ** 2 / 9000 + R.uniform(-6, 6)
    r = R.uniform(13, 24)
    col = R.choice(["#2f5a2a", "#3a6b30", "#284f25", "#46783a"])
    trees.append(f'<ellipse cx="{x:.0f}" cy="{y:.0f}" rx="{r * .8:.0f}" ry="{r * 1.25:.0f}" fill="{col}"/>')
    x += R.uniform(12, 22)
a("<g>" + "".join(trees) + "</g>")

# холмы, от дальних к ближним
a('<path d="M0 300 C 120 286, 260 300, 400 330 C 520 356, 610 372, 700 395 L 700 900 L 0 900 Z" fill="url(#vh1)"/>')
# полосы-террасы на левом холме
for i in range(5):
    y0 = 345 + i * 22
    a(f'<path d="M0 {y0} C 150 {y0 - 14}, 330 {y0 + 6}, 520 {y0 + 34}" stroke="#8cc650" stroke-opacity=".35" stroke-width="7" fill="none"/>')
a('<path d="M560 430 C 700 395, 860 300, 1060 282 C 1250 266, 1440 276, 1600 262 L 1600 900 L 560 900 Z" fill="url(#vh2)"/>')
a('<path d="M560 430 C 700 395, 860 300, 1060 282 C 1250 266, 1440 276, 1600 262 L 1600 330 C 1400 330, 1180 360, 900 420 Z" fill="url(#vhaze)" opacity=".35"/>')
a('<path d="M260 470 C 420 440, 600 446, 780 470 C 980 496, 1180 488, 1600 450 L 1600 900 L 260 900 Z" fill="url(#vh3)"/>')
a('<path d="M0 470 C 140 470, 240 500, 330 540 C 400 570, 470 600, 560 620 L 560 900 L 0 900 Z" fill="url(#vh4)"/>')
a('<path d="M720 900 C 860 760, 1080 660, 1600 600 L 1600 900 Z" fill="url(#vh6)"/>')
a('<path d="M0 600 C 240 640, 520 720, 820 770 C 1030 805, 1220 850, 1360 900 L 0 900 Z" fill="url(#vh5)"/>')
# мягкие тени на склонах
for cx, cy, rx, ry in [(250, 420, 260, 50), (1150, 360, 300, 40), (420, 700, 380, 60), (1350, 720, 300, 70)]:
    a(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="url(#vshade)" opacity=".5"/>')

# тропинка
a('<path d="M1600 880 C 1300 830, 980 760, 820 700 C 700 655, 600 640, 560 615 C 530 598, 540 585, 600 575 C 680 562, 760 560, 860 552" '
  'stroke="#e6d8b0" stroke-width="7" fill="none" stroke-linecap="round"/>')
a('<path d="M1600 880 C 1300 830, 980 760, 820 700 C 700 655, 600 640, 560 615 C 530 598, 540 585, 600 575 C 680 562, 760 560, 860 552" '
  'stroke="#c9b98d" stroke-width="2" fill="none" stroke-linecap="round" stroke-dasharray="2 9" opacity=".7"/>')


def house(x, y, w, h, roof, wall="#f3ead8"):
    return (f'<g><rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{wall}"/>'
            f'<rect x="{x}" y="{y + h - 2}" width="{w}" height="3" fill="#000" opacity=".12"/>'
            f'<polygon points="{x - 5},{y + 2} {x + w / 2},{y - h * .55} {x + w + 5},{y + 2}" fill="{roof}"/>'
            + "".join(f'<rect x="{x + 5 + i * 11}" y="{y + h * .35}" width="5" height="6" fill="#6b4b2e" opacity=".6"/>'
                      for i in range(int((w - 6) // 11)))
            + "</g>")


# хутор в долине и деревня справа
for args in [(495, 548, 70, 26, "#b9472f"), (560, 552, 58, 24, "#c55a3a"), (470, 562, 40, 18, "#7a4b3a"),
             (612, 566, 46, 18, "#b84a32"), (540, 575, 38, 14, "#d0623f")]:
    a(house(*args))
for args in [(1130, 500, 64, 26, "#6e4637"), (1200, 494, 72, 28, "#5e3d33"), (1280, 508, 54, 22, "#8a4a35"),
             (1340, 515, 48, 20, "#b9472f"), (1180, 525, 40, 16, "#c55a3a")]:
    a(house(*args))


def tree(x, y, s, kind):
    trunk = f'<rect x="{x - s * .08:.1f}" y="{y - s * .9:.1f}" width="{s * .16:.1f}" height="{s * .9:.1f}" fill="#5b4632"/>'
    shade = f'<ellipse cx="{x + s * .5:.1f}" cy="{y + 2:.1f}" rx="{s * .9:.1f}" ry="{s * .18:.1f}" fill="#2f5d1b" opacity=".22"/>'
    if kind == "blossom":
        cols = ["#ffffff", "#f5f7ee", "#eef3e2", "#fdf0f4"]
    elif kind == "pink":
        cols = ["#fbe3ea", "#f6cfdc", "#fff1f5", "#f9dbe5"]
    else:
        cols = ["#4f8f2f", "#5ea23a", "#3f7a28", "#6bb043"]
    blobs = "".join(
        f'<circle cx="{x + R.uniform(-.45, .45) * s:.1f}" cy="{y - s * 1.15 + R.uniform(-.35, .35) * s:.1f}" '
        f'r="{s * R.uniform(.38, .55):.1f}" fill="{R.choice(cols)}"/>' for _ in range(6))
    return f"<g>{shade}{trunk}{blobs}</g>"


spots = [
    # дальний правый холм
    (760, 395, 14, "blossom"), (820, 350, 13, "blossom"), (880, 330, 12, "green"), (940, 312, 11, "blossom"),
    (1010, 300, 11, "blossom"), (1180, 300, 12, "green"), (1260, 296, 12, "blossom"), (1330, 292, 11, "green"),
    (1420, 286, 12, "blossom"), (1500, 282, 12, "green"), (1560, 276, 13, "blossom"),
    # долина
    (430, 520, 18, "blossom"), (690, 548, 20, "green"), (735, 520, 16, "blossom"), (880, 480, 16, "blossom"),
    (980, 470, 15, "pink"), (1090, 480, 18, "blossom"), (1410, 470, 18, "green"), (1460, 455, 16, "blossom"),
    # левый склон
    (90, 470, 16, "blossom"), (180, 500, 18, "green"), (300, 560, 22, "blossom"), (230, 610, 26, "blossom"),
    # передний план
    (980, 620, 30, "blossom"), (1050, 600, 26, "blossom"), (1110, 610, 30, "green"), (1240, 560, 28, "blossom"),
    (1300, 575, 24, "pink"), (1500, 560, 30, "blossom"), (150, 760, 46, "blossom"), (640, 690, 22, "green"),
]
for sp in spots:
    a(tree(*sp))
a("</svg>")
sys.stdout.write("\n".join(out))
