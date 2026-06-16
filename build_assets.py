"""Generate all visual assets for the Hapi presentation."""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math, random

random.seed(99)

def get_fonts():
    paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ]
    thin_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    bold, thin = None, None
    for p in paths:
        try: bold = p; break
        except: pass
    for p in thin_paths:
        try: thin = p; break
        except: pass
    return bold or paths[0], thin or thin_paths[0]

BOLD_FONT, THIN_FONT = get_fonts()

PUR  = (107, 63, 255)
TEAL = (0, 194, 181)
DARK = (13, 17, 40)
WHITE = (255, 255, 255)
GRAY  = (240, 242, 248)
ORANGE = (240, 113, 50)

# ── 1. HAPI LOGO (improved) ──────────────────────────────────────────────────
def make_logo():
    W, H = 520, 140
    img = Image.new('RGBA', (W, H), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)

    # Stylised mark: two overlapping angled parallelograms (like the real hapi "N" wave)
    # Left bar (vertical)
    d.polygon([(10,20),(32,20),(32,108),(10,108)], fill=PUR)
    # Middle diagonal (top-left to bottom-right)
    d.polygon([(32,20),(55,20),(77,108),(54,108)], fill=PUR)
    # Right bar (vertical)
    d.polygon([(54,20),(78,20),(78,108),(54,108)], fill=PUR)

    try:
        fnt = ImageFont.truetype(BOLD_FONT, 72)
    except:
        fnt = ImageFont.load_default()

    d.text((96, 22), "hapi", fill=PUR, font=fnt)
    img.save('/home/user/clase13/hapi_logo.png')
    print("Logo done")

make_logo()

# ── 2. TITLE BG – rising stock chart, dark navy ───────────────────────────────
def make_title_bg():
    W, H = 1920, 680
    img = Image.new('RGB', (W, H), DARK)
    d = ImageDraw.Draw(img)

    # Gradient background
    for y in range(H):
        f = y / H
        r = int(13 + 12 * f)
        g = int(17 + 10 * f)
        b = int(40 + 20 * f)
        d.line([(0, y), (W, y)], fill=(r, g, b))

    # Subtle grid
    for x in range(0, W, 100):
        d.line([(x, 0), (x, H)], fill=(25, 32, 65), width=1)
    for y in range(0, H, 80):
        d.line([(0, y), (W, y)], fill=(25, 32, 65), width=1)

    # Rising chart line
    pts = []
    x, y = 50, 520
    for i in range(240):
        x = 50 + i * 8
        y = max(120, min(560, y + random.randint(-14, 18) - 2))  # slight upward bias
        pts.append((x, y))

    # Filled area under curve (semi-transparent teal)
    fill_pts = list(pts) + [(pts[-1][0], H), (pts[0][0], H)]
    overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.polygon(fill_pts, fill=(0, 194, 181, 30))
    img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')

    d2 = ImageDraw.Draw(img)
    d2.line(pts, fill=TEAL, width=3)
    for i in range(0, len(pts), 30):
        px, py = pts[i]
        d2.ellipse([(px-5, py-5), (px+5, py+5)], fill=TEAL)

    # Right side: glowing vertical accent
    for x0 in range(W - 6, W, 2):
        d2.line([(x0, 0), (x0, H)], fill=(*PUR, 80))

    img.save('/home/user/clase13/bg_title.png')
    print("Title bg done")

make_title_bg()

# ── 3. DARK SECTION DIVIDER BG ───────────────────────────────────────────────
def make_section_bg():
    W, H = 1920, 680
    img = Image.new('RGB', (W, H), (35, 38, 60))
    d = ImageDraw.Draw(img)
    # Subtle diagonal lines
    for i in range(-H, W, 80):
        d.line([(i, 0), (i + H, H)], fill=(42, 46, 72), width=1)
    # Left purple stripe
    for x in range(0, 8):
        d.line([(x, 0), (x, H)], fill=PUR)
    img.save('/home/user/clase13/bg_section.png')
    print("Section bg done")

make_section_bg()

# ── 4. USER GROWTH CHART ─────────────────────────────────────────────────────
def make_growth_chart():
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np

    fig, ax = plt.subplots(figsize=(10, 5.5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FC')

    dates = ['Oct\n2022', 'Nov\n2023', '2025', 'Goal\n2026']
    users = [10, 300, 500, 1000]
    colors = ['#6B3FFF', '#6B3FFF', '#00C2B5', '#E5E5E5']
    bar_colors = ['#6B3FFF', '#6B3FFF', '#00C2B5', '#CCCCCC']
    ec = ['#5530CC', '#5530CC', '#00A098', '#AAAAAA']

    bars = ax.bar(dates, users, color=bar_colors, edgecolor=ec, linewidth=1.2,
                  width=0.55, zorder=3)

    # Value labels on bars
    for bar, val, c in zip(bars, users, bar_colors):
        label = f"{val}K" if val < 1000 else "1M"
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 18,
                label, ha='center', va='bottom', fontsize=14, fontweight='bold',
                color='#333333')

    ax.set_ylim(0, 1200)
    ax.set_ylabel('Users (thousands)', fontsize=11, color='#555555')
    ax.set_title('Hapi User Growth', fontsize=16, fontweight='bold', color='#1A1A2E', pad=15)
    ax.spines[['top', 'right']].set_visible(False)
    ax.spines[['left', 'bottom']].set_color('#CCCCCC')
    ax.tick_params(colors='#555555', labelsize=11)
    ax.yaxis.grid(True, color='#E0E0E0', linestyle='--', zorder=0)
    ax.set_axisbelow(True)

    # Annotation arrow for goal
    ax.annotate('Target', xy=(3, 1000), xytext=(2.6, 1100),
                fontsize=10, color='#AAAAAA', fontstyle='italic',
                arrowprops=dict(arrowstyle='->', color='#AAAAAA'))

    plt.tight_layout()
    plt.savefig('/home/user/clase13/chart_growth.png', dpi=150, bbox_inches='tight',
                facecolor='white')
    plt.close()
    print("Growth chart done")

make_growth_chart()

# ── 5. REVENUE PIE CHART ─────────────────────────────────────────────────────
def make_revenue_chart():
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    labels = ['PFOF', 'Hapi Prime\n(subscription)', 'Deposit/\nWithdrawal fees',
              'Idle cash\ninterest', 'Crypto spread', 'Closing fees']
    sizes  = [35, 25, 15, 12, 8, 5]
    colors = ['#6B3FFF', '#00C2B5', '#F07132', '#28A745', '#FFC107', '#6C757D']

    fig, ax = plt.subplots(figsize=(7, 6))
    fig.patch.set_facecolor('white')
    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors,
        autopct='%1.0f%%', startangle=140,
        pctdistance=0.75,
        wedgeprops=dict(edgecolor='white', linewidth=2),
        textprops={'fontsize': 10}
    )
    for at in autotexts:
        at.set_color('white')
        at.set_fontweight('bold')
        at.set_fontsize(10)

    ax.set_title('Revenue Mix (estimated)', fontsize=13, fontweight='bold',
                 color='#1A1A2E', pad=10)
    plt.tight_layout()
    plt.savefig('/home/user/clase13/chart_revenue.png', dpi=150, bbox_inches='tight',
                facecolor='white')
    plt.close()
    print("Revenue chart done")

make_revenue_chart()

# ── 6. COMPETITIVE RADAR (simple bar) ────────────────────────────────────────
def make_competitor_chart():
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import numpy as np

    platforms = ['Hapi', 'Interactive\nBrokers', 'Charles\nSchwab', 'XTB', 'eToro', 'Trii']
    criteria = ['LATAM Focus', 'Ease of Use', 'Regulation', 'Low Cost', 'Education']
    scores = [
        [5, 4, 5, 5, 5],   # Hapi
        [2, 2, 5, 3, 2],   # IB
        [1, 3, 5, 5, 3],   # Schwab
        [3, 3, 3, 5, 2],   # XTB
        [3, 4, 3, 4, 2],   # eToro
        [4, 4, 2, 4, 3],   # Trii
    ]

    fig, ax = plt.subplots(figsize=(10, 5.5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FC')

    x = np.arange(len(criteria))
    width = 0.12
    colors = ['#6B3FFF', '#4A90D9', '#28A745', '#F07132', '#FFC107', '#6C757D']

    for i, (plat, sc, col) in enumerate(zip(platforms, scores, colors)):
        offset = (i - len(platforms)/2) * width + width/2
        bars = ax.bar(x + offset, sc, width, label=plat, color=col,
                      alpha=0.9, edgecolor='white', linewidth=0.8, zorder=3)

    ax.set_xticks(x)
    ax.set_xticklabels(criteria, fontsize=10)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['★', '★★', '★★★', '★★★★', '★★★★★'], fontsize=10)
    ax.set_ylim(0, 6)
    ax.set_title('Competitive Scoring (1–5)', fontsize=13, fontweight='bold', color='#1A1A2E')
    ax.spines[['top', 'right']].set_visible(False)
    ax.yaxis.grid(True, color='#E0E0E0', linestyle='--', zorder=0)
    ax.set_axisbelow(True)
    ax.legend(loc='upper right', fontsize=8, framealpha=0.9)

    # Highlight Hapi bar with border
    plt.tight_layout()
    plt.savefig('/home/user/clase13/chart_competitive.png', dpi=150, bbox_inches='tight',
                facecolor='white')
    plt.close()
    print("Competitive chart done")

make_competitor_chart()

print("\nAll assets generated successfully.")
