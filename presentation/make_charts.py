# -*- coding: utf-8 -*-
"""生成 PPT 所需的全部图表与示意图（matplotlib 渲染中文）。

配色遵循 cli-anything-wps 的 academic 预设：
  primary 深蓝 #1A3C8B / secondary 橙 #E67733 / accent 绿 #188050
数据为据公开行业资料整理的示意值，用于教学演示。
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

# ---- 中文字体 ----
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
font_manager.fontManager.addfont(FONT_PATH)
ZH = font_manager.FontProperties(fname=FONT_PATH).get_name()
plt.rcParams["font.family"] = ZH
plt.rcParams["axes.unicode_minus"] = False

ASSETS = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(ASSETS, exist_ok=True)

# ---- 调色板（academic 预设）----
BLUE   = "#1A3C8B"
ORANGE = "#E67733"
GREEN  = "#188050"
DARK   = "#222222"
GRAY   = "#8A8A8A"
LIGHT  = "#F5F8FC"
PALETTE = [BLUE, ORANGE, GREEN, "#5B8DEF", "#C8504B", "#9B59B6"]

DPI = 200


def _save(fig, name):
    path = os.path.join(ASSETS, name)
    fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white",
                edgecolor="none", pad_inches=0.12)
    plt.close(fig)
    print("saved", name)
    return path


# 1) 中国废旧轮胎产生量趋势（柱 + 利用率折线）
def chart_production():
    years = ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
    amount = [10.8, 11.6, 12.7, 13.9, 13.3, 15.2, 16.8, 18.5, 20.1]  # 百万吨
    rate   = [44, 47, 50, 53, 51, 55, 58, 61, 64]                    # 综合利用率 %
    fig, ax1 = plt.subplots(figsize=(7.4, 4.2))
    bars = ax1.bar(years, amount, color=BLUE, width=0.6, label="年产生量（百万吨）", zorder=3)
    ax1.set_ylabel("废旧轮胎产生量 / 百万吨", color=BLUE, fontsize=12)
    ax1.tick_params(axis="y", labelcolor=BLUE)
    ax1.set_ylim(0, 24)
    for b, v in zip(bars, amount):
        ax1.text(b.get_x() + b.get_width() / 2, v + 0.4, f"{v:.0f}",
                 ha="center", va="bottom", fontsize=9, color=BLUE)
    ax2 = ax1.twinx()
    ax2.plot(years, rate, color=ORANGE, marker="o", lw=2.5, label="综合利用率（%）", zorder=4)
    ax2.set_ylabel("综合利用率 / %", color=ORANGE, fontsize=12)
    ax2.tick_params(axis="y", labelcolor=ORANGE)
    ax2.set_ylim(30, 75)
    for x, y in zip(years, rate):
        ax2.text(x, y + 1.5, f"{y}%", ha="center", fontsize=8.5, color=ORANGE)
    ax1.set_title("中国废旧轮胎年产生量与综合利用率（2016–2024）",
                  fontsize=13, color=DARK, fontweight="bold", pad=12)
    ax1.grid(axis="y", ls="--", alpha=0.35, zorder=0)
    fig.tight_layout()
    _save(fig, "production_trend.png")


# 2) 资源化利用途径占比（环形图）
def chart_routes_pie():
    labels = ["胶粉/改性沥青", "再生橡胶", "热解（油/炭黑）", "轮胎翻新", "其它/填埋焚烧"]
    sizes = [34, 28, 16, 12, 10]
    colors = [BLUE, GREEN, ORANGE, "#5B8DEF", GRAY]
    fig, ax = plt.subplots(figsize=(5.6, 4.6))
    wedges, _ = ax.pie(sizes, colors=colors, startangle=90,
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.legend(wedges, [f"{l}  {s}%" for l, s in zip(labels, sizes)],
              loc="center left", bbox_to_anchor=(0.92, 0.5), fontsize=11, frameon=False)
    ax.text(0, 0.08, "资源化", ha="center", va="center", fontsize=15, fontweight="bold", color=DARK)
    ax.text(0, -0.16, "利用结构", ha="center", va="center", fontsize=13, color=GRAY)
    ax.set_title("废旧轮胎资源化利用途径占比", fontsize=13, color=DARK, fontweight="bold", pad=10)
    fig.tight_layout()
    _save(fig, "routes_pie.png")


# 3) 各国/地区回收利用率对比
def chart_country_rate():
    countries = ["欧盟", "日本", "美国", "韩国", "中国", "全球平均"]
    rate = [95, 92, 81, 90, 64, 70]
    colors = [GREEN if r >= 85 else (ORANGE if r >= 70 else BLUE) for r in rate]
    fig, ax = plt.subplots(figsize=(6.6, 4.0))
    bars = ax.barh(countries[::-1], rate[::-1], color=colors[::-1], height=0.62, zorder=3)
    for b, v in zip(bars, rate[::-1]):
        ax.text(v + 1, b.get_y() + b.get_height() / 2, f"{v}%",
                va="center", fontsize=11, color=DARK, fontweight="bold")
    ax.set_xlim(0, 105)
    ax.set_xlabel("废旧轮胎规范化回收利用率 / %", fontsize=11)
    ax.set_title("主要国家/地区废旧轮胎回收利用率对比", fontsize=13,
                 color=DARK, fontweight="bold", pad=10)
    ax.grid(axis="x", ls="--", alpha=0.35, zorder=0)
    fig.tight_layout()
    _save(fig, "country_rate.png")


# 4) 热解工艺流程示意图
def diagram_pyrolysis():
    fig, ax = plt.subplots(figsize=(8.6, 3.2))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 32)
    ax.axis("off")
    steps = [
        ("废旧轮胎\n破碎除杂", BLUE),
        ("进料\n密闭反应釜", "#5B8DEF"),
        ("无氧/缺氧\n热裂解 400–600℃", ORANGE),
        ("冷凝\n分离", GREEN),
        ("产物回收", BLUE),
    ]
    n = len(steps)
    bw, gap = 15, 4.0
    x0 = 3
    centers = []
    for i, (txt, c) in enumerate(steps):
        x = x0 + i * (bw + gap)
        box = FancyBboxPatch((x, 11), bw, 12, boxstyle="round,pad=0.3,rounding_size=1.2",
                             fc=c, ec="none", zorder=3)
        ax.add_patch(box)
        ax.text(x + bw / 2, 17, txt, ha="center", va="center", color="white",
                fontsize=10.5, fontweight="bold", zorder=4)
        centers.append(x + bw / 2)
        if i < n - 1:
            ax.add_patch(FancyArrowPatch((x + bw + 0.4, 17), (x + bw + gap - 0.4, 17),
                         arrowstyle="-|>", mutation_scale=18, lw=2, color=GRAY, zorder=2))
    # 产物输出
    products = [("热解油 ~45%", ORANGE), ("炭黑 ~33%", DARK), ("钢丝 ~12%", GRAY), ("可燃气 ~10%", GREEN)]
    px = centers[-1]
    for j, (ptxt, pc) in enumerate(products):
        ax.text(px + 9, 22 - j * 5.0, "● " + ptxt, ha="left", va="center",
                fontsize=9.5, color=pc, fontweight="bold")
    ax.set_title("废旧轮胎热裂解（Pyrolysis）工艺流程与产物分布", fontsize=13,
                 color=DARK, fontweight="bold")
    fig.tight_layout()
    _save(fig, "pyrolysis_flow.png")


# 5) 应用领域分布
def chart_applications():
    apps = ["橡胶改性沥青\n（道路）", "塑胶运动场地\n跑道", "再生橡胶\n制品", "建筑工程\n材料", "替代燃料\nTDF"]
    val = [32, 24, 20, 13, 11]
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    bars = ax.bar(apps, val, color=PALETTE[:5], width=0.62, zorder=3)
    for b, v in zip(bars, val):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.6, f"{v}%",
                ha="center", fontsize=11, color=DARK, fontweight="bold")
    ax.set_ylim(0, 38)
    ax.set_ylabel("胶粉/再生橡胶下游应用占比 / %", fontsize=11)
    ax.set_title("废旧轮胎再生材料主要应用领域分布", fontsize=13,
                 color=DARK, fontweight="bold", pad=10)
    ax.grid(axis="y", ls="--", alpha=0.35, zorder=0)
    fig.tight_layout()
    _save(fig, "applications.png")


# 6) 技术路线雷达对比
def chart_radar():
    dims = ["资源化程度", "经济效益", "环境友好", "技术成熟度", "处理规模", "附加值"]
    data = {
        "轮胎翻新": [60, 90, 85, 95, 40, 70],
        "胶粉/改性沥青": [85, 75, 88, 90, 80, 75],
        "热解": [95, 65, 70, 70, 85, 85],
        "再生胶": [80, 70, 60, 85, 75, 65],
    }
    colors = [GREEN, BLUE, ORANGE, "#9B59B6"]
    ang = np.linspace(0, 2 * np.pi, len(dims), endpoint=False).tolist()
    ang += ang[:1]
    fig, ax = plt.subplots(figsize=(5.8, 5.2), subplot_kw=dict(polar=True))
    for (name, vals), c in zip(data.items(), colors):
        v = vals + vals[:1]
        ax.plot(ang, v, color=c, lw=2, label=name)
        ax.fill(ang, v, color=c, alpha=0.08)
    ax.set_xticks(ang[:-1])
    ax.set_xticklabels(dims, fontsize=11)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8, color=GRAY)
    ax.set_ylim(0, 100)
    ax.set_title("四类资源化技术路线综合评价", fontsize=13, color=DARK,
                 fontweight="bold", pad=18)
    ax.legend(loc="upper right", bbox_to_anchor=(1.28, 1.12), fontsize=10, frameon=False)
    fig.tight_layout()
    _save(fig, "radar.png")


# 7) 环境效益：胶粉改性沥青 vs 传统（碳排放/资源）
def chart_benefit():
    cats = ["CO₂ 排放\n(相对)", "石油沥青\n消耗", "路面寿命", "噪声\n降低"]
    trad = [100, 100, 100, 100]
    rec  = [82, 80, 140, 130]
    x = np.arange(len(cats))
    w = 0.36
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    ax.bar(x - w / 2, trad, w, label="传统工艺（基准=100）", color=GRAY, zorder=3)
    ax.bar(x + w / 2, rec, w, label="掺胶粉/再生方案", color=GREEN, zorder=3)
    for i, v in enumerate(rec):
        ax.text(i + w / 2, v + 2, f"{v}", ha="center", fontsize=10,
                color=GREEN, fontweight="bold")
    ax.axhline(100, color=ORANGE, ls="--", lw=1.2, alpha=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(cats, fontsize=10.5)
    ax.set_ylabel("相对指标（基准 = 100）", fontsize=11)
    ax.set_ylim(0, 160)
    ax.set_title("橡胶改性沥青的环境与使用效益（相对传统工艺）", fontsize=12.5,
                 color=DARK, fontweight="bold", pad=10)
    ax.legend(fontsize=10, frameon=False)
    ax.grid(axis="y", ls="--", alpha=0.35, zorder=0)
    fig.tight_layout()
    _save(fig, "benefit.png")


# 8) 黑色污染危害示意（概念图）
def diagram_hazard():
    fig, ax = plt.subplots(figsize=(6.4, 4.6))
    ax.set_xlim(-0.8, 10.8)
    ax.set_ylim(0, 10)
    ax.axis("off")
    # 中心轮胎
    circ = plt.Circle((5, 5), 1.6, fc=DARK, ec=ORANGE, lw=4, zorder=3)
    ax.add_patch(circ)
    ax.add_patch(plt.Circle((5, 5), 0.7, fc="white", ec=DARK, lw=2, zorder=4))
    ax.text(5, 2.9, "废旧轮胎", ha="center", fontsize=12, color=DARK, fontweight="bold", zorder=5)
    hazards = [
        (5, 9.0, "占用土地·堆积如山", BLUE),
        (9.2, 6.5, "蚊虫滋生·传播疾病", GREEN),
        (9.2, 3.5, "易燃·火灾黑烟污染", ORANGE),
        (0.8, 3.5, "渗滤·土壤地下水污染", "#9B59B6"),
        (0.8, 6.5, "难降解·百年不腐", "#C8504B"),
    ]
    for x, y, txt, c in hazards:
        ax.add_patch(FancyBboxPatch((x - 1.75, y - 0.45), 3.5, 0.9,
                     boxstyle="round,pad=0.1,rounding_size=0.25",
                     fc=c, ec="none", alpha=0.92, zorder=3))
        ax.text(x, y, txt, ha="center", va="center", color="white",
                fontsize=9.5, fontweight="bold", zorder=4)
        ax.add_patch(FancyArrowPatch((5, 5), (x, y), arrowstyle="-",
                     lw=1.4, color=GRAY, alpha=0.6, zorder=1,
                     connectionstyle="arc3,rad=0.0"))
    ax.set_title('"黑色污染"——废旧轮胎的五大环境危害', fontsize=13,
                 color=DARK, fontweight="bold")
    fig.tight_layout()
    _save(fig, "hazard.png")


if __name__ == "__main__":
    chart_production()
    chart_routes_pie()
    chart_country_rate()
    diagram_pyrolysis()
    chart_applications()
    chart_radar()
    chart_benefit()
    diagram_hazard()
    print("\nAll charts generated in", ASSETS)
