#!/usr/bin/env python3
"""Generate modern trend charts for Laying Rate and Egg Weight (flock G2)."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch
from matplotlib.ticker import MultipleLocator

OUT = Path(__file__).resolve().parents[1] / "charts"
OUT.mkdir(parents=True, exist_ok=True)

# Theme — clean agricultural teal / amber (modern flat analytics)
BG = "#F7F8F5"
PANEL = "#FFFFFF"
INK = "#1C2B24"
MUTED = "#5F6F66"
GRID = "#E4E9E4"
LR_COLOR = "#1F7A5C"
LR_FILL = "#1F7A5C"
EW_COLOR = "#C67A2F"
EW_FILL = "#C67A2F"
ACCENT = "#2E6B8A"

LR_WEEKS = np.arange(19, 58)
LR_G2 = np.array(
    [
        10.5,
        23.53,
        27.59,
        35.57,
        44.4,
        48.46,
        53.84,
        64.15,
        69.71,
        73.94,
        83.2,
        83.86,
        86.51,
        85.32,
        89.95,
        91.53,
        89.68,
        80.95,
        79.23,
        78.31,
        79.76,
        79.37,
        73.81,
        71.3,
        70.9,
        71.0,
        68.12,
        67.06,
        65.21,
        62.96,
        56.61,
        48.0,
        42.86,
        43.25,
        44.71,
        43.12,
        43.52,
        48.15,
        56.35,
    ]
)
LR_STD = np.array(
    [
        1.09,
        4.74,
        2.68,
        3.48,
        2.93,
        4.38,
        6.1,
        3.65,
        7.44,
        6.6,
        5.28,
        5.21,
        3.37,
        6.85,
        4.52,
        3.27,
        4.49,
        7.4,
        1.99,
        2.19,
        1.73,
        3.99,
        4.71,
        3.66,
        3.66,
        3.31,
        1.68,
        1.05,
        2.38,
        4.0,
        3.14,
        2.0,
        2.49,
        1.98,
        2.55,
        2.38,
        2.0,
        3.55,
        2.17,
    ]
)

EW_WEEKS = np.arange(20, 31)
EW_WEIGHT = np.array(
    [66.18, 67.48, 67.83, 68.24, 69.34, 69.44, 71.53, 72.39, 71.77, 71.12, 72.44]
)
EW_SD = np.array([1.15, 0.94, 0.99, 0.73, 0.95, 0.64, 0.86, 0.71, 0.62, 0.63, 0.49])


def style_axes(ax, ylabel: str, xlabel: str = "Umur (minggu)") -> None:
    ax.set_facecolor(PANEL)
    ax.set_xlabel(xlabel, color=MUTED, fontsize=10, labelpad=8)
    ax.set_ylabel(ylabel, color=MUTED, fontsize=10, labelpad=8)
    ax.tick_params(colors=MUTED, labelsize=9)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color(GRID)
        ax.spines[spine].set_linewidth(1)
    ax.grid(axis="y", color=GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)


def add_title(fig, title: str, subtitle: str) -> None:
    fig.text(0.06, 0.94, title, fontsize=16, fontweight="600", color=INK, ha="left")
    fig.text(0.06, 0.90, subtitle, fontsize=9.5, color=MUTED, ha="left")


def annotate_peak(ax, x, y, label: str, color: str) -> None:
    ax.scatter([x], [y], s=42, color=color, zorder=5, edgecolors="white", linewidths=1.2)
    ax.annotate(
        label,
        xy=(x, y),
        xytext=(10, 12),
        textcoords="offset points",
        fontsize=8.5,
        color=color,
        fontweight="600",
        arrowprops=dict(arrowstyle="-", color=color, lw=0.8),
    )


def chart_laying_rate() -> Path:
    fig, ax = plt.subplots(figsize=(11, 5.2), dpi=160)
    fig.patch.set_facecolor(BG)
    add_title(
        fig,
        "Laying Rate G2 — Tren Produksi Telur",
        "Persentase produksi harian flok G2 · Week 19–57 · dengan ±1 SD",
    )
    fig.subplots_adjust(left=0.08, right=0.97, top=0.84, bottom=0.12)

    upper = LR_G2 + LR_STD
    lower = np.clip(LR_G2 - LR_STD, 0, None)
    ax.fill_between(LR_WEEKS, lower, upper, color=LR_FILL, alpha=0.12, zorder=1, label="±1 SD")
    ax.fill_between(LR_WEEKS, 0, LR_G2, color=LR_FILL, alpha=0.10, zorder=2)
    ax.plot(LR_WEEKS, LR_G2, color=LR_COLOR, lw=2.4, zorder=3, label="Laying Rate G2")
    ax.scatter(LR_WEEKS, LR_G2, s=16, color=LR_COLOR, zorder=4, edgecolors="white", linewidths=0.6)

    peak_i = int(np.argmax(LR_G2))
    annotate_peak(
        ax,
        LR_WEEKS[peak_i],
        LR_G2[peak_i],
        f"Puncak {round(LR_G2[peak_i])}% (Minggu {LR_WEEKS[peak_i]})",
        LR_COLOR,
    )

    style_axes(ax, "Tingkat Produksi Telur (%)")
    ax.set_xlim(18.5, 57.5)
    ax.set_ylim(0, 100)
    ax.xaxis.set_major_locator(MultipleLocator(2))
    ax.legend(frameon=False, loc="upper right", fontsize=8.5, labelcolor=MUTED)
    path = OUT / "laying_rate_g2.png"
    fig.savefig(path, facecolor=BG, bbox_inches="tight")
    plt.close(fig)
    return path


def chart_egg_weight() -> Path:
    fig, ax = plt.subplots(figsize=(10, 5.0), dpi=160)
    fig.patch.set_facecolor(BG)
    add_title(
        fig,
        "Rata-rata Bobot Telur G2",
        "Gram per telur · Week 20–30 · dengan ±1 SD",
    )
    fig.subplots_adjust(left=0.09, right=0.97, top=0.84, bottom=0.12)

    upper = EW_WEIGHT + EW_SD
    lower = EW_WEIGHT - EW_SD
    ax.fill_between(EW_WEEKS, lower, upper, color=EW_FILL, alpha=0.15, zorder=1, label="±1 SD")
    ax.fill_between(EW_WEEKS, 64.5, EW_WEIGHT, color=EW_FILL, alpha=0.10, zorder=2)
    ax.plot(EW_WEEKS, EW_WEIGHT, color=EW_COLOR, lw=2.4, zorder=3, label="Egg Weight G2")
    ax.scatter(EW_WEEKS, EW_WEIGHT, s=28, color=EW_COLOR, zorder=4, edgecolors="white", linewidths=0.8)

    peak_i = int(np.argmax(EW_WEIGHT))
    peak_x = float(EW_WEEKS[peak_i])
    peak_y = float(EW_WEIGHT[peak_i])
    ax.scatter(
        [peak_x],
        [peak_y],
        s=42,
        color=EW_COLOR,
        zorder=5,
        edgecolors="white",
        linewidths=1.2,
    )
    # Place label left of the peak, above the curve, fully inside the plot box
    ax.annotate(
        f"Bobot {peak_y:.2f} g (Minggu {int(peak_x)})",
        xy=(peak_x, peak_y),
        xytext=(26.2, 73.35),
        textcoords="data",
        fontsize=8.5,
        color=EW_COLOR,
        fontweight="600",
        ha="left",
        va="center",
        arrowprops=dict(
            arrowstyle="-",
            color=EW_COLOR,
            lw=0.8,
            connectionstyle="arc3,rad=0.05",
        ),
        clip_on=True,
        zorder=6,
    )

    avg = float(EW_WEIGHT.mean())
    ax.axhline(avg, color=MUTED, lw=1, ls=(0, (4, 3)), alpha=0.8, zorder=2)
    ax.text(
        EW_WEEKS[0] + 0.15,
        avg + 0.25,
        f"Rata-rata Bobot {avg:.1f} g",
        fontsize=8,
        color=MUTED,
    )

    style_axes(ax, "Bobot telur (g)")
    ax.set_xlim(19.5, 30.5)
    ax.set_ylim(64.5, 74)
    ax.xaxis.set_major_locator(MultipleLocator(1))
    ax.legend(frameon=False, loc="lower right", fontsize=8.5, labelcolor=MUTED)
    path = OUT / "egg_weight_g2.png"
    fig.savefig(path, facecolor=BG, bbox_inches="tight")
    plt.close(fig)
    return path


def chart_combined() -> Path:
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(11, 7.2), dpi=160, sharex=False, gridspec_kw={"hspace": 0.28}
    )
    fig.patch.set_facecolor(BG)
    add_title(
        fig,
        "Dashboard Tren Produksi — Flok G2",
        "Laying Rate (atas) & Bobot Telur (bawah) · data dari Excel Mas Pito",
    )
    fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.08)

    # KPIs strip
    kpis = [
        ("Puncak LR", f"{round(LR_G2.max())}%", f"Minggu {LR_WEEKS[int(np.argmax(LR_G2))]}"),
        ("LR terkini", f"{LR_G2[-1]:.1f}%", "W57"),
        ("Avg LR", f"{LR_G2.mean():.1f}%", "W19–57"),
        ("Peak bobot", f"{EW_WEIGHT.max():.2f} g", f"W{EW_WEEKS[int(np.argmax(EW_WEIGHT))]}"),
    ]
    for i, (label, value, meta) in enumerate(kpis):
        x0 = 0.08 + i * 0.23
        fig.patches.append(
            FancyBboxPatch(
                (x0, 0.905),
                0.21,
                0.055,
                transform=fig.transFigure,
                boxstyle="round,pad=0.008,rounding_size=0.008",
                facecolor=PANEL,
                edgecolor=GRID,
                linewidth=1,
                mutation_aspect=0.5,
                zorder=0,
            )
        )
        # skip fancy patch complexity — draw text KPIs simply
    # Clear accidental patches and draw simple KPI texts instead
    fig.patches.clear()
    for i, (label, value, meta) in enumerate(kpis):
        x = 0.10 + i * 0.23
        fig.text(x, 0.945, value, fontsize=12, fontweight="700", color=INK, ha="left")
        fig.text(x, 0.922, f"{label} · {meta}", fontsize=7.5, color=MUTED, ha="left")

    ax1.fill_between(LR_WEEKS, 0, LR_G2, color=LR_FILL, alpha=0.12, zorder=2)
    ax1.plot(LR_WEEKS, LR_G2, color=LR_COLOR, lw=2.2, zorder=3)
    ax1.scatter(LR_WEEKS, LR_G2, s=12, color=LR_COLOR, zorder=4, edgecolors="white", linewidths=0.5)
    peak_i = int(np.argmax(LR_G2))
    annotate_peak(
        ax1,
        LR_WEEKS[peak_i],
        LR_G2[peak_i],
        f"Puncak {round(LR_G2[peak_i])}% (Minggu {LR_WEEKS[peak_i]})",
        LR_COLOR,
    )
    style_axes(ax1, "Tingkat Produksi Telur (%)")
    ax1.set_xlim(18.5, 57.5)
    ax1.set_ylim(0, 100)
    ax1.set_title("Laying Rate G2 (%)", loc="left", fontsize=11, color=INK, pad=8)
    ax1.xaxis.set_major_locator(MultipleLocator(2))

    ax2.fill_between(EW_WEEKS, 64.5, EW_WEIGHT, color=EW_FILL, alpha=0.12, zorder=2)
    ax2.plot(EW_WEEKS, EW_WEIGHT, color=EW_COLOR, lw=2.2, zorder=3)
    ax2.scatter(EW_WEEKS, EW_WEIGHT, s=22, color=EW_COLOR, zorder=4, edgecolors="white", linewidths=0.6)
    style_axes(ax2, "Bobot telur (g)")
    ax2.set_xlim(19.5, 30.5)
    ax2.set_ylim(64.5, 74)
    ax2.set_title("Egg Weight G2 (g)", loc="left", fontsize=11, color=INK, pad=8)
    ax2.xaxis.set_major_locator(MultipleLocator(1))

    path = OUT / "dashboard_lr_egg_weight_g2.png"
    fig.savefig(path, facecolor=BG, bbox_inches="tight")
    plt.close(fig)
    return path


def main() -> None:
    paths = [chart_laying_rate(), chart_egg_weight(), chart_combined()]
    for p in paths:
        print(p)


if __name__ == "__main__":
    main()
