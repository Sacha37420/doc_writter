from __future__ import annotations

from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np


def render_chart_png(tpl: dict, content: dict) -> BytesIO:
    chart_type = content.get("chart_type", "bar")
    labels = content.get("labels", [])
    datasets = content.get("datasets", [])

    fig_def = tpl.get("figure", {})
    width = fig_def.get("width_inches", 6)
    height = fig_def.get("height_inches", 3.5)
    dpi = fig_def.get("dpi", 150)
    font_size = tpl.get("font_size", 8)
    show_grid = tpl.get("grid", True)
    colors = tpl.get("colors", ["#2980B9", "#E74C3C", "#27AE60", "#F39C12"])
    # Strip unresolved palette placeholders
    colors = [c for c in colors if not c.startswith("{{")]
    if not colors:
        colors = ["#2980B9", "#E74C3C", "#27AE60", "#F39C12"]

    fig, ax = plt.subplots(figsize=(width, height))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.rcParams.update({"font.size": font_size})

    if chart_type == "bar":
        n = len(datasets)
        x = np.arange(len(labels))
        bar_w = 0.8 / max(n, 1)
        for i, ds in enumerate(datasets):
            offset = (i - (n - 1) / 2) * bar_w
            ax.bar(x + offset, ds["values"], bar_w, label=ds.get("label", ""),
                   color=colors[i % len(colors)], edgecolor="white", linewidth=0.5)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=font_size)
        ax.tick_params(axis="y", labelsize=font_size)
        if show_grid:
            ax.yaxis.grid(True, linestyle="--", alpha=0.5)
            ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    elif chart_type == "barh":
        n = len(datasets)
        y = np.arange(len(labels))
        bar_h = 0.8 / max(n, 1)
        for i, ds in enumerate(datasets):
            offset = (i - (n - 1) / 2) * bar_h
            ax.barh(y + offset, ds["values"], bar_h, label=ds.get("label", ""),
                    color=colors[i % len(colors)], edgecolor="white", linewidth=0.5)
        ax.set_yticks(y)
        ax.set_yticklabels(labels, fontsize=font_size)
        ax.tick_params(axis="x", labelsize=font_size)
        if show_grid:
            ax.xaxis.grid(True, linestyle="--", alpha=0.5)
            ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    elif chart_type == "line":
        for i, ds in enumerate(datasets):
            ax.plot(labels, ds["values"], label=ds.get("label", ""),
                    color=colors[i % len(colors)], marker="o",
                    linewidth=1.5, markersize=4)
        ax.tick_params(labelsize=font_size)
        if show_grid:
            ax.yaxis.grid(True, linestyle="--", alpha=0.5)
            ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    elif chart_type == "pie":
        ds = datasets[0] if datasets else {}
        vals = ds.get("values", [])
        pie_labels = labels or [f"Part {i + 1}" for i in range(len(vals))]
        ax.pie(vals, labels=pie_labels, colors=colors[: len(vals)],
               autopct="%1.1f%%", textprops={"fontsize": font_size},
               startangle=90, wedgeprops={"edgecolor": "white", "linewidth": 0.8})
        ax.set_aspect("equal")

    if len(datasets) > 1 or (chart_type != "pie" and any(ds.get("label") for ds in datasets)):
        ax.legend(fontsize=font_size, framealpha=0.7)

    fig.tight_layout(pad=0.5)
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf
