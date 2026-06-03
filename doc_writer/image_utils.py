from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path


def caption_to_slug(caption: str) -> str:
    """Convert a caption string to a filename-safe slug."""
    s = caption.lower().strip()
    s = re.sub(r"[^\w\s-]", "", s)
    s = re.sub(r"[\s_]+", "_", s)
    return s[:40].strip("_") or "image"


def generate_placeholder_png(caption: str) -> bytes:
    """Generate a gray PNG image with the caption centered on it."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(5, 2.5))
    fig.patch.set_facecolor("#F5F5F5")
    ax.set_facecolor("#F5F5F5")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_edgecolor("#CCCCCC")

    ax.text(0.5, 0.58, caption, ha="center", va="center",
            fontsize=11, color="#666666", wrap=True, transform=ax.transAxes)
    ax.text(0.5, 0.22, "[ IMAGE PLACEHOLDER ]", ha="center", va="center",
            fontsize=8, color="#AAAAAA", fontstyle="italic", transform=ax.transAxes)

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=96, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()


_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")


def resolve_image_bytes(content: dict, assets_dir: Path, counter: int) -> tuple[bytes, str]:
    """
    Return (image_bytes, filename) for an image component.

    Lookup order:
    1. assets_dir / basename(content["path"])  — if path given
    2. assets_dir / (caption_slug + ext)        — for each common extension
    3. Generate placeholder PNG, save it, return it
    """
    assets_dir.mkdir(parents=True, exist_ok=True)

    caption = content.get("caption") or "?"
    path_hint = content.get("path")

    # 1. Explicit path → check assets by basename
    if path_hint:
        candidate = assets_dir / Path(path_hint).name
        if candidate.exists():
            return candidate.read_bytes(), candidate.name

    # 2. Caption slug → check assets for any common extension
    slug = caption_to_slug(caption)
    for ext in _IMAGE_EXTENSIONS:
        candidate = assets_dir / (slug + ext)
        if candidate.exists():
            return candidate.read_bytes(), candidate.name

    # 3. Generate and save a placeholder
    png_bytes = generate_placeholder_png(caption)
    filename = f"{slug}.png" if slug != "image" else f"image_{counter:03d}.png"
    (assets_dir / filename).write_bytes(png_bytes)
    return png_bytes, filename
