import json
import re
from pathlib import Path

_BUILTIN_TEMPLATES_DIR = Path(__file__).parent / "templates"


def _load_json(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _resolve(obj: object, palette: dict) -> object:
    """Recursively replace {{key}} placeholders with values from palette."""
    if isinstance(obj, str):
        return re.sub(r"\{\{(\w+)\}\}", lambda m: palette.get(m.group(1), m.group(0)), obj)
    if isinstance(obj, dict):
        return {k: _resolve(v, palette) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_resolve(item, palette) for item in obj]
    return obj


class TemplateLoader:
    def __init__(self, custom_dir: Path | str | None = None):
        self._dirs: list[Path] = []
        if custom_dir:
            self._dirs.append(Path(custom_dir))
        self._dirs.append(_BUILTIN_TEMPLATES_DIR)

    def _find(self, subdir: str, name: str) -> Path:
        for base in self._dirs:
            path = base / subdir / f"{name}.json"
            if path.exists():
                return path
        searched = ", ".join(str(d / subdir) for d in self._dirs)
        raise FileNotFoundError(f"Template '{name}' introuvable dans : {searched}")

    def load_singular(self, name: str) -> dict:
        return _load_json(self._find("singular", name))

    def load_combined(self, name: str) -> dict:
        return _load_json(self._find("combined", name))

    def load_page(self, name: str) -> dict:
        return _load_json(self._find("pages", name))

    def get_resolved_singular(self, name: str, palette: dict) -> dict:
        """Load a singular template and substitute all {{color_key}} placeholders."""
        return _resolve(self.load_singular(name), palette)
