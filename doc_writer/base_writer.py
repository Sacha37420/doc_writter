from abc import ABC, abstractmethod
from pathlib import Path
from .template_loader import TemplateLoader


class BaseDocumentWriter(ABC):
    def __init__(self, custom_templates_dir: Path | str | None = None):
        self._loader = TemplateLoader(custom_templates_dir)

    def _resolve_layout(self, combined_name: str, content: dict) -> list[dict]:
        """
        Load a combined template, resolve its color palette into all singular
        templates, then pair each component with the matching content entry.

        Returns a list of resolved rows ready for rendering.
        """
        combined = self._loader.load_combined(combined_name)
        palette = combined["color_palette"]

        # Normalize all content values to lists so we can pop in order
        queues: dict[str, list] = {
            k: (v if isinstance(v, list) else [v])
            for k, v in content.items()
        }
        # Work on mutable copies
        queues = {k: list(v) for k, v in queues.items()}

        resolved_rows = []
        for row in combined["layout"]["rows"]:
            resolved_cols = []
            for col in row["columns"]:
                resolved_comps = []
                for comp in col["components"]:
                    singular_name = comp["singular_template"]
                    tpl = self._loader.get_resolved_singular(singular_name, palette)

                    comp_content: dict = {}
                    if singular_name in queues and queues[singular_name]:
                        comp_content = dict(queues[singular_name].pop(0))

                    # Inject level from layout into content if not already set
                    if "level" in comp and "level" not in comp_content:
                        comp_content["level"] = comp["level"]

                    resolved_comps.append({"template": tpl, "content": comp_content})

                resolved_cols.append({
                    "width_pct": col["width_pct"],
                    "components": resolved_comps,
                })
            resolved_rows.append({
                "height": row.get("height", "auto"),
                "columns": resolved_cols,
            })

        return resolved_rows

    def add_combined(self, combined_name: str, content: dict) -> None:
        """Render a combined template block filled with the provided content."""
        resolved_rows = self._resolve_layout(combined_name, content)
        self._render_combined(resolved_rows)

    def add_page(self, page_name: str, sections_content: list[dict]) -> None:
        """
        Render a full page template.

        ``sections_content`` must contain one content dict per section defined
        in the page template, in the same order.
        """
        page = self._loader.load_page(page_name)
        self._setup_page(page)
        for i, section in enumerate(page["sections"]):
            content = sections_content[i] if i < len(sections_content) else {}
            self.add_combined(section["combined_template"], content)

    def write_from_json(self, source: "dict | str | Path") -> None:
        """
        Render a complete document from a JSON definition.

        ``source`` can be:
        - a ``dict``  — already-parsed document object
        - a ``str``   — file path to a .json file, or a raw JSON string
        - a ``Path``  — path to a .json file
        """
        import json as _json

        if isinstance(source, dict):
            doc = source
        else:
            p = Path(source)
            if p.exists():
                with open(p, encoding="utf-8") as f:
                    doc = _json.load(f)
            else:
                doc = _json.loads(str(source))

        for block in doc.get("blocks", []):
            block_type = block.get("type")
            if block_type == "page":
                self.add_page(block["template"], block.get("sections", []))
            elif block_type == "combined":
                self.add_combined(block["template"], block.get("content", {}))

    # ------------------------------------------------------------------
    # Abstract interface for subclasses
    # ------------------------------------------------------------------

    @abstractmethod
    def _setup_page(self, page_def: dict) -> None:
        """Apply page-level settings (margins, format, orientation…)."""

    @abstractmethod
    def _render_combined(self, resolved_rows: list[dict]) -> None:
        """Render a fully resolved list of layout rows."""

    @abstractmethod
    def save(self, output_path: str | Path) -> None:
        """Write the final document to disk."""
