"""W11-E API Surface audit (stdlib-only).

For a given source-root, walks every .py file, parses with ast, extracts:
  - module-level __all__ if present (authoritative)
  - else, module-level public names (no leading underscore) that are:
      * ClassDef
      * FunctionDef / AsyncFunctionDef
      * Assign (module constants, uppercase or dataclass-ish)
      * ImportFrom aliases IFF they appear in __all__ (explicit re-export)

Tags each name against FEATURES.md and module docstrings.

Usage:
    python audit.py <source_root_with_pkg_dir> <pkg_name> <features_md_path> <out_md>
"""

from __future__ import annotations

import ast
import os
import re
import sys
from dataclasses import dataclass


# Names we never treat as library surface, no matter where they appear.
STDLIB_NOISE = {
    "TYPE_CHECKING",
    "Iterable",
    "Iterator",
    "Sequence",
    "Mapping",
    "List",
    "Dict",
    "Tuple",
    "Set",
    "Optional",
    "Union",
    "Any",
    "Callable",
    "IO",
    "BinaryIO",
    "TextIO",
    "Generator",
    "cast",
    "overload",
    "annotations",
    "dataclass",
    "dataclasses",
    "field",
    "fields",
    "contextmanager",
    "Path",
    "StringIO",
    "BytesIO",
    "ABC",
    "abstractmethod",
    "ABCMeta",
    "Enum",
    "IntEnum",
    "auto",
    "wraps",
    "partial",
    "reduce",
    "lru_cache",
    "cached_property",
    "defaultdict",
    "OrderedDict",
    "deque",
    "Counter",
    "namedtuple",
    "ElementTree",
    "Element",
    "SubElement",
    "parse",
    "fromstring",
    "etree",
    "datetime",
    "date",
    "time",
    "timedelta",
    "re",
    "os",
    "sys",
    "json",
    "copy",
    "deepcopy",
    "copy2",
    "shutil",
    "warnings",
    "warn",
    "logging",
    "getLogger",
    "resources",
    "Final",
    "Literal",
    "ClassVar",
}


@dataclass
class Name:
    name: str
    kind: str  # class / function / constant / reexport
    has_docstring: bool = False
    line: int = 0


def parse_module(path: str) -> tuple[list[Name], list[str] | None, str]:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            src = fh.read()
    except Exception:
        return [], None, ""
    try:
        tree = ast.parse(src, filename=path)
    except SyntaxError:
        return [], None, ""

    module_doc = ast.get_docstring(tree) or ""

    all_list: list[str] | None = None
    reexports: list[Name] = []
    defined: list[Name] = []

    for node in tree.body:
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and t.id == "__all__":
                    if isinstance(node.value, (ast.List, ast.Tuple)):
                        try:
                            all_list = [
                                elt.value
                                for elt in node.value.elts
                                if isinstance(elt, ast.Constant)
                                and isinstance(elt.value, str)
                            ]
                        except Exception:
                            pass

        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if not node.name.startswith("_"):
                defined.append(
                    Name(
                        name=node.name,
                        kind="function",
                        has_docstring=bool(ast.get_docstring(node)),
                        line=node.lineno,
                    )
                )
        elif isinstance(node, ast.ClassDef):
            if not node.name.startswith("_"):
                defined.append(
                    Name(
                        name=node.name,
                        kind="class",
                        has_docstring=bool(ast.get_docstring(node)),
                        line=node.lineno,
                    )
                )
        elif isinstance(node, ast.Assign):
            for t in node.targets:
                if (
                    isinstance(t, ast.Name)
                    and not t.id.startswith("_")
                    and t.id != "__all__"
                    # only treat UPPER_CASE as a module constant; mixedCase likely local var
                    and (t.id.isupper() or t.id[0].isupper())
                ):
                    defined.append(Name(name=t.id, kind="constant", line=node.lineno))
        elif isinstance(node, ast.AnnAssign):
            if (
                isinstance(node.target, ast.Name)
                and not node.target.id.startswith("_")
                and (node.target.id.isupper() or node.target.id[0].isupper())
            ):
                defined.append(
                    Name(name=node.target.id, kind="constant", line=node.lineno)
                )
        elif isinstance(node, ast.ImportFrom):
            for alias in node.names:
                local = alias.asname or alias.name
                if local == "*" or local.startswith("_"):
                    continue
                if local in STDLIB_NOISE:
                    continue
                reexports.append(Name(name=local, kind="reexport", line=node.lineno))

    # Filter reexports: only keep if in __all__, or if the module has no __all__
    # (in which case everything not-underscore is implicitly public).
    if all_list is not None:
        kept_reexports = [r for r in reexports if r.name in all_list]
    else:
        kept_reexports = reexports

    return defined + kept_reexports, all_list, module_doc[:120]


def walk_package(src_root: str, pkg: str) -> dict[str, tuple[list[Name], list[str] | None, str]]:
    out: dict[str, tuple[list[Name], list[str] | None, str]] = {}
    pkg_dir = os.path.join(src_root, pkg)
    for root, dirs, files in os.walk(pkg_dir):
        dirs[:] = [d for d in dirs if d not in {"__pycache__", "tests", "templates"}]
        for f in files:
            if not f.endswith(".py") or f == "conftest.py":
                continue
            full = os.path.join(root, f)
            rel = os.path.relpath(full, src_root)
            mod = rel[:-3].replace(os.sep, ".")
            if mod.endswith(".__init__"):
                mod = mod[: -len(".__init__")]
            names, all_list, doc = parse_module(full)
            out[mod] = (names, all_list, doc)
    return out


def load_features_md(path: str) -> str:
    if not os.path.exists(path):
        return ""
    with open(path, "r", encoding="utf-8") as fh:
        return fh.read()


def tag_name(
    n: Name,
    mod: str,
    all_list: list[str] | None,
    features_text: str,
) -> str:
    if n.name.startswith("CT_") or n.name.startswith("ST_"):
        return "accidentally-public (oxml leak)"
    if n.name in {"qn", "nsmap", "register_element_cls", "OxmlElement", "parse_xml"}:
        return "accidentally-public (oxml helper)"
    if all_list is not None and n.name not in all_list:
        return "module-private (not in __all__)"
    # FEATURES.md name mention (case-sensitive word boundary to avoid "add" matching)
    if features_text and re.search(rf"(?<![A-Za-z0-9_]){re.escape(n.name)}(?![A-Za-z0-9_])", features_text):
        return "documented (FEATURES.md)"
    if n.kind in {"class", "function"} and n.has_docstring:
        return "documented (docstring)"
    if n.kind == "reexport":
        return "reexport (untagged)"
    return "undocumented"


def classify_module_visibility(mod: str, pkg: str) -> str:
    parts = mod.split(".")
    if parts == [pkg]:
        return "top-level"
    second = parts[1] if len(parts) > 1 else ""
    if second == "oxml":
        return "internal (oxml)"
    if second == "opc":
        return "semi-internal (opc)"
    if second == "parts":
        return "semi-internal (parts)"
    if second in {"reader", "writer", "descriptors", "xml", "_deprecated"}:
        return "internal (impl)"
    return "public"


def gen_markdown(
    lib: str,
    modules: dict[str, tuple[list[Name], list[str] | None, str]],
    features_text: str,
    out_path: str,
) -> tuple[int, dict[str, int], list[tuple[str, str, str]]]:
    lines: list[str] = []
    lines.append(f"# {lib} — Public API Surface Audit (W11-E)\n")
    lines.append(
        f"Regenerate with: `python3 scripts/w11_e_audit.py src {lib} FEATURES.md API_SURFACE.md`\n"
    )
    lines.append(
        "Generated by a stdlib-only (ast-based) walk of `src/`. For each `.py` "
        "we capture module-level classes, functions, UPPER_CASE constants, and "
        "explicit `ImportFrom` re-exports (only when inside `__all__` or when "
        "`__all__` is absent).\n"
    )
    lines.append(
        "Every name is tagged as one of: "
        "`documented (FEATURES.md)` (name appears in FEATURES.md), "
        "`documented (docstring)` (has a class/function docstring), "
        "`undocumented`, "
        "`module-private (not in __all__)` (module declares `__all__` and the "
        "name isn't in it), "
        "`reexport (untagged)` (appears as an import but isn't mentioned in "
        "FEATURES.md), "
        "or `accidentally-public (oxml leak/helper)` (a `CT_*`/`ST_*` class or "
        "oxml helper bled through a public import).\n"
    )

    total = 0
    tag_counts: dict[str, int] = {}
    candidates: list[tuple[int, str, str, str]] = []

    grouped: dict[str, list[tuple[str, list[Name], list[str] | None, str]]] = {}
    for mod in sorted(modules):
        names, all_list, doc = modules[mod]
        vis = classify_module_visibility(mod, lib)
        grouped.setdefault(vis, []).append((mod, names, all_list, doc))

    order = [
        "top-level",
        "public",
        "semi-internal (opc)",
        "semi-internal (parts)",
        "internal (impl)",
        "internal (oxml)",
    ]
    for vis in order:
        if vis not in grouped:
            continue
        lines.append(f"\n## Section: {vis}\n")
        for mod, names, all_list, doc in grouped[vis]:
            if not names:
                continue
            lines.append(f"\n### `{mod}`")
            if all_list is not None:
                lines.append(f"\n`__all__` declared ({len(all_list)} names)")
            if doc:
                lines.append(f"\n> {doc.strip()[:100]}")
            lines.append("\n| Name | Kind | Tag |")
            lines.append("|---|---|---|")
            seen: set[str] = set()
            for n in names:
                if n.name in seen:
                    continue
                seen.add(n.name)
                tag = tag_name(n, mod, all_list, features_text)
                lines.append(f"| `{n.name}` | {n.kind} | {tag} |")
                total += 1
                tag_counts[tag] = tag_counts.get(tag, 0) + 1

                # Candidate for top-10 "needs docs or deprecation":
                # - tag is undocumented
                # - name looks public (CapitalCamelCase or Capitalised identifier,
                #   NOT ALL_CAPS which is almost always a re-exported enum alias)
                # - in public/top-level module (not oxml, not parts, not impl)
                name_looks_classy = (
                    n.name[:1].isupper() and not n.name.isupper() and "_" not in n.name
                )
                if (
                    tag == "undocumented"
                    and vis in {"top-level", "public"}
                    and (
                        n.kind in {"class", "function"}
                        or (n.kind == "constant" and name_looks_classy)
                    )
                ):
                    score = 10 if n.kind == "class" else (8 if name_looks_classy else 5)
                    depth = len(mod.split("."))
                    if depth == 1:
                        score += 20
                    elif depth == 2:
                        score += 10
                    if name_looks_classy:
                        score += 2
                    candidates.append((score, mod, n.name, n.kind))

    candidates.sort(key=lambda t: -t[0])
    top10 = candidates[:10]

    lines.insert(
        3,
        (
            "\n## Summary\n\n"
            f"- Total names surveyed: **{total}** across **{sum(1 for v in modules.values() if v[0])}** non-empty modules\n"
            "- Tag breakdown:\n"
            + "\n".join(f"  - `{t}`: {c}" for t, c in sorted(tag_counts.items(), key=lambda x: -x[1]))
            + "\n"
        ),
    )

    lines.append("\n## Top 10 needs-documentation-or-deprecation candidates\n")
    lines.append(
        "Names that look public (capitalised class / top-level function) but "
        "appear in no `FEATURES.md` entry and carry no class/function docstring. "
        "Each is a candidate for either (a) a `HISTORY.rst` entry + `FEATURES.md` "
        "section, or (b) an underscore-prefix rename to mark it internal.\n"
    )
    lines.append("| # | Module | Name | Kind |")
    lines.append("|---|---|---|---|")
    for i, (_score, mod, name, kind) in enumerate(top10, 1):
        lines.append(f"| {i} | `{mod}` | `{name}` | {kind} |")

    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")

    return total, tag_counts, [(m, n, k) for _, m, n, k in top10]


def main():
    if len(sys.argv) != 5:
        print("usage: audit.py <src_root> <pkg> <features_md> <out_md>")
        sys.exit(2)
    src_root, pkg, features_md, out_md = sys.argv[1:]
    modules = walk_package(src_root, pkg)
    features_text = load_features_md(features_md)
    total, counts, top10 = gen_markdown(pkg, modules, features_text, out_md)
    print(f"{pkg}: {total} names, {len(modules)} modules -> {out_md}")
    for tag, c in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"  {tag}: {c}")
    print("Top 10 candidates:")
    for mod, name, kind in top10:
        print(f"  {mod}.{name}  ({kind})")


if __name__ == "__main__":
    main()
