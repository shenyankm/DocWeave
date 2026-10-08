"""OPC relationship paths shared by original-package editing and conversion."""

import posixpath
from urllib.parse import unquote, urlsplit


def relationships_path(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", filename + ".rels")


def resolve_target(part_name: str, target: str) -> str:
    """Resolve an internal URI inside the package, never on the filesystem."""
    uri = urlsplit(target)
    if not target or uri.scheme or uri.netloc or uri.query or uri.fragment:
        raise ValueError("Invalid internal relationship target")
    path = unquote(uri.path)
    if "\\" in path or any(ord(char) < 32 for char in path):
        raise ValueError("Unsafe internal relationship target")
    name = posixpath.normpath(path.lstrip("/") if path.startswith("/") else
                             posixpath.join(posixpath.dirname(part_name), path))
    if name in {".", ".."} or name.startswith("../"):
        raise ValueError("Relationship target escapes the package")
    return name
