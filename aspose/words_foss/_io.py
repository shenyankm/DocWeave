"""Bounded reads and same-directory atomic output for document conversions."""

from contextlib import contextmanager
import os
from pathlib import Path
import tempfile


class DocumentLoadWarning(UserWarning):
    """Source constructs cannot be retained by the light document model."""


MAX_INPUT_BYTES = 64 * 1024 * 1024
MAX_PART_BYTES = 64 * 1024 * 1024
MAX_EXPANDED_BYTES = 256 * 1024 * 1024
MAX_ZIP_ENTRIES = 10_000
MAX_IMAGE_PIXELS = 25_000_000


def check_input_size(size: int) -> None:
    if size > MAX_INPUT_BYTES:
        raise ValueError(f"Input exceeds {MAX_INPUT_BYTES} bytes")


def read_bounded(stream) -> bytes:
    data = bytearray()
    while chunk := stream.read(min(1024 * 1024, MAX_INPUT_BYTES + 1 - len(data))):
        data.extend(chunk)
        check_input_size(len(data))
    return bytes(data)


def validate_docx_archive(archive) -> None:
    entries = archive.infolist()
    if len(entries) > MAX_ZIP_ENTRIES:
        raise ValueError("DOCX has too many ZIP entries")
    if len({entry.filename for entry in entries}) != len(entries):
        raise ValueError("DOCX has duplicate ZIP entries")
    if sum(entry.file_size for entry in entries) > MAX_EXPANDED_BYTES:
        raise ValueError("DOCX expanded size exceeds the safety limit")
    for entry in entries:
        if entry.file_size > MAX_PART_BYTES:
            raise ValueError(f"DOCX part exceeds the safety limit: {entry.filename}")
        if entry.flag_bits & 1:
            raise ValueError("Encrypted ZIP entries are not supported")
        if (
            entry.filename.startswith("/")
            or ".." in entry.filename.split("/")
            or "\\" in entry.filename
        ):
            raise ValueError("Unsafe DOCX ZIP path")


def validate_image(data: bytes) -> None:
    if len(data) > MAX_PART_BYTES:
        raise ValueError("Image data exceeds the safety limit")
    from io import BytesIO
    from PIL import Image, UnidentifiedImageError
    from defusedxml.ElementTree import fromstring

    if data.lstrip().startswith(b"<"):
        root = fromstring(data)
        import base64

        for node in root.iter():
            for key, value in node.attrib.items():
                if not key.endswith("href"):
                    continue
                if value.startswith("#") and node.tag.rsplit("}", 1)[-1] != "image":
                    continue
                header, _, payload = value.partition(",")
                if header not in {
                    f"data:image/{kind};base64"
                    for kind in ("png", "jpeg", "gif", "webp", "bmp", "tiff")
                }:
                    raise ValueError(
                        "External SVG resources are not permitted; images must be inline rasters"
                    )
                embedded = base64.b64decode(payload, validate=True)
                with Image.open(BytesIO(embedded)) as image:
                    if image.width * image.height > MAX_IMAGE_PIXELS:
                        raise ValueError(f"Image exceeds {MAX_IMAGE_PIXELS} pixels")
        return
    try:
        with Image.open(BytesIO(data)) as image:
            if image.width * image.height > MAX_IMAGE_PIXELS:
                raise ValueError(f"Image exceeds {MAX_IMAGE_PIXELS} pixels")
    except Image.DecompressionBombError as exc:
        raise ValueError("Image exceeds the safety limit") from exc
    except UnidentifiedImageError:
        pass  # Vector/legacy image formats are handled by their renderer.


@contextmanager
def atomic_output(path):
    """Leave existing output intact on failure; never follow a destination symlink."""
    path = Path(path)
    if path.is_symlink():
        raise ValueError("Refusing to replace a symlink output")
    if path.exists() and not path.is_file():
        raise ValueError("Output must be a regular file")
    if path.exists() and not os.access(path, os.W_OK):
        raise PermissionError("Existing output is not writable")
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{path.stem}-", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    temporary = Path(name)
    try:
        yield temporary
        with temporary.open("rb+") as stream:
            os.fsync(stream.fileno())
        if path.exists():
            os.chmod(temporary, path.stat().st_mode & 0o777)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
