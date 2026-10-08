"""Input limits, resource access policy, and failure-safe output."""

import io
import warnings
from zipfile import ZipFile, ZIP_DEFLATED

import pytest
from defusedxml.common import EntitiesForbidden

import aspose.words_foss as aw
from aspose.words_foss import _io
from aspose.words_foss.docx_reader import DocumentReader
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning


def test_stream_read_is_bounded(monkeypatch):
    monkeypatch.setattr(_io, "MAX_INPUT_BYTES", 10)

    class Stream(io.BytesIO):
        def read(self, size=-1):
            assert size == 11
            return super().read(size)

    with pytest.raises(ValueError, match="Input exceeds"):
        aw.Document(Stream(b"x" * 50))


def test_file_and_bytes_limits(tmp_path, monkeypatch):
    monkeypatch.setattr(_io, "MAX_INPUT_BYTES", 10)
    source = tmp_path / "large.txt"
    source.write_bytes(b"x" * 11)
    with pytest.raises(ValueError, match="Input exceeds"):
        aw.Document(source)
    with pytest.warns(DeprecationWarning), pytest.raises(ValueError, match="Input exceeds"):
        aw.Document(data=b"x" * 11)


def archive_with(entries):
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for name, data in entries:
            archive.writestr(name, data)
    return buffer.getvalue()


@pytest.mark.parametrize(
    "limit,value,entries,message",
    [
        ("MAX_PART_BYTES", 10, [("word/document.xml", "x" * 11)], "part exceeds"),
        ("MAX_EXPANDED_BYTES", 10, [("word/document.xml", "x" * 11)], "expanded size"),
        ("MAX_ZIP_ENTRIES", 1, [("a", ""), ("b", "")], "too many"),
        ("MAX_ZIP_ENTRIES", 100, [("../outside", "")], "Unsafe"),
    ],
)
def test_docx_zip_limits(monkeypatch, limit, value, entries, message):
    monkeypatch.setattr(_io, limit, value)
    with pytest.raises(ValueError, match=message):
        DocumentReader().load_bytes(archive_with(entries))


def test_duplicate_zip_names():
    with pytest.warns(UserWarning, match="Duplicate name"):
        data = archive_with([("word/document.xml", ""), ("word/document.xml", "")])
    with pytest.raises(ValueError, match="duplicate"):
        DocumentReader().load_bytes(data)


def test_xml_entities_are_rejected():
    xml = '<!DOCTYPE x [<!ENTITY secret "unsafe">]><x>&secret;</x>'
    with pytest.raises(EntitiesForbidden):
        DocumentReader().load_bytes(archive_with([("word/document.xml", xml)]))


def test_unsupported_source_construct_warns():
    xml = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body></w:document>'
    )
    with pytest.warns(aw.loading.DocumentLoadWarning, match="footnoteReference"):
        aw.Document(io.BytesIO(archive_with([("word/document.xml", xml)])))


def test_local_markdown_images_require_opt_in_and_stay_in_directory(tmp_path):
    from PIL import Image

    folder = tmp_path / "documents"
    folder.mkdir()
    image = folder / "image.png"
    Image.new("RGB", (2, 2)).save(image)
    source = folder / "input.md"
    source.write_text("![image](image.png)")
    assert not aw.Document(source).get_child_nodes(aw.NodeType.SHAPE, True)
    options = aw.MarkdownLoadOptions()
    options.allow_local_images = True
    assert len(aw.Document(source, options).get_child_nodes(aw.NodeType.SHAPE, True)) == 1
    outside = tmp_path / "outside.png"
    Image.new("RGB", (2, 2)).save(outside)
    for target in ("../outside.png", str(outside)):
        source.write_text(f"![image]({target})")
        with pytest.raises(ValueError, match="inside the document directory"):
            aw.Document(source, options)
    (folder / "link.png").symlink_to(outside)
    source.write_text("![image](link.png)")
    with pytest.raises(ValueError, match="inside the document directory"):
        aw.Document(source, options)


def test_image_pixel_limit(monkeypatch):
    from PIL import Image

    data = io.BytesIO()
    Image.new("RGB", (10, 10)).save(data, format="PNG")
    monkeypatch.setattr(_io, "MAX_IMAGE_PIXELS", 50)
    with pytest.raises(ValueError, match="pixels"):
        _io.validate_image(data.getvalue())


def test_svg_cannot_load_external_resources():
    for uri in ("https://example.com/image.png", "file:///etc/passwd", "../secret.png"):
        with pytest.raises(ValueError, match="External SVG"):
            _io.validate_image(f'<svg><image href="{uri}"/></svg>'.encode())


def test_svg_inline_rasters_are_bounded(monkeypatch):
    import base64
    from PIL import Image

    image = io.BytesIO()
    Image.new("RGB", (2, 2)).save(image, format="PNG")
    encoded = base64.b64encode(image.getvalue()).decode()
    svg = f'<svg><image href="data:image/png;base64,{encoded}"/></svg>'.encode()
    _io.validate_image(svg)
    monkeypatch.setattr(_io, "MAX_IMAGE_PIXELS", 1)
    with pytest.raises(ValueError, match="pixels"):
        _io.validate_image(svg)


@pytest.mark.parametrize("suffix", ["pdf", "docx", "txt", "md"])
def test_failed_replace_preserves_existing_output(tmp_path, monkeypatch, suffix):
    output = tmp_path / f"report.{suffix}"
    output.write_bytes(b"existing output")
    source = tmp_path / "source.txt"
    source.write_text("中文 source")

    def fail(*args):
        raise PermissionError("simulated replace failure")

    monkeypatch.setattr(_io.os, "replace", fail)
    with pytest.raises(PermissionError, match="simulated"):
        aw.Document(source).save(output)
    assert output.read_bytes() == b"existing output"
    assert not list(tmp_path.glob(".report-*"))


def test_symlink_output_is_not_replaced(tmp_path):
    original = tmp_path / "original.txt"
    original.write_bytes(b"existing output")
    link = tmp_path / "link.txt"
    link.symlink_to(original)
    with pytest.raises(ValueError, match="symlink"):
        with _io.atomic_output(link):
            pass
    assert link.is_symlink() and original.read_bytes() == b"existing output"


def test_output_permissions_and_non_file_targets(tmp_path, monkeypatch):
    output = tmp_path / "report.txt"
    with _io.atomic_output(output) as temporary:
        temporary.write_text("new")
    if _io.os.name == "posix":
        assert output.stat().st_mode & 0o777 == 0o600
        output.chmod(0o640)
        with _io.atomic_output(output) as temporary:
            temporary.write_text("replacement")
        assert output.stat().st_mode & 0o777 == 0o640
    with pytest.raises(ValueError, match="regular file"):
        with _io.atomic_output(tmp_path):
            pass
    monkeypatch.setattr(_io.os, "access", lambda *args: False)
    with pytest.raises(PermissionError, match="not writable"):
        with _io.atomic_output(output):
            pass
    assert output.read_text() in ("new", "replacement")


def test_short_stream_reads_are_not_truncated():
    class Stream(io.BytesIO):
        def read(self, size=-1):
            return super().read(min(size, 2))

    assert _io.read_bounded(Stream(b"complete input")) == b"complete input"


def test_missing_glyph_can_be_fatal_without_destroying_output(tmp_path):
    source = tmp_path / "source.txt"
    source.write_text("中文 " + chr(0x1FAE0))
    output = tmp_path / "report.pdf"
    output.write_bytes(b"existing output")
    with warnings.catch_warnings():
        warnings.simplefilter("error", PdfMissingGlyphWarning)
        with pytest.raises(PdfMissingGlyphWarning, match="U\\+1FAE0"):
            aw.Document(source).save(output)
    assert output.read_bytes() == b"existing output"
