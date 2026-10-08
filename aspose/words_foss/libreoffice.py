"""Optional original-file Word-to-PDF conversion via installed LibreOffice.

This is separate from the LDM renderer. Native Office parsers still need an
OS/container sandbox for untrusted documents; a private profile is not a sandbox.
"""

from io import BytesIO
from pathlib import Path
import shutil
import tempfile
from zipfile import ZipFile

from aspose.words_foss._io import atomic_output, read_bounded, validate_docx_archive
from aspose.words_foss._process import run_process


def convert_to_pdf(source, output, *, timeout=60, _new_session=True):
    """Convert the original DOC/DOCX/RTF without a lossy LDM round trip."""
    source, output = Path(source).resolve(), Path(output)
    if source.suffix.lower() not in (".doc", ".docx", ".rtf"):
        raise ValueError("LibreOffice Word input must be DOC, DOCX or RTF")
    if output.suffix.lower() != ".pdf":
        raise ValueError("LibreOffice output must be PDF")
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        executable = str(mac_path) if mac_path.is_file() else None
    if executable is None:
        raise FileNotFoundError("Install LibreOffice and put soffice on PATH")
    with source.open("rb") as stream:
        data = read_bounded(stream)
    if source.suffix.lower() == ".docx":
        with ZipFile(BytesIO(data)) as archive:
            validate_docx_archive(archive)

    with atomic_output(output) as temporary:
        with tempfile.TemporaryDirectory(prefix=".libreoffice-", dir=output.parent) as directory:
            job = Path(directory).resolve()
            original = job / source.name
            original.write_bytes(data)
            profile = job / "profile"
            converted = job / "pdf"
            converted.mkdir()
            command = [
                executable,
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--norestore",
                "--nodefault",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                str(converted),
                str(original),
            ]
            log = run_process(command, timeout, new_session=_new_session)
            result = converted / (source.stem + ".pdf")
            if not result.is_file() or result.stat().st_size < 5:
                raise RuntimeError("LibreOffice produced no PDF: " + log)
            with result.open("rb") as stream:
                if stream.read(5) != b"%PDF-":
                    raise RuntimeError("LibreOffice produced an invalid PDF")
                stream.seek(max(0, result.stat().st_size - 1024))
                if b"%%EOF" not in stream.read():
                    raise RuntimeError("LibreOffice produced an incomplete PDF")
            shutil.copyfile(result, temporary)
