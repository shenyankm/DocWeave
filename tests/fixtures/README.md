# Flat OPC 26.9.0 inputs

`flat-opc-26.9.zip` contains four outputs from the official **Aspose.Words for Python
via .NET 26.9.0 trial**, macOS arm64, Python 3.13.15. The source is the self-created
[`commercial-26.9-dom.docx`](../../docs/benchmarks/corpus/commercial-26.9-dom.docx):
ALPHA (bold), BETA, and a one-cell table containing CELL. Trial watermarks and
generated resources remain in the inputs. These are generated documents, not
vendor binaries or implementation code. They do not exercise executable VBA.

Regenerate outside the checkout using the official environment (otherwise the
local `aspose` namespace may shadow the commercial package):

```python
from zipfile import ZipFile, ZIP_DEFLATED
import aspose.words as aw

source = "/absolute/path/to/docs/benchmarks/corpus/commercial-26.9-dom.docx"
names = ("FLAT_OPC", "FLAT_OPC_MACRO_ENABLED", "FLAT_OPC_TEMPLATE",
         "FLAT_OPC_TEMPLATE_MACRO_ENABLED")
with ZipFile("flat-opc-26.9.zip", "w", compression=ZIP_DEFLATED) as archive:
    for name in names:
        output = "/tmp/" + name + ".xml"
        aw.Document(source).save(output, getattr(aw.SaveFormat, name))
        archive.write(output, name + ".xml")
```

Regeneration may change generated metadata or ZIP timestamps. The committed
fixture and per-input hashes in [`flat-opc-loading.json`](../../docs/benchmarks/flat-opc-loading.json)
identify the exact tested bytes. Tests independently parse saved XML and compare
all part payloads after editing; they do not assert visual or full API equivalence.
