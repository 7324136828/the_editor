from __future__ import annotations

import json
import sys

from PySide6.QtCore import QCoreApplication
from PySide6.QtQml import QJSEngine


def main():
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    application = QCoreApplication([])
    engine = QJSEngine()
    source = sys.stdin.read(12_000_001)
    if len(source) > 12_000_000:
        print(json.dumps({"error": "Macro input exceeds the size limit."}))
        return
    result = engine.evaluate(source, "folio-macro.js")
    value = result.toString()
    if len(value) > 4_000_000:
        print(json.dumps({"error": "Macro output exceeds the size limit."}))
    else:
        print(json.dumps({"error" if result.isError() else "result": value}))
    application.quit()


if __name__ == "__main__":
    main()
