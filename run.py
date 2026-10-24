import os
from pathlib import Path
import tempfile

import uvicorn


def main():
    # Keep the database, exports and static files beside this entry point on both OSes.
    os.chdir(Path(__file__).resolve().parent)
    try:
        with tempfile.TemporaryFile(dir="."):
            pass
    except OSError as exc:
        raise SystemExit(
            "MeasureM needs a writable local directory. "
            "Move the application to a folder owned by your user."
        ) from exc
    uvicorn.run("app.main:app", host="127.0.0.1", port=8002, reload=False)

if __name__ == "__main__":
    main()
