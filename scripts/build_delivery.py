"""Build a source delivery ZIP from an explicit allowlist, excluding local data."""
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile


ROOT = Path(__file__).resolve().parents[1]
FIXED = [
    "run.py", "start_windows.bat", "requirements.txt", "requirements-windows.lock.txt",
    "WINDOWS_DEPLOYMENT.md", "DEPLOYMENT_VALIDATION.md",
    "examples/README.md", "examples/input-template.xlsx", "examples/demo-project.xlsx",
    "scripts/verify_install.py", "scripts/create_examples.py", "scripts/prepare_windows.ps1",
    "scripts/build_delivery.py",
]


def delivery_files():
    files = [ROOT / name for name in FIXED]
    files.extend(ROOT.glob("app/**/*.py"))
    files.extend(p for p in (ROOT / "app/static").rglob("*") if p.is_file())
    for path in files:
        if path.is_symlink() or not path.is_file():
            raise ValueError(f"Missing file or unexpected symlink: {path}")
        if any(part.startswith(".") or part == "__pycache__" for part in path.relative_to(ROOT).parts):
            raise ValueError(f"Unexpected hidden/cache file: {path}")
    return sorted(set(files))


def build(destination=None):
    files = delivery_files()
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    names = [p.relative_to(ROOT).as_posix() for p in files]
    dirty = bool(subprocess.check_output(["git", "status", "--porcelain", "--", *names], cwd=ROOT, text=True).strip())
    payload = {p.relative_to(ROOT).as_posix(): p.read_bytes() for p in files}
    manifest = {name: hashlib.sha256(data).hexdigest() for name, data in payload.items()}
    metadata = {"commit": commit, "uncommitted_delivery_files": dirty, "python": "CPython 3.14 x64", "windows_native_verified": False}
    destination = destination or ROOT / "dist" / f"MeasureM-{commit[:7]}{'-working' if dirty else ''}-windows-source.zip"
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in payload.items():
            archive.writestr(f"MeasureM/{name}", data)
        archive.writestr("MeasureM/MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
        archive.writestr("MeasureM/BUILD_INFO.json", json.dumps(metadata, indent=2) + "\n")
    sha = hashlib.sha256(destination.read_bytes()).hexdigest()
    destination.with_suffix(".zip.sha256").write_text(f"{sha}  {destination.name}\n", encoding="ascii")
    print(destination)
    print(f"SHA-256 {sha}")
    if dirty:
        print("Working-tree preview only. Rebuild after committing for final delivery.")
    return destination


if __name__ == "__main__":
    build()
