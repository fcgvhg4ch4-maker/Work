"""Recalculate an xlsx with headless LibreOffice and return the recalculated copy path."""
import shutil, subprocess, sys, tempfile
from pathlib import Path


def recalc(src: Path, outdir: Path | None = None) -> Path:
    outdir = Path(outdir or tempfile.mkdtemp(prefix="qis_recalc_"))
    outdir.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--calc",
                    "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, timeout=900)
    shutil.rmtree(profile, ignore_errors=True)
    return outdir / src.name


if __name__ == "__main__":
    print(recalc(Path(sys.argv[1]), Path(sys.argv[2]) if len(sys.argv) > 2 else None))
