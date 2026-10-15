import subprocess
import sys
from pathlib import Path


def build():
    entry_point = Path("src/my_qt_app/main.py")
    output_dir = Path("dist")
    output_dir.mkdir(exist_ok=True)

    cmd = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--enable-plugin=pyside6",
        "--include-package=openpyxl",
        "--include-package=pypdf",
        "--remove-output",
        "--assume-yes-for-downloads",
        f"--output-dir={output_dir}",
        "--output-filename=report-consolidator",
    ]

    if sys.platform == "win32":
        cmd.append("--windows-console-mode=disable")
    elif sys.platform.startswith("linux"):
        cmd.append("--disable-console")

    cmd.append(str(entry_point))

    print("Executing Nuitka build command:")
    print(" ".join(str(c) for c in cmd))
    subprocess.run(cmd, check=True)

if __name__ == "__main__":
    build()
