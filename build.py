"""Đóng gói Tool Tự Động Lọc Dữ Liệu thành một file .exe duy nhất.

Chạy: uv run python build.py

Các bước: tạo Icon.ico nhiều kích thước từ Icon.png, rồi gọi PyInstaller gói
run.py cùng index.html và Icon.png vào một file .exe kèm cửa sổ console.
Kết quả nằm ở dist/Tool-Loc-Du-Lieu.exe; đặt file này ở đâu thì options.json và
thư mục output/ được tạo cạnh nó ở đó.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image

BASE_DIR = Path(__file__).resolve().parent
ICON_PNG = BASE_DIR / "Icon.png"
ICON_ICO = BASE_DIR / "Icon.ico"
ENTRY = BASE_DIR / "run.py"

# PyInstaller tách tên file .exe theo dấu cách khi tự chạy lại chính nó, nên tên
# có dấu cách làm chương trình thoát ngay mà không báo lỗi. Tên file vì thế chỉ
# dùng chữ, số và gạch nối; tên hiện trên cửa sổ vẫn đầy đủ dấu tiếng Việt.
NAME = "Tool-Loc-Du-Lieu"

# Windows lấy biểu tượng ở nhiều kích thước tuỳ nơi hiển thị, từ thanh tác vụ tới
# cửa sổ lớn, nên file .ico phải chứa đủ các cỡ thay vì chỉ cỡ 256.
ICON_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def build_icon() -> Path:
    """Tạo Icon.ico từ Icon.png với đủ các kích thước Windows dùng."""
    if not ICON_PNG.exists():
        raise SystemExit(f"Không tìm thấy {ICON_PNG.name}.")

    image = Image.open(ICON_PNG).convert("RGBA")
    image.save(ICON_ICO, format="ICO", sizes=ICON_SIZES)
    print(f"Đã tạo {ICON_ICO.name} ({image.width}x{image.height}) từ {ICON_PNG.name}.")
    return ICON_ICO


def build_executable(icon: Path) -> Path:
    """Gói chương trình thành một file .exe duy nhất."""
    for stale in ("build", "dist"):
        shutil.rmtree(BASE_DIR / stale, ignore_errors=True)

    command = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--noconfirm",
        "--clean",
        "--name", NAME,
        "--icon", str(icon),
        # Chạy kèm cửa sổ console để thấy địa chỉ đang phục vụ và tắt bằng cách
        # đóng cửa sổ đó.
        # Giao diện và biểu tượng được đọc từ gói, không lấy từ đĩa.
        "--add-data", f"{ICON_PNG}{';' if sys.platform == 'win32' else ':'}.",
        "--add-data", f"{BASE_DIR / 'index.html'}{';' if sys.platform == 'win32' else ':'}.",
        # run.py chỉ nhập app, processor và các thư viện bên trong hàm, nên
        # PyInstaller không tự thấy; khai báo tường minh để chúng vào gói.
        "--hidden-import", "app",
        "--hidden-import", "processor",
        "--hidden-import", "flask",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.utils",
        "--hidden-import", "openpyxl.worksheet.worksheet",
        # _multiarray_umath.pyd nhập numpy._core._exceptions bằng API C, nên
        # PyInstaller không thấy và bỏ sót; thiếu module này thì numpy không
        # nhập được và cả chương trình dừng ngay khi khởi động.
        "--hidden-import", "numpy._core._exceptions",
        # pandas và openpyxl chỉ cần phần xử lý Excel.
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "pytest",
        "--exclude-module", "PyQt5",
        "--exclude-module", "tkinter",
        str(ENTRY),
    ]

    print("Đang đóng gói, quá trình này mất vài phút…")
    subprocess.run(command, cwd=BASE_DIR, check=True)

    suffix = ".exe" if sys.platform == "win32" else ""
    return BASE_DIR / "dist" / f"{NAME}{suffix}"


def main() -> None:
    icon = build_icon()
    executable = build_executable(icon)
    size_mb = executable.stat().st_size / (1024 * 1024)
    print(f"\nXong: {executable} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
