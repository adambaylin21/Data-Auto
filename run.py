"""Điểm khởi chạy của Tool Tự Động Lọc Dữ Liệu khi đóng thành một file .exe.

Chạy server nội bộ ở cổng còn trống, chờ tới khi server trả lời được rồi tự mở
trình duyệt tới giao diện. Vì mở từ file .exe nên giao diện nằm trong gói chứ
không còn cạnh app.py.

Chương trình chạy kèm cửa sổ console: cửa sổ hiện địa chỉ đang phục vụ để tiện
mở lại hoặc gửi cho người khác, và đóng cửa sổ đó là tắt chương trình.

Chạy khi phát triển: uv run python app.py
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime

HOST = "127.0.0.1"
PORT = 5001
WAIT_SECONDS = 3.0
LOG_NAME = "KhoiDong.log"

# Máy đã đổi bảng mã console (chcp) sang bảng mã không có dấu tiếng Việt thì
# print sẽ lỗi; thay ký tự không in được để chương trình không dừng vì việc này.
try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except (AttributeError, OSError):
    pass


def log_path() -> str:
    """Đường dẫn file nhật ký khởi động, nằm cạnh chương trình."""
    if getattr(sys, "frozen", False):
        directory = os.path.dirname(os.path.abspath(sys.executable))
    else:
        directory = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(directory, LOG_NAME)


def write_log(message: str) -> None:
    """Ghi một dòng nhật ký; lỗi ghi thì bỏ qua để không chặn chương trình."""
    line = f"{datetime.now().isoformat(timespec='seconds')} {message}\n"
    try:
        with open(log_path(), "a", encoding="utf-8") as handle:
            handle.write(line)
    except OSError:
        pass


def choose_port(preferred: int = PORT) -> int:
    """Trả về cổng còn trống, bắt đầu từ cổng mong muốn.

    Lần chạy trước có thể chưa thoát hẳn, nên cổng cũ bận thì lấy cổng kế tiếp
    thay vì báo lỗi.
    """
    for port in range(preferred, preferred + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                probe.bind((HOST, port))
            except OSError:
                continue
            return port
    return 0


def wait_until_ready(url: str, timeout: float = WAIT_SECONDS) -> bool:
    """Chờ tới khi giao diện trả lời, tránh mở trình duyệt khi server chưa lên."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=0.5):
                return True
        except (urllib.error.URLError, OSError):
            time.sleep(0.05)
    return False


def open_browser(url: str) -> None:
    """Mở trình duyệt tới địa chỉ giao diện ngay khi server sẵn sàng."""
    wait_until_ready(url)
    write_log(f"Đã sẵn sàng, mở trình duyệt tại {url}")
    try:
        webbrowser.open(url)
    except webbrowser.Error as exc:
        write_log(f"Không mở được trình duyệt: {exc}")


def announce(address: str) -> None:
    """In địa chỉ và cách tắt lên cửa sổ console, kèm nhật ký KhoiDong.log."""
    lines = [
        "Tool Tự Động Lọc Dữ Liệu đang chạy.",
        f"Địa chỉ: {address}",
        "Đóng cửa sổ này để tắt chương trình.",
    ]
    banner = "=" * 60
    print(banner)
    for line in lines:
        print(f"  {line}")
    print(banner, flush=True)
    for line in lines:
        write_log(line)


def main() -> None:
    port = choose_port()
    address = f"http://{HOST}:{port}/"

    announce(address)
    write_log(f"Khởi động, phục vụ tại {address}")
    threading.Thread(target=open_browser, args=(address,), daemon=True).start()

    import app

    app.app.run(host=HOST, port=port, debug=False, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()
