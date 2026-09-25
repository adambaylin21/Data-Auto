"""Backend cho Tool Tự Động Lọc Dữ Liệu.

Nhận file Excel từ giao diện, đọc danh sách sheet và tên cột, lưu các quy tắc
lọc vào options.json, và chạy xử lý: lọc sheet của Input 1 rồi ghi sang sheet
của Input 2, kết quả lưu thành file mới trong thư mục output/.

options.json là dữ liệu của người dùng: không bị xoá khi reset giao diện hay
chạy lại chương trình. Chỉ nút "Xoá" trong popup mới làm trống nội dung file.

Chạy: uv run python app.py
"""

from __future__ import annotations

import io
import json
import os
import tempfile
from datetime import datetime

import pandas as pd
from flask import Flask, jsonify, request, send_from_directory

import processor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
OPTIONS_PATH = os.path.join(BASE_DIR, "options.json")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DEFAULT_OPTIONS = {
    "header_row": None,
    "rules": [],
    "write_mode": processor.WRITE_OVERWRITE,
    "from_row": None,
}

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024


def column_letter(position: int) -> str:
    """Đổi vị trí cột thành ký tự kiểu Excel: 1 -> A, 26 -> Z, 27 -> AA."""
    letter = ""
    while position > 0:
        position, remainder = divmod(position - 1, 26)
        letter = chr(ord("A") + remainder) + letter
    return letter


def read_excel(upload):
    """Đọc file tải lên thành (luồng dữ liệu, danh sách tên sheet).

    Trả về (giá trị, None) khi thành công, hoặc (None, (json, status)) khi lỗi.
    """
    if upload is None or not upload.filename:
        return None, (jsonify(error="Không nhận được file."), 400)

    extension = os.path.splitext(upload.filename)[1].lower()
    if extension not in EXCEL_EXTENSIONS:
        return None, (
            jsonify(
                error=f"Định dạng {extension or 'không xác định'} không được hỗ trợ. "
                      f"Hãy chọn file .xlsx, .xlsm hoặc .xls."
            ),
            400,
        )

    data = upload.read()
    if not data:
        return None, (jsonify(error="File rỗng."), 400)

    try:
        with pd.ExcelFile(io.BytesIO(data)) as workbook:
            sheets = list(workbook.sheet_names)
    except Exception as exc:
        return None, (jsonify(error=f"Không đọc được file Excel: {exc}"), 422)

    if not sheets:
        return None, (jsonify(error="File không có sheet nào."), 422)

    return (io.BytesIO(data), sheets), None


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.get("/<path:filename>")
def assets(filename: str):
    return send_from_directory(BASE_DIR, filename)


@app.post("/api/sheets")
def list_sheets():
    """Trả về danh sách sheet của file Excel được tải lên."""
    parsed, error = read_excel(request.files.get("file"))
    if error:
        return error

    _, sheets = parsed
    return jsonify(filename=request.files["file"].filename, sheets=sheets)


@app.post("/api/columns")
def list_columns():
    """Trả về danh sách tên cột của sheet được chọn trong file tải lên.

    Tên cột đọc theo chiều ngang của hàng chứa tên cột và dừng ở ô trắng đầu
    tiên, nên các cột phía sau ô trắng không được tính.
    """
    parsed, error = read_excel(request.files.get("file"))
    if error:
        return error

    stream, sheets = parsed
    sheet = request.form.get("sheet") or ""
    if sheet not in sheets:
        return jsonify(error=f"Sheet “{sheet}” không có trong file."), 400

    try:
        row_number = int(request.form.get("header_row") or 1)
    except ValueError:
        return jsonify(error="Số hàng chứa tên cột không hợp lệ."), 400
    if not 1 <= row_number <= 20:
        return jsonify(error="Số hàng chứa tên cột phải nằm trong khoảng 1 đến 20."), 400

    try:
        frame = pd.read_excel(stream, sheet_name=sheet, header=None)
    except Exception as exc:
        return jsonify(error=f"Không đọc được sheet “{sheet}”: {exc}"), 422

    if row_number > len(frame):
        return jsonify(columns=[])

    columns: list[dict] = []
    for index, value in enumerate(frame.iloc[row_number - 1]):
        if pd.isna(value):
            break
        text = str(value).strip()
        if not text:
            break
        columns.append({"letter": column_letter(index + 1), "name": text})

    return jsonify(columns=columns)


@app.get("/api/options")
def get_options():
    """Trả về các quy tắc đã lưu, để giao diện hiện lại sau khi tải trang."""
    saved = read_options()
    return jsonify(
        header_row=saved.get("header_row"),
        rules=saved.get("rules") or [],
        write_mode=saved.get("write_mode") or processor.WRITE_OVERWRITE,
        from_row=saved.get("from_row"),
    )


def read_options() -> dict:
    """Đọc options.json; trả về giá trị mặc định khi file thiếu hoặc hỏng."""
    if not os.path.exists(OPTIONS_PATH):
        return dict(DEFAULT_OPTIONS)

    try:
        with open(OPTIONS_PATH, encoding="utf-8") as handle:
            saved = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return dict(DEFAULT_OPTIONS)

    if not isinstance(saved, dict):
        return dict(DEFAULT_OPTIONS)
    return saved


def write_options(header_row, rules: list[dict], write_mode: str, from_row=None) -> None:
    """Ghi options.json qua file tạm rồi đổi tên, tránh mất dữ liệu khi ghi dở."""
    document = {
        "header_row": header_row,
        "rules": rules,
        "write_mode": write_mode,
        "from_row": from_row,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }

    directory = os.path.dirname(OPTIONS_PATH) or "."
    handle = tempfile.NamedTemporaryFile(
        "w", encoding="utf-8", dir=directory, prefix=".options-", suffix=".tmp", delete=False
    )
    try:
        with handle:
            json.dump(document, handle, ensure_ascii=False, indent=2)
        os.replace(handle.name, OPTIONS_PATH)
    except OSError:
        if os.path.exists(handle.name):
            os.unlink(handle.name)
        raise


def clean_rules(raw) -> list[dict]:
    """Kiểm tra và chuẩn hoá danh sách quy tắc do giao diện gửi lên."""
    cleaned = []
    for index, rule in enumerate(raw, start=1):
        if not isinstance(rule, dict):
            raise ValueError(f"Quy tắc thứ {index} không hợp lệ.")

        item = {
            "type": str(rule.get("type") or ""),
            "column": str(rule.get("column") or ""),
            "match": str(rule.get("match") or ""),
            "keyword": str(rule.get("keyword") or ""),
        }
        # "Trống Không" là quy tắc tìm ô trống nên không cần từ khoá
        required = ("type", "column", "match")
        if not all(item[field] for field in required):
            raise ValueError(f"Quy tắc thứ {index} còn thiếu thông tin.")
        if item["match"] != "Trống Không" and not item["keyword"]:
            raise ValueError(f"Quy tắc thứ {index} chưa có từ khoá.")
        cleaned.append(item)
    return cleaned


@app.post("/api/options")
def save_options():
    """Lưu hàng chứa tên cột, các quy tắc và cách thức ghi vào options.json."""
    payload = request.get_json(silent=True) or {}

    header_row = payload.get("header_row")
    if header_row is not None:
        if isinstance(header_row, bool) or not isinstance(header_row, int):
            return jsonify(error="Số hàng chứa tên cột không hợp lệ."), 400
        if not 1 <= header_row <= 20:
            return jsonify(error="Số hàng chứa tên cột phải nằm trong khoảng 1 đến 20."), 400

    saved = read_options()

    # Không gửi rules thì giữ nguyên quy tắc cũ, tránh ghi đè bằng danh sách rỗng
    if "rules" in payload:
        rules = payload.get("rules")
        if not isinstance(rules, list):
            return jsonify(error="Danh sách quy tắc không hợp lệ."), 400
        try:
            cleaned = clean_rules(rules)
        except ValueError as exc:
            return jsonify(error=str(exc)), 400
    else:
        cleaned = saved.get("rules") or []

    write_mode = payload.get("write_mode") or saved.get("write_mode") or processor.WRITE_OVERWRITE
    if write_mode not in (processor.WRITE_OVERWRITE, processor.WRITE_UPDATE):
        return jsonify(error="Cách thức ghi không hợp lệ."), 400

    from_row = payload.get("from_row", saved.get("from_row"))
    if from_row is not None:
        if isinstance(from_row, bool) or not isinstance(from_row, int):
            return jsonify(error="Số hàng bắt đầu ghi không hợp lệ."), 400
        if not 1 <= from_row <= 100000:
            return jsonify(error="Số hàng bắt đầu ghi phải là số nguyên dương."), 400

    try:
        write_options(header_row, cleaned, write_mode, from_row)
    except OSError as exc:
        return jsonify(error=f"Không ghi được options.json: {exc}"), 500

    return jsonify(
        saved=True,
        header_row=header_row,
        rules=cleaned,
        write_mode=write_mode,
        from_row=from_row,
    )


@app.post("/api/options/reset")
def reset_options():
    """Làm trống nội dung options.json để tạo cấu hình mới."""
    try:
        write_options(None, [], processor.WRITE_OVERWRITE, None)
    except OSError as exc:
        return jsonify(error=f"Không xoá được options.json: {exc}"), 500

    return jsonify(reset=True)


@app.post("/api/process")
def run_process():
    """Lọc sheet của Input 1 theo quy tắc rồi ghi sang sheet của Input 2."""
    input1 = request.files.get("input1")
    input2 = request.files.get("input2")
    if not input1 or not input1.filename:
        return jsonify(error="Chưa chọn file cho Input 1."), 400
    if not input2 or not input2.filename:
        return jsonify(error="Chưa chọn file cho Input 2."), 400

    for upload, label in ((input1, "Input 1"), (input2, "Input 2")):
        extension = os.path.splitext(upload.filename)[1].lower()
        if extension not in EXCEL_EXTENSIONS:
            return jsonify(error=f"{label} phải là file .xlsx, .xlsm hoặc .xls."), 400
        if extension == ".xls":
            return jsonify(error=f"{label} là định dạng .xls cũ, hãy lưu lại thành .xlsx."), 400

    sheet1 = request.form.get("sheet1") or ""
    sheet2 = request.form.get("sheet2") or ""
    if not sheet1 or not sheet2:
        return jsonify(error="Chưa chọn sheet cho Input 1 hoặc Input 2."), 400

    write_mode = request.form.get("write_mode") or processor.WRITE_OVERWRITE
    if write_mode not in (processor.WRITE_OVERWRITE, processor.WRITE_UPDATE):
        return jsonify(error="Cách thức ghi không hợp lệ."), 400

    apply_filter = (request.form.get("remove_extra") or "").lower() in ("true", "1", "on")

    saved = read_options()
    header_row = saved.get("header_row")
    if not isinstance(header_row, int) or not 1 <= header_row <= 20:
        return jsonify(error="Chưa thiết lập hàng chứa tên cột trong popup."), 400

    # Giao diện gửi số hàng trực tiếp; không có thì lấy mốc đã lưu trong options.json
    raw_from_row = (request.form.get("from_row") or "").strip()
    if raw_from_row:
        try:
            from_row = int(raw_from_row)
        except ValueError:
            return jsonify(error="Số hàng bắt đầu ghi không hợp lệ."), 400
        if from_row < 1:
            return jsonify(error="Số hàng bắt đầu ghi phải là số nguyên dương."), 400
    else:
        from_row = saved.get("from_row")
        if not isinstance(from_row, int) or isinstance(from_row, bool):
            from_row = None

    rules = saved.get("rules") or []

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    # Đặt tên theo file Input 2 vì đây chính là file được ghi thêm dữ liệu
    stem = os.path.splitext(os.path.basename(input2.filename))[0]
    extension = os.path.splitext(input2.filename)[1].lower()
    if extension not in (".xlsx", ".xlsm"):
        extension = ".xlsx"
    output_name = f"{stem}_ketqua{extension}"
    output_path = os.path.join(OUTPUT_DIR, output_name)

    with tempfile.TemporaryDirectory() as workdir:
        path1 = os.path.join(workdir, "input1" + extension)
        path2 = os.path.join(workdir, "input2" + os.path.splitext(input2.filename)[1].lower())
        input1.save(path1)
        input2.save(path2)

        try:
            stats = processor.process(
                input1=path1,
                sheet1=sheet1,
                input2=path2,
                sheet2=sheet2,
                output_path=output_path,
                header_row=header_row,
                rules=rules,
                write_mode=write_mode,
                apply_filter=apply_filter,
                from_row=from_row,
            )
        except processor.ProcessError as exc:
            return jsonify(error=exc.message), exc.status
        except Exception as exc:  # noqa: BLE001 - hiện lỗi đọc được cho người dùng
            return jsonify(error=f"Không xử lý được file: {exc}"), 500

    return jsonify(
        ok=True,
        output=os.path.relpath(output_path, BASE_DIR),
        download="/" + os.path.relpath(output_path, BASE_DIR).replace(os.sep, "/"),
        applied_filter=apply_filter and bool(rules),
        stats=stats,
    )


@app.errorhandler(413)
def file_too_large(_):
    return jsonify(error="File vượt quá giới hạn 64 MB."), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=True)
