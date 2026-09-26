"""Xử lý file Excel: lọc dữ liệu theo quy tắc trong options.json rồi ghi sang Input 2.

Toàn bộ hàm ở đây thuần Python, không phụ thuộc Flask, để có thể kiểm thử riêng.

Quy tắc lọc: mỗi quy tắc gồm type (Giữ/Loại), column, match (Chứa/Bằng/Trống Không)
và keyword. Nhiều quy tắc cùng loại là danh sách để chọn (OR); giữa hai nhóm là giao:

    dòng được giữ  ⇔  khớp ÍT NHẤT MỘT quy tắc "Giữ"  AND  không khớp quy tắc "Loại" nào

Khi không có quy tắc "Giữ" nào thì mọi dòng đều qua bước Giữ.
"""

from __future__ import annotations

import copy
import datetime
import os
import re
from dataclasses import dataclass, field

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

KEEP = "Giữ"
DROP = "Loại"
CONTAINS = "Chứa"
EQUALS = "Bằng"
BLANK = "Trống Không"

WRITE_OVERWRITE = "overwrite"
WRITE_UPDATE = "update"


class ProcessError(Exception):
    """Lỗi đọc hiểu được, thông báo hiện thẳng lên giao diện."""

    def __init__(self, message: str, status: int = 422):
        super().__init__(message)
        self.message = message
        self.status = status


@dataclass
class FilterResult:
    kept_rows: list[int] = field(default_factory=list)
    total_rows: int = 0
    per_rule: dict[str, int] = field(default_factory=dict)

    @property
    def kept(self) -> int:
        return len(self.kept_rows)

    @property
    def removed(self) -> int:
        return self.total_rows - self.kept


# --------------------------------------------------------------------------- #
# Đọc cấu trúc sheet
# --------------------------------------------------------------------------- #


def normalise(value) -> str:
    """Đưa giá trị ô về chuỗi đã cắt khoảng trắng để so khớp."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def header_map(sheet: Worksheet, header_row: int) -> dict[str, int]:
    """Trả về {tên cột: số thứ tự cột} của hàng tiêu đề.

    Đọc theo chiều ngang và dừng ở ô trắng đầu tiên, giống hệt cách giao diện
    lấy danh sách cột, để tên cột lúc chạy luôn khớp tên cột lúc thiết lập.
    """
    if header_row > sheet.max_row:
        raise ProcessError(
            f"Sheet “{sheet.title}” chỉ có {sheet.max_row} hàng, "
            f"không có hàng tiêu đề {header_row}."
        )

    mapping: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        name = normalise(sheet.cell(row=header_row, column=column).value)
        if not name:
            break
        mapping.setdefault(name, column)

    if not mapping:
        raise ProcessError(
            f"Hàng {header_row} của sheet “{sheet.title}” không có tên cột nào."
        )
    return mapping


def column_letter_map(sheet: Worksheet, header_row: int) -> dict[str, int]:
    """Trả về {ký tự cột: số thứ tự cột} của hàng tiêu đề."""
    return {get_column_letter(column): column for column in header_map(sheet, header_row).values()}


def name_map(sheet: Worksheet, header_row: int) -> dict[str, int]:
    """Trả về {tên cột đã chuẩn hoá: số thứ tự cột}, để so khớp không phân biệt hoa thường."""
    return {name.casefold(): column for name, column in header_map(sheet, header_row).items()}


def header_names(sheet: Worksheet, header_row: int) -> dict[str, int]:
    """Như name_map nhưng trả về rỗng khi sheet chưa có hàng mang số đó."""
    if header_row is None or header_row > sheet.max_row:
        return {}
    return name_map(sheet, header_row)


def available_columns(sheet: Worksheet, header_row: int) -> str:
    """Danh sách cột hiện có của sheet, dạng "A Tên khách hàng, B Tên hàng"."""
    pairs = [
        (name, column) for name, column in header_map(sheet, header_row).items()
    ]
    return ", ".join(f"{get_column_letter(column)} {name}" for name, column in pairs)


def saved_reference(spec: dict, keys: tuple[str, ...]) -> str:
    """Mô tả cột đã lưu trong options.json, dạng "B — Tên hàng", để báo lỗi."""
    values = [str(spec.get(key) or "").strip() for key in keys]
    values = [value for value in values if value]
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    return " — ".join(values)


def column_index(letter: str) -> int | None:
    """Đổi ký tự cột thành số thứ tự: "A" → 1, "AA" → 27. Trả về None nếu sai.

    Chỉ nhận chữ cái ASCII, tối đa 3 ký tự như Excel, để tên cột tiếng Việt
    (ví dụ "ĐVT") không bị hiểu nhầm thành ký tự cột.
    """
    text = str(letter or "").strip().upper()
    if not text or len(text) > 3 or not text.isascii() or not text.isalpha():
        return None
    position = 0
    for character in text:
        position = position * 26 + (ord(character) - ord("A") + 1)
    return position if position <= 16384 else None


def looks_like_letter(value: str) -> bool:
    """Chuỗi có phải ký tự cột kiểu Excel (A, B, …, AA) không."""
    return column_index(value) is not None


def column_reference(spec: dict, keys: tuple[str, ...]) -> tuple[str, str]:
    """Tách ký tự cột và tên cột đã lưu của một mục trong options.json."""
    letter = ""
    name = ""
    for key in keys:
        value = str(spec.get(key) or "").strip()
        if not value:
            continue
        if not letter and looks_like_letter(value):
            letter = value
        elif not name:
            name = value
    return letter, name


def resolve_column(
    sheet: Worksheet,
    header_row: int,
    spec: dict,
    keys: tuple[str, ...],
    label: str,
) -> int:
    """Tìm số thứ tự cột mà một mục trong options.json nhắc tới.

    Mục được lưu kèm cả ký tự cột (A, B, C…) và tên cột. Ký tự cột là mốc chính
    vì nó chỉ đúng một cột, nhưng chỉ dùng khi ở ký tự đó vẫn còn đúng tên cột đã
    lưu. Cột bị chèn thêm ở trước nên tên cột chuyển sang ký tự khác thì theo tên,
    tránh lấy nhầm sang cột khác nghĩa.

    options.json cũ chỉ có tên cột thì chạy hoàn toàn theo tên.
    """
    letters = column_letter_map(sheet, header_row)
    names = name_map(sheet, header_row)
    letter, name = column_reference(spec, keys)
    saved = saved_reference(spec, keys)

    if letter and letter.upper() in letters:
        position = letters[letter.upper()]
        current = normalise(sheet.cell(row=header_row, column=position).value)
        if not name or name.casefold() == current.casefold():
            return position
        if name.casefold() in names:
            return names[name.casefold()]
        raise ProcessError(
            f"{label}: cột {letter.upper()} của sheet “{sheet.title}” (hàng tiêu đề "
            f"{header_row}) bây giờ là “{current}”, không còn là “{name}” như đã lưu. "
            f"Cột đã lưu: “{saved}”. Các cột hiện có: {available_columns(sheet, header_row)}."
        )

    if name and name.casefold() in names:
        return names[name.casefold()]

    raise ProcessError(
        f"{label}: cột “{saved}” không còn trong sheet “{sheet.title}” "
        f"(hàng tiêu đề {header_row}). "
        f"Các cột hiện có: {available_columns(sheet, header_row)}."
    )


def resolve_columns(
    sheet: Worksheet,
    header_row: int,
    rules: list[dict],
    sheet_name: str,
) -> dict[int, int]:
    """Tìm cột của từng quy tắc lọc, báo lỗi rõ ràng nếu không thấy."""
    columns: dict[int, int] = {}
    for index, rule in enumerate(rules, start=1):
        columns[index] = resolve_column(
            sheet, header_row, rule, ("column", "name"), f"Quy tắc lọc thứ {index}"
        )
    return columns


# --------------------------------------------------------------------------- #
# So khớp và lọc
# --------------------------------------------------------------------------- #


def is_blank(value) -> bool:
    """Ô được coi là trống khi bỏ trống, chỉ có khoảng trắng, hoặc bằng 0.

    File kế toán thường ghi 0 vào ô không có số liệu thay vì để trống, nên
    "Trống Không" phải hiểu cả số 0, nếu không quy tắc sẽ không khớp dòng nào.
    """
    if value is None:
        return True
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return value == 0
    return normalise(value) == ""


def cell_matches(value, rule: dict) -> bool:
    """Một ô có thoả điều kiện của quy tắc không."""
    text = normalise(value)
    match = rule.get("match")

    if match == BLANK:
        return is_blank(value)
    if match == EQUALS:
        return text.casefold() == str(rule.get("keyword") or "").strip().casefold()
    if match == CONTAINS:
        keyword = str(rule.get("keyword") or "").strip().casefold()
        return keyword in text.casefold()
    return False


def filter_rows(sheet: Worksheet, header_row: int, rules: list[dict]) -> FilterResult:
    """Lọc các dòng dữ liệu của sheet theo danh sách quy tắc."""
    positions = resolve_columns(sheet, header_row, rules, sheet.title)

    keep_rules: list[tuple[int, dict]] = []
    drop_rules: list[tuple[int, dict]] = []
    labels: dict[int, str] = {}
    result = FilterResult()

    for index, rule in enumerate(rules, start=1):
        name = normalise(sheet.cell(row=header_row, column=positions[index]).value)
        labels[index] = describe(rule, name)
        result.per_rule[labels[index]] = 0
        if rule.get("type") == KEEP:
            keep_rules.append((index, rule))
        elif rule.get("type") == DROP:
            drop_rules.append((index, rule))

    last_row = last_table_row(sheet, header_row)
    for row in range(header_row + 1, last_row + 1):
        result.total_rows += 1

        matched_keep = [
            index for index, rule in keep_rules
            if cell_matches(sheet.cell(row=row, column=positions[index]).value, rule)
        ]
        matched_drop = [
            index for index, rule in drop_rules
            if cell_matches(sheet.cell(row=row, column=positions[index]).value, rule)
        ]

        for index in matched_keep + matched_drop:
            result.per_rule[labels[index]] += 1

        passes_keep = bool(matched_keep) if keep_rules else True
        if passes_keep and not matched_drop:
            result.kept_rows.append(row)

    return result


def last_table_row(sheet: Worksheet, header_row: int) -> int:
    """Hàng cuối cùng còn dữ liệu, bỏ qua các hàng trắng ở cuối sheet."""
    for row in range(sheet.max_row, header_row, -1):
        if any(sheet.cell(row=row, column=column).value not in (None, "")
               for column in range(1, sheet.max_column + 1)):
            return row
    return header_row


def describe(rule: dict, column_name: str) -> str:
    """Câu mô tả quy tắc, khớp với câu hiện trên giao diện."""
    text = f"Cột {column_name} {str(rule['type']).lower()} các hàng {rule['match']}"
    keyword = str(rule.get("keyword") or "")
    return f"{text} {keyword}".strip()


# --------------------------------------------------------------------------- #
# Ghi sang sheet đích
# --------------------------------------------------------------------------- #


def is_date_value(value) -> bool:
    """Ô đang chứa giá trị ngày hoặc giờ."""
    return isinstance(value, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta))


def is_date_format(number_format: str) -> bool:
    """Định dạng số của ô dùng để hiện ngày hoặc giờ."""
    if not number_format:
        return False
    # Bỏ phần định dạng màu và điều kiện trong ngoặc vuông để chỉ còn mã định dạng.
    code = re.sub(r"\[[^\]]*\]", "", number_format).lower()
    code = re.sub(r'"[^"]*"', "", code)
    return any(token in code for token in ("yy", "dd", "mm", "hh", "ss", "am/pm"))


def spare_number_format(source, target) -> None:
    """Đặt lại định dạng số của ô đích khi ô đó chưa dành cho kiểu giá trị đang ghi.

    Cách hiển thị phụ thuộc hoàn toàn vào định dạng số của ô, và định dạng của
    nguồn là thứ đúng với giá trị của nguồn. Nhưng sheet đích đặt sẵn định dạng
    cho cả cột, không phải cho từng ô, nên có ô đang mang định dạng của việc khác:
    ô “Tổng cộng” ở cột ngày mang định dạng số, còn ô ngày nằm dưới bảng mang
    định dạng ngày. Gặp ô như vậy thì lấy định dạng của nguồn.

    Ô đích có định dạng ngày mà giá trị đang ghi cũng là ngày thì giữ nguyên định
    dạng của đích, tránh đổi kiểu ngày/tháng của cả cột.
    """
    if not is_date_value(source.value):
        return

    target_format = target.number_format
    source_format = source.number_format

    # Ô ngày giữ nguyên cách hiển thị đã có ở đích; chỉ đổi khi đích đang mang
    # định dạng số của việc khác (ví dụ ô “Tổng cộng”) hoặc định dạng chung chung.
    if is_date_format(target_format) and target_format != "General":
        return

    if source_format != target_format:
        target.number_format = source_format


def copy_cell(source, target, keep_target_style: bool = False) -> None:
    """Sao chép giá trị và định dạng của một ô.

    Định dạng phải gán bằng các đối tượng style thật (font, fill, border,
    alignment, protection). Gán thẳng `_style` sẽ chỉ chép chỉ số tra cứu, mà
    chỉ số đó thuộc bảng style của file nguồn nên khi mở file kết quả sẽ trỏ sai.

    keep_target_style giữ định dạng sẵn có của ô đích, chỉ lấy giá trị. Sheet
    đích như bảng COM đã đặt sẵn định dạng số cho từng cột (số lượng, đơn giá,
    doanh số) nên chép định dạng của sheet nguồn đè lên sẽ làm sai cách hiển thị.
    Ô đích chưa có định dạng nào thì vẫn lấy định dạng của nguồn.
    """
    # Đọc định dạng của đích trước khi gán giá trị: gán một giá trị ngày vào ô
    # đang chứa chữ làm openpyxl đặt lại định dạng số của ô theo mặc định của
    # kiểu ngày, xoá mất định dạng sẵn có của bảng đích.
    target_format = target.number_format

    target.value = source.value
    if keep_target_style and target.has_style:
        target.number_format = target_format
        spare_number_format(source, target)
        return
    if source.has_style:
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.protection = copy.copy(source.protection)
        target.number_format = source.number_format
    spare_number_format(source, target)
    if source.hyperlink:
        target.hyperlink = copy.copy(source.hyperlink)
    if source.comment:
        target.comment = copy.copy(source.comment)


def copy_columns(source: Worksheet, target: Worksheet, count: int) -> None:
    """Chép độ rộng cột và định dạng cột từ sheet nguồn sang sheet đích."""
    for column in range(1, count + 1):
        letter = get_column_letter(column)
        if letter in source.column_dimensions:
            target.column_dimensions[letter] = copy.copy(source.column_dimensions[letter])


def merge_positions(
    sheet: Worksheet,
    header_row: int,
    rules: list[dict],
    side: str,
) -> dict[int, int]:
    """Số thứ tự cột của mỗi quy tắc ghép, tra theo ký tự cột và tên cột đã lưu."""
    name_key = f"{side}_name"
    positions: dict[int, int] = {}
    for index, rule in enumerate(rules, start=1):
        position = resolve_column(
            sheet,
            header_row,
            {side: rule.get(side), name_key: rule.get(name_key)},
            (side, name_key),
            f"Quy tắc ghép cột thứ {index}",
        )
        positions[index] = position
    return positions


def merge_names(
    sheet: Worksheet,
    header_row: int,
    rules: list[dict],
    side: str,
    positions: dict[int, int],
) -> dict[int, str]:
    """Tên cột hiện tại của mỗi quy tắc ghép, dùng cho dòng thông báo kết quả."""
    names: dict[int, str] = {}
    for index, rule in enumerate(rules, start=1):
        column = positions.get(index)
        current = normalise(sheet.cell(row=header_row, column=column).value) if column else ""
        names[index] = current or str(rule.get(f"{side}_name") or rule.get(side) or "").strip()
    return names


def merge_columns(
    source: Worksheet,
    target: Worksheet,
    source_header_row: int,
    target_header_row: int,
    rules: list[dict],
    row_map: dict[int, int] | None = None,
) -> list[dict]:
    """Lấy dữ liệu ở cột nguồn ghi sang cột đích, theo từng cặp cột đã chọn.

    Mỗi quy tắc lưu ký tự cột và tên cột của cả hai bên. Cột được xác định theo
    ký tự rồi đối chiếu tên cột ở hàng tiêu đề của từng sheet, nên ghép cột không
    phụ thuộc thứ tự cột giữa hai file.

    row_map ánh xạ dòng nguồn sang dòng đích; chỉ những dòng đã được ghi mới
    nhận dữ liệu, nên cột đích khớp đúng với các dòng thực sự có dữ liệu và phần
    giữ nguyên phía trên không bị ghi nửa vời.
    """
    source_columns = merge_positions(source, source_header_row, rules, "source")
    target_columns = merge_positions(target, target_header_row, rules, "target")
    source_names = merge_names(source, source_header_row, rules, "source", source_columns)
    target_names = merge_names(target, target_header_row, rules, "target", target_columns)
    applied: list[dict] = []

    if row_map is None:
        last = min(last_table_row(source, source_header_row), last_table_row(target, target_header_row))
        row_map = {row: row for row in range(target_header_row + 1, last + 1)}

    for index, rule in enumerate(rules, start=1):
        source_column = source_columns[index]
        target_column = target_columns[index]
        source_name = source_names[index]
        target_name = target_names[index]

        filled = 0
        for source_row, target_row in row_map.items():
            value = source.cell(row=source_row, column=source_column).value
            if value is None or (isinstance(value, str) and not value.strip()):
                continue
            target.cell(row=target_row, column=target_column).value = value
            filled += 1

        applied.append({
            "source": f"{get_column_letter(source_column)} — {source_name}",
            "target": f"{get_column_letter(target_column)} — {target_name}",
            "source_column": get_column_letter(source_column),
            "target_column": get_column_letter(target_column),
            "source_name": source_name,
            "target_name": target_name,
            "filled": filled,
        })

    return applied


def write_rows(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    start_row: int,
    column_map: dict[int, int],
) -> None:
    """Ghi các dòng đã chọn từ source sang target, giữ định dạng từng ô.

    column_map cho biết mỗi cột nguồn ghi vào cột nào của sheet đích, theo tên
    cột. Cột nguồn không có tên trùng ở sheet đích thì để trống, không ghi lệch
    sang cột khác. Ô đích đã có định dạng sẵn thì giữ định dạng đó, vì sheet
    đích là bảng báo cáo đã đặt sẵn định dạng số cho từng cột.

    Chiều cao hàng của sheet đích được giữ nguyên: bảng báo cáo đặt sẵn chiều
    cao cho hàng dữ liệu, còn sheet nguồn giãn hàng theo nội dung nên lấy chiều
    cao từ nguồn sẽ làm hàng kết quả cao lệch.
    """
    for offset, row in enumerate(rows):
        destination = start_row + offset
        for source_column, target_column in column_map.items():
            copy_cell(
                source.cell(row=row, column=source_column),
                target.cell(row=destination, column=target_column),
                keep_target_style=True,
            )


def existing_header_row(target: Worksheet, source_header_row: int) -> int:
    """Tìm hàng tiêu đề đang có sẵn ở sheet đích.

    Sheet đích thường có sẵn phần đầu trang (tiêu đề báo cáo, tên nhân viên,
    tháng) rồi mới tới hàng tên cột, nên không được ghi đè từ hàng 1. Ưu tiên
    hàng có sẵn tên cột trùng với sheet nguồn; nếu không thấy thì dùng đúng số
    hàng tiêu đề như ở sheet nguồn.
    """
    for row in range(1, min(target.max_row, source_header_row + 20) + 1):
        first = normalise(target.cell(row=row, column=1).value)
        if first and first.casefold().startswith("ngày"):
            return row
    return source_header_row


def destination_header_row(target: Worksheet, source: Worksheet, header_row: int) -> int:
    """Tìm hàng tên cột của sheet đích.

    Thứ tự dò: hàng có cột thứ hai trùng tên cột thứ hai của sheet nguồn (sheet
    đích có thể xếp cột khác thứ tự nên chỉ so một cột), rồi tới hàng có tên cột
    trùng cột đầu của nguồn, cuối cùng mới tới cách dò cũ theo chữ "Ngày".
    """
    source_map = header_map(source, header_row)
    names = list(source_map)
    anchors = ([names[1]] if len(names) > 1 else []) + ([names[0]] if names else [])
    anchors = [anchor.casefold() for anchor in anchors]

    for row in range(1, min(target.max_row, header_row + 20) + 1):
        folded = {
            normalise(target.cell(row=row, column=column).value).casefold()
            for column in range(1, target.max_column + 1)
        }
        folded.discard("")
        if any(anchor in folded for anchor in anchors):
            return row

    return existing_header_row(target, header_row)


def cell_is_blank(cell) -> bool:
    """Ô trống theo nghĩa không có gì để chép: chưa có giá trị và chưa có định dạng."""
    return cell.value is None and not cell.has_style


def merge_row_dimensions(source: Worksheet, target: Worksheet, start_row: int, count: int) -> None:
    """Chép chiều cao hàng cho vùng vừa ghi khi sheet đích chưa quy định hàng đó.

    Sheet đích là bảng báo cáo đã đặt chiều cao cho toàn bộ hàng dữ liệu và cho
    hàng trống phía dưới, nên chỉ những hàng đích chưa có quy định mới lấy chiều
    cao từ nguồn; hàng đã có thì giữ nguyên.
    """
    for offset in range(count):
        row = start_row + offset
        if row in source.row_dimensions and row not in target.row_dimensions:
            target.row_dimensions[row] = copy.copy(source.row_dimensions[row])


def build_column_map(
    source: Worksheet,
    header_row: int,
    target: Worksheet,
    target_header_row: int,
    target_names: dict[str, int] | None = None,
) -> dict[int, int]:
    """Ánh xạ cột nguồn sang cột đích theo tên cột, không theo thứ tự cột.

    Sheet đích thường xếp cột khác thứ tự và có thêm cột riêng (công thức, chỉ
    tiêu theo dõi). Ghi theo vị trí sẽ đẩy dữ liệu sang cột khác nghĩa, nên ở
    đây chỉ những cột có tên trùng mới được ghi, theo đúng tên đó.
    """
    target_names = target_names or header_names(target, target_header_row)
    column_map: dict[int, int] = {}

    for source_column in range(1, source.max_column + 1):
        name = normalise(source.cell(row=header_row, column=source_column).value)
        if not name:
            continue
        target_column = target_names.get(name.casefold())
        if target_column is None:
            continue
        if cell_is_blank(target.cell(row=target_header_row, column=target_column)):
            continue
        column_map[source_column] = target_column

    return column_map


def remap_merges(
    source: Worksheet,
    header_row: int,
    rows: list[int],
    target: Worksheet,
    start_row: int,
    keep_header: bool,
) -> None:
    """Dựng lại các vùng gộp của sheet nguồn theo vị trí mới ở sheet đích.

    Giữ nguyên các vùng gộp đang có ở sheet đích (phần đầu trang). Vùng gộp của
    sheet nguồn chỉ được dựng lại khi nằm trong tập dòng được giữ và không đè
    lên vùng đã có.
    """
    position: dict[int, int] = {}
    if keep_header:
        position[header_row] = start_row - 1
    for offset, row in enumerate(rows):
        position[row] = start_row + offset

    for area in source.merged_cells.ranges:
        if area.max_row - area.min_row != 0 or area.min_row not in position:
            continue
        destination = position[area.min_row]
        address = (
            f"{get_column_letter(area.min_col)}{destination}:"
            f"{get_column_letter(area.max_col)}{destination}"
        )
        clash = any(
            existing.min_row <= destination <= existing.max_row for existing in target.merged_cells.ranges
        )
        if clash:
            continue
        target.merge_cells(address)


def is_formula(value) -> bool:
    """Ô đang chứa công thức (dạng chuỗi "=..." hoặc công thức mảng)."""
    return (isinstance(value, str) and value.startswith("=")) or hasattr(value, "text")


def last_content_row(sheet: Worksheet, header_row: int) -> int:
    """Hàng cuối cùng còn dữ liệu ở bất kỳ cột nào, bỏ qua hàng trắng ở cuối sheet."""
    for row in range(sheet.max_row, header_row, -1):
        if any(
            sheet.cell(row=row, column=column).value not in (None, "")
            for column in range(1, sheet.max_column + 1)
        ):
            return row
    return header_row


def table_extent(
    sheet: Worksheet,
    header_row: int,
    columns: list[int],
    target_row_count: int,
) -> int:
    """Hàng cuối cùng của bảng đích, kể cả hàng tổng cộng nằm dưới dữ liệu.

    Vùng dữ liệu cũ được dò theo một cột chỉ chứa dữ liệu thật (cột công thức
    trải sẵn xuống dưới bảng nên dò theo cột đó sẽ ra tận hàng công thức cuối).
    Phép dò dừng ở hàng cuối của phần sắp ghi, nên phần bảng cũ nằm dưới dữ liệu
    mới luôn được tính là hàng thừa.

    Khi phép dò chạm đúng hàng cuối của phần sắp ghi, bảng cũ có thể còn hàng
    “Tổng cộng” nằm dưới dữ liệu; hàng cuối cùng còn nội dung của sheet mới là
    hàng cuối của bảng, vì phần dưới bảng chỉ có công thức nên không được tính.
    """
    limit = min(sheet.max_row, header_row + target_row_count)

    last_data = header_row
    for row in range(limit, header_row, -1):
        if any(sheet.cell(row=row, column=column).value not in (None, "") for column in columns):
            last_data = row
            break

    if last_data == header_row:
        return header_row

    if last_data < limit:
        return last_data

    return max(last_data, last_content_row(sheet, header_row))


def clear_data_band(
    sheet: Worksheet,
    first_row: int,
    last_row: int,
) -> None:
    """Xoá giá trị dữ liệu thừa trong một vùng hàng, giữ nguyên công thức và biểu mẫu.

    Sheet đích dạng bảng COM có sẵn cột công thức, viền và chiều cao hàng trải
    xuống hết bảng. Xoá hàng bằng delete_rows sẽ cuốn công thức, định dạng số và
    chiều cao hàng lên theo, làm bảng kết quả mất các cột tính toán, nên ở đây
    chỉ xoá giá trị.
    """
    if last_row < first_row:
        return

    for row in range(first_row, min(last_row, sheet.max_row) + 1):
        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=column)
            if is_formula(cell.value):
                continue
            cell.value = None


def prepare_write(
    source: Worksheet,
    header_row: int,
    target: Worksheet,
) -> tuple[int, dict[int, int], bool]:
    """Chuẩn bị ghi: tìm hàng tên cột của đích, ánh xạ cột và xác định hàng bắt đầu.

    Trả về (hàng tên cột của đích, ánh xạ cột, sheet đích đã có hàng tên cột).
    Sheet đích chưa có hàng tiêu đề nào thì chép hàng tiêu đề của nguồn vào và
    ghi theo thứ tự cột, vì không có tên cột để đối chiếu.
    """
    header_row_target = destination_header_row(target, source, header_row)
    target_names = header_names(target, header_row_target)

    if target_names:
        column_map = build_column_map(source, header_row, target, header_row_target, target_names)
        if not column_map:
            source_names = [
                normalise(source.cell(row=header_row, column=column).value)
                for column in range(1, source.max_column + 1)
            ]
            source_names = [name for name in source_names if name]
            raise ProcessError(
                f"Không có cột nào của sheet nguồn trùng tên với cột ở sheet đích "
                f"(hàng tên cột {header_row_target}) nên không ghi được dữ liệu. "
                f"Cột nguồn: {', '.join(source_names) or 'không có'}. "
                f"Cột đích: {available_columns(target, header_row_target) or 'không có'}."
            )
        return header_row_target, column_map, True

    column_map = {
        column: column
        for column in range(1, source.max_column + 1)
        if normalise(source.cell(row=header_row, column=column).value)
    }
    for source_column, target_column in column_map.items():
        copy_cell(
            source.cell(row=header_row, column=source_column),
            target.cell(row=header_row_target, column=target_column),
        )
    return header_row_target, column_map, False


def mapped_destination_columns(column_map: dict[int, int]) -> list[int]:
    """Các cột đích được ghi, xếp theo thứ tự trái sang phải."""
    return sorted(column_map)


def overwrite_sheet(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    from_row: int | None = None,
) -> dict:
    """Ghi dữ liệu đã lọc vào sheet đích, bắt đầu ngay dưới hàng tên cột.

    "Ghi đè" là ghi lại vùng dữ liệu: mọi hàng từ hàng đầu tiên của bảng đích
    trở xuống đều được thay bằng dữ liệu đã lọc, nên dữ liệu cũ còn thừa không
    ở lại. Phần đầu trang của đích (tiêu đề báo cáo, tên nhân viên, tháng) và
    hàng tên cột của chính đích được giữ nguyên; sheet đích chưa có hàng tên cột
    thì ghi thêm hàng tiêu đề của nguồn.

    Chỉ dữ liệu được ghi đè. Cột đích không có tên trùng với cột nguồn vẫn giữ
    nguyên công thức và định dạng sẵn có, vì bảng COM có sẵn các cột tính toán
    tham chiếu tới cột dữ liệu vừa ghi.

    from_row là số hàng bắt đầu ghi ở sheet đích; bỏ trống thì ghi ngay sau hàng
    tên cột. Mốc này chỉ có tác dụng khi nằm dưới hàng dữ liệu đầu tiên của đích:
    ghi đè thì luôn lấp đầy bảng từ hàng đầu tiên, nên một mốc nằm giữa bảng
    không thể để lại khoảng trống phía trên.
    """
    header_row_target, column_map, has_header = prepare_write(source, header_row, target)
    columns = mapped_destination_columns(column_map)
    first_data_row = header_row_target + 1
    last_written_row = first_data_row + len(rows) - 1
    start_row = first_data_row

    if from_row is not None:
        if from_row <= header_row_target:
            raise ProcessError(
                f"Hàng bắt đầu ghi ({from_row}) phải nằm dưới hàng tên cột "
                f"(hàng {header_row_target}) của sheet đích."
            )
        if from_row > first_data_row:
            start_row = from_row

    # Vùng dữ liệu cũ được dò theo một cột chỉ chứa dữ liệu thật: cột công thức
    # trải sẵn xuống dưới bảng nên dò theo cột đó sẽ ra tận hàng công thức cuối.
    # Hàng cuối của bảng đích gồm cả hàng “Tổng cộng” nằm dưới dữ liệu, vì hàng
    # đó cũng là dữ liệu cũ và phải bị xoá khi bảng mới ngắn hơn.
    old_extent = table_extent(target, header_row_target, columns, len(rows))
    clear_to = max(old_extent, last_written_row)

    clear_data_band(target, start_row, clear_to)
    write_rows(source, target, rows, header_row, start_row=start_row, column_map=column_map)
    merge_row_dimensions(source, target, start_row, len(rows))
    remap_merges(source, header_row, rows, target, start_row, keep_header=has_header)
    return {"written": len(rows), "header_row": header_row_target, "start_row": start_row}


def write_from_mark(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    from_row: int,
    column_map: dict[int, int],
) -> None:
    """Ghi lại toàn bộ vùng dữ liệu từ mốc from_row trở xuống.

    Hàng cũ phải bị xoá hết tới hàng cuối của vùng sắp ghi, không chỉ tới vùng dữ
    liệu cũ: dữ liệu mới ngắn hơn thì phần bảng cũ nằm dưới vẫn còn nguyên nếu chỉ
    xoá tới hàng cuối vùng dữ liệu cũ.
    """
    last_written_row = from_row + len(rows) - 1
    old_extent = table_extent(target, from_row - 1, mapped_destination_columns(column_map), len(rows))
    clear_to = max(old_extent, last_written_row)

    clear_data_band(target, from_row, clear_to)
    write_rows(source, target, rows, header_row, start_row=from_row, column_map=column_map)
    merge_row_dimensions(source, target, from_row, len(rows))
    remap_merges(source, header_row, rows, target, from_row, keep_header=False)


def update_sheet(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    from_row: int | None = None,
    take_from_row: int | None = None,
) -> dict:
    """Ghi dữ liệu vào sheet đích, đối chiếu theo cột khoá.

    Khoá là cột dữ liệu trái nhất của bảng đích (thường là mã hàng). Không truyền
    from_row thì hành vi là cập nhật: dòng trùng khoá được ghi đè tại chỗ, dòng
    mới thêm vào cuối.

    Nghĩa là: mốc nằm dưới hàng dữ liệu đầu tiên thì ghi đè từ mốc đó xuống; mốc
    nằm ngay hàng đầu tiên (hoặc giữa bảng, bị kéo lên hàng đầu tiên) thì ghi đè
    cả bảng; không có mốc thì cập nhật theo cột khoá, dòng trùng khoá ghi đè tại
    chỗ và dòng mới thêm vào cuối.
    """
    header_row_target, column_map, _ = prepare_write(source, header_row, target)

    # Khoá là cột dữ liệu trái nhất của bảng đích (thường là mã hàng), không
    # phải cột của quy tắc lọc đầu tiên: quy tắc lọc thường đặt ở cột phân loại
    # nên giá trị lặp lại nhiều lần và không dùng làm khoá được.
    key_source_column = min(column_map, key=lambda column: column_map[column])
    key_name = normalise(source.cell(row=header_row, column=key_source_column).value)
    key_column = column_map[key_source_column]
    first_data_row = header_row_target + 1

    if from_row is not None and from_row <= header_row_target:
        raise ProcessError(
            f"Hàng bắt đầu ghi ({from_row}) phải nằm dưới hàng tên cột "
            f"(hàng {header_row_target}) của sheet đích."
        )

    if from_row is not None and from_row > first_data_row:
        write_from_mark(source, target, rows, header_row, from_row, column_map)
        return {
            "written": len(rows),
            "overwritten_from": from_row,
            "take_from_row": take_from_row,
            "key_column": key_name,
            "start_row": from_row,
        }

    # Mốc nằm ngay hàng dữ liệu đầu tiên (hoặc nằm giữa bảng nên bị kéo lên đó)
    # nghĩa là ghi đè cả bảng. Đi theo đường đối chiếu khoá sẽ báo lỗi khi dữ liệu
    # nguồn có khoá trùng, trong khi ghi đè cả bảng không cần khoá nào cả.
    if from_row is not None:
        return overwrite_sheet(source, target, rows, header_row)

    existing: dict[str, int] = {}
    old_extent = table_extent(
        target,
        header_row_target,
        mapped_destination_columns(column_map),
        target.max_row - header_row_target,
    )
    old_extent = max(old_extent, first_data_row - 1)
    for row in range(header_row_target + 1, old_extent + 1):
        value = normalise(target.cell(row=row, column=key_column).value)
        if value:
            existing.setdefault(value, row)

    inserted = 0
    updated = 0
    appended = old_extent
    seen: dict[str, int] = {}
    row_map: dict[int, int] = {}

    for row in rows:
        key = normalise(source.cell(row=row, column=key_source_column).value)
        if not key:
            raise ProcessError(
                f"Dòng {row} không có giá trị ở cột khoá “{key_name}” nên không thể cập nhật."
            )
        if key in seen:
            raise ProcessError(
                f"Cột khoá “{key_name}” có giá trị trùng “{key}” ở dòng {seen[key]} và {row} "
                f"trong dữ liệu nguồn."
            )
        seen[key] = row

        if key in existing:
            destination = existing[key]
            updated += 1
        else:
            appended += 1
            destination = appended
            existing[key] = destination
            inserted += 1

        # Cập nhật theo khoá rải dòng đi khắp bảng chứ không ghi liền mạch, nên
        # phải ghi lại đích thực của từng dòng để bước ghép cột đi theo đúng dòng.
        row_map[row] = destination

        for source_column, target_column in column_map.items():
            copy_cell(
                source.cell(row=row, column=source_column),
                target.cell(row=destination, column=target_column),
                keep_target_style=True,
            )
        if row in source.row_dimensions and destination not in target.row_dimensions:
            target.row_dimensions[destination] = copy.copy(source.row_dimensions[row])

    # Ghi bằng khoá có thể dùng ít hàng hơn số hàng đang có ở sheet đích;
    # xoá dữ liệu cũ bên dưới để file kết quả không còn dòng thừa, nhưng giữ
    # công thức và định dạng của bảng đích. Vùng xoá tính theo hàng cuối của
    # bảng đích — kể cả hàng “Tổng cộng” — và phải phủ tới hàng cuối của phần
    # vừa ghi, nếu không hàng cũ nằm ngoài vùng dữ liệu cũ sẽ ở lại.
    columns = mapped_destination_columns(column_map)
    reference_row = appended - 1 if appended > header_row_target else header_row_target
    clear_to = max(table_extent(target, reference_row, columns, appended - reference_row), appended)
    clear_data_band(target, appended + 1, clear_to)

    return {
        "updated": updated,
        "inserted": inserted,
        "key_column": key_name,
        "start_row": header_row_target + 1,
        "row_map": row_map,
    }


# --------------------------------------------------------------------------- #
# Điều phối
# --------------------------------------------------------------------------- #


def load(path: str, keep_vba: bool = False):
    try:
        return load_workbook(path, keep_vba=keep_vba)
    except Exception as exc:
        raise ProcessError(f"Không mở được file “{os.path.basename(path)}”: {exc}") from exc


def require_sheet(workbook, name: str, label: str) -> Worksheet:
    if name not in workbook.sheetnames:
        available = ", ".join(f"“{sheet}”" for sheet in workbook.sheetnames)
        raise ProcessError(f"Sheet “{name}” không có trong {label}. Các sheet hiện có: {available}.")
    return workbook[name]


def process(
    input1: str,
    sheet1: str,
    input2: str,
    sheet2: str,
    output_path: str,
    header_row: int,
    rules: list[dict],
    write_mode: str = WRITE_OVERWRITE,
    apply_filter: bool = True,
    from_row: int | None = None,
    take_from_row: int | None = None,
    merge_rules: list[dict] | None = None,
    merge_header1: int | None = None,
    merge_header2: int | None = None,
) -> dict:
    """Lọc sheet1 của input1 rồi ghi sang sheet2 của input2, lưu thành output_path.

    take_from_row là số hàng bắt đầu lấy ở sheet nguồn (Input 1); bỏ trống thì
    lấy từ ngay dưới hàng tên cột. from_row là số hàng bắt đầu ghi ở sheet đích;
    bỏ trống thì ghi ngay dưới hàng tên cột. Ghi đè luôn ghi từ hàng dữ liệu đầu
    tiên của bảng đích trở xuống, nên from_row chỉ có tác dụng khi nằm dưới hàng
    đó; ở chế độ cập nhật theo khoá thì from_row chuyển sang kiểu ghi đè từ mốc
    đó xuống.

    merge_rules ghép cột: mỗi quy tắc lấy cột "source" ở sheet nguồn ghi sang
    cột "target" ở sheet đích, chạy sau khi dữ liệu đã được ghi. merge_header1 và
    merge_header2 là hàng tên cột của hai sheet khi ghép, vì sheet đích có thể
    đặt tên cột ở hàng khác sheet nguồn.
    """
    source_book = load(input1, keep_vba=input1.lower().endswith(".xlsm"))
    source = require_sheet(source_book, sheet1, "Input 1")

    if apply_filter and rules:
        result = filter_rows(source, header_row, rules)
        rows = result.kept_rows
    else:
        result = FilterResult()
        rows = list(range(header_row + 1, last_table_row(source, header_row) + 1))
        result.total_rows = len(rows)
        result.kept_rows = rows

    if take_from_row is not None:
        if take_from_row <= header_row:
            raise ProcessError(
                f"Hàng lấy dữ liệu ({take_from_row}) phải nằm dưới hàng tên cột "
                f"(hàng {header_row}) của sheet nguồn."
            )
        rows = [row for row in rows if row >= take_from_row]

    target_book = load(input2, keep_vba=input2.lower().endswith(".xlsm"))
    target = require_sheet(target_book, sheet2, "Input 2")

    if write_mode == WRITE_UPDATE:
        detail = update_sheet(
            source, target, rows, header_row, from_row=from_row, take_from_row=take_from_row
        )
    else:
        detail = overwrite_sheet(source, target, rows, header_row, from_row=from_row)
        # Sheet đích đã có sẵn phần đầu trang thì giữ nguyên định dạng của nó,
        # chỉ chép độ rộng cột khi sheet đích chưa quy định cột nào.
        if not target.column_dimensions:
            copy_columns(source, target, source.max_column)

    merged = []
    if merge_rules:
        # Ghép cột chạy trên đúng những dòng vừa được ghi, nên cột đích khớp
        # với dữ liệu và phần giữ nguyên phía trên không bị ghi nửa vời.
        # Mỗi bên có hàng tên cột riêng vì sheet đích có thể đặt ở hàng khác.
        target_header_row = merge_header2 or destination_header_row(target, source, header_row)
        source_header_row = merge_header1 or header_row
        # Cập nhật theo khoá ghi mỗi dòng vào đúng dòng có khoá đó trong bảng
        # đích, không ghi liền mạch, nên phải dùng lại ánh xạ dòng của bước ghi.
        start_row = detail.get("start_row", target_header_row + 1)
        row_map = detail.get("row_map") or {
            source_row: start_row + offset for offset, source_row in enumerate(rows)
        }
        merged = merge_columns(
            source, target, source_header_row, target_header_row, merge_rules, row_map=row_map
        )

    target_book.save(output_path)

    return {
        "total_rows": result.total_rows,
        "kept_rows": result.kept,
        "removed_rows": result.removed,
        "per_rule": result.per_rule,
        "write_mode": write_mode,
        "detail": detail,
        "merged_columns": merged,
    }
