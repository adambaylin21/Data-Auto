"""Xử lý file Excel: lọc dữ liệu theo quy tắc trong options.json rồi ghi sang Input 2.

Toàn bộ hàm ở đây thuần Python, không phụ thuộc Flask, để có thể kiểm thử riêng.

Quy tắc lọc: mỗi quy tắc gồm type (Giữ/Loại), column, match (Chứa/Bằng/Trống Không)
và keyword. Nhiều quy tắc cùng loại là danh sách để chọn (OR); giữa hai nhóm là giao:

    dòng được giữ  ⇔  khớp ÍT NHẤT MỘT quy tắc "Giữ"  AND  không khớp quy tắc "Loại" nào

Khi không có quy tắc "Giữ" nào thì mọi dòng đều qua bước Giữ.
"""

from __future__ import annotations

import copy
import os
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


def resolve_columns(mapping: dict[str, int], rules: list[dict], sheet_name: str, header_row: int) -> None:
    """Báo lỗi nếu quy tắc nhắc tới cột không có trong sheet."""
    missing = [rule["column"] for rule in rules if rule["column"] not in mapping]
    if not missing:
        return

    unique = list(dict.fromkeys(missing))
    listed = ", ".join(f"“{name}”" for name in unique)
    available = ", ".join(f"“{name}”" for name in mapping)
    raise ProcessError(
        f"Không tìm thấy cột {listed} trong sheet “{sheet_name}” "
        f"(hàng tiêu đề {header_row}). Các cột hiện có: {available}."
    )


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


def split_rules(rules: list[dict]) -> tuple[list[dict], list[dict]]:
    keep_rules = [rule for rule in rules if rule.get("type") == KEEP]
    drop_rules = [rule for rule in rules if rule.get("type") == DROP]
    return keep_rules, drop_rules


def filter_rows(sheet: Worksheet, header_row: int, rules: list[dict]) -> FilterResult:
    """Lọc các dòng dữ liệu của sheet theo danh sách quy tắc."""
    mapping = header_map(sheet, header_row)
    resolve_columns(mapping, rules, sheet.title, header_row)

    keep_rules, drop_rules = split_rules(rules)
    result = FilterResult(per_rule={describe(rule): 0 for rule in rules})

    last_row = last_data_row(sheet, header_row)
    for row in range(header_row + 1, last_row + 1):
        result.total_rows += 1

        matched_keep = [rule for rule in keep_rules if cell_matches(cell(sheet, row, mapping, rule), rule)]
        matched_drop = [rule for rule in drop_rules if cell_matches(cell(sheet, row, mapping, rule), rule)]

        for rule in matched_keep + matched_drop:
            result.per_rule[describe(rule)] += 1

        passes_keep = bool(matched_keep) if keep_rules else True
        if passes_keep and not matched_drop:
            result.kept_rows.append(row)

    return result


def cell(sheet: Worksheet, row: int, mapping: dict[str, int], rule: dict):
    column = mapping.get(rule["column"])
    return None if column is None else sheet.cell(row=row, column=column).value


def last_data_row(sheet: Worksheet, header_row: int) -> int:
    """Hàng cuối cùng còn dữ liệu, bỏ qua các hàng trắng ở cuối sheet."""
    for row in range(sheet.max_row, header_row, -1):
        if any(sheet.cell(row=row, column=column).value not in (None, "") for column in range(1, sheet.max_column + 1)):
            return row
    return header_row


def describe(rule: dict) -> str:
    """Câu mô tả quy tắc, khớp với câu hiện trên giao diện."""
    text = f"Cột {rule['column']} {str(rule['type']).lower()} các hàng {rule['match']}"
    keyword = str(rule.get("keyword") or "")
    return f"{text} {keyword}".strip()


# --------------------------------------------------------------------------- #
# Ghi sang sheet đích
# --------------------------------------------------------------------------- #


def copy_cell(source, target) -> None:
    """Sao chép giá trị và định dạng của một ô.

    Định dạng phải gán bằng các đối tượng style thật (font, fill, border,
    alignment, protection). Gán thẳng `_style` sẽ chỉ chép chỉ số tra cứu, mà
    chỉ số đó thuộc bảng style của file nguồn nên khi mở file kết quả sẽ trỏ sai.
    """
    target.value = source.value
    if source.has_style:
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.protection = copy.copy(source.protection)
        target.number_format = source.number_format
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


def write_rows(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    start_row: int,
    copy_height: bool = True,
) -> None:
    """Ghi các dòng đã chọn từ source sang target, giữ định dạng từng ô."""
    for offset, row in enumerate(rows):
        destination = start_row + offset
        for column in range(1, source.max_column + 1):
            copy_cell(source.cell(row=row, column=column), target.cell(row=destination, column=column))
        if copy_height and row in source.row_dimensions:
            target.row_dimensions[destination] = copy.copy(source.row_dimensions[row])


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


def clear_data_below(sheet: Worksheet, keep_until: int) -> None:
    """Xoá dữ liệu từ sau hàng giữ lại, chừa phần đầu trang và hàng tiêu đề."""
    if sheet.max_row > keep_until:
        sheet.delete_rows(keep_until + 1, sheet.max_row - keep_until)


def overwrite_sheet(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    from_row: int | None = None,
) -> dict:
    """Ghi các dòng đã lọc vào sheet đích, giữ nguyên phần đầu trang của đích.

    Sheet đích giữ tiêu đề báo cáo, tên nhân viên, tháng và hàng tên cột của
    chính nó; chỉ dữ liệu bên dưới hàng tên cột được thay bằng dữ liệu đã lọc.
    Nếu sheet đích chưa có hàng tên cột thì ghi thêm hàng tiêu đề của nguồn.

    from_row là số hàng bắt đầu ghi ở sheet đích; bỏ trống thì ghi ngay sau
    hàng tên cột. Các hàng phía trên from_row được giữ nguyên.
    """
    header_row_target = existing_header_row(target, header_row)
    has_header = normalise(target.cell(row=header_row_target, column=1).value) != ""

    if has_header:
        start_row = header_row_target + 1
    else:
        start_row = header_row_target
        copy_cell(source.cell(row=header_row, column=1), target.cell(row=start_row, column=1))

    if from_row is not None:
        if from_row <= header_row_target:
            raise ProcessError(
                f"Hàng bắt đầu ghi ({from_row}) phải nằm dưới hàng tên cột "
                f"(hàng {header_row_target}) của sheet đích."
            )
        start_row = from_row

    clear_data_below(target, start_row - 1)
    write_rows(source, target, rows, header_row, start_row=start_row)
    remap_merges(source, header_row, rows, target, start_row, keep_header=has_header)
    return {"written": len(rows), "header_row": header_row_target}


def update_sheet(
    source: Worksheet,
    target: Worksheet,
    rows: list[int],
    header_row: int,
    mapping: dict[str, int],
    from_row: int | None = None,
) -> dict:
    """Cập nhật sheet đích theo cột khoá: dòng có khoá trùng thì ghi đè, chưa có thì thêm.

    Khoá là cột đầu tiên của hàng tiêu đề (thường là mã hàng).

    from_row là số hàng bắt đầu ghi ở sheet đích. Chỉ dòng từ from_row trở xuống
    được đối chiếu khoá, nên dữ liệu đã có phía trên được giữ nguyên; dòng mới
    cũng chỉ được thêm từ from_row trở xuống.
    """
    key_column = min(mapping.values())
    key_name = next(name for name, column in mapping.items() if column == key_column)

    header_row_target = existing_header_row(target, header_row)
    if from_row is not None and from_row <= header_row_target:
        raise ProcessError(
            f"Hàng bắt đầu ghi ({from_row}) phải nằm dưới hàng tên cột "
            f"(hàng {header_row_target}) của sheet đích."
        )

    scan_from = from_row or header_row_target + 1
    existing: dict[str, int] = {}
    for row in range(scan_from, last_data_row(target, header_row_target) + 1):
        value = normalise(target.cell(row=row, column=key_column).value)
        if value:
            existing.setdefault(value, row)

    inserted = 0
    updated = 0
    appended = max(last_data_row(target, header_row_target), scan_from - 1)
    seen: dict[str, int] = {}

    for row in rows:
        key = normalise(source.cell(row=row, column=key_column).value)
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

        for column in range(1, source.max_column + 1):
            copy_cell(source.cell(row=row, column=column), target.cell(row=destination, column=column))

    return {"updated": updated, "inserted": inserted, "key_column": key_name}


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
) -> dict:
    """Lọc sheet1 của input1 rồi ghi sang sheet2 của input2, lưu thành output_path.

    from_row là số hàng bắt đầu ghi ở sheet đích; bỏ trống thì ghi ngay dưới
    hàng tên cột như trước.
    """
    source_book = load(input1, keep_vba=input1.lower().endswith(".xlsm"))
    source = require_sheet(source_book, sheet1, "Input 1")

    mapping = header_map(source, header_row)

    if apply_filter and rules:
        resolve_columns(mapping, rules, source.title, header_row)
        result = filter_rows(source, header_row, rules)
        rows = result.kept_rows
    else:
        result = FilterResult()
        rows = list(range(header_row + 1, last_data_row(source, header_row) + 1))
        result.total_rows = len(rows)
        result.kept_rows = rows

    target_book = load(input2, keep_vba=input2.lower().endswith(".xlsm"))
    target = require_sheet(target_book, sheet2, "Input 2")

    if write_mode == WRITE_UPDATE:
        detail = update_sheet(source, target, rows, header_row, mapping, from_row=from_row)
    else:
        detail = overwrite_sheet(source, target, rows, header_row, from_row=from_row)
        # Sheet đích đã có sẵn phần đầu trang thì giữ nguyên định dạng của nó,
        # chỉ chép độ rộng cột khi sheet đích chưa quy định cột nào.
        if not target.column_dimensions:
            copy_columns(source, target, source.max_column)

    target_book.save(output_path)

    return {
        "total_rows": result.total_rows,
        "kept_rows": result.kept,
        "removed_rows": result.removed,
        "per_rule": result.per_rule,
        "write_mode": write_mode,
        "detail": detail,
    }
