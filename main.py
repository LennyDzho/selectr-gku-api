from __future__ import annotations

import math
import os
import re
import sys
import tempfile
from zipfile import BadZipFile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
from typing import Callable, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import from_excel
from openpyxl.utils.exceptions import InvalidFileException


def _application_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APPLICATION_DIRECTORY = _application_directory()
INPUT_DIRECTORY = APPLICATION_DIRECTORY / "input"
OUTPUT_FILE = APPLICATION_DIRECTORY / "output" / "output.xlsx"

CREATED_AT = "Дата создания"
VISIT_START = "Дата начала"
LOGIN = "Логин"
ROUTE = "Номер маршрута"
BOARD = "Бортовой номер"
TICKET_ID = "Уникальный номер билета"
TICKET_TYPE = "Название типа билета"

REQUIRED_HEADERS = (
    CREATED_AT,
    VISIT_START,
    LOGIN,
    ROUTE,
    BOARD,
    TICKET_ID,
    TICKET_TYPE,
)
OUTPUT_HEADERS = (
    "Дата",
    LOGIN,
    ROUTE,
    BOARD,
    "Кол-во проверок",
    "Кол-во проверок ТС",
)

ZERO_TICKET_ID = "00000000000000000000"
SOCIAL_CARD_TYPE = "социальная карта москвича"


class ReportError(Exception):
    """Base error for an expected report-generation failure."""


class WorkbookSchemaError(ReportError):
    """Raised when an input workbook cannot be mapped to the required schema."""


class NoValidDataError(ReportError):
    """Raised when no report can be produced from the available input."""


@dataclass(frozen=True, slots=True)
class Record:
    timestamp: datetime
    login: str
    route: str
    board: str
    ticket_id: str
    visit_start: datetime
    ticket_type: str = ""
    source_name: str = ""
    source_row: int = 0


@dataclass(frozen=True, slots=True)
class Visit:
    start: datetime
    login: str
    route: str
    board: str
    check_count: int
    visit_number: int
    visit_total: int
    counted_records: tuple[Record, ...] = ()


@dataclass(frozen=True, slots=True)
class DailyReportRow:
    day: date
    login: str
    route: str
    board: str
    check_count: int
    vehicle_check_count: int


@dataclass(frozen=True, slots=True)
class InspectorDay:
    day: date
    login: str
    check_count: int
    vehicle_check_count: int
    counted_records: tuple[Record, ...]


@dataclass(frozen=True, slots=True)
class ReportSections:
    main_rows: tuple[DailyReportRow, ...]
    low_visits: tuple[Visit, ...]
    low_days: tuple[InspectorDay, ...]


@dataclass(frozen=True, slots=True)
class ProcessingSummary:
    input_files: int
    processed_files: int
    skipped_files: int
    valid_rows: int
    skipped_rows: int
    visits: int
    checks: int
    main_visits: int = 0
    main_checks: int = 0
    low_visits: int = 0
    low_checks: int = 0
    low_days: int = 0
    low_day_checks: int = 0


WarningHandler = Callable[[str], None]


def warn(message: str) -> None:
    print(f"Предупреждение: {message}", file=sys.stderr)


def _normalise_header(value: object) -> str:
    return "" if value is None else str(value).strip()


def _cell_to_text(cell: Cell) -> str:
    """Convert an identifier cell to text without stripping existing leading zeroes."""
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            return ""
        text = str(int(value)) if value.is_integer() else format(value, ".15g")
    else:
        return str(value).strip()

    # A numeric cell formatted as 00000000 is an Excel representation of a
    # text-like identifier. Restore the display zeroes when that is unambiguous.
    number_format = str(cell.number_format or "").split(";", 1)[0]
    if re.fullmatch(r"0+", number_format) and not text.startswith("-"):
        text = text.zfill(len(number_format))
    return text


def _parse_timestamp(cell: Cell, workbook_epoch: datetime) -> datetime:
    value = cell.value
    parsed: datetime

    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime.combine(value, time.min)
    elif isinstance(value, (int, float)) and not isinstance(value, bool) and cell.is_date:
        converted = from_excel(value, workbook_epoch)
        parsed = (
            converted
            if isinstance(converted, datetime)
            else datetime.combine(converted, time.min)
        )
    elif isinstance(value, str):
        text_value = value.strip()
        if not text_value:
            raise ValueError("пустая дата")

        normalised = text_value.replace(",", ".")
        try:
            parsed = datetime.fromisoformat(normalised)
        except ValueError:
            parsed = _parse_timestamp_with_known_formats(normalised)
    else:
        raise ValueError(f"неподдерживаемое значение даты: {value!r}")

    if parsed.tzinfo is not None:
        raise ValueError("дата с часовым поясом не поддерживается")
    return parsed


def _parse_timestamp_with_known_formats(value: str) -> datetime:
    formats = (
        "%d.%m.%Y %H:%M:%S.%f",
        "%d.%m.%Y %H:%M:%S",
        "%Y/%m/%d %H:%M:%S.%f",
        "%Y/%m/%d %H:%M:%S",
    )
    for timestamp_format in formats:
        try:
            return datetime.strptime(value, timestamp_format)
        except ValueError:
            continue
    raise ValueError(f"не удалось распознать дату {value!r}")


def read_records(
    workbook_path: Path,
    warning_handler: WarningHandler = warn,
) -> tuple[list[Record], int]:
    """Read valid records from the active sheet and return them with a skip count."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows()
        header_row = next(rows, None)
        if header_row is None:
            raise WorkbookSchemaError("активный лист пуст")

        header_positions: dict[str, int] = {}
        duplicate_headers: set[str] = set()
        for index, cell in enumerate(header_row):
            header = _normalise_header(cell.value)
            if header not in REQUIRED_HEADERS:
                continue
            if header in header_positions:
                duplicate_headers.add(header)
            else:
                header_positions[header] = index

        if duplicate_headers:
            duplicates = ", ".join(sorted(duplicate_headers))
            raise WorkbookSchemaError(f"повторяются обязательные колонки: {duplicates}")

        missing_headers = [
            header for header in REQUIRED_HEADERS if header not in header_positions
        ]
        if missing_headers:
            raise WorkbookSchemaError(
                "нет обязательных колонок: " + ", ".join(missing_headers)
            )

        records: list[Record] = []
        skipped_rows = 0
        for row_number, row in enumerate(rows, start=2):
            cells = {header: row[index] for header, index in header_positions.items()}

            # Styled but otherwise empty trailing rows are not data errors.
            if all(cell.value is None for cell in cells.values()):
                continue

            errors: list[str] = []
            try:
                timestamp = _parse_timestamp(cells[CREATED_AT], workbook.epoch)
            except (TypeError, ValueError) as error:
                timestamp = None
                errors.append(f"«{CREATED_AT}»: {error}")

            try:
                visit_start = _parse_timestamp(cells[VISIT_START], workbook.epoch)
            except (TypeError, ValueError) as error:
                visit_start = None
                errors.append(f"«{VISIT_START}»: {error}")

            login = _cell_to_text(cells[LOGIN])
            route = _cell_to_text(cells[ROUTE])
            board = _cell_to_text(cells[BOARD])
            ticket_id = _cell_to_text(cells[TICKET_ID])
            ticket_type = _cell_to_text(cells[TICKET_TYPE])

            for field_name, field_value in (
                (LOGIN, login),
                (ROUTE, route),
                (BOARD, board),
                (TICKET_ID, ticket_id),
            ):
                if not field_value:
                    errors.append(f"пустое поле «{field_name}»")

            if errors:
                skipped_rows += 1
                warning_handler(
                    f"{workbook_path.name}, строка {row_number}: " + "; ".join(errors)
                )
                continue

            assert timestamp is not None
            assert visit_start is not None
            records.append(
                Record(
                    timestamp=timestamp,
                    login=login,
                    route=route,
                    board=board,
                    ticket_id=ticket_id,
                    visit_start=visit_start,
                    ticket_type=ticket_type,
                    source_name=workbook_path.name,
                    source_row=row_number,
                )
            )

        return records, skipped_rows
    finally:
        workbook.close()


def _record_sort_key(record: Record) -> tuple:
    return (
        record.timestamp,
        record.source_name.casefold(),
        record.source_name,
        record.source_row,
        record.visit_start,
        record.login.casefold(),
        record.login,
        record.route.casefold(),
        record.route,
        record.board.casefold(),
        record.board,
        record.ticket_id,
        record.ticket_type,
    )


def _counted_passenger_records(records: Sequence[Record]) -> list[Record]:
    """Apply the existing social/zero rules before cross-visit deduplication."""
    social_card_seconds = {
        record.timestamp.replace(microsecond=0)
        for record in records
        if record.ticket_id != ZERO_TICKET_ID
        and " ".join(record.ticket_type.split()).casefold() == SOCIAL_CARD_TYPE
    }
    selected: list[Record] = []
    ticket_ids: set[str] = set()
    for record in sorted(records, key=_record_sort_key):
        if record.ticket_id == ZERO_TICKET_ID:
            if record.timestamp.replace(microsecond=0) not in social_card_seconds:
                selected.append(record)
        elif record.ticket_id not in ticket_ids:
            selected.append(record)
            ticket_ids.add(record.ticket_id)
    return selected


def count_passenger_checks(records: Sequence[Record]) -> int:
    """Count one visit in isolation; build_visits also deduplicates by day."""
    return len(_counted_passenger_records(records))


def build_visits(records: Iterable[Record]) -> list[Visit]:
    """Keep source visits; credit tickets once per day/login/route/board."""
    groups: dict[
        tuple[date, str, str, str], dict[datetime, list[Record]]
    ] = defaultdict(lambda: defaultdict(list))
    for record in records:
        key = (record.visit_start.date(), record.login, record.route, record.board)
        groups[key][record.visit_start].append(record)

    candidates: list[Record] = []
    for starts in groups.values():
        for session in starts.values():
            candidates.extend(_counted_passenger_records(session))

    counted_by_visit: dict[
        tuple[datetime, str, str, str], list[Record]
    ] = defaultdict(list)
    daily_tickets: set[tuple[date, str, str, str, str]] = set()
    for record in sorted(candidates, key=_record_sort_key):
        if record.ticket_id != ZERO_TICKET_ID:
            ticket_key = (
                record.visit_start.date(), record.login, record.route,
                record.board, record.ticket_id,
            )
            if ticket_key in daily_tickets:
                continue
            daily_tickets.add(ticket_key)
        visit_key = (record.visit_start, record.login, record.route, record.board)
        counted_by_visit[visit_key].append(record)

    visits: list[Visit] = []
    for (_day, login, route, board), starts in groups.items():
        sessions = sorted(starts.items())
        visit_total = len(sessions)
        for visit_number, (visit_start, session) in enumerate(sessions, start=1):
            counted = tuple(counted_by_visit[(visit_start, login, route, board)])
            visits.append(
                Visit(
                    start=visit_start,
                    login=login,
                    route=route,
                    board=board,
                    check_count=len(counted),
                    visit_number=visit_number,
                    visit_total=visit_total,
                    counted_records=counted,
                )
            )

    visits.sort(
        key=lambda visit: (
            visit.start,
            visit.login.casefold(),
            visit.route.casefold(),
            visit.board.casefold(),
            visit.visit_number,
        )
    )
    return visits


def build_daily_report_rows(visits: Iterable[Visit]) -> list[DailyReportRow]:
    """Combine every visit of one inspector/route/board into one daily row."""
    grouped_visits: dict[tuple[date, str, str, str], list[Visit]] = defaultdict(list)
    for visit in visits:
        key = (visit.start.date(), visit.login, visit.route, visit.board)
        grouped_visits[key].append(visit)

    report_rows = [
        DailyReportRow(
            day=day,
            login=login,
            route=route,
            board=board,
            check_count=sum(visit.check_count for visit in daily_visits),
            vehicle_check_count=len(daily_visits),
        )
        for (day, login, route, board), daily_visits in grouped_visits.items()
    ]
    report_rows.sort(
        key=lambda row: (
            row.day,
            row.login.casefold(),
            row.route.casefold(),
            row.board.casefold(),
        )
    )
    return report_rows


def build_report_sections(visits: Iterable[Visit]) -> ReportSections:
    """Exclude visits with at most two checks and days below thirty checks."""
    visits = list(visits)
    grouped_days: dict[tuple[date, str], list[Visit]] = defaultdict(list)
    for visit in visits:
        grouped_days[(visit.start.date(), visit.login)].append(visit)

    low_day_keys: set[tuple[date, str]] = set()
    low_days: list[InspectorDay] = []
    for (day, login), day_visits in grouped_days.items():
        check_count = sum(visit.check_count for visit in day_visits)
        if check_count < 30:
            low_day_keys.add((day, login))
            counted = sorted(
                (
                    record for visit in day_visits
                    for record in visit.counted_records
                ),
                key=_record_sort_key,
            )
            low_days.append(
                InspectorDay(day, login, check_count, len(day_visits), tuple(counted))
            )
    low_days.sort(key=lambda item: (item.day, item.login.casefold(), item.login))

    low_visits = tuple(
        sorted(
            (visit for visit in visits if visit.check_count <= 2),
            key=lambda visit: (
                visit.start, visit.login.casefold(), visit.route.casefold(),
                visit.board.casefold(), visit.login, visit.route, visit.board,
            ),
        )
    )
    eligible = [
        visit for visit in visits
        if visit.check_count > 2
        and (visit.start.date(), visit.login) not in low_day_keys
    ]
    return ReportSections(
        main_rows=tuple(build_daily_report_rows(eligible)),
        low_visits=low_visits,
        low_days=tuple(low_days),
    )


def _format_report_sheet(
    worksheet,
    widths: Sequence[int],
    *,
    datetime_columns: Sequence[int] = (),
    text_columns: Sequence[int] = (2, 3, 4),
) -> None:
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in worksheet.iter_rows(min_row=2):
        row[0].number_format = "yyyy-mm-dd"
        for column in datetime_columns:
            row[column - 1].number_format = "yyyy-mm-dd hh:mm:ss.000"
        for column in text_columns:
            cell = row[column - 1]
            cell.number_format = "@"
            if isinstance(cell.value, str):
                cell.data_type = "s"
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_cells, width in zip(worksheet.columns, widths):
        worksheet.column_dimensions[column_cells[0].column_letter].width = width


def write_report(
    report_rows: Sequence[DailyReportRow],
    output_path: Path,
    *,
    low_visits: Sequence[Visit] = (),
    low_days: Sequence[InspectorDay] = (),
) -> None:
    """Write the main and both exclusion sheets, replacing only on success."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Проверки"
    worksheet.append(OUTPUT_HEADERS)
    for report_row in report_rows:
        worksheet.append(
            (
                report_row.day,
                report_row.login,
                report_row.route,
                report_row.board,
                report_row.check_count,
                report_row.vehicle_check_count,
            )
        )
    _format_report_sheet(worksheet, (13, 20, 18, 18, 20, 21))

    low_sheet = workbook.create_sheet(">=2")
    low_sheet.append((*OUTPUT_HEADERS, VISIT_START))
    for visit in low_visits:
        low_sheet.append(
            (
                visit.start.date(),
                visit.login,
                visit.route,
                visit.board,
                visit.check_count,
                1,
                visit.start,
            )
        )
    _format_report_sheet(
        low_sheet, (13, 20, 18, 18, 20, 21, 27), datetime_columns=(7,)
    )

    low_day_sheet = workbook.create_sheet(">30")
    low_day_sheet.append(
        (*OUTPUT_HEADERS, CREATED_AT, VISIT_START, TICKET_ID, TICKET_TYPE)
    )
    day_blocks: list[tuple[int, int]] = []
    for inspector_day in low_days:
        first_row = low_day_sheet.max_row + 1
        # A zero-check day still needs a row for its date, login and totals.
        for index, record in enumerate(inspector_day.counted_records or (None,)):
            low_day_sheet.append(
                (
                    inspector_day.day if index == 0 else None,
                    inspector_day.login if index == 0 else None,
                    record.route if record is not None else None,
                    record.board if record is not None else None,
                    inspector_day.check_count if index == 0 else None,
                    inspector_day.vehicle_check_count if index == 0 else None,
                    record.timestamp if record is not None else None,
                    record.visit_start if record is not None else None,
                    record.ticket_id if record is not None else None,
                    record.ticket_type if record is not None else None,
                )
            )
        day_blocks.append((first_row, low_day_sheet.max_row))
    _format_report_sheet(
        low_day_sheet,
        (13, 20, 18, 18, 20, 21, 27, 27, 27, 32),
        datetime_columns=(7, 8),
        text_columns=(2, 3, 4, 9, 10),
    )
    for row in low_day_sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for first_row, last_row in day_blocks:
        if first_row < last_row:
            # Route/board stay alongside each card, including when they change.
            for column in (1, 2, 5, 6):
                low_day_sheet.merge_cells(
                    start_row=first_row,
                    end_row=last_row,
                    start_column=column,
                    end_column=column,
                )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=".output-",
            suffix=".xlsx",
            dir=output_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        workbook.save(temporary_path)
        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def generate_report(
    input_directory: Path = INPUT_DIRECTORY,
    output_file: Path = OUTPUT_FILE,
    warning_handler: WarningHandler = warn,
) -> ProcessingSummary:
    input_files = sorted(
        (
            path
            for path in input_directory.glob("*.xlsx")
            if path.is_file() and not path.name.startswith("~$")
        ),
        key=lambda path: path.name.casefold(),
    )

    if not input_files:
        raise NoValidDataError(f"в папке {input_directory} нет файлов .xlsx")

    all_records: list[Record] = []
    skipped_rows = 0
    processed_files = 0
    skipped_files = 0

    for workbook_path in input_files:
        try:
            records, file_skipped_rows = read_records(
                workbook_path, warning_handler=warning_handler
            )
        except (
            OSError,
            ValueError,
            BadZipFile,
            InvalidFileException,
            WorkbookSchemaError,
        ) as error:
            skipped_files += 1
            warning_handler(f"{workbook_path.name}: файл пропущен: {error}")
            continue

        processed_files += 1
        skipped_rows += file_skipped_rows
        all_records.extend(records)

    if not all_records:
        raise NoValidDataError("во входных файлах нет валидных строк")

    visits = build_visits(all_records)
    sections = build_report_sections(visits)
    write_report(
        sections.main_rows, output_file,
        low_visits=sections.low_visits, low_days=sections.low_days,
    )

    return ProcessingSummary(
        input_files=len(input_files),
        processed_files=processed_files,
        skipped_files=skipped_files,
        valid_rows=len(all_records),
        skipped_rows=skipped_rows,
        visits=len(visits),
        checks=sum(visit.check_count for visit in visits),
        main_visits=sum(row.vehicle_check_count for row in sections.main_rows),
        main_checks=sum(row.check_count for row in sections.main_rows),
        low_visits=len(sections.low_visits),
        low_checks=sum(visit.check_count for visit in sections.low_visits),
        low_days=len(sections.low_days),
        low_day_checks=sum(day.check_count for day in sections.low_days),
    )


def main() -> int:
    try:
        summary = generate_report()
    except (ReportError, OSError) as error:
        print(f"Ошибка: {error}", file=sys.stderr)
        return 1

    print(f"Входных файлов: {summary.input_files}")
    print(f"Обработано файлов: {summary.processed_files}")
    print(f"Пропущено файлов: {summary.skipped_files}")
    print(f"Валидных строк: {summary.valid_rows}")
    print(f"Пропущено строк: {summary.skipped_rows}")
    print(f"Проверок ТС до исключений: {summary.visits}")
    print(f"Проверок пассажиров после дневного объединения: {summary.checks}")
    print(f"Проверки: {summary.main_visits} ТС, {summary.main_checks} пассажирских проверок")
    print(f">=2: {summary.low_visits} ТС с 0, 1 или 2 пассажирскими проверками")
    print(f">30: {summary.low_days} дней инспекторов с менее чем 30 проверками")
    print(f"Отчёт сохранён: {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
