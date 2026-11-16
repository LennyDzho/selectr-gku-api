from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell

from main import (
    CREATED_AT,
    InspectorDay,
    NoValidDataError,
    OUTPUT_HEADERS,
    TICKET_ID,
    TICKET_TYPE,
    VISIT_START,
    ZERO_TICKET_ID,
    Record,
    WorkbookSchemaError,
    build_daily_report_rows,
    build_report_sections,
    build_visits,
    count_passenger_checks,
    generate_report,
    read_records,
    write_report,
)


INPUT_HEADERS = (
    "Дата создания",
    "Логин",
    "Номер маршрута",
    "Бортовой номер",
    "Уникальный номер билета",
    "Название типа билета",
    VISIT_START,
)
LOW_VISIT_HEADERS = OUTPUT_HEADERS + (VISIT_START,)
LOW_DAY_HEADERS = OUTPUT_HEADERS + (
    CREATED_AT, VISIT_START, TICKET_ID, TICKET_TYPE
)


def record(
    timestamp: str,
    ticket_id: str,
    board: str = "00193554",
    ticket_type: str = "",
    *,
    visit_start: str | None = None,
    login: str = "biryukovlb",
    route: str = "1542",
    source_name: str = "",
    source_row: int = 0,
) -> Record:
    parsed_timestamp = datetime.fromisoformat(timestamp)
    return Record(
        timestamp=parsed_timestamp,
        login=login,
        route=route,
        board=board,
        ticket_id=ticket_id,
        visit_start=(
            datetime.fromisoformat(visit_start)
            if visit_start is not None
            else parsed_timestamp.replace(hour=0, minute=0, second=0, microsecond=0)
        ),
        ticket_type=ticket_type,
        source_name=source_name,
        source_row=source_row,
    )


def visit_records(
    count: int,
    *,
    start: str = "2026-08-26 08:00:00",
    login: str = "inspector",
    board: str = "00193554",
    route: str = "1542",
    prefix: str = "TICKET",
) -> list[Record]:
    return [
        record(
            start,
            f"{prefix}-{index}",
            board=board,
            visit_start=start,
            login=login,
            route=route,
        )
        for index in range(count)
    ]


def input_rows(records: list[Record]) -> list[tuple[object, ...]]:
    return [
        (
            item.timestamp,
            item.login,
            item.route,
            item.board,
            item.ticket_id,
            item.ticket_type,
            item.visit_start,
        )
        for item in records
    ]


class CheckCountingTests(unittest.TestCase):
    def test_social_card_absorbs_only_zero_rows_from_the_same_second(self) -> None:
        records = [
            record("2026-08-26 12:59:33", ZERO_TICKET_ID),
            record("2026-08-26 12:59:33", ZERO_TICKET_ID),
            record(
                "2026-08-26 12:59:33",
                "34139E11412806",
                ticket_type="Социальная карта москвича",
            ),
            record(
                "2026-08-26 12:59:33",
                "34139E11412806",
                ticket_type="Социальная карта москвича",
            ),
            record(
                "2026-08-26 12:59:33",
                "34139E11412803",
                ticket_type="Тройка",
            ),
        ]

        self.assertEqual(count_passenger_checks(records), 2)

    def test_social_card_does_not_absorb_another_nonzero_ticket(self) -> None:
        records = [
            record(
                "2026-08-26 12:59:33",
                "34139E11412806",
                ticket_type="Социальная карта москвича",
            ),
            record(
                "2026-08-26 12:59:33",
                "34139E11412803",
                ticket_type="Тройка",
            ),
        ]

        self.assertEqual(count_passenger_checks(records), 2)

    def test_nonzero_ticket_is_counted_once_despite_large_time_gaps(self) -> None:
        records = [
            record("2026-08-26 08:00:00.000", "TICKET"),
            record("2026-08-26 08:00:01.000", "TICKET"),
            record("2026-08-26 08:00:02.001", "TICKET"),
            record("2026-08-26 09:00:00", "TICKET"),
        ]

        self.assertEqual(count_passenger_checks(records), 1)

    def test_each_zero_ticket_row_is_counted(self) -> None:
        records = [
            record("2026-08-26 08:00:00", ZERO_TICKET_ID),
            record("2026-08-26 08:00:00", ZERO_TICKET_ID),
            record("2026-08-26 08:00:01", ZERO_TICKET_ID),
        ]

        self.assertEqual(count_passenger_checks(records), 3)

    def test_social_card_uses_calendar_seconds_not_elapsed_seconds(self) -> None:
        for zero_time, social_time, expected in (
            ("12:59:33.100", "12:59:33.900", 1),
            ("12:59:33.900", "12:59:34.000", 2),
        ):
            with self.subTest(zero_time=zero_time, social_time=social_time):
                records = [
                    record(f"2026-08-26 {zero_time}", ZERO_TICKET_ID),
                    record(
                        f"2026-08-26 {social_time}",
                        "SOCIAL",
                        ticket_type="Социальная карта москвича",
                    ),
                ]
                self.assertEqual(count_passenger_checks(records), expected)

    def test_social_card_type_requires_exact_normalised_name(self) -> None:
        for ticket_type, expected in (
            (" СОЦИАЛЬНАЯ   КАРТА МОСКВИЧА ", 1),
            ("Социальная карта москвича ММ", 2),
            ("", 2),
        ):
            with self.subTest(ticket_type=ticket_type):
                records = [
                    record("2026-08-26 12:59:33", ZERO_TICKET_ID),
                    record(
                        "2026-08-26 12:59:33", "SOCIAL", ticket_type=ticket_type
                    ),
                ]
                self.assertEqual(count_passenger_checks(records), expected)

    def test_zero_ticket_cannot_trigger_the_social_card_exception(self) -> None:
        records = [
            record("2026-08-26 12:59:33", ZERO_TICKET_ID),
            record(
                "2026-08-26 12:59:33",
                ZERO_TICKET_ID,
                ticket_type="Социальная карта москвича",
            ),
        ]
        self.assertEqual(count_passenger_checks(records), 2)


class VisitGroupingTests(unittest.TestCase):
    def test_same_ticket_on_different_vehicles_is_counted_independently(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 08:00:00", "A", board="00193554"),
                record("2026-08-26 08:00:00", "A", board="00193555"),
            ]
        )

        self.assertEqual(len(visits), 2)
        self.assertEqual(sum(visit.check_count for visit in visits), 2)
        self.assertEqual([visit.check_count for visit in visits], [1, 1])

    def test_same_source_start_keeps_one_visit_despite_large_pauses(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 08:00:00", "A", visit_start="2026-08-26 07:55:00"
                ),
                record(
                    "2026-08-26 08:09:59", "A", visit_start="2026-08-26 07:55:00"
                ),
                record(
                    "2026-08-26 09:00:00", "A", visit_start="2026-08-26 07:55:00"
                ),
            ]
        )

        self.assertEqual(len(visits), 1)
        self.assertEqual(visits[0].start, datetime(2026, 8, 26, 7, 55))
        self.assertEqual(visits[0].check_count, 1)

    def test_one_source_visit_can_cross_midnight(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 23:59:59", "A", visit_start="2026-08-26 23:58:00"
                ),
                record(
                    "2026-08-27 00:00:00", "A", visit_start="2026-08-26 23:58:00"
                ),
            ]
        )

        self.assertEqual(len(visits), 1)
        self.assertEqual(visits[0].check_count, 1)
        report_rows = build_daily_report_rows(visits)
        self.assertEqual(len(report_rows), 1)
        self.assertEqual(report_rows[0].day, datetime(2026, 8, 26).date())
        self.assertEqual(report_rows[0].vehicle_check_count, 1)

    def test_different_starts_create_visits_even_one_second_apart(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 08:00:00", "A", visit_start="2026-08-26 07:59:00"
                ),
                record(
                    "2026-08-26 08:00:01", "A", visit_start="2026-08-26 08:00:00"
                ),
            ]
        )
        self.assertEqual(len(visits), 2)
        self.assertEqual([visit.check_count for visit in visits], [1, 0])
        self.assertEqual([visit.visit_total for visit in visits], [2, 2])

    def test_source_start_milliseconds_distinguish_visits(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 08:00:01",
                    "A",
                    visit_start="2026-08-26 08:00:00.100",
                ),
                record(
                    "2026-08-26 08:00:01",
                    "A",
                    visit_start="2026-08-26 08:00:00.101",
                ),
            ]
        )
        self.assertEqual(len(visits), 2)
        self.assertEqual([visit.start.microsecond for visit in visits], [100000, 101000])
        self.assertEqual(sum(visit.check_count for visit in visits), 1)

    def test_source_start_is_also_scoped_to_login_route_and_board(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 08:00:00", "A"),
                record("2026-08-26 08:00:00", "A", login="other-inspector"),
                record("2026-08-26 08:00:00", "A", route="other-route"),
                record("2026-08-26 08:00:00", "A", board="other-board"),
            ]
        )
        self.assertEqual(len(visits), 4)
        self.assertEqual(sum(visit.check_count for visit in visits), 4)

    def test_social_card_cannot_absorb_zero_rows_from_another_visit(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 08:01:00",
                    ZERO_TICKET_ID,
                    visit_start="2026-08-26 07:55:00",
                ),
                record(
                    "2026-08-26 08:01:00",
                    "SOCIAL",
                    ticket_type="Социальная карта москвича",
                    visit_start="2026-08-26 08:00:00",
                ),
            ]
        )
        self.assertEqual(len(visits), 2)
        self.assertEqual([visit.check_count for visit in visits], [1, 1])

    def test_daily_report_combines_all_visits_of_the_same_vehicle(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 08:00:00", "A"),
                record("2026-08-26 08:01:00", "B"),
                record(
                    "2026-08-26 08:11:00", "C", visit_start="2026-08-26 08:10:00"
                ),
            ]
        )

        report_rows = build_daily_report_rows(visits)

        self.assertEqual(len(report_rows), 1)
        self.assertEqual(report_rows[0].check_count, 3)
        self.assertEqual(report_rows[0].vehicle_check_count, 2)

    def test_daily_rows_sort_by_source_day_login_route_and_board(self) -> None:
        visits = build_visits(
            [
                record("2026-08-27 08:00:00", "A", login="a", route="1", board="01"),
                record("2026-08-26 08:00:00", "A", login="b", route="1", board="01"),
                record("2026-08-26 08:00:00", "A", login="a", route="2", board="01"),
                record("2026-08-26 08:00:00", "A", login="a", route="1", board="02"),
                record("2026-08-26 08:00:00", "A", login="a", route="1", board="01"),
            ]
        )
        report_rows = build_daily_report_rows(visits)
        self.assertEqual(
            [(row.day.isoformat(), row.login, row.route, row.board) for row in report_rows],
            [
                ("2026-08-26", "a", "1", "01"),
                ("2026-08-26", "a", "1", "02"),
                ("2026-08-26", "a", "2", "01"),
                ("2026-08-26", "b", "1", "01"),
                ("2026-08-27", "a", "1", "01"),
            ],
        )


class DailyTicketUniquenessTests(unittest.TestCase):
    def test_daily_ticket_identity_includes_both_route_and_board(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-26 08:00:00", "A", route="1", board="001",
                    visit_start="2026-08-26 08:00:00",
                ),
                record(
                    "2026-08-26 09:00:00", "A", route="1", board="002",
                    visit_start="2026-08-26 09:00:00",
                ),
                record(
                    "2026-08-26 10:00:00", "A", route="2", board="001",
                    visit_start="2026-08-26 10:00:00",
                ),
                record(
                    "2026-08-26 11:00:00", "A", route="1", board="001",
                    visit_start="2026-08-26 11:00:00",
                ),
            ]
        )
        self.assertEqual(len(visits), 4)
        self.assertEqual([visit.check_count for visit in visits], [1, 1, 1, 0])
        self.assertEqual(sum(visit.check_count for visit in visits), 3)

    def test_first_scan_owns_the_ticket_even_if_another_visit_started_earlier(self) -> None:
        first_scan = record(
            "2026-08-26 09:00:00", "A",
            visit_start="2026-08-26 08:00:00",
        )
        later_scan = record(
            "2026-08-26 10:00:00", "A",
            visit_start="2026-08-26 07:00:00",
        )
        visits = build_visits([later_scan, first_scan])
        self.assertEqual([visit.check_count for visit in visits], [0, 1])
        self.assertEqual(visits[0].counted_records, ())
        self.assertEqual(visits[1].counted_records, (first_scan,))

    def test_duplicate_ties_are_resolved_by_source_file_then_source_row(self) -> None:
        records = [
            record(
                "2026-08-26 12:00:00", "A",
                visit_start="2026-08-26 09:00:00", source_name="b.xlsx", source_row=2,
            ),
            record(
                "2026-08-26 12:00:00", "A",
                visit_start="2026-08-26 10:00:00", source_name="a.xlsx", source_row=9,
            ),
            record(
                "2026-08-26 12:00:00", "A",
                visit_start="2026-08-26 11:00:00", source_name="a.xlsx", source_row=3,
            ),
        ]
        for ordered in (records, list(reversed(records))):
            with self.subTest(reversed=ordered is not records):
                visits = build_visits(ordered)
                self.assertEqual([visit.check_count for visit in visits], [0, 0, 1])
                self.assertEqual(visits[-1].counted_records, (records[-1],))

    def test_daily_identity_uses_visit_day_and_does_not_cross_logins(self) -> None:
        visits = build_visits(
            [
                record(
                    "2026-08-27 00:01:00", "A", login="first",
                    visit_start="2026-08-26 23:55:00",
                ),
                record(
                    "2026-08-27 00:01:00", "A", login="first",
                    visit_start="2026-08-27 00:00:00",
                ),
                record(
                    "2026-08-27 00:01:00", "A", login="second",
                    visit_start="2026-08-26 23:55:00",
                ),
            ]
        )
        self.assertEqual(len(visits), 3)
        self.assertEqual(sum(visit.check_count for visit in visits), 3)

    def test_raw_duplicate_social_card_still_absorbs_zeros_only_in_its_own_visit(self) -> None:
        first = record(
            "2026-08-26 08:00:00", "SOCIAL",
            visit_start="2026-08-26 07:55:00", ticket_type="Undefined type",
        )
        repeated_social = record(
            "2026-08-26 09:00:00", "SOCIAL",
            visit_start="2026-08-26 08:55:00",
            ticket_type="Социальная карта москвича",
        )
        suppressed = record(
            "2026-08-26 09:00:00", ZERO_TICKET_ID,
            visit_start="2026-08-26 08:55:00",
        )
        separate_visit_zero = record(
            "2026-08-26 09:00:00", ZERO_TICKET_ID,
            visit_start="2026-08-26 08:56:00",
        )
        visits = build_visits(
            [suppressed, repeated_social, first, suppressed, separate_visit_zero]
        )
        self.assertEqual([visit.check_count for visit in visits], [1, 0, 1])
        self.assertEqual(visits[0].counted_records, (first,))
        self.assertEqual(visits[1].counted_records, ())
        self.assertEqual(visits[2].counted_records, (separate_visit_zero,))

    def test_equal_zero_rows_remain_separate_counted_detail_records(self) -> None:
        zero = record("2026-08-26 08:00:00", ZERO_TICKET_ID)
        visits = build_visits([zero, zero])
        self.assertEqual(visits[0].check_count, 2)
        self.assertEqual(visits[0].counted_records, (zero, zero))

    def test_social_card_on_other_route_or_vehicle_does_not_absorb_zero_rows(self) -> None:
        for other_route, other_board in (("1", "002"), ("2", "001")):
            with self.subTest(route=other_route, board=other_board):
                visits = build_visits(
                    [
                        record(
                            "2026-08-26 08:00:00", "SOCIAL", route="1", board="001",
                            ticket_type="Социальная карта москвича",
                        ),
                        record(
                            "2026-08-26 08:00:00", "SOCIAL",
                            route=other_route, board=other_board,
                            ticket_type="Undefined type",
                        ),
                        record(
                            "2026-08-26 08:00:00", ZERO_TICKET_ID,
                            route=other_route, board=other_board,
                        ),
                    ]
                )
                by_vehicle = {(visit.route, visit.board): visit for visit in visits}
                self.assertEqual(by_vehicle[("1", "001")].check_count, 1)
                self.assertEqual(by_vehicle[(other_route, other_board)].check_count, 2)


class ReportSectionTests(unittest.TestCase):
    def test_day_threshold_is_below_thirty_and_uses_counted_checks(self) -> None:
        for count in (29, 30):
            with self.subTest(count=count):
                records = visit_records(count)
                # Repeated product rows must not make a 29-check day pass 30.
                visits = build_visits(records + records)
                sections = build_report_sections(visits)
                self.assertEqual(sections.low_visits, ())
                if count == 29:
                    self.assertEqual(sections.main_rows, ())
                    self.assertEqual(len(sections.low_days), 1)
                    self.assertEqual(sections.low_days[0].check_count, 29)
                    self.assertEqual(len(sections.low_days[0].counted_records), 29)
                else:
                    self.assertEqual(sections.low_days, ())
                    self.assertEqual(len(sections.main_rows), 1)
                    self.assertEqual(sections.main_rows[0].check_count, 30)

    def test_visit_threshold_excludes_zero_one_and_two_but_includes_three(self) -> None:
        records = [
            *visit_records(1, prefix="A"),
            *visit_records(2, start="2026-08-26 09:00:00", prefix="B"),
            *visit_records(3, start="2026-08-26 10:00:00", prefix="C"),
            *visit_records(24, start="2026-08-26 11:00:00", prefix="D"),
            record(
                "2026-08-26 12:00:00", "A-0", login="inspector",
                visit_start="2026-08-26 12:00:00",
            ),
        ]
        sections = build_report_sections(build_visits(records))
        self.assertEqual([visit.check_count for visit in sections.low_visits], [1, 2, 0])
        self.assertEqual(sections.low_days, ())
        self.assertEqual(len(sections.main_rows), 1)
        self.assertEqual(sections.main_rows[0].check_count, 27)
        self.assertEqual(sections.main_rows[0].vehicle_check_count, 2)

    def test_day_total_is_computed_before_low_visits_are_removed(self) -> None:
        visits = build_visits(
            [
                *visit_records(28, prefix="large"),
                *visit_records(2, start="2026-08-26 09:00:00", prefix="small"),
            ]
        )
        sections = build_report_sections(visits)
        self.assertEqual(sections.low_days, ())
        self.assertEqual(len(sections.low_visits), 1)
        self.assertEqual(sections.low_visits[0].check_count, 2)
        self.assertEqual(len(sections.main_rows), 1)
        self.assertEqual(sections.main_rows[0].check_count, 28)
        self.assertEqual(sections.main_rows[0].vehicle_check_count, 1)

    def test_day_threshold_combines_boards_and_routes_for_one_login(self) -> None:
        sections = build_report_sections(
            build_visits(
                [
                    *visit_records(15, board="001", route="1", prefix="first"),
                    *visit_records(15, board="002", route="2", prefix="second"),
                ]
            )
        )
        self.assertEqual(sections.low_days, ())
        self.assertEqual(len(sections.main_rows), 2)
        self.assertEqual(sum(row.check_count for row in sections.main_rows), 30)

    def test_day_threshold_does_not_combine_other_logins_or_other_days(self) -> None:
        sections = build_report_sections(
            build_visits(
                [
                    *visit_records(15, login="first", prefix="A"),
                    *visit_records(15, login="second", prefix="B"),
                    *visit_records(15, login="first", start="2026-08-27 08:00:00", prefix="C"),
                ]
            )
        )
        self.assertEqual(sections.main_rows, ())
        self.assertEqual(sections.low_visits, ())
        self.assertEqual(len(sections.low_days), 3)
        self.assertEqual([day.check_count for day in sections.low_days], [15, 15, 15])

    def test_zero_rows_can_make_a_thirty_check_day(self) -> None:
        zero = record("2026-08-26 08:00:00", ZERO_TICKET_ID)
        sections = build_report_sections(build_visits([zero] * 30))
        self.assertEqual(sections.low_days, ())
        self.assertEqual(sections.low_visits, ())
        self.assertEqual(sections.main_rows[0].check_count, 30)

    def test_low_day_and_low_visit_exceptions_can_overlap_without_entering_main(self) -> None:
        sections = build_report_sections(
            build_visits(
                [
                    *visit_records(1, prefix="small"),
                    *visit_records(2, start="2026-08-26 09:00:00", prefix="large"),
                ]
            )
        )
        self.assertEqual(sections.main_rows, ())
        self.assertEqual(len(sections.low_visits), 2)
        self.assertEqual(len(sections.low_days), 1)
        self.assertEqual([visit.check_count for visit in sections.low_visits], [1, 2])
        self.assertEqual(sections.low_days[0].check_count, 3)
        self.assertEqual(sections.low_days[0].vehicle_check_count, 2)
        self.assertEqual(len(sections.low_days[0].counted_records), 3)

    def test_empty_sections_are_valid(self) -> None:
        sections = build_report_sections([])
        self.assertEqual(sections.main_rows, ())
        self.assertEqual(sections.low_visits, ())
        self.assertEqual(sections.low_days, ())


class ReportIntegrationTests(unittest.TestCase):
    def test_multiple_files_produce_one_sorted_report_and_skip_bad_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()

            self._write_input(
                input_directory / "later.xlsx",
                [
                    (
                        "2026-08-27 10:00:00",
                        "user2",
                        "2",
                        "00000002",
                        "B",
                        "Undefined type",
                        "2026-08-27 09:55:00",
                    ),
                    (
                        "not-a-date",
                        "user2",
                        "2",
                        "00000002",
                        "C",
                        "Undefined type",
                        "2026-08-27 09:55:00",
                    ),
                ],
            )
            self._write_input(
                input_directory / "earlier.xlsx",
                [
                    (
                        "2026-08-26 09:00:00",
                        "user1",
                        "1",
                        "00000001",
                        "A",
                        "Undefined type",
                        "2026-08-26 08:55:00",
                    )
                ],
            )

            warnings: list[str] = []
            summary = generate_report(input_directory, output_file, warnings.append)

            self.assertEqual(summary.input_files, 2)
            self.assertEqual(summary.valid_rows, 2)
            self.assertEqual(summary.skipped_rows, 1)
            self.assertEqual(summary.visits, 2)
            self.assertEqual(summary.checks, 2)
            self.assertEqual(summary.main_visits, 0)
            self.assertEqual(summary.main_checks, 0)
            self.assertEqual(summary.low_visits, 2)
            self.assertEqual(summary.low_checks, 2)
            self.assertEqual(summary.low_days, 2)
            self.assertEqual(summary.low_day_checks, 2)
            self.assertEqual(len(warnings), 1)
            self.assertIn("Дата создания", warnings[0])
            self.assertTrue(output_file.exists())
            self.assertEqual(list(output_file.parent.glob("*.xlsx")), [output_file])

            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["Проверки", ">=2", ">30"])
                worksheet = workbook["Проверки"]
                self.assertEqual(
                    tuple(cell.value for cell in worksheet[1]), OUTPUT_HEADERS
                )
                self.assertEqual(worksheet.max_row, 1)
                low_visits = workbook[">=2"]
                self.assertEqual(tuple(c.value for c in low_visits[1]), LOW_VISIT_HEADERS)
                self.assertEqual(low_visits["B2"].value, "user1")
                self.assertEqual(low_visits["B3"].value, "user2")
                self.assertEqual(low_visits["C2"].value, "1")
                self.assertEqual(low_visits["D2"].value, "00000001")
                self.assertEqual(low_visits["E2"].value, 1)
                self.assertEqual(low_visits["F2"].value, 1)
                self.assertEqual(low_visits["G2"].value, datetime(2026, 8, 26, 8, 55))
                low_days = workbook[">30"]
                self.assertEqual(tuple(c.value for c in low_days[1]), LOW_DAY_HEADERS)
                self.assertEqual(low_days.max_row, 3)
                self.assertEqual(low_days["B2"].value, "user1")
                self.assertEqual(low_days["B3"].value, "user2")
                self.assertEqual(low_days["I2"].value, "A")
                self.assertEqual(low_days["I3"].value, "B")
                self.assertEqual(list(low_days.merged_cells.ranges), [])
            finally:
                workbook.close()

    def test_multiple_files_merge_source_visits_and_keep_one_daily_row(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            self._write_input(
                input_directory / "first.xlsx",
                [
                    (
                        "2026-08-26 08:00:00",
                        "user",
                        "1542",
                        "00193554",
                        "A",
                        "Тройка",
                        "2026-08-26 07:55:00",
                    ),
                ],
            )
            self._write_input(
                input_directory / "second.xlsx",
                [
                    (
                        "2026-08-26 08:20:00",
                        "user",
                        "1542",
                        "00193554",
                        "A",
                        "Кошелек",
                        "2026-08-26 07:55:00",
                    ),
                    (
                        "2026-08-26 08:22:00",
                        "user",
                        "1542",
                        "00193554",
                        "A",
                        "Тройка",
                        "2026-08-26 08:21:00",
                    ),
                ],
            )
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.input_files, 2)
            self.assertEqual(summary.valid_rows, 3)
            self.assertEqual(summary.visits, 2)
            self.assertEqual(summary.checks, 1)
            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook["Проверки"]
                self.assertEqual(worksheet.max_row, 1)
                self.assertEqual(workbook[">=2"].max_row, 3)
                self.assertEqual(workbook[">=2"]["E2"].value, 1)
                self.assertEqual(workbook[">=2"]["E3"].value, 0)
                low_days = workbook[">30"]
                self.assertEqual(low_days.max_row, 2)
                self.assertEqual(low_days["E2"].value, 1)
                self.assertEqual(low_days["F2"].value, 2)
                self.assertEqual(low_days["I2"].value, "A")
                self.assertEqual(low_days["G2"].value, datetime(2026, 8, 26, 8))
                self.assertEqual(list(low_days.merged_cells.ranges), [])
            finally:
                workbook.close()

    def test_missing_or_duplicate_source_start_header_is_rejected(self) -> None:
        for headers in (INPUT_HEADERS[:-1], INPUT_HEADERS + (VISIT_START,)):
            with self.subTest(headers=headers):
                with tempfile.TemporaryDirectory() as temporary_directory:
                    path = Path(temporary_directory) / "bad-header.xlsx"
                    self._write_input(path, [], headers=headers)
                    with self.assertRaisesRegex(WorkbookSchemaError, VISIT_START):
                        read_records(path)

    def test_invalid_source_start_headers_are_skipped_without_losing_valid_files(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            valid_row = (
                "2026-08-26 08:00:00",
                "user",
                "1542",
                "00193554",
                "A",
                "Тройка",
                "2026-08-26 07:55:00",
            )
            self._write_input(input_directory / "valid.xlsx", [valid_row])
            self._write_input(input_directory / "~$temporary.xlsx", [valid_row])
            self._write_input(
                input_directory / "missing-start.xlsx", [], headers=INPUT_HEADERS[:-1]
            )
            self._write_input(
                input_directory / "duplicate-start.xlsx",
                [],
                headers=INPUT_HEADERS + (VISIT_START,),
            )
            warnings: list[str] = []
            summary = generate_report(input_directory, output_file, warnings.append)
            self.assertEqual(summary.input_files, 3)
            self.assertEqual(summary.processed_files, 1)
            self.assertEqual(summary.skipped_files, 2)
            self.assertEqual(summary.checks, 1)
            self.assertEqual(len(warnings), 2)
            self.assertTrue(all(VISIT_START in warning for warning in warnings))
            self.assertTrue(output_file.exists())

    def test_invalid_or_blank_source_start_is_skipped_with_field_and_row(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "invalid-start.xlsx"
            self._write_input(
                path,
                [
                    (
                        "2026-08-26 08:00:00",
                        "user",
                        "1542",
                        "00193554",
                        "A",
                        "Тройка",
                        visit_start,
                    )
                    for visit_start in ("not-a-date", None, " ")
                ],
            )
            warnings: list[str] = []
            records, skipped = read_records(path, warnings.append)
            self.assertEqual(records, [])
            self.assertEqual(skipped, 3)
            self.assertEqual(len(warnings), 3)
            for row_number, warning in enumerate(warnings, start=2):
                self.assertIn(path.name, warning)
                self.assertIn(f"строка {row_number}", warning)
                self.assertIn(VISIT_START, warning)

    def test_both_timestamps_can_be_excel_dates_and_keep_milliseconds(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "excel-dates.xlsx"
            timestamp = datetime(2026, 8, 26, 8, 0, 0, 456000)
            visit_start = datetime(2026, 8, 26, 7, 55, 0, 123000)
            self._write_input(
                path,
                [(timestamp, "user", "1542", "00193554", "A", "", visit_start)],
            )
            records, skipped = read_records(path)
            self.assertEqual(skipped, 0)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0].timestamp, timestamp)
            self.assertEqual(records[0].visit_start, visit_start)

    def test_no_valid_rows_do_not_replace_an_existing_report(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            output_file.parent.mkdir()
            output_file.write_bytes(b"existing report")
            self._write_input(
                input_directory / "invalid.xlsx",
                [
                    (
                        "not-a-date",
                        "user",
                        "1",
                        "00000001",
                        "A",
                        "Undefined type",
                        "2026-08-26 07:55:00",
                    )
                ],
            )

            with self.assertRaises(NoValidDataError):
                generate_report(input_directory, output_file, lambda _message: None)

            self.assertEqual(output_file.read_bytes(), b"existing report")

    def test_no_valid_source_starts_do_not_replace_an_existing_report(self) -> None:
        for invalid_start in (None, "not-a-date"):
            with self.subTest(invalid_start=invalid_start):
                with tempfile.TemporaryDirectory() as temporary_directory:
                    root = Path(temporary_directory)
                    input_directory = root / "input"
                    output_file = root / "output" / "output.xlsx"
                    input_directory.mkdir()
                    output_file.parent.mkdir()
                    output_file.write_bytes(b"existing report")
                    self._write_input(
                        input_directory / "invalid.xlsx",
                        [
                            (
                                "2026-08-26 08:00:00",
                                "user",
                                "1542",
                                "00193554",
                                "A",
                                "Тройка",
                                invalid_start,
                            )
                        ],
                    )
                    with self.assertRaises(NoValidDataError):
                        generate_report(input_directory, output_file, lambda _: None)
                    self.assertEqual(output_file.read_bytes(), b"existing report")

    def test_main_filters_visits_before_daily_aggregation_and_reports_category_totals(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            records = [
                *visit_records(29, prefix="large"),
                *visit_records(1, start="2026-08-26 09:00:00", prefix="small"),
                *visit_records(29, login="low-day-inspector", prefix="other"),
            ]
            self._write_input(input_directory / "checks.xlsx", input_rows(records))
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.visits, 3)
            self.assertEqual(summary.checks, 59)
            self.assertEqual(summary.main_visits, 1)
            self.assertEqual(summary.main_checks, 29)
            self.assertEqual(summary.low_visits, 1)
            self.assertEqual(summary.low_checks, 1)
            self.assertEqual(summary.low_days, 1)
            self.assertEqual(summary.low_day_checks, 29)
            workbook = load_workbook(output_file, data_only=True)
            try:
                main = workbook["Проверки"]
                self.assertEqual(main.max_row, 2)
                self.assertEqual(main["B2"].value, "inspector")
                self.assertEqual(main["E2"].value, 29)
                self.assertEqual(main["F2"].value, 1)
                low_visits = workbook[">=2"]
                self.assertEqual(low_visits.max_row, 2)
                self.assertEqual(low_visits["B2"].value, "inspector")
                self.assertEqual(low_visits["E2"].value, 1)
                self.assertEqual(low_visits["F2"].value, 1)
                self.assertEqual(low_visits["G2"].value, datetime(2026, 8, 26, 9))
                low_days = workbook[">30"]
                self.assertEqual(low_days.max_row, 30)
                self.assertEqual(low_days["B2"].value, "low-day-inspector")
                self.assertEqual(low_days["E2"].value, 29)
            finally:
                workbook.close()

    def test_main_combines_kept_visits_into_one_day_row_and_keeps_empty_exception_sheets(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            records = [
                *visit_records(15, prefix="first"),
                *visit_records(15, start="2026-08-26 09:00:00", prefix="second"),
                record(
                    "2026-08-26 09:00:01", "first-0", login="inspector",
                    visit_start="2026-08-26 09:00:00",
                ),
            ]
            self._write_input(input_directory / "checks.xlsx", input_rows(records))
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.valid_rows, 31)
            self.assertEqual(summary.checks, 30)
            self.assertEqual(summary.main_checks, 30)
            self.assertEqual(summary.main_visits, 2)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["Проверки", ">=2", ">30"])
                self.assertEqual(workbook["Проверки"].max_row, 2)
                self.assertEqual(workbook["Проверки"]["E2"].value, 30)
                self.assertEqual(workbook["Проверки"]["F2"].value, 2)
                self.assertEqual(workbook[">=2"].max_row, 1)
                self.assertEqual(workbook[">30"].max_row, 1)
            finally:
                workbook.close()

    def test_exactly_two_checks_go_to_visit_exceptions_and_overlap_a_low_day(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            records = [
                *visit_records(28, login="active", prefix="large"),
                *visit_records(
                    2, login="active", start="2026-08-26 09:00:00", prefix="small"
                ),
                *visit_records(
                    2, login="low", start="2026-08-26 10:00:00", prefix="low"
                ),
            ]
            self._write_input(input_directory / "two-check-visits.xlsx", input_rows(records))
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.visits, 3)
            self.assertEqual(summary.checks, 32)
            self.assertEqual(summary.main_visits, 1)
            self.assertEqual(summary.main_checks, 28)
            self.assertEqual(summary.low_visits, 2)
            self.assertEqual(summary.low_checks, 4)
            self.assertEqual(summary.low_days, 1)
            self.assertEqual(summary.low_day_checks, 2)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["Проверки", ">=2", ">30"])
                main = workbook["Проверки"]
                self.assertEqual(main.max_row, 2)
                self.assertEqual(main["B2"].value, "active")
                self.assertEqual(main["E2"].value, 28)
                self.assertEqual(main["F2"].value, 1)
                low_visits = workbook[">=2"]
                self.assertEqual(low_visits.max_row, 3)
                self.assertEqual([low_visits.cell(i, 2).value for i in (2, 3)], ["active", "low"])
                self.assertEqual([low_visits.cell(i, 5).value for i in (2, 3)], [2, 2])
                self.assertEqual([low_visits.cell(i, 6).value for i in (2, 3)], [1, 1])
                self.assertEqual(low_visits["G2"].value, datetime(2026, 8, 26, 9))
                self.assertEqual(low_visits["G3"].value, datetime(2026, 8, 26, 10))
                low_days = workbook[">30"]
                self.assertEqual(low_days.max_row, 3)
                self.assertEqual(low_days["B2"].value, "low")
                self.assertEqual(low_days["E2"].value, 2)
                self.assertEqual(low_days["F2"].value, 1)
                self.assertEqual([low_days.cell(i, 9).value for i in (2, 3)], ["low-0", "low-1"])
            finally:
                workbook.close()

    def test_thirty_check_day_with_only_two_check_visits_has_an_empty_main(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            records = [
                item
                for hour in range(8, 23)
                for item in visit_records(
                    2, start=f"2026-08-26 {hour:02d}:00:00", prefix=f"visit-{hour}"
                )
            ]
            self._write_input(input_directory / "all-small.xlsx", input_rows(records))
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.visits, 15)
            self.assertEqual(summary.checks, 30)
            self.assertEqual(summary.main_visits, 0)
            self.assertEqual(summary.main_checks, 0)
            self.assertEqual(summary.low_visits, 15)
            self.assertEqual(summary.low_checks, 30)
            self.assertEqual(summary.low_days, 0)
            self.assertEqual(summary.low_day_checks, 0)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook["Проверки"].max_row, 1)
                self.assertEqual(workbook[">=2"].max_row, 16)
                self.assertEqual(workbook[">30"].max_row, 1)
                self.assertTrue(all(workbook[">=2"].cell(i, 5).value == 2 for i in range(2, 17)))
            finally:
                workbook.close()

    def test_low_day_card_details_are_unique_and_do_not_double_the_summary_counts(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            first = record(
                "2026-08-26 08:01:00", "A", ticket_type="Тройка",
                visit_start="2026-08-26 08:00:00",
            )
            duplicate = record(
                "2026-08-26 08:01:01", "A", ticket_type="Кошелек",
                visit_start="2026-08-26 08:00:00",
            )
            zero = record(
                "2026-08-26 08:02:00", ZERO_TICKET_ID,
                visit_start="2026-08-26 08:00:00",
            )
            social = record(
                "2026-08-26 08:03:00", "SOCIAL",
                ticket_type="Социальная карта москвича",
                visit_start="2026-08-26 08:00:00",
            )
            suppressed = record(
                "2026-08-26 08:03:00", ZERO_TICKET_ID,
                visit_start="2026-08-26 08:00:00",
            )
            later_visit_duplicate = record(
                "2026-08-26 09:00:00", "A", visit_start="2026-08-26 09:00:00"
            )
            self._write_input(
                input_directory / "details.xlsx",
                input_rows([first, duplicate, zero, zero, social, social, suppressed, later_visit_duplicate]),
            )
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.checks, 4)
            self.assertEqual(summary.visits, 2)
            self.assertEqual(summary.main_checks, 0)
            self.assertEqual(summary.low_visits, 1)
            self.assertEqual(summary.low_checks, 0)
            self.assertEqual(summary.low_days, 1)
            self.assertEqual(summary.low_day_checks, 4)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook["Проверки"].max_row, 1)
                self.assertEqual(workbook[">=2"]["E2"].value, 0)
                worksheet = workbook[">30"]
                self.assertEqual(tuple(c.value for c in worksheet[1]), LOW_DAY_HEADERS)
                self.assertEqual(worksheet.max_column, 10)
                self.assertEqual(worksheet.max_row, 5)
                self.assertEqual(worksheet["E2"].value, 4)
                self.assertEqual(worksheet["F2"].value, 2)
                self.assertEqual(
                    {str(merged) for merged in worksheet.merged_cells.ranges},
                    {f"{column}2:{column}5" for column in "ABEF"},
                )
                self.assertEqual([worksheet.cell(i, 9).value for i in range(2, 6)], ["A", ZERO_TICKET_ID, ZERO_TICKET_ID, "SOCIAL"])
                self.assertEqual(worksheet["J2"].value, "Тройка")
                self.assertEqual(worksheet["G2"].value, first.timestamp)
                self.assertEqual(worksheet["H2"].value, first.visit_start)
                for row_number in range(2, 6):
                    self.assertEqual(worksheet.cell(row_number, 3).value, "1542")
                    self.assertEqual(worksheet.cell(row_number, 4).value, "00193554")
                    self.assertIsInstance(worksheet.cell(row_number, 7).value, datetime)
                    self.assertIsInstance(worksheet.cell(row_number, 8).value, datetime)
                self.assertEqual(sum(worksheet.cell(i, 5).value or 0 for i in range(2, 6)), 4)
                self.assertEqual(sum(worksheet.cell(i, 6).value or 0 for i in range(2, 6)), 2)
                for row in worksheet.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell, MergedCell):
                            continue
                        self.assertIsNone(cell.fill.patternType)
                        self.assertFalse(cell.font.bold)
                        self.assertEqual(cell.alignment.horizontal, "center")
                        self.assertEqual(cell.alignment.vertical, "center")
                self.assertTrue(worksheet["A1"].font.bold)
                self.assertEqual(worksheet["A1"].fill.patternType, "solid")
                for coordinate in ("C2", "D2", "I3", "J2"):
                    self.assertEqual(worksheet[coordinate].number_format, "@")
                self.assertTrue(worksheet["A2"].is_date)
                self.assertTrue(worksheet["G2"].is_date)
                self.assertTrue(worksheet["H2"].is_date)
            finally:
                workbook.close()

    def test_direct_writer_always_creates_three_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            report_rows = build_daily_report_rows(build_visits(visit_records(1)))
            write_report(report_rows, output_file)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["Проверки", ">=2", ">30"])
                self.assertEqual(workbook["Проверки"].max_row, 2)
                self.assertEqual(tuple(c.value for c in workbook[">=2"][1]), LOW_VISIT_HEADERS)
                self.assertEqual(tuple(c.value for c in workbook[">30"][1]), LOW_DAY_HEADERS)
                self.assertEqual(workbook[">=2"].max_row, 1)
                self.assertEqual(workbook[">30"].max_row, 1)
            finally:
                workbook.close()

    def test_low_day_keeps_same_ticket_details_for_different_vehicles(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_directory = root / "input"
            output_file = root / "output" / "output.xlsx"
            input_directory.mkdir()
            first_vehicle = record(
                "2026-08-26 08:00:00", "A", board="00193554",
                visit_start="2026-08-26 07:55:00",
            )
            second_vehicle = record(
                "2026-08-26 08:30:00", "A", board="00193555",
                visit_start="2026-08-26 08:20:00",
            )
            later_first_vehicle = record(
                "2026-08-26 09:00:00", "A", board="00193554",
                visit_start="2026-08-26 08:55:00",
            )
            self._write_input(
                input_directory / "two-vehicles.xlsx",
                input_rows([later_first_vehicle, second_vehicle, first_vehicle]),
            )
            summary = generate_report(input_directory, output_file)
            self.assertEqual(summary.valid_rows, 3)
            self.assertEqual(summary.visits, 3)
            self.assertEqual(summary.checks, 2)
            self.assertEqual(summary.low_visits, 3)
            self.assertEqual(summary.low_day_checks, 2)
            workbook = load_workbook(output_file, data_only=True)
            try:
                self.assertEqual(workbook["Проверки"].max_row, 1)
                self.assertEqual(workbook[">=2"].max_row, 4)
                worksheet = workbook[">30"]
                self.assertEqual(worksheet.max_row, 3)
                self.assertEqual(worksheet["E2"].value, 2)
                self.assertEqual(worksheet["F2"].value, 3)
                self.assertEqual([worksheet.cell(i, 9).value for i in (2, 3)], ["A", "A"])
                self.assertEqual([worksheet.cell(i, 4).value for i in (2, 3)], ["00193554", "00193555"])
                self.assertEqual(worksheet["G2"].value, first_vehicle.timestamp)
                self.assertEqual(worksheet["G3"].value, second_vehicle.timestamp)
                self.assertIsNone(worksheet["E3"].value)
                self.assertEqual(
                    {str(merged) for merged in worksheet.merged_cells.ranges},
                    {f"{column}2:{column}3" for column in "ABEF"},
                )
            finally:
                workbook.close()

    def test_low_day_single_card_shares_its_row_with_totals_and_keeps_all_fields(self) -> None:
        item = record(
            "2026-08-26 08:01:00.456", "00000000000000001234",
            ticket_type="Тройка", visit_start="2026-08-26 08:00:00.123",
        )
        sections = build_report_sections(build_visits([item]))
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            write_report(
                sections.main_rows, output_file,
                low_visits=sections.low_visits, low_days=sections.low_days,
            )
            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook[">30"]
                self.assertEqual(worksheet.max_row, 2)
                self.assertEqual(worksheet.max_column, 10)
                self.assertEqual(list(worksheet.merged_cells.ranges), [])
                self.assertEqual(
                    tuple(cell.value for cell in worksheet[2]),
                    (
                        datetime(2026, 8, 26), item.login, item.route, item.board,
                        1, 1, item.timestamp, item.visit_start,
                        item.ticket_id, item.ticket_type,
                    ),
                )
                self.assertEqual(worksheet["I2"].data_type, "s")
            finally:
                workbook.close()

    def test_low_day_block_keeps_each_cards_route_and_vehicle_unmerged(self) -> None:
        records = [
            record("2026-08-26 08:00:00", "A", route="001", board="00000001"),
            record("2026-08-26 09:00:00", "B", route="002", board="00000002"),
        ]
        sections = build_report_sections(build_visits(records))
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            write_report(sections.main_rows, output_file, low_days=sections.low_days)
            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook[">30"]
                self.assertEqual(worksheet.max_row, 3)
                self.assertEqual((worksheet["C2"].value, worksheet["D2"].value), ("001", "00000001"))
                self.assertEqual((worksheet["C3"].value, worksheet["D3"].value), ("002", "00000002"))
                self.assertEqual((worksheet["I2"].value, worksheet["I3"].value), ("A", "B"))
                self.assertEqual(
                    {str(merged) for merged in worksheet.merged_cells.ranges},
                    {f"{column}2:{column}3" for column in "ABEF"},
                )
            finally:
                workbook.close()

    def test_low_day_without_card_records_keeps_one_empty_detail_row(self) -> None:
        empty_day = InspectorDay(
            day=datetime(2026, 8, 26).date(), login="empty-day",
            check_count=0, vehicle_check_count=2, counted_records=(),
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            write_report((), output_file, low_days=(empty_day,))
            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook[">30"]
                self.assertEqual(worksheet.max_row, 2)
                self.assertEqual(tuple(cell.value for cell in worksheet[1]), LOW_DAY_HEADERS)
                self.assertEqual(
                    tuple(cell.value for cell in worksheet[2]),
                    (datetime(2026, 8, 26), "empty-day", None, None, 0, 2, None, None, None, None),
                )
                self.assertEqual(list(worksheet.merged_cells.ranges), [])
            finally:
                workbook.close()

    def test_low_day_blocks_for_multiple_days_do_not_merge_into_each_other(self) -> None:
        sections = build_report_sections(
            build_visits(
                [
                    *visit_records(2, prefix="first-day"),
                    *visit_records(1, start="2026-08-27 09:00:00", prefix="second-day"),
                ]
            )
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            write_report(sections.main_rows, output_file, low_days=sections.low_days)
            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook[">30"]
                self.assertEqual(worksheet.max_row, 4)
                self.assertEqual(
                    {str(merged) for merged in worksheet.merged_cells.ranges},
                    {f"{column}2:{column}3" for column in "ABEF"},
                )
                self.assertEqual(worksheet["A2"].value, datetime(2026, 8, 26))
                self.assertEqual(worksheet["E2"].value, 2)
                self.assertEqual(worksheet["F2"].value, 1)
                self.assertEqual(worksheet["A4"].value, datetime(2026, 8, 27))
                self.assertEqual(worksheet["B4"].value, "inspector")
                self.assertEqual(worksheet["E4"].value, 1)
                self.assertEqual(worksheet["F4"].value, 1)
                self.assertEqual([worksheet.cell(i, 9).value for i in range(2, 5)], ["first-day-0", "first-day-1", "second-day-0"])
                self.assertEqual(worksheet["G4"].value, datetime(2026, 8, 27, 9))
            finally:
                workbook.close()

    @staticmethod
    def _write_input(
        path: Path,
        rows: list[tuple[object, ...]],
        *,
        headers: tuple[str, ...] = INPUT_HEADERS,
    ) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(headers)
        for row in rows:
            worksheet.append(row)
        workbook.save(path)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
