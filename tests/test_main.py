from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from main import (
    NoValidDataError,
    OUTPUT_HEADERS,
    ZERO_TICKET_ID,
    Record,
    build_visits,
    count_passenger_checks,
    generate_report,
)


def record(
    timestamp: str,
    ticket_id: str,
    board: str = "00193554",
    ticket_type: str = "",
) -> Record:
    return Record(
        timestamp=datetime.fromisoformat(timestamp),
        login="biryukovlb",
        route="1542",
        board=board,
        ticket_id=ticket_id,
        ticket_type=ticket_type,
    )


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

    def test_nonzero_ticket_is_deduplicated_only_within_one_second(self) -> None:
        records = [
            record("2026-08-26 08:00:00.000", "TICKET"),
            record("2026-08-26 08:00:01.000", "TICKET"),
            record("2026-08-26 08:00:02.001", "TICKET"),
        ]

        self.assertEqual(count_passenger_checks(records), 2)

    def test_each_zero_ticket_row_is_counted(self) -> None:
        records = [
            record("2026-08-26 08:00:00", ZERO_TICKET_ID),
            record("2026-08-26 08:00:00", ZERO_TICKET_ID),
            record("2026-08-26 08:00:01", ZERO_TICKET_ID),
        ]

        self.assertEqual(count_passenger_checks(records), 3)


class VisitGroupingTests(unittest.TestCase):
    def test_same_ticket_and_second_on_different_vehicles_are_separate(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 08:00:00", "A", board="00193554"),
                record("2026-08-26 08:00:00", "A", board="00193555"),
            ]
        )

        self.assertEqual(len(visits), 2)
        self.assertEqual(sum(visit.check_count for visit in visits), 2)

    def test_ten_minute_gap_starts_a_new_visit(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 08:00:00", "A"),
                record("2026-08-26 08:09:59", "B"),
                record("2026-08-26 08:19:59", "C"),
            ]
        )

        self.assertEqual(len(visits), 2)
        self.assertEqual([visit.visit_total for visit in visits], [2, 2])
        self.assertEqual([visit.check_count for visit in visits], [2, 1])

    def test_calendar_date_always_splits_visits(self) -> None:
        visits = build_visits(
            [
                record("2026-08-26 23:59:59", "A"),
                record("2026-08-27 00:00:00", "B"),
            ]
        )

        self.assertEqual(len(visits), 2)
        self.assertEqual([visit.visit_total for visit in visits], [1, 1])


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
                    ),
                    (
                        "not-a-date",
                        "user2",
                        "2",
                        "00000002",
                        "C",
                        "Undefined type",
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
            self.assertEqual(len(warnings), 1)
            self.assertTrue(output_file.exists())
            self.assertEqual(list(output_file.parent.glob("*.xlsx")), [output_file])

            workbook = load_workbook(output_file, data_only=True)
            try:
                worksheet = workbook["Проверки"]
                self.assertEqual(
                    tuple(cell.value for cell in worksheet[1]), OUTPUT_HEADERS
                )
                self.assertEqual(worksheet["C2"].value, "user1")
                self.assertEqual(worksheet["C3"].value, "user2")
                self.assertEqual(worksheet["D2"].value, "1")
                self.assertEqual(worksheet["E2"].value, "00000001")
            finally:
                workbook.close()

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
                    )
                ],
            )

            with self.assertRaises(NoValidDataError):
                generate_report(input_directory, output_file, lambda _message: None)

            self.assertEqual(output_file.read_bytes(), b"existing report")

    @staticmethod
    def _write_input(
        path: Path, rows: list[tuple[str, str, str, str, str, str]]
    ) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(
            (
                "Дата создания",
                "Логин",
                "Номер маршрута",
                "Бортовой номер",
                "Уникальный номер билета",
                "Название типа билета",
            )
        )
        for row in rows:
            worksheet.append(row)
        workbook.save(path)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
