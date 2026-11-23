"""Tests for the Evotec driving-date extraction logic (no API calls)."""
import unittest
from unittest.mock import patch

import app


class DrivingDateTests(unittest.TestCase):
    def test_date_formats(self):
        self.assertEqual(app.parse_report_date("09.07.2026 00:00"), "2026-07-09")
        self.assertEqual(app.parse_report_date("2026-07-09T23:59:59"), "2026-07-09")
        self.assertEqual(app.parse_report_date("09/07/2026"), "2026-07-09")
        self.assertEqual(app.parse_report_date("2026-07-09"), "2026-07-09")

    def test_single_day_report_header(self):
        rows = [["Period: 09.07.2026 00:00 - 09.07.2026 23:59"], ["Licence plate", "Total (km)"], ["AB12345", "7.3"]]
        self.assertEqual(app.excel_trip_dates(rows, rows[1], rows[2:], 2), ["2026-07-09"])

    def test_separate_date_per_row(self):
        rows = [["Date", "Licence plate"], ["09.07.2026", "AB12345"], ["10.07.2026", "CD12345"]]
        self.assertEqual(app.excel_trip_dates(rows, rows[0], rows[1:], 1), ["2026-07-09", "2026-07-10"])

    def test_ambiguous_period_rejected(self):
        rows = [["Period: 09.07.2026 - 10.07.2026"], ["Licence plate"], ["AB12345"]]
        with self.assertRaisesRegex(ValueError, "flere forskellige datoer"):
            app.excel_trip_dates(rows, rows[1], rows[2:], 2)

    def test_enriched_report_has_date_and_keeps_distance(self):
        rows = [["Period: 09.07.2026"], ["Licence plate", "Vehicle title", "05:00", "Total (km)"], ["AB12345", "Van", "77.3", "77.3"]]
        with patch.object(app, "read_xlsx_rows", return_value=rows), \
             patch.object(app, "load_units", return_value=({}, {})), \
             patch.object(app, "load_groups", return_value={}):
            headers, data, stats = app.enrich(b"not needed", {"manual": "", "group_mode": "none"})
        self.assertIn("Dato", headers)
        self.assertEqual(data[0]["Dato"], "2026-07-09")
        self.assertEqual(data[0]["05:00"], "77.3")
        self.assertEqual(data[0]["Total (km)"], "77.3")


if __name__ == "__main__":
    unittest.main()
