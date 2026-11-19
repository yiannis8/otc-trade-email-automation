import io
import unittest
from datetime import datetime

from openpyxl import Workbook

from trade_mapper import load_trade_file, map_trades


class ExcelImportTests(unittest.TestCase):
    def workbook_bytes(self, rows, formats):
        workbook = Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        for address, number_format in formats.items():
            sheet[address].number_format = number_format
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def test_percentage_cells_map_to_correct_trade_terms(self):
        # Synthetic field/value export with the same Excel types as Marex.
        rows = [
            ["Bid", 0.015],
            ["Size", 1000000],
            ["Currency", "USD"],
            ["Assets", "SX5E Index, SPX Index"],
            ["Reference Date", datetime(2026, 11, 20)],
            ["Maturity Date", datetime(2031, 11, 20)],
            ["KO Barrier", 1],
            ["Coupon Barrier", 0.6],
            ["KI Barrier", 0.6],
            ["KI Monitoring type", "EUROPEAN"],
            ["Put Type", "KI_PUT"],
            ["Put Strike", 1],
            ["Memory", False],
            ["Coupons per observation", 0.03],
            ["Coupon Dates", "22-Feb-27,20-May-27,20-Aug-27"],
            ["KO Dates", "20-May-27,20-Aug-27,22-Nov-27"],
        ]
        formats = {f"B{row}": "0.00%" for row in (1, 7, 8, 9, 12, 14)}
        raw = self.workbook_bytes(rows, formats)
        trades = load_trade_file(raw, "synthetic.xlsx")
        self.assertEqual(len(trades), 1)
        result = map_trades(trades)[0]
        expected = {
            "Autocall Trigger Level (%)": "100",
            "Coupon Trigger Level (%)": "60",
            "Barrier Level": "60",
            "Strike Level (%)": "100",
            "Coupon p.a. (%)": "12",
            "Strike Date": "20/11/2026",
            "Tenor (m)": "60",
            "First Observation in (m)": "6",
            "Memory": "No",
            "Size": "1000000",
        }
        for key, value in expected.items():
            with self.subTest(field=key):
                self.assertEqual(result.fields[key], value)
        self.assertEqual(result.source["Bid"], "1.5%")

    def test_matrix_preserves_percentage_units_and_trade_ids(self):
        raw = self.workbook_bytes(
            [[None, "Trade A", "Trade B"], ["Bid", 0.015, 0.02],
             ["Currency", "USD", "EUR"]],
            {"B2": "0.00%", "C2": "0%"},
        )
        results = map_trades(load_trade_file(raw, "synthetic.xlsm"))
        self.assertEqual([r.trade_id for r in results], ["Trade A", "Trade B"])
        self.assertEqual([r.fields["Reoffer / Upfront (%)"] for r in results], ["1.5", "2"])

    def test_unformatted_and_literal_percent_values_are_not_scaled(self):
        for number_format in ("General", '0.00"%"', r"0.00\%"):
            with self.subTest(number_format=number_format):
                raw = self.workbook_bytes(
                    [["Bid", 0.6], ["Currency", "USD"]],
                    {"B1": number_format},
                )
                result = map_trades(load_trade_file(raw, "synthetic.xlsx"))[0]
                self.assertEqual(result.fields["Reoffer / Upfront (%)"], "0.6")

    def test_csv_percentages_keep_existing_units(self):
        for value in ("0.6", "60", "60%"):
            with self.subTest(value=value):
                raw = f"Bid,{value}\nCurrency,USD\n".encode()
                result = map_trades(load_trade_file(raw, "synthetic.csv"))[0]
                self.assertEqual(result.fields["Reoffer / Upfront (%)"], value.rstrip("%"))


if __name__ == "__main__":
    unittest.main()
