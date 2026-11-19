import json
import runpy
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch


class PasteInputTests(unittest.TestCase):
    def test_pasted_excel_builds_downloads_without_using_file_uploader(self):
        streamlit = MagicMock()
        streamlit.radio.return_value = "Paste from Excel"
        streamlit.text_area.return_value = (
            "Bid\t1.50%\r\n"
            "LV Mid\t\r\n"
            "Currency\tUSD\r\n"
            "Size\t1,000,000\r\n"
            "Assets\tSX5E Index, SPX Index\r\n"
            "Reference Date\t20-Nov-26\r\n"
            "Maturity Date\t20-Nov-31\r\n"
            "Coupons per observation\t3.00%\r\n"
            "Coupon Dates\t22-Feb-27,20-May-27,20-Aug-27\r\n"
        )
        streamlit.selectbox.return_value = "Trade 1"
        streamlit.text_input.side_effect = lambda label, value: value
        streamlit.columns.return_value = [MagicMock() for _ in range(3)]
        with patch.dict("sys.modules", {"streamlit": streamlit, "pandas": MagicMock()}):
            runpy.run_path(str(Path(__file__).with_name("app.py")))

        streamlit.file_uploader.assert_not_called()
        streamlit.success.assert_called_once_with("Recognised 1 trade.")
        downloads = streamlit.download_button.call_args_list
        self.assertEqual(len(downloads), 3)
        mapping = json.loads(downloads[2].kwargs["data"])
        fields = mapping["trades"][0]["fields"]
        self.assertEqual(fields["Coupon p.a. (%)"], "12")
        self.assertEqual(fields["Size"], "1000000")
        self.assertEqual(fields["Bloomberg Ticker 2"], "SPX")
        self.assertEqual(fields["Strike Date"], "20/11/2026")
        self.assertEqual(fields["Tenor (m)"], "60")
        self.assertIn(b"MIME-Version", downloads[0].kwargs["data"])


if __name__ == "__main__":
    unittest.main()
