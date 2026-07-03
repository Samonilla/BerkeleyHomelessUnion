import os
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from smallclaims import fill_forms


class FillFormsTests(unittest.TestCase):
    def test_sc112a_does_not_receive_damages_calculation(self):
        case = {
            "plaintiff": {"name": "Jane Doe"},
            "defendant": fill_forms.DEFENDANT_DEFAULTS["city_of_oakland"],
            "claim": {
                "amount": "1000",
                "reason": "Property was destroyed during a sweep.",
                "incident_date": "01/01/2024",
                "govt_claim_filed_date": "01/15/2024",
                "damages_calculation": "Clothing $500 + emotional distress $500",
            },
            "filing": {"filing_date": "02/01/2024"},
        }

        with patch.object(fill_forms, "_write_pdf") as write_pdf:
            fill_forms.fill_sc112a(case, "template.pdf", "output.pdf")

        values = write_pdf.call_args.args[2]
        self.assertNotIn(
            "SC-112A[0].Page1[0].List3[0].Lic[0].FillText12[0]",
            values,
        )


if __name__ == "__main__":
    unittest.main()
