import pandas as pd
import pytest

from app.services.overloaded_summary import count_valid_permit_vehicles


def test_count_valid_permit_vehicles_counts_known_vardict_formats():
    df = pd.DataFrame(
        {
            "Vardict": [
                "Vehicle has a valid permit App-0388434",
                (
                    "Special Release:Vehicle With Permit; "
                    "Truck has a valid Permit No: App-0390912 ; "
                    "Authorized by: XXXXXXXXXX"
                ),
                (
                    "Special Release:Vehicle With Permit; "
                    "Permit No: App-0390756 ; Authorized by: DUNCAN ODHIAMBO"
                ),
                "No valid permit noted",
                "Minimum axle overload of 160 [KG] : ",
                None,
            ]
        }
    )

    assert count_valid_permit_vehicles(df) == 3


def test_count_valid_permit_vehicles_does_not_double_count_one_row():
    df = pd.DataFrame(
        {
            "Vardict": [
                (
                    "Vehicle has a valid permit App-0388434; "
                    "Special Release:Vehicle With Permit; Permit No: App-0390756"
                )
            ]
        }
    )

    assert count_valid_permit_vehicles(df) == 1


def test_count_valid_permit_vehicles_requires_vardict_column():
    with pytest.raises(ValueError, match="Missing required column: Vardict"):
        count_valid_permit_vehicles(pd.DataFrame({"Other": []}))
