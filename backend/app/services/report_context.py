from dataclasses import dataclass


@dataclass
class ReportContext:
    report_date: str
    station: str
    bound: str

    wideload_count: int = 0
    overloaded_valid_permit_count: int = 0