import bisect
from datetime import datetime, timedelta

class My_Calendar:
    def __init__(self):
        # Set report date to most recent Sunday
        self.FISCAL_PERIODS = [
            (datetime(2024, 12, 29), "Jan"),
            (datetime(2025, 1, 26),  "Feb"),
            (datetime(2025, 2, 23),  "Mar"),
            (datetime(2025, 3, 30),  "Apr"),
            (datetime(2025, 4, 27),  "May"),
            (datetime(2025, 5, 25),  "Jun"),
            (datetime(2025, 6, 29),  "Jul"),
            (datetime(2025, 7, 27),  "Aug"),
            (datetime(2025, 8, 24),  "Sep"),
            (datetime(2025, 9, 28),  "Oct"),
            (datetime(2025, 10, 26), "Nov"),
            (datetime(2025, 11, 23), "Dec")
        ]


        today = datetime.today()
        days_since_sunday = (today.weekday() + 1) % 7  # weekday(): Mon=0, Sun=6
        last_sunday = today - timedelta(days=days_since_sunday)
        self.report_date = last_sunday
        self.report_date_str = self._format_date(last_sunday)

    def get_relative_months(self):
        starts = [d[0] for d in self.FISCAL_PERIODS]
        idx = bisect.bisect_right(starts, self.report_date) - 1

        if idx < 0:
            raise ValueError("The set report date is before the start of this fiscal year. Please either update the fiscal calendar or update the report date.")

        result = {}
        # Wrap around if going before the first month
        for i in range(1, 6):
            relative_idx = (idx - i) % len(self.FISCAL_PERIODS)
            result[-i] = self.FISCAL_PERIODS[relative_idx][1]
        return result

    def get_report_date_str(self) -> str:
        return self.report_date_str

    def get_next_fiscal_month(self):
        starts = [d[0] for d in self.FISCAL_PERIODS]
        idx = bisect.bisect_right(starts, self.report_date)
        return self._format_date(self.FISCAL_PERIODS[idx][0])

    def get_this_fiscal_month(self):

        starts = [d[0] for d in self.FISCAL_PERIODS]
        idx = bisect.bisect_right(starts, self.report_date) - 1
        if idx >= 0:
            return self._format_date(self.FISCAL_PERIODS[idx][0])
        else:
            raise ValueError("The set report date is before the start of this fiscal year. Please either update the fiscal calendar or update the report date.")

    def set_report_date(self, date: str):
        """
        Sets the report date to the given string.
        Accepts MM/DD/YYYY or M/D/YYYY and converts it to the correct format.
        Raises ValueError if invalid.
        """
        try:
            parsed = datetime.strptime(date.strip(), "%m/%d/%Y")
        except ValueError:
            try:
                parsed = datetime.strptime(date.strip(), "%m/%d/%y")  # allow 2-digit year
            except ValueError:
                raise ValueError("Invalid date format. Use M/D/YYYY or MM/DD/YYYY.")
        self.report_date = parsed
        self.report_date_str = self._format_date(parsed)

    def _format_date(self, dt: datetime) -> str:
        """Formats date without leading zeros in month/day."""
        return f"{dt.month}/{dt.day}/{dt.year}"