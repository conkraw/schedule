"""Per-build report reuse. Never cached globally or written to GitHub.

Calculations still use the existing validated public functions. An entire
archive's time rows used to be recalculated for *every* individual document.
This batch computes each table once, then indexes it by preceptor. It is local
to one synchronous ZIP build and must not survive changes to its source scan.
"""
from collections import defaultdict

from schedule_app.services.educational_time import teaching_time_rows
from schedule_app.services.opd_archive import OPDArchiveError


class TeachingReportBatch:
    def __init__(self, scan, years):
        self._scan = scan
        self._years = tuple(sorted({int(year) for year in years}))
        self._tables = {}
        self._by_preceptor = {}
        self._chair = None

    def _check(self, scan, years):
        if scan is not self._scan or not set(map(int, years)).issubset(self._years):
            raise OPDArchiveError("Prepared report data does not match this scan/period. Rebuild the reports.")

    def rows(self, scan, years, kind="overall"):
        self._check(scan, years)
        if kind not in ("overall", "work_type", "monthly"):
            raise ValueError("Unknown teaching table kind")
        if kind not in self._tables:
            self._tables[kind] = teaching_time_rows(
                scan, self._years, by_work_type=kind != "overall", monthly=kind == "monthly")
            indexed = defaultdict(list)
            for row in self._tables[kind]:
                indexed[row["preceptor_name"]].append(row)
            self._by_preceptor[kind] = dict(indexed)
        # Internal batch rows describe the full build's years. Callers wanting a
        # smaller set filter the labels, exactly as the individual writer does.
        return self._tables[kind]

    def preceptor_rows(self, scan, name, years):
        self._check(scan, years)
        values = []
        for kind in ("overall", "work_type", "monthly"):
            self.rows(scan, years, kind)
            values.append(self._by_preceptor[kind].get(name, []))
        return tuple(values)

    def chair_summaries(self, scan, years):
        self._check(scan, years)
        if tuple(sorted({int(year) for year in years})) != self._years:
            raise OPDArchiveError("The chair report must use the whole prepared reporting period.")
        if self._chair is None:
            from schedule_app.reports.chair_summary import teaching_chair_summary_data
            self._chair = teaching_chair_summary_data(scan, self._years)
        return self._chair
