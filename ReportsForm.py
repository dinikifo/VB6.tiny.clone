import json
import sys
from PySide6.QtWidgets import QApplication

try:
    from tinyvb_webext import runtime, interpreter, gui
except ImportError:
    # Fallback to local modules in the same folder
    import runtime, interpreter, gui


FORM_JSON = """{
  "name": "ReportsForm",
  "title": "Account Statistics (Five-Number Summary)",
  "width": 700,
  "height": 400,
  "controls": [
    {
      "id": "lblAccount",
      "type": "Label",
      "text": "Account code",
      "x": 20,
      "y": 20,
      "width": 120,
      "height": 24
    },
    {
      "id": "txtAccount",
      "name": "txtAccount",
      "type": "TextBox",
      "x": 150,
      "y": 20,
      "width": 160,
      "height": 24
    },
    {
      "id": "lblFrom",
      "type": "Label",
      "text": "From period (YYYY-MM)",
      "x": 20,
      "y": 60,
      "width": 160,
      "height": 24
    },
    {
      "id": "txtFromPeriod",
      "name": "txtFromPeriod",
      "type": "TextBox",
      "x": 190,
      "y": 60,
      "width": 120,
      "height": 24
    },
    {
      "id": "lblTo",
      "type": "Label",
      "text": "To period (YYYY-MM)",
      "x": 330,
      "y": 60,
      "width": 160,
      "height": 24
    },
    {
      "id": "txtToPeriod",
      "name": "txtToPeriod",
      "type": "TextBox",
      "x": 500,
      "y": 60,
      "width": 120,
      "height": 24
    },
    {
      "id": "btnRunReport",
      "name": "btnRunReport",
      "type": "Button",
      "text": "Run report",
      "x": 530,
      "y": 20,
      "width": 120,
      "height": 30,
      "events": {
        "click": "btnRunReport_Click"
      }
    },
    {
      "id": "lblReport",
      "type": "Label",
      "text": "Report",
      "x": 20,
      "y": 110,
      "width": 100,
      "height": 24
    },
    {
      "id": "lstReport",
      "name": "lstReport",
      "type": "ListBox",
      "x": 20,
      "y": 140,
      "width": 660,
      "height": 220
    }
  ]
}"""


SCRIPT_TEXT = r"""
Dim currentStatsId

Sub ReportsForm_Load()
    currentStatsId = ""
End Sub

Sub btnRunReport_Click()
    Dim acc, pFrom, pTo
    Dim minVal, q1, med, q3, maxVal

    acc = txtAccount.Text
    pFrom = txtFromPeriod.Text
    pTo   = txtToPeriod.Text

    lstReport.Clear

    If acc = "" Then
        MsgBox "Please enter an account code."
        Exit Sub
    End If

    If pFrom = "" Or pTo = "" Then
        MsgBox "Please enter both From and To periods (e.g. 2025-01)."
        Exit Sub
    End If

    currentStatsId = StatsPrepare(acc, pFrom, pTo)

    If currentStatsId = "" Then
        lstReport.Add "No data for account " & acc & " between " & pFrom & " and " & pTo & "."
        Exit Sub
    End If

    minVal = StatsGet(currentStatsId, "min")
    q1     = StatsGet(currentStatsId, "q1")
    med    = StatsGet(currentStatsId, "median")
    q3     = StatsGet(currentStatsId, "q3")
    maxVal = StatsGet(currentStatsId, "max")

    lstReport.Add "Five-number summary for " & acc & " (" & pFrom & " to " & pTo & "):"
    lstReport.Add ""
    lstReport.Add "  Min:     " & minVal
    lstReport.Add "  Q1:      " & q1
    lstReport.Add "  Median:  " & med
    lstReport.Add "  Q3:      " & q3
    lstReport.Add "  Max:     " & maxVal
End Sub
"""


def _find_account_id(app_data, code):
    ledger = app_data.get("ledger", {})
    accounts = ledger.get("accounts", [])
    for a in accounts:
        if isinstance(a, dict) and a.get("code") == code:
            return a.get("id")
    return None


def build_account_series(app_data, account_code, period_from, period_to):
    ledger = app_data.get("ledger", {})
    postings = ledger.get("postings", [])

    acc_id = _find_account_id(app_data, account_code)
    if acc_id is None:
        return []

    sums = {}
    for p in postings:
        if not isinstance(p, dict):
            continue
        if p.get("accountId") != acc_id:
            continue
        period = str(p.get("period", ""))
        if period_from and period < period_from:
            continue
        if period_to and period > period_to:
            continue
        try:
            amt = float(p.get("amount", 0.0))
        except (TypeError, ValueError):
            amt = 0.0
        sums[period] = sums.get(period, 0.0) + amt

    return sorted(sums.items(), key=lambda kv: kv[0])


def five_number_summary(values):
    if not values:
        return None

    vals = sorted(float(v) for v in values)
    n = len(vals)
    min_val = vals[0]
    max_val = vals[-1]

    # Median
    if n % 2 == 1:
        median = vals[n // 2]
    else:
        median = (vals[n // 2 - 1] + vals[n // 2]) / 2.0

    # Simple Q1/Q3: medians of lower/upper halves
    if n == 1:
        q1 = median
        q3 = median
    else:
        mid = n // 2
        if n % 2 == 0:
            lower = vals[:mid]
            upper = vals[mid:]
        else:
            lower = vals[:mid]
            upper = vals[mid + 1:]

        def _med(seq):
            m = len(seq)
            if m == 0:
                return median
            if m % 2 == 1:
                return seq[m // 2]
            return (seq[m // 2 - 1] + seq[m // 2]) / 2.0

        q1 = _med(lower)
        q3 = _med(upper)

    return {
        "min": min_val,
        "q1": q1,
        "median": median,
        "q3": q3,
        "max": max_val,
    }


def run_stats_analysis(app_data, account_code, period_from, period_to):
    series = build_account_series(app_data, account_code, period_from, period_to)
    values = [v for _, v in series]
    if not values:
        return None

    summary = five_number_summary(values)

    analysis_id = f"stats:{account_code}:{period_from}:{period_to}"
    analysis = app_data.setdefault("analysis", {})
    analysis[analysis_id] = {
        "type": "five_number",
        "account": account_code,
        "period_from": period_from,
        "period_to": period_to,
        "series": series,
        "summary": summary,
    }
    return analysis_id


def get_stats_value(app_data, analysis_id, key):
    analysis = app_data.get("analysis", {})
    entry = analysis.get(analysis_id)
    if not isinstance(entry, dict):
        return 0.0
    if entry.get("type") != "five_number":
        return 0.0
    summary = entry.get("summary", {})
    return summary.get(key, 0.0)


class StatsVBInterpreter(interpreter.VBInterpreter):
    def _call_function(self, name, args):
        upper = name.upper()

        if upper == "STATSPREPARE":
            app_data = self.ctx.get_var("AppData")
            account_code = str(args[0]) if len(args) > 0 else ""
            period_from = str(args[1]) if len(args) > 1 else ""
            period_to = str(args[2]) if len(args) > 2 else ""
            if not account_code:
                return ""
            analysis_id = run_stats_analysis(app_data, account_code, period_from, period_to)
            return analysis_id or ""

        if upper == "STATSGET":
            app_data = self.ctx.get_var("AppData")
            analysis_id = str(args[0]) if len(args) > 0 else ""
            key = str(args[1]) if len(args) > 1 else ""
            return get_stats_value(app_data, analysis_id, key)

        return super()._call_function(name, args)


def main():
    form_def = json.loads(FORM_JSON)

    ctx = runtime.VBContext()
    app_data = runtime.load_app_data()
    ctx.set_var("AppData", app_data)

    interp = StatsVBInterpreter(ctx)
    interp.load_source(SCRIPT_TEXT)

    app = QApplication(sys.argv)
    form = gui.VBForm(form_def, interp, ctx)

    form.show()
    exit_code = app.exec()

    runtime.save_app_data(app_data)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
