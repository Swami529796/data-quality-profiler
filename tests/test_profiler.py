"""Tests for profile_data.py. Run with: python -m pytest"""
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import make_sample_data as sample  # noqa: E402
import profile_data as dq  # noqa: E402

DATA = ROOT / "data"


def column(values, name="value", **options):
    return dq.profile_column(name, values, dq.Options(**options))


def count(col, mask):
    return int(col["masks"][mask].sum())


def test_single_cells_are_read_as_the_right_kind():
    assert dq.classify("42") == ("integer", 42, "")
    assert dq.classify(" 3.50 ") == ("decimal", 3.5, "")
    assert dq.classify("Rs 1,250.00") == ("decimal", 1250.0, "formatted_number")
    assert dq.classify("1,25,000") == ("integer", 125000, "formatted_number")     # Indian digit grouping
    assert dq.classify("00123")[0] == "text"                                      # a code, not a number
    assert dq.classify("Yes") == ("boolean", True, "")
    assert dq.classify("N/A") == ("missing", None, "placeholder")
    assert dq.classify("   ")[0] == "missing"
    assert dq.classify("2026-03-05") == ("date", pd.Timestamp(2026, 3, 5), "yyyy-mm-dd")
    assert dq.classify("05/03/2026")[1] == pd.Timestamp(2026, 3, 5)               # day first by default
    assert dq.classify("05/03/2026", date_order="mdy")[1] == pd.Timestamp(2026, 5, 3)
    assert dq.classify("31/02/2026")[0] == "text"                                 # not a real date
    assert dq.classify("42", from_excel=True) == ("integer", 42, "text_number")


def test_numeric_column_statistics_and_iqr_outlier():
    # 9 values: Q1 = 11, Q3 = 13, so the fences are 8 and 16 and only 100 lies outside
    col = column(["10", "11", "11", "12", "12", "13", "13", "14", "100", "", ""])
    assert col["type"] == "integer"
    assert (col["missing"], round(col["missing_pct"], 2), col["distinct"]) == (2, 18.18, 6)
    assert (col["min"], col["max"], col["median"]) == (10, 100, 12)
    assert col["mean"] == pytest.approx(196 / 9)
    assert (col["fence_low"], col["fence_high"]) == (8, 16)
    assert count(col, "outlier") == 1


def test_pattern_and_identifier_are_detected():
    codes = [f"INV-{1000 + i}" for i in range(30)]
    assert dq.shape("INV-1004") == "AAA-9999"
    clean = column(codes, name="invoice_no")
    assert (clean["type"], clean["unique"], clean["pattern"]) == ("identifier", True, "AAA-9999")

    broken = column(codes[:28] + ["inv1028", "INV-1000"], name="invoice_no")
    assert broken["type"] == "identifier" and not broken["unique"]
    assert count(broken, "pattern") == 1 and count(broken, "duplicate") == 1


def test_untidy_values_are_counted():
    city = column(["Pune", "Pune", "PUNE", "pune ", "Delhi", " Delhi", "Kochi", "New  Delhi"])
    assert count(city, "case") == 2                 # PUNE and pune differ from the usual spelling Pune
    assert count(city, "whitespace") == 3           # 'pune ', ' Delhi', 'New  Delhi'

    amount = column(["100", "250.50", "Rs 1,200", "75", "abc", "N/A"] + ["10"] * 14)
    assert amount["type"] == "decimal"
    assert (count(amount, "number_as_text"), count(amount, "mixed"), count(amount, "placeholder")) == (1, 1, 1)
    assert amount["missing"] == 1

    dates = column(["2026-01-05"] * 18 + ["15/01/2026", "soon"])
    assert dates["type"] == "date" and dates["max"] == date(2026, 1, 15)
    assert (count(dates, "date_format"), count(dates, "mixed")) == (1, 1)

    assert column(["INR"] * 5)["constant"]


def test_score_follows_the_documented_formula():
    # 10 rows: 2 blank; of the 8 values, 1 is not a number and 1 is a number stored as text
    # (7 of 8 values are numbers, so the type threshold is lowered from 0.90 for this tiny column)
    col = column(["5", "6", "7", "8", "9", "10", "Rs 1,000", "abc", "", ""], type_threshold=0.8)
    dq.score_column(col)
    assert col["parts"] == {"completeness": 80.0, "validity": 87.5, "uniqueness": None, "consistency": 87.5}
    assert col["score"] == pytest.approx((0.35 * 80 + 0.35 * 87.5 + 0.15 * 87.5) / 0.85)

    frame = pd.DataFrame({"a": ["1", "2", "3", "3"], "b": ["x", "y", "z", "z"]})
    result = dq.profile_table(frame)                # one exact duplicate row out of four
    assert result["table"]["duplicate_rows"] == 1
    assert result["score"] == pytest.approx(100 * 0.75)


def test_every_rule_type_on_a_small_table(tmp_path):
    (tmp_path / "master.csv").write_text("code\nA1\nA2\nA3\n", encoding="utf-8")
    frame = pd.DataFrame({
        "id": ["1", "2", "2", ""],
        "code": ["A1", "A2", "A9", "A3"],
        "status": ["Open", "Closed", "open", "Lost"],
        "amount": ["100", "-5", "abc", "250"],
        "start": ["2026-01-01", "2026-02-01", "2026-03-10", "2027-01-01"],
        "end": ["2026-01-05", "2026-01-20", "2026-03-10", ""],
    })
    rules = [
        {"column": "id", "check": "not_null"},
        {"column": "id", "check": "unique"},
        {"column": "status", "check": "allowed_values", "values": ["Open", "Closed"]},
        {"column": "amount", "check": "range", "min": 0, "max": 200},
        {"column": "code", "check": "regex", "pattern": r"A[1-3]"},
        {"column": "start", "check": "not_in_future"},
        {"column": "start", "check": "less_or_equal", "other": "end"},
        {"column": "code", "check": "exists_in", "file": "master.csv", "other": "code"},
        {"column": "no_such_column", "check": "not_null"},
        {"column": "id", "check": "is_prime"},
    ]
    result = dq.profile_table(frame, dq.Options(as_of=date(2026, 6, 30), type_threshold=0.6),
                              rules=rules, rules_dir=tmp_path)
    failed = [rule["failed"] for rule in result["rules"]]
    assert failed == [1, 1, 2, 2, 1, 1, 1, 1, 0, 0]
    assert result["rules"][3]["unchecked"] == 1                       # 'abc' cannot be compared with a range
    assert "row 4" in result["rules"][2]["examples"]                  # 'open' sits on file row 4
    assert [rule["status"] for rule in result["rules"][-2:]] == ["Not run", "Not run"]
    assert "not in the file" in result["rules"][-2]["reason"] and "unknown check" in result["rules"][-1]["reason"]
    assert sum(issue["issue"].startswith("Rule could not be run") for issue in result["issues"]) == 2


def test_sample_file_end_to_end_finds_every_planted_problem(tmp_path, capsys):
    planted = sample.PLANTED
    frame, from_excel = dq.load_table(DATA / "orders.csv")
    rules, key, as_of = dq.load_rules(DATA / "rules.yaml")
    result = dq.profile_table(frame, dq.Options(as_of=as_of), from_excel, key, rules, DATA)
    cols = {col["name"]: col for col in result["columns"]}
    table = result["table"]

    assert (table["rows"], table["columns"]) == (sample.N_ORDERS + 13, 15)
    assert table["blank_rows"] == planted["blank_rows"]
    assert table["duplicate_rows"] == planted["exact_duplicate_rows"]
    assert table["duplicate_keys"] == planted["order_id_reused"] + planted["exact_duplicate_rows"]
    assert table["empty_columns"] == ["remarks"] and table["constant_columns"] == ["currency"]

    assert cols["order_id"]["type"] == "identifier" and not cols["order_id"]["unique"]
    assert cols["order_id"]["pattern"] == "AAA-999999"
    assert count(cols["order_id"], "pattern") == planted["order_id_wrong_shape"]
    assert cols["order_date"]["type"] == "date"
    assert count(cols["order_date"], "date_format") == planted["order_date_other_format"]
    assert count(cols["order_date"], "mixed") == planted["order_date_unreadable"]
    assert cols["ship_date"]["missing"] == planted["ship_date_blank"]
    assert count(cols["customer_name"], "whitespace") == planted["name_stray_spaces"]
    assert cols["email"]["missing"] == planted["email_blank"]
    assert count(cols["city"], "case") == planted["city_wrong_case"]
    assert cols["quantity"]["type"] == "integer"
    assert count(cols["quantity"], "outlier") == planted["quantity_extreme"]
    assert cols["unit_price_rs"]["type"] == "decimal"
    assert count(cols["unit_price_rs"], "number_as_text") == planted["price_as_text"]
    assert count(cols["unit_price_rs"], "mixed") == planted["price_not_a_number"]
    assert cols["discount_pct"]["missing"] == planted["discount_blank"] + planted["discount_placeholder"]
    assert count(cols["discount_pct"], "placeholder") == planted["discount_placeholder"]
    assert (cols["product_code"]["pattern"], cols["prepaid"]["type"]) == ("AAA-9999", "boolean")

    failed = {rule["name"]: rule["failed"] for rule in result["rules"]}
    assert failed == {
        "order_id not_null": 0,
        "order_id unique": planted["order_id_reused"] + planted["exact_duplicate_rows"],
        "category is one of 5 allowed values": planted["category_not_allowed"],
        "quantity in range (min 1)": planted["quantity_not_positive"],
        "discount_pct in range (min 0, max 50)": 0,
        "email matches the pattern": planted["email_invalid"],
        "order_date not_in_future": planted["order_date_future"],
        "order_date is on or before ship_date": planted["ship_before_order"],
        "customer_id exists in customer_master.csv.customer_id": planted["customer_not_in_master"],
    }
    assert 0 < result["score"] < 100
    assert [issue["severity"] for issue in result["issues"]] == sorted(
        (issue["severity"] for issue in result["issues"]), key=dq.SEVERITY_ORDER.get)

    # the command line writes a self-contained report, the workbook and the charts
    assert dq.main(["--out", str(tmp_path)]) == 0
    assert "Rules: 9 given, 2 passed, 7 failed" in capsys.readouterr().out
    html = (tmp_path / "orders_profile.html").read_text(encoding="utf-8")
    assert "Overall quality score" in html and "http" not in html and "<script" not in html
    workbook = load_workbook(tmp_path / "orders_profile.xlsx")
    assert workbook.sheetnames == ["Summary", "Columns", "Issues", "Rules", "Failed Rows"]
    assert workbook["Failed Rows"].max_row - 1 == len(result["failed_rows"])
    assert all((tmp_path / name).stat().st_size > 0
               for name in ("column_scores.png", "problem_values.png", "rule_results.png"))


def test_a_clean_file_scores_close_to_100():
    rng = np.random.default_rng(7)
    clean = sample.make_orders(rng, sample.make_master(rng), rows=500, plant_problems=False)
    result = dq.profile_table(clean, key=["order_id"])
    assert result["score"] >= 99.5
    assert all(col["score"] >= 99 for col in result["columns"])
    assert not [issue for issue in result["issues"] if issue["severity"] in ("High", "Medium")]


def test_an_unseen_file_needs_no_configuration(tmp_path, capsys):
    """A different subject, different columns, and an Excel file with real number and date cells."""
    rng = np.random.default_rng(1)
    staff = pd.DataFrame({
        "Emp No": range(5001, 5061),
        "Joined": pd.date_range("2019-01-07", periods=60, freq="17D"),
        "Department": rng.choice(["Finance", "Sales", "Operations"], 60),
        "Monthly Salary Rs": rng.integers(30, 90, 60) * 1000.0 + 0.5,
        "Notes": [None] * 60,
    }).astype(object)
    staff.loc[3, "Monthly Salary Rs"] = "45,000"        # typed as text
    staff.loc[4, "Monthly Salary Rs"] = "pending"
    staff.loc[5, "Department"] = "finance"
    staff.loc[6, "Emp No"] = 5001                       # repeats the first employee number
    path = tmp_path / "staff list.xlsx"
    staff.to_excel(path, index=False, sheet_name="Staff")

    frame, from_excel = dq.load_table(path)
    assert from_excel
    cols = {col["name"]: col for col in dq.profile_table(frame, from_excel=True)["columns"]}
    assert [cols[name]["type"] for name in frame.columns] == ["identifier", "date", "text", "decimal", "empty"]
    assert count(cols["Emp No"], "duplicate") == 1 and count(cols["Department"], "case") == 1
    assert (count(cols["Monthly Salary Rs"], "number_as_text"), count(cols["Monthly Salary Rs"], "mixed")) == (1, 1)

    assert dq.main([str(path), "--out", str(tmp_path / "out")]) == 0
    assert "Rules:" not in capsys.readouterr().out
    assert (tmp_path / "out" / "staff list_profile.html").exists()

    semicolons = tmp_path / "stock.csv"
    semicolons.write_text("sku;qty\nAB-1;4\nAB-2;x\n", encoding="utf-8")
    assert dq.main([str(semicolons), "--out", str(tmp_path / "out")]) == 0
    assert dq.main([str(tmp_path / "missing.csv")]) == 1
