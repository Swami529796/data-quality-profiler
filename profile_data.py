"""Data Quality Profiler.

Point it at any CSV or Excel file and it tells you what state the data is in,
before anyone builds a report on it. No setup is needed:

    for every column   the type it really holds (integer, decimal, date, text, boolean, identifier),
                       % missing, distinct values, most common values, min / max / mean / median,
                       date range, text length, the usual shape of the values (ORD-100234 -> AAA-999999),
                       stray spaces, mixed capital letters, numbers and dates stored as text,
                       values that do not fit the column type, constant columns, outliers (IQR rule)
    for the table      row count, exact duplicate rows, duplicate keys, empty rows and columns
    optional rules     business rules written in a small YAML file: not_null, unique, allowed_values,
                       range, regex, not_in_future, less_or_equal (column A <= column B) and
                       exists_in (the value must exist in a column of another file)
    a score            0 to 100 per column and overall, from completeness, validity, uniqueness
                       and consistency. It is a summary to compare files with, not a guarantee.

Output (in the output folder): one self-contained HTML report, an Excel workbook with the
same tables plus a Failed Rows sheet, three PNG charts, and a summary on the console.

Usage:
    python profile_data.py                                   (profiles the bundled sample with its rules)
    python profile_data.py my_file.csv                       (any file, no configuration)
    python profile_data.py my_file.xlsx --sheet Data --key invoice_no
    python profile_data.py my_file.csv --rules my_rules.yaml --out reports

Author : Swaminathan Narayanan
Repo   : https://github.com/Swami529796/data-quality-profiler
"""
from __future__ import annotations

import argparse
import math
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import yaml  # noqa: E402
from jinja2 import Environment, FileSystemLoader, select_autoescape  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

HERE = Path(__file__).resolve().parent
SAMPLE_FILE = HERE / "data" / "orders.csv"
SAMPLE_RULES = HERE / "data" / "rules.yaml"

NAVY, LIGHT_BLUE, AMBER, GREY = "#1F4E78", "#9DB7D5", "#E08A00", "#595959"

# Share of each part in a column's score. Uniqueness only counts where it is expected.
WEIGHTS = {"completeness": 0.35, "validity": 0.35, "uniqueness": 0.15, "consistency": 0.15}

NULL_TOKENS = {"na", "n/a", "#n/a", "null", "none", "nan", "nil", "-", "--", "?"}
BOOLEAN_WORDS = {"true": True, "yes": True, "y": True, "false": False, "no": False, "n": False}
PATTERN_SHARE = 0.60          # a shape is reported as "the pattern" when this share of values has it
PATTERN_MIN_VALUES = 10       # ...and at least this many different values share it
IDENTIFIER_SHARE = 0.90       # distinct values / filled values needed to call a column identifier-like
IDENTIFIER_MIN_ROWS = 20
OUTLIER_MIN_VALUES = 8
MAX_FAILED_ROWS = 50_000      # cap on the Failed Rows sheet; the report says so if it is reached
SEVERITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}

INT_RE = re.compile(r"[+-]?\d+")
DEC_RE = re.compile(r"[+-]?(?:\d+\.\d*|\.\d+)(?:[eE][+-]?\d+)?")
FORMATTED_RE = re.compile(r"(?P<cur>(?:Rs\.?|INR|₹|\$)\s*)?(?P<num>[+-]?(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d+)?)",
                          re.IGNORECASE)
DATE_HINT_RE = re.compile(r"\d{1,4}[-/. ][0-9A-Za-z]{1,9}[-/. ,]+\d{1,4}")
TIME_TAIL_RE = re.compile(r"^(.*?)[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$")
NAME_HINT_RE = re.compile(r"(?:^|[^a-z])(?:id|no|num|number|code|key|ref)(?:$|[^a-z])", re.IGNORECASE)
NAME_HINT_CAMEL_RE = re.compile(r"[a-z](?:ID|Id|No)$")

MISSING = ("missing", None, "")


@dataclass
class Options:
    type_threshold: float = 0.90      # share of filled values that must agree before a type is assigned
    pattern_threshold: float = 0.90   # share that must follow one shape before the rest are flagged
    iqr_k: float = 1.5                # outlier fence: quartile -/+ k x interquartile range
    date_order: str = "dmy"           # how to read 03/04/2026: dmy = 3 April, mdy = 4 March
    top: int = 5                      # most common values listed per column
    as_of: date = field(default_factory=date.today)


# --------------------------------------------------------------------------- reading cells

def date_formats(order: str) -> list[tuple[str, str]]:
    first, label = ("%d", "dd") if order == "dmy" else ("%m", "mm")
    second, label2 = ("%m", "mm") if order == "dmy" else ("%d", "dd")
    return [
        ("yyyy-mm-dd", "%Y-%m-%d"), (f"{label}/{label2}/yyyy", f"{first}/{second}/%Y"),
        (f"{label}-{label2}-yyyy", f"{first}-{second}-%Y"), (f"{label}.{label2}.yyyy", f"{first}.{second}.%Y"),
        ("dd-mmm-yyyy", "%d-%b-%Y"), ("dd mmm yyyy", "%d %b %Y"), ("dd mmmm yyyy", "%d %B %Y"),
        ("mmm dd, yyyy", "%b %d, %Y"), ("yyyy/mm/dd", "%Y/%m/%d"), (f"{label}/{label2}/yy", f"{first}/{second}/%y"),
    ]


def parse_date(text: str, order: str = "dmy") -> tuple[pd.Timestamp, str] | None:
    """Read a date written as text. Returns the date and the name of the format it was written in."""
    tail = TIME_TAIL_RE.match(text)
    day_part = tail.group(1) if tail else text
    if not DATE_HINT_RE.fullmatch(day_part):
        return None
    for label, fmt in date_formats(order):
        try:
            parsed = datetime.strptime(day_part, fmt)
        except ValueError:
            continue
        if 1900 <= parsed.year <= 2100:
            return pd.Timestamp(parsed), label
    return None


def is_blank(value) -> bool:
    if value is None or value is pd.NaT:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return isinstance(value, str) and not value.strip()


def show(value) -> str:
    """A cell as the user would read it."""
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date().isoformat() if value.time() == datetime.min.time() else str(value)
    return str(value)


def classify(value, date_order: str = "dmy", from_excel: bool = False) -> tuple[str, object, str]:
    """What one cell holds: (kind, parsed value, note).

    kind is missing, integer, decimal, boolean, date or text. The note records how it was
    written: 'placeholder' (N/A and friends), 'formatted_number' ("Rs 1,250.00"),
    'text_number' (a number typed as text in Excel), or the date format.
    """
    if is_blank(value):
        return MISSING
    if isinstance(value, (bool, np.bool_)):
        return "boolean", bool(value), ""
    if isinstance(value, (int, np.integer)):
        return "integer", int(value), ""
    if isinstance(value, (float, np.floating)):
        return ("integer", int(value), "") if float(value).is_integer() else ("decimal", float(value), "")
    if isinstance(value, (datetime, date)):
        return "date", pd.Timestamp(value), "date cell"

    text = str(value).strip()
    if text.lower() in NULL_TOKENS:
        return "missing", None, "placeholder"
    as_text = "text_number" if from_excel else ""
    if INT_RE.fullmatch(text):
        digits = text.lstrip("+-")
        if (len(digits) > 1 and digits[0] == "0") or len(digits) > 15:
            return "text", text, ""              # 00123 is a code, not the number 123
        return "integer", int(text), as_text
    if DEC_RE.fullmatch(text):
        return "decimal", float(text), as_text
    formatted = FORMATTED_RE.fullmatch(text)
    if formatted and ("," in text or formatted.group("cur")):
        number = float(formatted.group("num").replace(",", ""))
        whole = "." not in formatted.group("num")
        return ("integer", int(number), "formatted_number") if whole else ("decimal", number, "formatted_number")
    if text.lower() in BOOLEAN_WORDS:
        return "boolean", BOOLEAN_WORDS[text.lower()], ""
    parsed = parse_date(text, date_order)
    if parsed:
        return "date", parsed[0], parsed[1]
    return "text", text, ""


def shape(text: str) -> str:
    """Character-class mask: capital letter -> A, small letter -> a, digit -> 9, the rest unchanged."""
    return "".join("9" if c.isdigit() else "A" if c.isupper() else "a" if c.islower() else c for c in text)


# --------------------------------------------------------------------------- one column

def profile_column(name: str, values: list, opts: Options, from_excel: bool = False) -> dict:
    """Everything the profiler can say about one column without being told anything."""
    cache: dict = {}
    cells = []
    for value in values:
        try:
            key = (type(value), value)
            if key not in cache:
                cache[key] = classify(value, opts.date_order, from_excel)
            cells.append(cache[key])
        except TypeError:                         # unhashable cell
            cells.append(classify(value, opts.date_order, from_excel))

    n = len(values)
    kind = np.array([c[0] for c in cells], dtype=object)
    note = np.array([c[2] for c in cells], dtype=object)
    parsed = [c[1] for c in cells]
    missing = kind == "missing"
    filled = ~missing
    nn = int(filled.sum())
    text = np.array(["" if gone else show(value).strip() for value, gone in zip(values, missing)], dtype=object)
    is_number = (kind == "integer") | (kind == "decimal")
    is_date = kind == "date"
    none = np.zeros(n, dtype=bool)

    # ---- type: the kind that at least `type_threshold` of the filled cells agree on
    share = (lambda mask: mask.sum() / nn) if nn else (lambda mask: 0.0)
    if nn == 0:
        base = "empty"
    elif share(is_number) >= opts.type_threshold:
        base = "decimal" if (kind == "decimal").any() else "integer"
    elif share(is_date) >= opts.type_threshold:
        base = "date"
    elif share(kind == "boolean") >= opts.type_threshold:
        base = "boolean"
    else:
        base = "text"

    # ---- cells that do not fit the type
    mixed, mixed_note = none.copy(), ""
    fits = {"integer": is_number, "decimal": is_number, "date": is_date, "boolean": kind == "boolean"}
    if base in fits:
        mixed = filled & ~fits[base]
    elif base == "text" and share(kind == "text") < opts.type_threshold:
        groups = {"number": is_number, "date": is_date, "boolean": kind == "boolean", "text": kind == "text"}
        largest = max(groups, key=lambda g: groups[g].sum())
        if largest != "text":                     # mostly numbers or dates, but too many exceptions to call it one
            mixed = filled & ~groups[largest]
            mixed_note = "no single type: " + ", ".join(
                f"{share(mask):.0%} {group}" for group, mask in sorted(groups.items(), key=lambda g: -g[1].sum())
                if mask.any())

    # ---- numbers and dates written differently from the rest of the column
    number_as_text, date_format, main_date_format = none.copy(), none.copy(), ""
    if base in ("integer", "decimal"):
        has_number_cells = bool((is_number & (note == "")).any())
        number_as_text = is_number & ((note == "formatted_number") | ((note == "text_number") & has_number_cells))
    if base == "date":
        formats = pd.Series(note[is_date]).value_counts()
        main_date_format = formats.index[0]
        date_format = is_date & (note != main_date_format)

    # ---- spaces and capital letters
    is_string = np.array([isinstance(value, str) for value in values], dtype=bool) & filled
    whitespace = np.array([bool(s) and (v != v.strip() or "  " in v) for v, s in zip(values, is_string)], dtype=bool)
    case = none.copy()
    if is_string.any():
        words = pd.Series(text[is_string], index=np.flatnonzero(is_string))
        folded = words.str.casefold()
        spellings = words.groupby(folded).agg(lambda group: group.value_counts().index[0])
        wrong = words[words != folded.map(spellings)]
        case[wrong.index] = True

    # ---- distinct values, most common values, shape of the values
    filled_text = pd.Series(text[filled], index=np.flatnonzero(filled))
    counts = filled_text.value_counts()
    distinct = int(len(counts))
    top_values = [(value, int(count)) for value, count in counts.head(opts.top).items()]
    duplicate = none.copy()
    duplicate[filled_text[filled_text.duplicated(keep="first")].index] = True

    pattern, pattern_share, pattern_break, top_shapes = "", 0.0, none.copy(), []
    if base in ("text", "integer") and nn:
        shapes = filled_text.map(shape)
        shape_counts = shapes.value_counts()
        top_shapes = [(s, int(c)) for s, c in shape_counts.head(3).items()]
        top_shape, top_share = shape_counts.index[0], shape_counts.iloc[0] / nn
        if base == "text" and top_share >= PATTERN_SHARE and \
                filled_text[shapes == top_shape].nunique() >= PATTERN_MIN_VALUES:
            pattern, pattern_share = top_shape, float(top_share)
            if top_share >= opts.pattern_threshold:
                pattern_break[shapes[shapes != top_shape].index] = True

    # ---- identifier-like: nearly every value is different and the values look like codes
    identifier = False
    if base in ("text", "integer") and nn >= IDENTIFIER_MIN_ROWS and distinct / nn >= IDENTIFIER_SHARE:
        if NAME_HINT_RE.search(name) or NAME_HINT_CAMEL_RE.search(name):
            identifier = True
        elif base == "text":
            identifier = top_share >= 0.8 and "9" in top_shape and " " not in top_shape
        else:
            lengths = {len(value.lstrip("+-")) for value in counts.index}
            identifier = len(lengths) == 1 and min(lengths) >= 4

    column = {
        "name": name, "rows": n, "filled": nn, "missing": n - nn, "missing_pct": 100 * (n - nn) / n if n else 0.0,
        "base_type": base, "type": "identifier" if identifier else base, "identifier": identifier,
        "unique": identifier and not duplicate.any(), "uniqueness_applies": identifier,
        "distinct": distinct, "top_values": top_values, "constant": distinct == 1 and nn > 1,
        "pattern": pattern, "pattern_share": pattern_share, "top_shapes": top_shapes,
        "mixed_note": mixed_note, "main_date_format": main_date_format,
        "kind": kind, "parsed": parsed, "text": text, "raw": np.array([str(value) for value in values], dtype=object),
        "masks": {"missing": missing, "placeholder": note == "placeholder", "mixed": mixed,
                  "number_as_text": number_as_text, "date_format": date_format, "pattern": pattern_break,
                  "whitespace": whitespace, "case": case, "duplicate": duplicate, "outlier": none.copy(),
                  "rule": none.copy()},
    }

    if base in ("integer", "decimal") and not identifier:
        numbers = np.array([float(parsed[i]) for i in np.flatnonzero(is_number)])
        column.update(min=numbers.min(), max=numbers.max(), mean=numbers.mean(), median=float(np.median(numbers)))
        if len(numbers) >= OUTLIER_MIN_VALUES:
            q1, q3 = np.percentile(numbers, [25, 75])
            low, high = q1 - opts.iqr_k * (q3 - q1), q3 + opts.iqr_k * (q3 - q1)
            column.update(q1=q1, q3=q3, fence_low=low, fence_high=high)
            if q3 > q1:
                column["masks"]["outlier"][np.flatnonzero(is_number)[(numbers < low) | (numbers > high)]] = True
    if base == "date":
        dates = [parsed[i] for i in np.flatnonzero(is_date)]
        column.update(min=min(dates).date(), max=max(dates).date())
    if base == "text" or identifier:
        lengths = filled_text.str.len()
        column.update(length_min=int(lengths.min()), length_max=int(lengths.max()))
    return column


def score_column(column: dict) -> None:
    """Add the four part-scores and the 0-100 column score (formula in the README)."""
    masks, n, nn = column["masks"], column["rows"], column["filled"]
    parts: dict[str, float | None] = {"completeness": 100 * nn / n if n else 100.0,
                                      "validity": None, "uniqueness": None, "consistency": None}
    if nn:
        invalid = masks["mixed"] | masks["pattern"] | masks["rule"]
        untidy = masks["whitespace"] | masks["case"] | masks["number_as_text"] | masks["date_format"]
        parts["validity"] = 100 * (1 - invalid.sum() / nn)
        parts["consistency"] = 100 * (1 - untidy.sum() / nn)
        if column["uniqueness_applies"]:
            parts["uniqueness"] = 100 * (1 - masks["duplicate"].sum() / nn)
    used = {part: value for part, value in parts.items() if value is not None}
    column["parts"] = parts
    column["score"] = sum(WEIGHTS[part] * value for part, value in used.items()) / sum(WEIGHTS[part] for part in used)


# --------------------------------------------------------------------------- business rules

def describe_rule(rule: dict) -> str:
    check, column = rule.get("check", "?"), rule.get("column", "?")
    if check == "allowed_values":
        return f"{column} is one of {len(rule.get('values') or [])} allowed values"
    if check == "range":
        bounds = [f"{word} {rule[word]}" for word in ("min", "max") if rule.get(word) is not None]
        return f"{column} in range ({', '.join(bounds)})"
    if check == "regex":
        return f"{column} matches the pattern"
    if check == "less_or_equal":
        return f"{column} <= {rule.get('other', '?')}"
    if check == "exists_in":
        return f"{column} exists in {rule.get('file', '?')}.{rule.get('other', column)}"
    return f"{column} {check}"


def run_rule(rule, columns: dict[str, dict], row_numbers: np.ndarray, opts: Options, rules_dir: Path) -> dict:
    """Check one rule. A rule that cannot be run is reported with the reason, never skipped quietly."""
    result = {"name": "", "column": "", "check": "", "severity": "High", "status": "Not run", "checked": 0,
              "failed": 0, "unchecked": 0, "pct": 0.0, "examples": "", "reason": "", "mask": None}
    try:
        if not isinstance(rule, dict):
            raise ValueError(f"a rule must have 'column' and 'check', got: {rule!r}")
        result.update(column=str(rule.get("column", "")), check=str(rule.get("check", "")),
                      name=str(rule.get("name") or describe_rule(rule)))
        severity = str(rule.get("severity", "High")).title()
        if severity not in SEVERITY_ORDER:
            raise ValueError(f"severity must be High, Medium or Low, not '{rule.get('severity')}'")
        result["severity"] = severity
        if result["column"] not in columns:
            raise ValueError(f"column '{result['column']}' is not in the file")
        column = columns[result["column"]]
        kind, text, parsed = column["kind"], column["text"], column["parsed"]
        filled = kind != "missing"
        check = result["check"]
        unchecked = np.zeros(len(kind), dtype=bool)
        considered = filled

        def as_numbers(col: dict) -> tuple[np.ndarray, np.ndarray]:
            ok = (col["kind"] == "integer") | (col["kind"] == "decimal")
            return np.array([float(v) if good else np.nan for v, good in zip(col["parsed"], ok)]), ok

        def as_dates(col: dict) -> tuple[np.ndarray, np.ndarray]:
            ok = col["kind"] == "date"
            stamps = [v if good else pd.NaT for v, good in zip(col["parsed"], ok)]
            return pd.to_datetime(pd.Series(stamps, dtype="datetime64[ns]")).to_numpy(), ok

        if check == "not_null":
            considered = np.ones(len(kind), dtype=bool)
            failed = ~filled
        elif check == "unique":
            failed = column["masks"]["duplicate"].copy()
            column["uniqueness_applies"] = True
        elif check == "allowed_values":
            if not isinstance(rule.get("values"), list) or not rule["values"]:
                raise ValueError("allowed_values needs a list under 'values'")
            fold = (lambda s: s.casefold()) if rule.get("ignore_case") else (lambda s: s)
            allowed = {fold(str(value).strip()) for value in rule["values"]}
            failed = filled & np.array([fold(value) not in allowed for value in text], dtype=bool)
        elif check == "regex":
            if not rule.get("pattern"):
                raise ValueError("regex needs a 'pattern'")
            compiled = re.compile(str(rule["pattern"]))
            failed = filled & np.array([compiled.fullmatch(value) is None for value in text], dtype=bool)
        elif check == "range":
            if rule.get("min") is None and rule.get("max") is None:
                raise ValueError("range needs 'min', 'max' or both")
            on_dates = column["base_type"] == "date"
            values, ok = as_dates(column) if on_dates else as_numbers(column)
            convert = (lambda bound: np.datetime64(pd.Timestamp(str(bound)))) if on_dates else float
            failed = np.zeros(len(kind), dtype=bool)
            if rule.get("min") is not None:
                failed |= ok & (values < convert(rule["min"]))
            if rule.get("max") is not None:
                failed |= ok & (values > convert(rule["max"]))
            unchecked = filled & ~ok
        elif check == "not_in_future":
            values, ok = as_dates(column)
            failed = ok & (values > np.datetime64(pd.Timestamp(opts.as_of)))
            unchecked = filled & ~ok
        elif check == "less_or_equal":
            if rule.get("other") not in columns:
                raise ValueError(f"'other' column '{rule.get('other')}' is not in the file")
            other = columns[rule["other"]]
            reader = as_dates if column["base_type"] == "date" else as_numbers
            (left, left_ok), (right, right_ok) = reader(column), reader(other)
            considered = filled & (other["kind"] != "missing")
            failed = left_ok & right_ok & (left > right)
            unchecked = considered & ~(left_ok & right_ok)
        elif check == "exists_in":
            if not rule.get("file"):
                raise ValueError("exists_in needs a 'file'")
            lookup_path = rules_dir / str(rule["file"])
            if not lookup_path.is_file():
                raise ValueError(f"lookup file not found: {rule['file']}")
            lookup, _ = load_table(lookup_path, rule.get("sheet"))
            lookup_column = str(rule.get("other") or result["column"])
            if lookup_column not in lookup.columns:
                raise ValueError(f"column '{lookup_column}' is not in {rule['file']}")
            known = {show(value).strip() for value in lookup[lookup_column] if not is_blank(value)}
            failed = filled & np.array([value not in known for value in text], dtype=bool)
        else:
            raise ValueError(f"unknown check '{check}'")

        checked = int(considered.sum() - unchecked.sum())
        result.update(checked=checked, failed=int(failed.sum()), unchecked=int(unchecked.sum()),
                      pct=100 * failed.sum() / checked if checked else 0.0, mask=failed,
                      status="Fail" if failed.any() else "Pass", examples=examples(text, failed, row_numbers, 5))
        if check not in ("not_null", "unique"):          # those two are scored as completeness / uniqueness
            column["masks"]["rule"] |= failed
    except (ValueError, re.error, OSError) as error:
        result["reason"] = str(error)
    return result


# --------------------------------------------------------------------------- the whole table

def examples(text: np.ndarray, mask: np.ndarray, row_numbers: np.ndarray, limit: int = 3) -> str:
    """A few different failing values with the file row they first appear on (row 1 is the header)."""
    seen: dict[str, int] = {}
    for position in np.flatnonzero(mask):
        seen.setdefault(text[position], int(row_numbers[position]))
        if len(seen) == limit:
            break
    return "; ".join(f"'{value}' (row {row})" if value else f"blank (row {row})" for value, row in seen.items())


def missing_severity(pct: float) -> str:
    return "High" if pct > 20 else "Medium" if pct > 5 else "Low"


def column_issues(column: dict, row_numbers: np.ndarray, opts: Options) -> list[dict]:
    masks, nn, name = column["masks"], column["filled"], column["name"]
    found = []

    def add(severity: str, issue: str, mask: np.ndarray | None = None, count: int | None = None,
            share_of: int | None = None, note: str = "") -> None:
        count = int(mask.sum()) if count is None else count
        if count:
            base = share_of or nn or 1
            found.append({"severity": severity, "column": name, "issue": issue, "count": count,
                          "pct": 100 * count / base,
                          "examples": note or (examples(column["text"], mask, row_numbers) if mask is not None else "")})

    if nn == 0:
        add("High", "Column is completely empty", count=column["rows"], share_of=column["rows"], note="no values at all")
        return found
    add(missing_severity(column["missing_pct"]), "Missing values", count=column["missing"], share_of=column["rows"],
        note="rows " + ", ".join(str(r) for r in row_numbers[masks["missing"]][:5]) +
             (" ..." if column["missing"] > 5 else ""))
    add("Low", "Placeholder text used instead of a blank (counted as missing)", count=int(masks["placeholder"].sum()),
        share_of=column["rows"], note="rows " + ", ".join(str(r) for r in row_numbers[masks["placeholder"]][:5]))
    add("High", f"Values that do not fit the column type ({column['mixed_note'] or column['base_type']})", masks["mixed"])
    add("Medium", "Numbers stored as text", masks["number_as_text"])
    add("Medium", f"Dates stored as text or not in the column's usual format ({column['main_date_format']})",
        masks["date_format"])
    add("Medium", f"Values that break the usual pattern {column['pattern']}", masks["pattern"])
    if column["identifier"]:
        add("High", "Identifier-like column has repeated values", masks["duplicate"])
    add("Medium", "Same value written with different capital letters", masks["case"])
    add("Low", "Leading, trailing or double spaces", masks["whitespace"],
        note=examples(column["raw"], masks["whitespace"], row_numbers))
    if masks["outlier"].any():
        add("Low", f"Outliers by the IQR rule (outside {number(column['fence_low'])} to {number(column['fence_high'])})",
            masks["outlier"])
    if column["constant"]:
        add("Low", "Column holds a single value on every row", count=nn, note=f"'{column['top_values'][0][0]}'")
    return found


def profile_table(frame: pd.DataFrame, opts: Options | None = None, from_excel: bool = False,
                  key: list[str] | None = None, rules: list | None = None, rules_dir: Path = HERE) -> dict:
    """Profile a whole table and (optionally) check it against rules. Returns plain dicts and lists."""
    opts = opts or Options()
    key = list(key or [])
    blank_rows = frame.map(is_blank).all(axis=1) if len(frame.columns) else pd.Series(False, index=frame.index)
    data = frame[~blank_rows]                       # blank rows are reported once, not counted in every column
    row_numbers = np.asarray(data.index) + 2        # row 1 of the file is the header
    as_text = data.map(lambda value: "" if is_blank(value) else show(value))
    duplicate_rows = as_text.duplicated(keep="first").to_numpy() if len(data) else np.zeros(0, dtype=bool)

    columns = {name: profile_column(name, data[name].tolist(), opts, from_excel) for name in data.columns}
    issues: list[dict] = []

    def table_issue(severity: str, issue: str, count: int, total: int, note: str) -> None:
        if count:
            issues.append({"severity": severity, "column": "(table)", "issue": issue, "count": count,
                           "pct": 100 * count / total if total else 0.0, "examples": note})

    def rows_note(mask: np.ndarray) -> str:
        return "rows " + ", ".join(str(r) for r in row_numbers[mask][:5]) + (" ..." if mask.sum() > 5 else "")

    # ---- key
    key_duplicates = np.zeros(len(data), dtype=bool)
    missing_keys = [name for name in key if name not in columns]
    if missing_keys:
        table_issue("High", f"Key column not found in the file: {', '.join(missing_keys)}", 1, 1, "key check not run")
    elif len(key) == 1:
        columns[key[0]]["uniqueness_applies"] = True
        key_duplicates = columns[key[0]]["masks"]["duplicate"]
    elif key:
        has_key = np.all([columns[name]["kind"] != "missing" for name in key], axis=0)
        combined = pd.Series(["\x1f".join(parts) for parts in zip(*(columns[name]["text"] for name in key))])
        key_duplicates = (combined[has_key].duplicated(keep="first").reindex(combined.index, fill_value=False)).to_numpy()

    # ---- rules, then scores (rule failures count towards validity)
    rule_results = [run_rule(rule, columns, row_numbers, opts, rules_dir) for rule in (rules or [])]
    for column in columns.values():
        score_column(column)
        issues.extend(column_issues(column, row_numbers, opts))
    for result in rule_results:
        if result["status"] == "Fail":
            issues.append({"severity": result["severity"], "column": result["column"],
                           "issue": f"Rule failed: {result['name']}", "count": result["failed"],
                           "pct": result["pct"], "examples": result["examples"]})
        elif result["status"] == "Not run":
            issues.append({"severity": "High", "column": result["column"] or "(rules)",
                           "issue": f"Rule could not be run: {result['name'] or 'rule'}", "count": 1, "pct": 0.0,
                           "examples": result["reason"]})

    table_issue("High", "Exact duplicate rows (every column the same as an earlier row)",
                int(duplicate_rows.sum()), len(data), rows_note(duplicate_rows) if duplicate_rows.any() else "")
    table_issue("High", f"Duplicate key ({' + '.join(key)}): rows repeating an earlier key",
                int(key_duplicates.sum()), len(data), rows_note(key_duplicates) if key_duplicates.any() else "")
    table_issue("Low", "Completely blank rows (left out of the column figures)", int(blank_rows.sum()), len(frame),
                "rows " + ", ".join(str(i + 2) for i in frame.index[blank_rows][:5]))
    issues.sort(key=lambda issue: (SEVERITY_ORDER[issue["severity"]], -issue["count"], issue["column"]))

    # ---- overall score
    scored = list(columns.values())
    parts = {}
    for part in WEIGHTS:
        values = [column["parts"][part] for column in scored if column["parts"][part] is not None]
        parts[part] = float(np.mean(values)) if values else None
    duplicate_share = duplicate_rows.sum() / len(data) if len(data) else 0.0
    overall = float(np.mean([column["score"] for column in scored])) * (1 - duplicate_share) if scored else 0.0

    # ---- every failing cell, for the Failed Rows sheet
    failed = []
    labels = {"mixed": "Does not fit the column type", "number_as_text": "Number stored as text",
              "date_format": "Date stored as text or in another format", "pattern": "Breaks the usual pattern",
              "case": "Different capital letters", "whitespace": "Stray spaces", "outlier": "Outlier (IQR rule)",
              "placeholder": "Placeholder used for a blank"}
    for column in scored:
        marks = dict(labels, duplicate="Repeated identifier") if column["identifier"] else labels
        for mask_name, label in marks.items():
            for position in np.flatnonzero(column["masks"][mask_name]):
                failed.append((int(row_numbers[position]), column["name"], str(data[column["name"]].iloc[position]),
                               "Profile", label))
    for result in rule_results:
        if result["mask"] is not None:
            for position in np.flatnonzero(result["mask"]):
                failed.append((int(row_numbers[position]), result["column"], columns[result["column"]]["text"][position],
                               "Rule", result["name"]))
    for position in np.flatnonzero(duplicate_rows):
        failed.append((int(row_numbers[position]), "(whole row)", "", "Profile", "Exact duplicate of an earlier row"))
    failed_rows = pd.DataFrame(failed, columns=["Row", "Column", "Value", "Source", "Check"])
    failed_rows = failed_rows.sort_values(["Row", "Column", "Check"], kind="stable").reset_index(drop=True)

    return {
        "table": {"rows": len(frame), "columns": len(frame.columns), "blank_rows": int(blank_rows.sum()),
                  "rows_profiled": len(data), "duplicate_rows": int(duplicate_rows.sum()),
                  "empty_columns": [c["name"] for c in scored if c["filled"] == 0],
                  "constant_columns": [c["name"] for c in scored if c["constant"]],
                  "key": key, "duplicate_keys": int(key_duplicates.sum()) if key and not missing_keys else None},
        "score": overall, "parts": parts, "columns": scored, "issues": issues, "rules": rule_results,
        "failed_rows": failed_rows, "options": opts,
    }


# --------------------------------------------------------------------------- reading files

def load_table(path: Path, sheet: str | None = None) -> tuple[pd.DataFrame, bool]:
    """Read a CSV or Excel file without letting pandas guess types. Returns (table, came from Excel)."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        frame = pd.read_excel(path, sheet_name=sheet if sheet is not None else 0, dtype=object)
        from_excel = True
    elif suffix in (".csv", ".txt", ".tsv"):
        frame = None
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                with open(path, encoding=encoding, newline="") as handle:
                    header = handle.readline()
                counts = {delimiter: header.count(delimiter) for delimiter in ",;\t|"}
                delimiter = max(counts, key=counts.get) if any(counts.values()) else ","
                frame = pd.read_csv(path, dtype=str, keep_default_na=False, na_filter=False, sep=delimiter,
                                    encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if frame is None:
            raise ValueError("the file is not UTF-8 or Windows-1252 text")
        from_excel = False
    else:
        raise ValueError(f"unsupported file type '{suffix}' (use .csv, .txt, .tsv, .xlsx or .xlsm)")
    frame.columns = [str(name).strip() for name in frame.columns]
    return frame.reset_index(drop=True), from_excel


def load_rules(path: Path) -> tuple[list, list[str], date | None]:
    """Read the rules file: (rules, key columns, as-of date or None)."""
    config = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if isinstance(config, list):
        config = {"rules": config}
    if not isinstance(config, dict) or not isinstance(config.get("rules", []), list):
        raise ValueError("the rules file must contain a list called 'rules'")
    key = config.get("key") or []
    key = [key] if isinstance(key, str) else [str(name) for name in key]
    as_of = config.get("as_of")
    if as_of is not None and not isinstance(as_of, date):
        as_of = date.fromisoformat(str(as_of))
    return config.get("rules") or [], key, as_of


# --------------------------------------------------------------------------- output

def number(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, np.integer)) or float(value).is_integer():
        return f"{int(value):,}"
    return f"{value:,.2f}"


def column_view(column: dict) -> dict:
    """One column as a row of plain text and numbers, shared by the HTML, Excel and console output."""
    base = column["base_type"]
    if "mean" in column:
        detail = (f"{number(column['min'])} to {number(column['max'])}; mean {number(round(column['mean'], 2))}, "
                  f"median {number(column['median'])}")
    elif base == "date":
        detail = f"{column['min']} to {column['max']}"
    elif "length_min" in column:
        detail = f"length {column['length_min']} to {column['length_max']}"
    else:
        detail = ""
    kind = column["type"]
    if column["identifier"]:
        kind += " (unique)" if column["unique"] else f" ({int(column['masks']['duplicate'].sum())} repeats)"
    masks = column["masks"]
    flags =[f"{int(masks[m].sum())} {label}" for label, m in
             (("mixed types", "mixed"), ("numbers as text", "number_as_text"), ("date formats", "date_format"),
              ("pattern breaks", "pattern"), ("capital letters", "case"), ("spaces", "whitespace"),
              ("outliers", "outlier"), ("rule failures", "rule")) if masks[m].any()]
    if column["constant"]:
        flags.append("constant")
    return {
        "name": column["name"], "type": kind, "rows": column["rows"], "missing": column["missing"],
        "missing_pct": round(column["missing_pct"], 2), "distinct": column["distinct"], "detail": detail,
        "pattern": f"{column['pattern']} ({column['pattern_share']:.0%})" if column["pattern"] else "",
        "top": ", ".join(f"{value} ({count:,})" for value, count in column["top_values"][:3]),
        "flags": ", ".join(flags), "score": round(column["score"], 1),
        "parts": {part: (None if value is None else round(value, 1)) for part, value in column["parts"].items()},
        "outliers": int(masks["outlier"].sum()),
    }


def summary_rows(result: dict, source: str) -> list[tuple[str, str]]:
    table, parts = result["table"], result["parts"]
    rows = [("File", source), ("Rows in the file", f"{table['rows']:,}"), ("Columns", f"{table['columns']:,}"),
            ("Completely blank rows", f"{table['blank_rows']:,}"),
            ("Exact duplicate rows", f"{table['duplicate_rows']:,}"),
            ("Completely empty columns", ", ".join(table["empty_columns"]) or "none"),
            ("Constant columns", ", ".join(table["constant_columns"]) or "none")]
    if table["key"]:
        rows.append((f"Duplicate keys ({' + '.join(table['key'])})",
                     "not checked" if table["duplicate_keys"] is None else f"{table['duplicate_keys']:,}"))
    rows.append(("Overall quality score", f"{result['score']:.1f} / 100"))
    rows += [(part.title(), "not applicable" if value is None else f"{value:.1f}") for part, value in parts.items()]
    return rows


def write_excel(result: dict, source: str, path: Path) -> None:
    views = [column_view(column) for column in result["columns"]]
    columns = pd.DataFrame([{
        "Column": v["name"], "Type": v["type"], "Rows": v["rows"], "Missing": v["missing"],
        "Missing %": v["missing_pct"], "Distinct": v["distinct"], "Range / length": v["detail"],
        "Pattern": v["pattern"], "Most common values": v["top"], "Findings": v["flags"], "Score": v["score"],
        **{part.title(): v["parts"][part] for part in WEIGHTS}} for v in views])
    issues = pd.DataFrame([{"Severity": i["severity"], "Column": i["column"], "Issue": i["issue"], "Rows": i["count"],
                            "% of values": round(i["pct"], 2), "Examples": i["examples"]} for i in result["issues"]],
                          columns=["Severity", "Column", "Issue", "Rows", "% of values", "Examples"])
    rules = pd.DataFrame([{"Rule": r["name"], "Column": r["column"], "Check": r["check"], "Status": r["status"],
                           "Checked": r["checked"], "Failed": r["failed"], "% failed": round(r["pct"], 2),
                           "Could not be checked": r["unchecked"], "Examples": r["examples"], "Reason": r["reason"]}
                          for r in result["rules"]],
                         columns=["Rule", "Column", "Check", "Status", "Checked", "Failed", "% failed",
                                  "Could not be checked", "Examples", "Reason"])
    sheets = {"Summary": pd.DataFrame(summary_rows(result, source), columns=["Measure", "Value"]),
              "Columns": columns, "Issues": issues, "Rules": rules,
              "Failed Rows": result["failed_rows"].head(MAX_FAILED_ROWS)}
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, sheet_data in sheets.items():
            sheet_data.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", start_color="1F4E78")
            sheet.freeze_panes = "A2"
            for cells in sheet.columns:
                longest = max(len(str(cell.value)) if cell.value is not None else 0 for cell in cells)
                sheet.column_dimensions[cells[0].column_letter].width = min(longest + 2, 60)


def write_html(result: dict, source: str, path: Path) -> None:
    environment = Environment(loader=FileSystemLoader(HERE / "templates"), autoescape=select_autoescape(["html"]),
                              trim_blocks=True, lstrip_blocks=True)
    table = result["table"]
    counts = {level: sum(1 for issue in result["issues"] if issue["severity"] == level) for level in SEVERITY_ORDER}
    html = environment.get_template("report.html").render(
        source=source, generated=datetime.now().strftime("%d %b %Y %H:%M"), table=table, score=result["score"],
        parts=result["parts"], weights=WEIGHTS, columns=[column_view(c) for c in result["columns"]],
        issues=result["issues"], issue_counts=counts, rules=result["rules"], options=result["options"],
        failed_rows=len(result["failed_rows"]))
    path.write_text(html, encoding="utf-8")


def finish_chart(fig, ax, title: str, xlabel: str, path: Path) -> Path:
    ax.set_xlabel(xlabel, color=GREY)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(colors=GREY)
    fig.suptitle(title, x=0.015, ha="left", fontsize=12, fontweight="bold", color=NAVY)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def write_charts(result: dict, folder: Path) -> list[Path]:
    """Three charts: score per column, problem values per column, rows failing each rule (if rules were given)."""
    written = []
    views = sorted((column_view(c) for c in result["columns"]), key=lambda v: v["score"])
    if views:
        below = sum(1 for v in views if v["score"] < 90)
        fig, ax = plt.subplots(figsize=(9, max(2.5, 0.42 * len(views) + 1.4)))
        bars = ax.barh([v["name"] for v in views], [v["score"] for v in views],
                       color=[AMBER if v["score"] < 90 else NAVY for v in views])
        for bar, view in zip(bars, views):
            ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f"{view['score']:.1f}",
                    va="center", fontsize=9, color=GREY)
        ax.set_xlim(0, 108)
        written.append(finish_chart(
            fig, ax, f"{below} of {len(views)} columns score below 90; the file scores {result['score']:.1f} overall",
            "Quality score (0 to 100)", folder / "column_scores.png"))

    # problem values per column, split by the part of the score they lower (empty columns are left out)
    groups = {"Missing": (LIGHT_BLUE, ["missing"]), "Not valid": (AMBER, ["mixed", "pattern", "rule"]),
              "Not consistent": (NAVY, ["whitespace", "case", "number_as_text", "date_format"]),
              "Repeated": (GREY, ["duplicate"])}
    rows = []
    for column in result["columns"]:
        if not column["filled"]:
            continue
        counts = {}
        for group, (_, names) in groups.items():
            mask = np.any([column["masks"][name] for name in names], axis=0)
            counts[group] = int(mask.sum()) if group != "Repeated" or column["uniqueness_applies"] else 0
        if sum(counts.values()):
            rows.append((column["name"], counts))
    if rows:
        rows.sort(key=lambda row: sum(row[1].values()))
        clean = sum(1 for column in result["columns"] if column["filled"]) - len(rows)
        fig, ax = plt.subplots(figsize=(9, max(2.5, 0.42 * len(rows) + 1.6)))
        left = np.zeros(len(rows))
        for group, (colour, _) in groups.items():
            widths = np.array([counts[group] for _, counts in rows])
            if widths.any():
                ax.barh([name for name, _ in rows], widths, left=left, color=colour, label=group)
                left += widths
        for position, total in enumerate(left):
            ax.text(total + left.max() * 0.012, position, f"{int(total):,}", va="center", fontsize=9, color=GREY)
        ax.set_xlim(0, left.max() * 1.1)
        ax.legend(frameon=False, loc="lower right", fontsize=9)
        worst, worst_counts = rows[-1]
        written.append(finish_chart(
            fig, ax, f"{worst} has the most problem values ({sum(worst_counts.values()):,}); "
                     f"{clean} filled column{'s have' if clean != 1 else ' has'} none",
            "Values with a problem (a value can fall in more than one group)", folder / "problem_values.png"))

    ran = sorted((r for r in result["rules"] if r["status"] != "Not run"), key=lambda r: r["failed"])
    if ran:
        failing = sum(1 for r in ran if r["failed"])
        fig, ax = plt.subplots(figsize=(9, max(2.5, 0.45 * len(ran) + 1.4)))
        bars = ax.barh([r["name"] for r in ran], [r["failed"] for r in ran],
                       color=[AMBER if r["failed"] else LIGHT_BLUE for r in ran])
        top = max(r["failed"] for r in ran) or 1
        for bar, rule in zip(bars, ran):
            ax.text(bar.get_width() + top * 0.015, bar.get_y() + bar.get_height() / 2,
                    f"{rule['failed']:,}" if rule["failed"] else "pass", va="center", fontsize=9, color=GREY)
        ax.set_xlim(0, top * 1.12)
        written.append(finish_chart(fig, ax, f"{failing} of {len(ran)} business rules are broken by at least one row",
                                    "Rows failing the rule", folder / "rule_results.png"))
    return written


def print_summary(result: dict, source: str) -> None:
    table, parts = result["table"], result["parts"]
    print(f"Data quality profile: {source}")
    print(f"Rows: {table['rows']:,}   Columns: {table['columns']}   Blank rows: {table['blank_rows']}   "
          f"Exact duplicate rows: {table['duplicate_rows']}"
          + (f"   Duplicate keys ({' + '.join(table['key'])}): {table['duplicate_keys']}"
             if table["duplicate_keys"] is not None else ""))
    print(f"Overall quality score: {result['score']:.1f} / 100   ("
          + ", ".join(f"{part} {value:.1f}" for part, value in parts.items() if value is not None) + ")")
    print()
    views = [column_view(column) for column in result["columns"]]
    width = max([len(v["name"]) for v in views] + [6])
    print(f"{'Column':<{width}}  {'Type':<24} {'Missing %':>9} {'Distinct':>9} {'Score':>6}")
    for v in views:
        print(f"{v['name']:<{width}}  {v['type']:<24} {v['missing_pct']:>9.2f} {v['distinct']:>9,} {v['score']:>6.1f}")
    print()
    counts = {level: sum(1 for issue in result["issues"] if issue["severity"] == level) for level in SEVERITY_ORDER}
    print(f"Issues found: {len(result['issues'])}   (" + ", ".join(f"{level} {n}" for level, n in counts.items()) + ")")
    for issue in result["issues"]:
        print(f"  [{issue['severity']:<6}] {issue['column']}: {issue['issue']} - {issue['count']:,} "
              f"({issue['pct']:.2f}%)")
    if result["rules"]:
        print()
        status = [rule["status"] for rule in result["rules"]]
        print(f"Rules: {len(status)} given, {status.count('Pass')} passed, {status.count('Fail')} failed, "
              f"{status.count('Not run')} could not be run")
        for rule in result["rules"]:
            extra = f", {rule['unchecked']} could not be checked" if rule["unchecked"] else ""
            line = f"{rule['failed']:,} of {rule['checked']:,} rows fail{extra}" if rule["status"] != "Not run" \
                else rule["reason"]
            print(f"  [{rule['status']:<7}] {rule['name']}: {line}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Profile the quality of any CSV or Excel file.")
    parser.add_argument("file", type=Path, nargs="?", help="CSV or Excel file (default: the bundled sample)")
    parser.add_argument("--rules", type=Path, help="YAML file of business rules (optional)")
    parser.add_argument("--key", help="column(s) that should identify a row, comma separated")
    parser.add_argument("--sheet", help="Excel sheet name (default: the first sheet)")
    parser.add_argument("--out", type=Path, default=Path("output"), help="folder for the report files")
    parser.add_argument("--charts-dir", type=Path, help="folder for the PNG charts (default: same as --out)")
    parser.add_argument("--ignore", default="", help="columns to leave out, comma separated")
    parser.add_argument("--as-of", type=date.fromisoformat, help="'today' for not_in_future, as YYYY-MM-DD")
    parser.add_argument("--date-order", choices=["dmy", "mdy"], default="dmy", help="how to read 03/04/2026")
    parser.add_argument("--type-threshold", type=float, default=0.90,
                        help="share of values that must agree before a column type is assigned")
    parser.add_argument("--pattern-threshold", type=float, default=0.90,
                        help="share of values that must follow one shape before the others are flagged")
    parser.add_argument("--iqr-k", type=float, default=1.5, help="outlier fence multiplier for the IQR rule")
    parser.add_argument("--top", type=int, default=5, help="most common values kept per column")
    args = parser.parse_args(argv)

    if args.file is None:                            # no file given: run the bundled sample with its rules
        args.file = SAMPLE_FILE
        args.rules = args.rules or SAMPLE_RULES
    started = time.perf_counter()
    if not args.file.is_file():
        print(f"File not found: {args.file}", file=sys.stderr)
        return 1
    try:
        frame, from_excel = load_table(args.file, args.sheet)
    except Exception as error:                       # unreadable input is reported, not guessed at
        print(f"Could not read {args.file.name}: {error}", file=sys.stderr)
        return 1
    if frame.empty or not len(frame.columns):
        print(f"{args.file.name} has no data rows to profile.", file=sys.stderr)
        return 1

    rules, key, as_of = [], [], None
    if args.rules:
        try:
            rules, key, as_of = load_rules(args.rules)
        except (OSError, ValueError, yaml.YAMLError) as error:
            print(f"Could not read the rules file {args.rules}: {error}", file=sys.stderr)
            return 1
    if args.key:
        key = [name.strip() for name in args.key.split(",") if name.strip()]
    ignore = [name.strip() for name in args.ignore.split(",") if name.strip()]
    unknown = [name for name in ignore if name not in frame.columns]
    if unknown:
        print(f"Note: --ignore names columns that are not in the file: {', '.join(unknown)}")
    frame = frame.drop(columns=[name for name in ignore if name in frame.columns])

    opts = Options(type_threshold=args.type_threshold, pattern_threshold=args.pattern_threshold, iqr_k=args.iqr_k,
                   date_order=args.date_order, top=args.top, as_of=args.as_of or as_of or date.today())
    result = profile_table(frame, opts, from_excel, key, rules, args.rules.parent if args.rules else HERE)

    args.out.mkdir(parents=True, exist_ok=True)
    charts_dir = args.charts_dir or args.out
    charts_dir.mkdir(parents=True, exist_ok=True)
    html_path = args.out / f"{args.file.stem}_profile.html"
    excel_path = args.out / f"{args.file.stem}_profile.xlsx"
    write_html(result, args.file.name, html_path)
    write_excel(result, args.file.name, excel_path)
    charts = write_charts(result, charts_dir)

    print_summary(result, args.file.name)
    print()
    if len(result["failed_rows"]) > MAX_FAILED_ROWS:
        print(f"Note: the Failed Rows sheet shows the first {MAX_FAILED_ROWS:,} of {len(result['failed_rows']):,} entries")
    print(f"HTML report:  {html_path.as_posix()}")
    print(f"Excel file:   {excel_path.as_posix()}   ({min(len(result['failed_rows']), MAX_FAILED_ROWS):,} failed-row entries)")
    print(f"Charts:       {', '.join(chart.as_posix() for chart in charts)}")
    print(f"Time taken:   {time.perf_counter() - started:.1f} sec")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
