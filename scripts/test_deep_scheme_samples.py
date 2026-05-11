from __future__ import annotations

import json
import re
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from core.activity_classifier import classify_activity_types
from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import build_quote_rows, extract_activity_sections
from generate_deep_scheme_samples import METADATA_PATH, SAMPLE_DIR, main as generate_samples


RULES_PATH = QUOTE_APP_DIR / "data" / "rules_config.json"
PRICE_DB_PATH = QUOTE_APP_DIR / "data" / "price_db.xlsx"
MIN_SCHEME_CHARS = 3000
MIN_ACTIVITY_CONTENT_CHARS = 100


def _main_names(sections: list[dict]) -> list[str]:
    return [str(section.get("name", "")) for section in sections if str(section.get("section_level", "")) == "main"]


def _assert_contains(label: str, actual: list[str], expected: list[str]) -> None:
    missing = [name for name in expected if name not in actual]
    if missing:
        raise AssertionError(f"{label} missing main sections: {missing}; actual={actual}")


def _assert_not_contains(label: str, actual: list[str], forbidden: list[str]) -> None:
    leaked = [name for name in forbidden if name in actual]
    if leaked:
        raise AssertionError(f"{label} leaked forbidden sections: {leaked}; actual={actual}")


def _activity_content_lengths(text: str) -> list[int]:
    return [len(match.strip()) for match in re.findall(r"（3）内容：(.+)", text)]


def main() -> None:
    generate_samples()
    metadata = json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    rules = load_rules(RULES_PATH)

    total_quote_rows = 0
    total_unassigned = 0
    total_activity_descriptions = 0

    for case in metadata:
        path = SAMPLE_DIR / str(case["file"])
        text = path.read_text(encoding="utf-8")
        compact_length = len("".join(text.split()))
        if compact_length < MIN_SCHEME_CHARS:
            raise AssertionError(f"{path.name} effective length too short: {compact_length}")

        content_lengths = _activity_content_lengths(text)
        if not content_lengths:
            raise AssertionError(f"{path.name} has no activity content descriptions")
        short_lengths = [length for length in content_lengths if length < MIN_ACTIVITY_CONTENT_CHARS]
        if short_lengths:
            raise AssertionError(f"{path.name} has short activity descriptions: {short_lengths}")

        extracted_rows = extract_quote_items(text, rules)
        sections = extract_activity_sections(text, extracted_rows)
        quote_rows = build_quote_rows(extracted_rows, PRICE_DB_PATH, text, activity_sections=sections)
        main_names = _main_names(sections)
        unassigned_count = sum(1 for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块")
        priced_count = sum(1 for row in quote_rows if float(row.get("单价") or 0) != 0)
        activity_types = [row.get("活动类型", "") for row in classify_activity_types(text)]

        _assert_contains(path.name, main_names, list(case["expected_main"]))
        _assert_not_contains(path.name, main_names, list(case["forbidden_main"]))
        if not quote_rows:
            raise AssertionError(f"{path.name} produced no quote rows")
        if unassigned_count:
            raise AssertionError(f"{path.name} has unassigned quote rows: {unassigned_count}")
        if priced_count:
            raise AssertionError(f"{path.name} filled real unit prices: {priced_count}")

        total_quote_rows += len(quote_rows)
        total_unassigned += unassigned_count
        total_activity_descriptions += len(content_lengths)
        print(
            f"PASS {path.name}: chars={compact_length} descriptions={len(content_lengths)} "
            f"min_desc={min(content_lengths)} main={len(main_names)} quote_rows={len(quote_rows)} "
            f"unassigned={unassigned_count} types={'/'.join(activity_types) or '无'}"
        )

    print(
        f"ALL DEEP SCHEME SAMPLE TESTS PASSED: cases={len(metadata)} "
        f"activity_descriptions={total_activity_descriptions} quote_rows={total_quote_rows} "
        f"unassigned={total_unassigned}"
    )


if __name__ == "__main__":
    main()
