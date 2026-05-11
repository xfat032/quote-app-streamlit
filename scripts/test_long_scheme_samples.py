from __future__ import annotations

import json
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
QUOTE_APP_DIR = ROOT_DIR / "quote_app"
if str(QUOTE_APP_DIR) not in sys.path:
    sys.path.insert(0, str(QUOTE_APP_DIR))

from core.extractor import extract_quote_items
from core.normalizer import load_rules
from core.quote_builder import build_quote_rows, extract_activity_sections


RULES_PATH = QUOTE_APP_DIR / "data" / "rules_config.json"
PRICE_DB_PATH = QUOTE_APP_DIR / "data" / "price_db.xlsx"
SAMPLE_DIR = ROOT_DIR / "scripts" / "samples" / "generated_long_regression"
METADATA_PATH = SAMPLE_DIR / "metadata.json"
MIN_EFFECTIVE_CHARS = 3000


def _names_by_level(sections: list[dict], level: str) -> list[str]:
    return [str(section.get("name", "")) for section in sections if str(section.get("section_level", "")) == level]


def _assert_contains(label: str, actual: list[str], expected: list[str]) -> None:
    missing = [name for name in expected if name not in actual]
    if missing:
        raise AssertionError(f"{label} missing: {missing}; actual={actual}")


def _assert_not_contains(label: str, actual: list[str], forbidden: list[str]) -> None:
    leaked = [name for name in forbidden if name in actual]
    if leaked:
        raise AssertionError(f"{label} leaked: {leaked}; actual={actual}")


def main() -> None:
    metadata = json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    rules = load_rules(RULES_PATH)
    total_quote_rows = 0
    total_unassigned = 0

    for case in metadata:
        path = SAMPLE_DIR / str(case["file"])
        text = path.read_text(encoding="utf-8")
        effective_chars = len("".join(text.split()))
        if effective_chars < MIN_EFFECTIVE_CHARS:
            raise AssertionError(f"{path.name} too short: {effective_chars}")

        extracted_rows = extract_quote_items(text, rules)
        sections = extract_activity_sections(text, extracted_rows)
        quote_rows = build_quote_rows(extracted_rows, PRICE_DB_PATH, text, activity_sections=sections)
        main_names = _names_by_level(sections, "main")
        sub_names = _names_by_level(sections, "sub")
        unassigned_count = sum(1 for row in quote_rows if str(row.get("quote_section", "")) == "未归属板块")

        _assert_contains(path.name, main_names, list(case["expected_main"]))
        _assert_not_contains(path.name, main_names, list(case["forbidden_main"]))
        if not quote_rows:
            raise AssertionError(f"{path.name} produced no quote rows")
        if unassigned_count:
            raise AssertionError(f"{path.name} has unassigned quote rows: {unassigned_count}")

        total_quote_rows += len(quote_rows)
        total_unassigned += unassigned_count
        print(
            f"PASS {path.name}: chars={effective_chars} main={len(main_names)} "
            f"sub={len(sub_names)} quote_rows={len(quote_rows)} unassigned={unassigned_count}"
        )

    print(
        f"ALL LONG SCHEME SAMPLE TESTS PASSED: cases={len(metadata)} "
        f"quote_rows={total_quote_rows} unassigned={total_unassigned}"
    )


if __name__ == "__main__":
    main()
