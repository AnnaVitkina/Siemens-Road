"""
Transform Zones and Address Zones dataframes into a zones import text file.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd

from project_paths import OUTPUT_DIR, PROCESSING_DIR

OUTPUT_COLUMNS = ["Zone name", "country", "postal code", "exclude"]

GERMAN_CHAR_VARIANTS: dict[str, list[str]] = {
    "ß": ["ß", "ss", "s", ""],
    "ü": ["ü", "u", "ue", ""],
    "Ü": ["Ü", "U", "Ue", ""],
    "ö": ["ö", "o", "oe", ""],
    "Ö": ["Ö", "O", "Oe", ""],
    "ä": ["ä", "a", "ae", ""],
    "Ä": ["Ä", "A", "Ae", ""],
}


def _normalize_zip(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _normalize_zone_value(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)):
        numeric = float(value)
        if numeric == int(numeric):
            return str(int(numeric))
    return str(value).strip()


def _format_zone_name(iso: object, zone: object) -> str:
    iso_text = "" if pd.isna(iso) else str(iso).strip()
    zone_text = _normalize_zone_value(zone)
    return f"{iso_text} Zone {zone_text}"


def _transliterate_to_ascii(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(character for character in normalized if not unicodedata.combining(character))


def _generate_german_variations(text: str) -> set[str]:
    if not text:
        return {""}

    def expand(start: int) -> set[str]:
        if start >= len(text):
            return {""}

        character = text[start]
        if character not in GERMAN_CHAR_VARIANTS:
            rest_variants = expand(start + 1)
            return {character + suffix for suffix in rest_variants}

        results: set[str] = set()
        for replacement in GERMAN_CHAR_VARIANTS[character]:
            for suffix in expand(start + 1):
                results.add(replacement + suffix)
        return results

    return {_transliterate_to_ascii(value) for value in expand(0)}


STREET_SUFFIX_VARIANTS = ["str", "strae", "strase", "strasse"]

STREET_SUFFIX_PATTERNS: tuple[tuple[re.Pattern[str], bool], ...] = (
    (re.compile(r"(?i)(?<=[a-zA-Z])straße"), True),
    (re.compile(r"(?i)(?<=[a-zA-Z])strasse"), False),
    (re.compile(r"(?i)(?<=[a-zA-Z])strase"), False),
    (re.compile(r"(?i)(?<=[a-zA-Z])strae"), False),
    (re.compile(r"(?i)(?<=[a-zA-Z])str(?=\d|$)"), False),
)


def _normalize_street_text(text: str) -> str:
    normalized_chars: list[str] = []
    lowercase_next = False

    for character in text:
        if character == "-":
            lowercase_next = True
            continue
        if character in " \t.,;:/\\|":
            continue
        if lowercase_next and character.isalpha():
            normalized_chars.append(character.lower())
            lowercase_next = False
            continue
        normalized_chars.append(character)
        lowercase_next = False

    return re.sub(r"[^0-9A-Za-zß]", "", "".join(normalized_chars))


def _find_street_suffix_match(text: str) -> tuple[re.Match[str], bool] | None:
    for pattern, includes_eszett_variant in STREET_SUFFIX_PATTERNS:
        match = pattern.search(text)
        if match is not None:
            return match, includes_eszett_variant
    return None


def _expand_street_suffix_variations(text: str) -> set[str]:
    match_info = _find_street_suffix_match(text)
    if match_info is None:
        return {text}

    match, includes_eszett_variant = match_info
    prefix = text[: match.start()]
    suffix = text[match.end() :]

    results = {prefix + variant + suffix for variant in STREET_SUFFIX_VARIANTS}
    if includes_eszett_variant:
        results.add(prefix + "traße" + suffix)

    return results


def _generate_street_variations(street_text: str) -> set[str]:
    results: set[str] = set()
    for german_variant in _generate_german_variations(street_text):
        normalized = _normalize_street_text(german_variant)
        if not normalized:
            continue
        results.update(_expand_street_suffix_variations(normalized))
    return results


def _join_values(values: set[str]) -> str:
    return ", ".join(sorted(value for value in values if value))


def _build_address_postal_codes(zip_value: object, street_value: object) -> set[str]:
    zip_text = _normalize_zip(zip_value)
    street_text = "" if pd.isna(street_value) else str(street_value).strip()
    street_variations = _generate_street_variations(street_text)

    if not zip_text:
        return set()

    if not street_variations:
        return {zip_text}

    return {f"{zip_text}_{street}" for street in street_variations}


def _build_address_lookup(address_zones_df: pd.DataFrame) -> dict[str, list[set[str]]]:
    lookup: dict[str, list[set[str]]] = {}

    for _, row in address_zones_df.iterrows():
        iso = "" if pd.isna(row["ISO"]) else str(row["ISO"]).strip()
        if not iso:
            continue
        postal_codes = _build_address_postal_codes(row["ZIP"], row["Street"])
        if not postal_codes:
            continue
        lookup.setdefault(iso, []).append(postal_codes)

    return lookup


def _compute_exclude(
    iso: str,
    zip_prefixes: list[str],
    address_lookup: dict[str, list[set[str]]],
) -> str:
    excluded: set[str] = set()
    for postal_codes in address_lookup.get(iso, []):
        sample = next(iter(postal_codes))
        zip_part = sample.split("_", 1)[0]
        if any(zip_part.startswith(prefix) for prefix in zip_prefixes):
            excluded.update(postal_codes)

    return _join_values(excluded)


def _process_corporate_zones(
    zones_df: pd.DataFrame,
    address_lookup: dict[str, list[set[str]]],
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    grouped = zones_df.groupby(["ISO", "Zone"], sort=False, dropna=False)

    for (iso, zone), group in grouped:
        iso_text = "" if pd.isna(iso) else str(iso).strip()
        zone_name = _format_zone_name(iso, zone)
        zip_prefixes = sorted(
            {
                _normalize_zip(value)
                for value in group["ZIP"].tolist()
                if _normalize_zip(value)
            },
            key=lambda value: (len(value), value),
        )

        rows.append(
            {
                "Zone name": zone_name,
                "country": iso_text,
                "postal code": ", ".join(zip_prefixes),
                "exclude": _compute_exclude(iso_text, zip_prefixes, address_lookup),
            }
        )

    return rows


def _process_address_zones(address_zones_df: pd.DataFrame) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    grouped_postal_codes: dict[tuple[str, str, str, str], set[str]] = {}

    for _, row in address_zones_df.iterrows():
        iso_text = "" if pd.isna(row["ISO"]) else str(row["ISO"]).strip()
        zone_text = _normalize_zone_value(row["Zone"])
        zip_text = _normalize_zip(row["ZIP"])
        street_text = "" if pd.isna(row["Street"]) else str(row["Street"]).strip()
        if not iso_text or not zip_text:
            continue

        key = (iso_text, zone_text, zip_text, street_text)
        grouped_postal_codes.setdefault(key, set()).update(
            _build_address_postal_codes(row["ZIP"], row["Street"])
        )

    for (iso_text, zone_text, _zip_text, _street_text), postal_codes in grouped_postal_codes.items():
        rows.append(
            {
                "Zone name": f"{iso_text} Zone {zone_text}",
                "country": iso_text,
                "postal code": _join_values(postal_codes),
                "exclude": "",
            }
        )

    return rows


def process_zones_dataframes(
    zones_df: pd.DataFrame,
    address_zones_df: pd.DataFrame,
    used_zone_names: set[str] | None = None,
) -> tuple[pd.DataFrame, int, int]:
    address_lookup = _build_address_lookup(address_zones_df)
    corporate_rows = _process_corporate_zones(zones_df, address_lookup)
    address_rows = _process_address_zones(address_zones_df)
    processed = pd.DataFrame(corporate_rows + address_rows, columns=OUTPUT_COLUMNS)
    processed["_is_corporate"] = [True] * len(corporate_rows) + [False] * len(address_rows)

    if used_zone_names is not None:
        processed = processed[processed["Zone name"].isin(used_zone_names)].reset_index(drop=True)

    corporate_count = int(processed["_is_corporate"].sum()) if len(processed) else 0
    address_count = len(processed) - corporate_count
    processed = processed.drop(columns=["_is_corporate"])
    return processed, corporate_count, address_count


def _load_rate_card_zone_names(processing_file: Path) -> set[str]:
    from build_rate_card_matrix import extract_rate_card_zone_names, load_extracted_rate_card

    rate_card = load_extracted_rate_card(processing_file)
    return extract_rate_card_zone_names(rate_card)


def write_zones_txt(
    zones_df: pd.DataFrame,
    address_zones_df: pd.DataFrame,
    output_path: Path,
    used_zone_names: set[str] | None = None,
    rate_card_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, int, int]:
    if used_zone_names is None and rate_card_df is not None:
        from build_rate_card_matrix import extract_rate_card_zone_names

        used_zone_names = extract_rate_card_zone_names(rate_card_df)

    processed, corporate_count, address_count = process_zones_dataframes(
        zones_df,
        address_zones_df,
        used_zone_names=used_zone_names,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8", newline="") as handle:
        handle.write("\t".join(OUTPUT_COLUMNS) + "\n")
        for _, row in processed.iterrows():
            handle.write("\t".join(str(row[column]) for column in OUTPUT_COLUMNS) + "\n")

    return processed, corporate_count, address_count


def load_zones_from_processing_file(processing_file: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    zones_df = pd.read_excel(processing_file, sheet_name="Zones", header=0)
    address_zones_df = pd.read_excel(processing_file, sheet_name="Address Zones", header=0)
    return zones_df, address_zones_df


def build_zones_from_processing_file(
    processing_file: Path,
    output_file: Path | None = None,
    used_zone_names: set[str] | None = None,
) -> tuple[pd.DataFrame, int, int]:
    zones_df, address_zones_df = load_zones_from_processing_file(processing_file)
    if used_zone_names is None:
        used_zone_names = _load_rate_card_zone_names(processing_file)
    if output_file is None:
        output_file = OUTPUT_DIR / f"{processing_file.stem}_zones.txt"
    return write_zones_txt(
        zones_df,
        address_zones_df,
        output_file,
        used_zone_names=used_zone_names,
    )


def list_processing_files() -> list[Path]:
    if not PROCESSING_DIR.exists():
        return []
    return sorted(PROCESSING_DIR.glob("*_extracted.xlsx"))


def run_interactive_zones_export() -> pd.DataFrame:
    processing_files = list_processing_files()
    if not processing_files:
        raise FileNotFoundError(f"No extracted files found in {PROCESSING_DIR}")

    print("Extracted files in processing folder:")
    for index, file_path in enumerate(processing_files, start=1):
        print(f"  {index}. {file_path.name}")

    while True:
        raw = input("Enter file number to export zones from: ").strip()
        if raw.isdigit():
            choice_index = int(raw) - 1
            if 0 <= choice_index < len(processing_files):
                selected_file = processing_files[choice_index]
                break
        print("Invalid choice. Try again.")

    output_file = OUTPUT_DIR / f"{selected_file.stem}_zones.txt"
    processed, corporate_count, address_count = build_zones_from_processing_file(
        selected_file,
        output_file,
    )
    print(f"\nSaved zones file to: {output_file}")
    print(f"Total rows: {len(processed)} (filtered to zones used in Rate Card)")
    print(f"Corporate zone rows: {corporate_count}")
    print(f"Address zone rows: {address_count}")
    return processed


if __name__ == "__main__":
    run_interactive_zones_export()
