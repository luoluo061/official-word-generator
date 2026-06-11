from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from list_features import parse_feature_catalog  # noqa: E402
from profile_resolver import PROFILES_DIR, resolve_relative  # noqa: E402


VALID_STATUSES = {"production", "draft", "planned"}
REFERENCE_PROFILE_ID = "general_official"
CORE_FILES = (
    "profile.md",
    "features.json",
    "format_rules.json",
    "validation_rules.json",
)
REFERENCE_FILES = (
    "example.md",
    "expected_validation_report.md",
    "feature_selection.md",
)
FEATURE_LIST_KEYS = (
    "enabled_features",
    "disabled_features",
    "overridable_features",
    "planned_features",
)


def load_json(path: Path, errors: list[str]) -> dict[str, Any]:
    if not path.exists():
        errors.append(f"Missing JSON file: {path.name}")
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        errors.append(f"Invalid JSON in {path.name}: {exc}")
        return {}


def unique_duplicates(values: list[str]) -> list[str]:
    seen = set()
    duplicates = set()
    for value in values:
        if value in seen:
            duplicates.add(value)
        seen.add(value)
    return sorted(duplicates)


def feature_catalog_maps() -> tuple[set[str], dict[str, str]]:
    features, _warnings = parse_feature_catalog()
    feature_ids = {feature.feature_id for feature in features}
    statuses = {feature.feature_id: feature.status for feature in features}
    return feature_ids, statuses


def profile_status(features: dict[str, Any], format_rules: dict[str, Any], validation_rules: dict[str, Any]) -> str:
    return (
        features.get("status")
        or format_rules.get("status")
        or validation_rules.get("status")
        or "draft"
    )


def resolve_template(profile_dir: Path, features: dict[str, Any]) -> Path:
    template_value = features.get("template") or features.get("template_path")
    resolved = resolve_relative(profile_dir, template_value)
    if resolved:
        return resolved
    candidate = profile_dir / "template.docx"
    if candidate.exists():
        return candidate
    return profile_dir / "template_placeholder.md"


def validate_feature_references(
    features: dict[str, Any],
    catalog_ids: set[str],
    catalog_statuses: dict[str, str],
    profile_status: str,
    errors: list[str],
    warnings: list[str],
) -> None:
    feature_map = features.get("features", {})
    if not isinstance(feature_map, dict):
        errors.append("features.json field `features` must be an object.")
        feature_map = {}

    all_refs: list[str] = list(feature_map.keys())
    for key in FEATURE_LIST_KEYS:
        value = features.get(key)
        if value is None:
            warnings.append(f"features.json is missing `{key}`; scaffolded profiles should include it.")
            continue
        if not isinstance(value, list):
            errors.append(f"features.json field `{key}` must be a list.")
            continue
        for feature_id in unique_duplicates(value):
            warnings.append(f"Duplicate feature_id inside `{key}`: {feature_id}")
        all_refs.extend(value)

    unknown = sorted({feature_id for feature_id in all_refs if feature_id not in catalog_ids})
    for feature_id in unknown:
        errors.append(f"Unknown feature_id referenced in features.json: {feature_id}")

    enabled_features = set(features.get("enabled_features") or [])
    for feature_id, config in feature_map.items():
        if isinstance(config, dict) and config.get("enabled"):
            enabled_features.add(feature_id)

    planned_enabled = sorted(
        feature_id
        for feature_id in enabled_features
        if catalog_statuses.get(feature_id) == "planned"
    )
    for feature_id in planned_enabled:
        message = f"Planned feature is enabled: {feature_id}"
        if profile_status == "production":
            errors.append(message)
        else:
            warnings.append(message)


def validate_profile(profile_id: str) -> dict[str, Any]:
    blocking_errors: list[str] = []
    warnings: list[str] = []
    info: list[str] = []

    profile_dir = PROFILES_DIR / profile_id
    if not profile_dir.exists() or not profile_dir.is_dir():
        return {
            "profile_id": profile_id,
            "ok": False,
            "status": "missing",
            "blocking_errors": [f"Profile directory does not exist: {profile_dir}"],
            "warnings": warnings,
            "info": info,
        }

    for filename in CORE_FILES:
        if not (profile_dir / filename).exists():
            blocking_errors.append(f"Missing required file: {filename}")

    features = load_json(profile_dir / "features.json", blocking_errors)
    format_rules = load_json(profile_dir / "format_rules.json", blocking_errors)
    validation_rules = load_json(profile_dir / "validation_rules.json", blocking_errors)

    status = profile_status(features, format_rules, validation_rules)
    if status not in VALID_STATUSES:
        blocking_errors.append(f"Invalid profile status `{status}`; must be production, draft, or planned.")

    is_reference = bool(features.get("is_reference", profile_id == REFERENCE_PROFILE_ID))
    if is_reference and profile_id != REFERENCE_PROFILE_ID:
        blocking_errors.append("Only general_official may use is_reference=true.")
    if profile_id == REFERENCE_PROFILE_ID and not is_reference:
        warnings.append("general_official should keep is_reference=true or omit it for backward compatibility.")

    for filename in REFERENCE_FILES:
        path = profile_dir / filename
        if not path.exists():
            if status == "production":
                blocking_errors.append(f"Production profile missing required reference file: {filename}")
            else:
                warnings.append(f"Draft/planned profile missing recommended file: {filename}")

    catalog_ids, catalog_statuses = feature_catalog_maps()
    validate_feature_references(features, catalog_ids, catalog_statuses, status, blocking_errors, warnings)

    template_path = resolve_template(profile_dir, features)
    template_exists = template_path.exists()
    template_is_placeholder = template_path.name == "template_placeholder.md"
    template_is_docx = template_path.suffix.lower() == ".docx"

    if not template_exists:
        if status == "production":
            blocking_errors.append(f"Production profile template path does not exist: {template_path}")
        else:
            warnings.append(f"Template path does not exist yet: {template_path}")

    if status == "production":
        if template_is_placeholder:
            blocking_errors.append("Production profile must not use template_placeholder.md.")
        if not template_is_docx or not template_exists:
            blocking_errors.append("Production profile requires a real template.docx or valid template path.")

    enabled_statuses = {}
    feature_map = features.get("features", {}) if isinstance(features.get("features"), dict) else {}
    for feature_id, config in feature_map.items():
        if isinstance(config, dict) and config.get("enabled"):
            enabled_statuses[feature_id] = catalog_statuses.get(feature_id, "unknown")

    if status == "production":
        partial_or_planned = sorted(
            feature_id
            for feature_id, feature_status in enabled_statuses.items()
            if feature_status in {"partial", "planned"}
        )
        for feature_id in partial_or_planned:
            blocking_errors.append(f"Production profile must not enable partial/planned feature: {feature_id} ({enabled_statuses[feature_id]})")
    elif status == "draft":
        info.append("Draft profile is not allowed for formal delivery.")
    elif status == "planned":
        info.append("Planned profile is design-only and cannot generate formal deliverables.")

    allow_formal_delivery = status == "production" and not blocking_errors
    if status == "draft" and allow_formal_delivery:
        blocking_errors.append("Draft profile must not allow formal delivery.")

    return {
        "profile_id": profile_id,
        "status": status,
        "ok": not blocking_errors,
        "allow_formal_delivery": allow_formal_delivery,
        "template_path": str(template_path),
        "template_exists": template_exists,
        "template_is_docx": template_is_docx,
        "blocking_errors": blocking_errors,
        "warnings": warnings,
        "info": info,
    }


def validate_all_profiles() -> list[dict[str, Any]]:
    if not PROFILES_DIR.exists():
        return []
    return [
        validate_profile(path.name)
        for path in sorted(PROFILES_DIR.iterdir())
        if path.is_dir()
    ]


def render_result(result: dict[str, Any]) -> str:
    lines = [
        f"Profile: {result.get('profile_id')}",
        f"Status: {result.get('status')}",
        f"OK: {result.get('ok')}",
        f"Allow Formal Delivery: {result.get('allow_formal_delivery')}",
        f"Template: {result.get('template_path')}",
        "blocking_errors:",
    ]
    errors = result.get("blocking_errors") or []
    lines.extend(f"- {item}" for item in errors) if errors else lines.append("- none")
    lines.append("warnings:")
    warnings = result.get("warnings") or []
    lines.extend(f"- {item}" for item in warnings) if warnings else lines.append("- none")
    lines.append("info:")
    info = result.get("info") or []
    lines.extend(f"- {item}" for item in info) if info else lines.append("- none")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate profile configuration integrity.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--profile", help="Profile id to validate.")
    group.add_argument("--all", action="store_true", help="Validate all profiles.")
    parser.add_argument("--json", action="store_true", help="Output JSON.")
    args = parser.parse_args()

    if args.all:
        results = validate_all_profiles()
        if args.json:
            print(json.dumps(results, ensure_ascii=False, indent=2))
        else:
            for index, result in enumerate(results):
                if index:
                    print("")
                print(render_result(result))
        raise SystemExit(0 if all(result["ok"] for result in results) else 1)

    result = validate_profile(args.profile)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(render_result(result))
    raise SystemExit(0 if result["ok"] else 1)


if __name__ == "__main__":
    main()
