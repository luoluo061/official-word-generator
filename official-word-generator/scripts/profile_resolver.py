from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
PROFILES_DIR = SKILL_DIR / "profiles"
FEATURE_CATALOG = SKILL_DIR / "references" / "feature_catalog.md"


class ProfileError(ValueError):
    pass


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def read_feature_statuses(catalog_path: Path = FEATURE_CATALOG) -> dict[str, str]:
    if not catalog_path.exists():
        return {}
    statuses: dict[str, str] = {}
    pattern = re.compile(r"^\|\s*`([^`]+)`\s*\|.*\|\s*(implemented|partial|planned)\s*\|$")
    for line in catalog_path.read_text(encoding="utf-8").splitlines():
        match = pattern.match(line.strip())
        if match:
            statuses[match.group(1)] = match.group(2)
    return statuses


def resolve_relative(base: Path, value: str | None) -> Path | None:
    if not value:
        return None
    path = Path(value)
    if path.is_absolute():
        return path
    return (base / path).resolve()


def profile_dirs() -> list[Path]:
    if not PROFILES_DIR.exists():
        return []
    return sorted(path for path in PROFILES_DIR.iterdir() if path.is_dir())


def infer_profile_status(profile_dir: Path, features: dict[str, Any], format_rules: dict[str, Any], validation_rules: dict[str, Any], template_path: Path | None) -> str:
    explicit = (
        features.get("status")
        or format_rules.get("status")
        or validation_rules.get("status")
    )
    if explicit in {"production", "draft", "planned"}:
        if template_path and template_path.suffix.lower() == ".docx" and template_path.exists() and explicit == "draft":
            return "draft"
        return explicit
    if template_path and template_path.suffix.lower() == ".docx" and template_path.exists():
        return "production"
    if template_path and template_path.name == "template_placeholder.md":
        return "draft"
    return "planned"


def summarize_features(features: dict[str, Any], catalog_statuses: dict[str, str] | None = None) -> dict[str, list[str]]:
    catalog_statuses = catalog_statuses or read_feature_statuses()
    summary = {"implemented": [], "partial": [], "planned": [], "unknown": []}
    for feature_id, config in sorted(features.get("features", {}).items()):
        if not isinstance(config, dict) or not config.get("enabled"):
            continue
        status = catalog_statuses.get(feature_id, "unknown")
        summary.setdefault(status, []).append(feature_id)
    return summary


def resolve_profile(profile_id: str) -> dict[str, Any]:
    profile_dir = PROFILES_DIR / profile_id
    if not profile_dir.exists() or not profile_dir.is_dir():
        available = ", ".join(path.name for path in profile_dirs()) or "none"
        raise ProfileError(f"Unknown profile '{profile_id}'. Available profiles: {available}")

    profile_md = profile_dir / "profile.md"
    features_path = profile_dir / "features.json"
    format_rules_path = profile_dir / "format_rules.json"
    validation_rules_path = profile_dir / "validation_rules.json"
    placeholder_path = profile_dir / "template_placeholder.md"

    missing = [
        str(path.relative_to(SKILL_DIR))
        for path in (profile_md, features_path, format_rules_path, validation_rules_path)
        if not path.exists()
    ]
    if missing:
        raise ProfileError(f"Profile '{profile_id}' is missing required files: {', '.join(missing)}")

    features = read_json(features_path)
    format_rules = read_json(format_rules_path)
    validation_rules = read_json(validation_rules_path)

    template_value = features.get("template")
    template_path = resolve_relative(profile_dir, template_value)
    if template_path is None:
        candidate = profile_dir / "template.docx"
        template_path = candidate if candidate.exists() else placeholder_path

    if template_path.name == "template_placeholder.md" and not template_path.exists() and placeholder_path.exists():
        template_path = placeholder_path

    profile_status = infer_profile_status(profile_dir, features, format_rules, validation_rules, template_path)
    feature_summary = summarize_features(features)

    format_rules_source = format_rules.get("format_rules_source")
    generator_rules_path = resolve_relative(profile_dir, format_rules_source)
    if generator_rules_path is None or not generator_rules_path.exists():
        generator_rules_path = format_rules_path

    return {
        "profile_id": profile_id,
        "profile_dir": str(profile_dir),
        "profile_md": str(profile_md),
        "features_path": str(features_path),
        "format_rules_path": str(format_rules_path),
        "validation_rules_path": str(validation_rules_path),
        "template_path": str(template_path) if template_path else None,
        "template_exists": bool(template_path and template_path.exists()),
        "template_is_docx": bool(template_path and template_path.suffix.lower() == ".docx"),
        "template_placeholder_path": str(placeholder_path) if placeholder_path.exists() else None,
        "generator_rules_path": str(generator_rules_path),
        "status": profile_status,
        "display_name": features.get("display_name", profile_id),
        "features": features,
        "format_rules": format_rules,
        "validation_rules": validation_rules,
        "feature_summary": feature_summary,
        "allow_formal_delivery": profile_status == "production",
    }


def list_profiles() -> list[dict[str, Any]]:
    profiles = []
    for profile_dir in profile_dirs():
        try:
            info = resolve_profile(profile_dir.name)
        except ProfileError as exc:
            profiles.append(
                {
                    "profile_id": profile_dir.name,
                    "display_name": profile_dir.name,
                    "status": "invalid",
                    "error": str(exc),
                }
            )
            continue
        profiles.append(
            {
                "profile_id": info["profile_id"],
                "display_name": info["display_name"],
                "status": info["status"],
                "template": info["template_path"],
                "allow_formal_delivery": info["allow_formal_delivery"],
            }
        )
    return profiles


def inspect_profile(profile_id: str) -> dict[str, Any]:
    info = resolve_profile(profile_id)
    checks = {
        "profile_dir_exists": Path(info["profile_dir"]).exists(),
        "profile_md_readable": Path(info["profile_md"]).exists(),
        "features_json_readable": Path(info["features_path"]).exists(),
        "format_rules_json_readable": Path(info["format_rules_path"]).exists(),
        "validation_rules_json_readable": Path(info["validation_rules_path"]).exists(),
        "template_exists": info["template_exists"],
        "template_is_docx": info["template_is_docx"],
    }
    info["checks"] = checks
    info["ok"] = all(
        checks[key]
        for key in (
            "profile_dir_exists",
            "profile_md_readable",
            "features_json_readable",
            "format_rules_json_readable",
            "validation_rules_json_readable",
        )
    ) and (info["status"] != "production" or (checks["template_exists"] and checks["template_is_docx"]))
    return info


def print_profile_list() -> None:
    for item in list_profiles():
        if item.get("error"):
            print(f"- {item['profile_id']}：{item['display_name']}，invalid，{item['error']}")
        else:
            print(f"- {item['profile_id']}：{item['display_name']}，{item['status']}")


def print_inspection(profile_id: str) -> None:
    info = inspect_profile(profile_id)
    print(f"Profile: {info['profile_id']}")
    print(f"Display name: {info['display_name']}")
    print(f"Status: {info['status']}")
    print(f"Allow formal delivery: {info['allow_formal_delivery']}")
    print(f"Profile path: {info['profile_dir']}")
    print(f"Template: {info['template_path']}")
    print(f"Generator rules: {info['generator_rules_path']}")
    print("Checks:")
    for key, value in info["checks"].items():
        print(f"- {key}: {value}")
    print("Enabled feature status:")
    for status in ("implemented", "partial", "planned", "unknown"):
        values = info["feature_summary"].get(status, [])
        print(f"- {status}: {', '.join(values) if values else 'none'}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Resolve official-word-generator document profiles.")
    parser.add_argument("--list", action="store_true", help="List available profiles.")
    parser.add_argument("--profile", help="Profile id to inspect or resolve.")
    parser.add_argument("--inspect", action="store_true", help="Inspect a profile and print file/status checks.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    args = parser.parse_args()

    if args.list:
        data = list_profiles()
        if args.json:
            print(json.dumps(data, ensure_ascii=False, indent=2))
        else:
            print_profile_list()
        return

    if not args.profile:
        parser.error("Use --list or provide --profile.")

    data = inspect_profile(args.profile) if args.inspect else resolve_profile(args.profile)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    elif args.inspect:
        print_inspection(args.profile)
    else:
        print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
