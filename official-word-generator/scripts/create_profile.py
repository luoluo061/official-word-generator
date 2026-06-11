from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from dataclasses import asdict
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from list_features import Feature, parse_feature_catalog  # noqa: E402
from profile_resolver import ProfileError, resolve_profile  # noqa: E402


SKILL_DIR = SCRIPT_DIR.parent
PROFILES_DIR = SKILL_DIR / "profiles"
PROFILE_ID_RE = re.compile(r"^[a-z][a-z0-9_]*$")


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def validate_profile_id(profile_id: str) -> None:
    if not PROFILE_ID_RE.match(profile_id):
        raise ValueError("profile id must be lowercase snake_case and start with a letter.")


def repair_windows_mojibake(text: str) -> str:
    """Repair common PowerShell/code-page mojibake for Chinese CLI arguments."""
    suspicious = ("Ã", "Â", "æ", "è", "é", "ç", "²", "Ê", "Ä", "£", "°")
    if not any(marker in text for marker in suspicious):
        return text
    candidates = []
    for source_encoding in ("cp1252", "latin1"):
        for target_encoding in ("utf-8", "gb18030", "gbk"):
            try:
                repaired = text.encode(source_encoding).decode(target_encoding)
            except UnicodeError:
                continue
            candidates.append(repaired)
    for candidate in candidates:
        if any("\u4e00" <= ch <= "\u9fff" for ch in candidate):
            return candidate
    return text


def feature_enabled_from_reference(reference_features: dict, feature: Feature) -> bool:
    config = reference_features.get("features", {}).get(feature.feature_id)
    if isinstance(config, dict):
        return bool(config.get("enabled")) and feature.status != "planned"
    return False


def render_feature_selection(features: list[Feature], reference_features: dict, profile_id: str, display_name: str) -> str:
    lines = [
        f"# Feature Selection: {display_name}",
        "",
        f"- profile_id: `{profile_id}`",
        "- purpose: review and edit this checklist before finalizing `features.json`.",
        "- note: implemented features may be enabled; partial features require review; planned features are listed for future planning and should not be enabled for production.",
        "- workflow: human requirements -> Codex maps to feature_id -> update features.json -> create_template_docx.py builds template.docx candidate.",
        "",
    ]
    current = None
    for feature in features:
        if feature.category != current:
            current = feature.category
            lines.extend([f"## {current}", ""])

        checked = "x" if feature_enabled_from_reference(reference_features, feature) else " "
        suffix = ""
        if feature.status == "partial":
            suffix = " - 可启用但需复核"
        elif feature.status == "planned":
            checked = " "
            suffix = " - planned，不可作为 production 能力"
        lines.append(
            f"- [{checked}] `{feature.feature_id}`：{feature.name}（{feature.status}）{suffix}"
        )
        lines.append(f"  - 说明：{feature.description}")
        lines.append(f"  - 默认启用：{feature.default_enabled}；可覆盖：{feature.overridable}；适用：{feature.applicable_documents}；依赖：{feature.dependencies}")
    lines.append("")
    return "\n".join(lines)


def build_features_json(profile_id: str, display_name: str, reference_features: dict, features: list[Feature]) -> dict:
    reference_map = reference_features.get("features", {})
    feature_map: dict[str, dict] = {}
    enabled_features: list[str] = []
    disabled_features: list[str] = []
    overridable_features: list[str] = []
    planned_features: list[str] = []

    for feature in features:
        ref_config = reference_map.get(feature.feature_id, {})
        enabled = bool(ref_config.get("enabled")) if isinstance(ref_config, dict) else False
        if feature.status == "planned":
            enabled = False
            planned_features.append(feature.feature_id)
        elif enabled:
            enabled_features.append(feature.feature_id)
        else:
            disabled_features.append(feature.feature_id)

        overridable = bool(ref_config.get("overridable", feature.overridable == "是")) if isinstance(ref_config, dict) else feature.overridable == "是"
        if overridable:
            overridable_features.append(feature.feature_id)
        feature_map[feature.feature_id] = {
            "enabled": enabled,
            "overridable": overridable,
            "status": feature.status,
        }

    return {
        "profile_id": profile_id,
        "display_name": display_name,
        "status": "draft",
        "is_reference": False,
        "template": "template_placeholder.md",
        "enabled_features": enabled_features,
        "disabled_features": disabled_features,
        "overridable_features": overridable_features,
        "planned_features": planned_features,
        "features": feature_map,
    }


def build_profile_md(profile_id: str, display_name: str, source_profile: str) -> str:
    return f"""# Profile: {profile_id}

## 中文名称

{display_name}

## Profile 定位

- 状态：`draft`
- 类型：待配置模板
- is_reference：`false`
- 来源 reference profile：`{source_profile}`

## 用途

用于 `{display_name}` 文档类型的模板化 Word 生成。当前 profile 由脚手架创建，尚未绑定正式 `template.docx`，不能作为 production profile 交付。

## 后续处理

1. 查看并修改 `feature_selection.md`。
2. 根据功能选择更新 `features.json`。
3. 替换 `template_placeholder.md` 为正式 `template.docx`，或在 `features.json` 配置稳定模板路径。
4. 按模板更新 `format_rules.json` 和 `validation_rules.json`。
5. 完善 `example.md`。
6. 运行 inspect / generate / validate。
7. 验证通过后再考虑升级为 `production`。
"""


def build_placeholder(title: str, body: str) -> str:
    return f"# {title}\n\n{body}\n"


def create_profile(profile_id: str, display_name: str, source_profile: str) -> Path:
    validate_profile_id(profile_id)
    display_name = repair_windows_mojibake(display_name)
    target_dir = PROFILES_DIR / profile_id
    if target_dir.exists():
        raise FileExistsError(f"profile already exists: {target_dir}")

    source_info = resolve_profile(source_profile)
    source_dir = Path(source_info["profile_dir"])
    if not source_dir.exists():
        raise FileNotFoundError(f"source profile not found: {source_profile}")

    features, warnings = parse_feature_catalog()
    if not features:
        raise ValueError("No features found in references/feature_catalog.md")

    target_dir.mkdir(parents=True)
    reference_features = read_json(source_dir / "features.json")

    format_rules = read_json(source_dir / "format_rules.json")
    format_rules["profile_id"] = profile_id
    format_rules["status"] = "draft"
    write_json(target_dir / "format_rules.json", format_rules)

    validation_rules = read_json(source_dir / "validation_rules.json")
    validation_rules["profile_id"] = profile_id
    validation_rules["status"] = "draft"
    write_json(target_dir / "validation_rules.json", validation_rules)

    (target_dir / "profile.md").write_text(build_profile_md(profile_id, display_name, source_profile), encoding="utf-8")
    write_json(target_dir / "features.json", build_features_json(profile_id, display_name, reference_features, features))
    (target_dir / "template_placeholder.md").write_text(
        build_placeholder(
            "Template Placeholder",
            "No approved Word template is bundled yet.\n\nAdd `template.docx` here or configure a stable template path in `features.json` before production use.",
        ),
        encoding="utf-8",
    )
    (target_dir / "example.md").write_text(
        build_placeholder(
            f"{display_name}测试文档",
            "这是脚手架生成的示例 Markdown。请根据该文档类型补充标题层级、正文、表格、图片、签署区或附件等代表性内容。",
        ),
        encoding="utf-8",
    )
    (target_dir / "expected_validation_report.md").write_text(
        build_placeholder(
            "Expected Validation Report Placeholder",
            "当前 profile 仍为 draft。请在配置 template、format_rules、validation_rules 并通过验证后，用实际通过的 validation report 替换本文件。",
        ),
        encoding="utf-8",
    )
    (target_dir / "feature_selection.md").write_text(
        render_feature_selection(features, reference_features, profile_id, display_name),
        encoding="utf-8",
    )
    if warnings:
        (target_dir / "feature_catalog_warnings.md").write_text(
            "# Feature Catalog Warnings\n\n" + "\n".join(f"- {warning}" for warning in warnings) + "\n",
            encoding="utf-8",
        )
    return target_dir


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a draft document profile from a reference profile.")
    parser.add_argument("--profile", required=True, help="New profile id, lowercase snake_case.")
    parser.add_argument("--name", required=True, help="Chinese display name.")
    parser.add_argument("--from", dest="source_profile", default="general_official", help="Reference profile id.")
    args = parser.parse_args()

    try:
        target_dir = create_profile(args.profile, args.name, args.source_profile)
    except (ValueError, FileExistsError, FileNotFoundError, ProfileError) as exc:
        parser.error(str(exc))

    print(f"Created draft profile: {target_dir}")
    print("Next steps:")
    print("- Run list_features.py --markdown to review available Word capabilities.")
    print("- Review feature_selection.md.")
    print("- Adjust features.json using feature_id values from references/feature_catalog.md.")
    print(f"- Run create_template_docx.py --profile {args.profile} --from {args.source_profile} to create a draft template.docx candidate.")
    print("- Update format_rules.json and validation_rules.json.")
    print("- Run profile_resolver.py --inspect, generate_docx.py --profile, and validate_docx.py --profile.")


if __name__ == "__main__":
    main()
