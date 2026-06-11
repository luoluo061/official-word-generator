from __future__ import annotations

import argparse
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
FEATURE_CATALOG = SKILL_DIR / "references" / "feature_catalog.md"

STATUS_VALUES = {"implemented", "partial", "planned"}
REQUIRED_FIELDS = (
    "feature_id",
    "name",
    "description",
    "default_enabled",
    "applicable_documents",
    "dependencies",
    "overridable",
    "validation",
    "status",
)


@dataclass
class Feature:
    category_id: str
    category: str
    feature_id: str
    name: str
    description: str
    default_enabled: str
    applicable_documents: str
    dependencies: str
    overridable: str
    validation: str
    status: str


def _strip_cell(value: str) -> str:
    value = value.strip()
    value = re.sub(r"<br\s*/?>", " ", value)
    value = value.replace("`", "").strip()
    return value


def _category_key(category_heading: str) -> str:
    mapping = {
        "页面设置": "page",
        "字体段落": "paragraph",
        "标题层级": "heading",
        "目录": "toc",
        "页眉页脚页码": "page_number",
        "表格": "table",
        "图片与图题": "image",
        "文种结构": "structure",
        "内容整理": "content",
        "校验": "validation",
    }
    for zh, key in mapping.items():
        if zh in category_heading:
            return key
    cleaned = re.sub(r"^\d+\.\s*", "", category_heading).strip()
    return re.sub(r"\W+", "_", cleaned).strip("_").lower() or "uncategorized"


def parse_feature_catalog(path: Path = FEATURE_CATALOG) -> tuple[list[Feature], list[str]]:
    warnings: list[str] = []
    features: list[Feature] = []
    current_category = "Uncategorized"
    current_category_id = "uncategorized"

    if not path.exists():
        return [], [f"feature catalog not found: {path}"]

    for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        heading = re.match(r"^##\s+(.+)$", line)
        if heading:
            current_category = heading.group(1).strip()
            current_category_id = _category_key(current_category)
            continue

        if not line.startswith("| `"):
            continue

        cells = [_strip_cell(cell) for cell in line.strip().strip("|").split("|")]
        if len(cells) < 9:
            warnings.append(f"line {lineno}: expected 9 table cells, got {len(cells)}")
            cells += [""] * (9 - len(cells))

        feature = Feature(
            category_id=current_category_id,
            category=current_category,
            feature_id=cells[0],
            name=cells[1],
            description=cells[2],
            default_enabled=cells[3],
            applicable_documents=cells[4],
            dependencies=cells[5],
            overridable=cells[6],
            validation=cells[7],
            status=cells[8],
        )
        missing = [field for field in REQUIRED_FIELDS if not getattr(feature, field)]
        if missing:
            warnings.append(f"line {lineno} {feature.feature_id or '<missing feature_id>'}: missing {', '.join(missing)}")
        if feature.status and feature.status not in STATUS_VALUES:
            warnings.append(f"line {lineno} {feature.feature_id}: unknown status '{feature.status}'")
        features.append(feature)

    return features, warnings


def filter_features(features: list[Feature], category: str | None = None, status: str | None = None) -> list[Feature]:
    result = features
    if category:
        category_l = category.lower()
        result = [
            feature
            for feature in result
            if category_l in feature.category_id.lower()
            or category_l in feature.category.lower()
            or category_l in feature.feature_id.lower().split(".")[0]
        ]
    if status:
        result = [feature for feature in result if feature.status == status]
    return result


def render_text(features: list[Feature], warnings: list[str]) -> str:
    lines: list[str] = []
    current = None
    for feature in features:
        if feature.category != current:
            current = feature.category
            lines.extend(["", current])
        lines.append(
            f"- {feature.feature_id}: {feature.name} | {feature.status} | "
            f"default={feature.default_enabled} | overridable={feature.overridable} | "
            f"docs={feature.applicable_documents} | depends={feature.dependencies} | {feature.description}"
        )
    if warnings:
        lines.extend(["", "Warnings:"])
        lines.extend(f"- {warning}" for warning in warnings)
    return "\n".join(lines).lstrip()


def render_markdown(features: list[Feature], warnings: list[str]) -> str:
    lines = ["# Word Feature List", ""]
    current = None
    for feature in features:
        if feature.category != current:
            current = feature.category
            lines.extend([f"## {current}", ""])
            lines.append("| feature_id | 中文名称 | 说明 | 状态 | 默认启用 | 可覆盖 | 适用文档 | 依赖 |")
            lines.append("| --- | --- | --- | --- | --- | --- | --- | --- |")
        lines.append(
            f"| `{feature.feature_id}` | {feature.name} | {feature.description} | {feature.status} | "
            f"{feature.default_enabled} | {feature.overridable} | {feature.applicable_documents} | {feature.dependencies} |"
        )
    if warnings:
        lines.extend(["", "## Warnings", ""])
        lines.extend(f"- {warning}" for warning in warnings)
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="List Word feature capabilities from references/feature_catalog.md.")
    parser.add_argument("--category", help="Filter by category id/name, such as table, toc, page, heading.")
    parser.add_argument("--status", choices=sorted(STATUS_VALUES), help="Filter by implementation status.")
    parser.add_argument("--markdown", action="store_true", help="Output a Markdown feature list.")
    parser.add_argument("--json", action="store_true", help="Output JSON.")
    args = parser.parse_args()

    features, warnings = parse_feature_catalog()
    features = filter_features(features, args.category, args.status)

    if args.json:
        print(json.dumps({"features": [asdict(feature) for feature in features], "warnings": warnings}, ensure_ascii=False, indent=2))
    elif args.markdown:
        print(render_markdown(features, warnings))
    else:
        print(render_text(features, warnings))


if __name__ == "__main__":
    main()
