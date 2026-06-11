"""Validate Markdown outlines before generating PPT decks.

用法 / Usage:
    python scripts/outline_lint.py outline.md
    python scripts/outline_lint.py outline.md --json
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")


def lint_outline(path: str | Path, *, min_slides: int | None = None, max_slides: int | None = None) -> dict:
    source = Path(path)
    text = source.read_text(encoding="utf-8")
    headings: list[dict] = []
    issues: list[dict] = []
    previous_level = 0
    last_heading_line = 0
    content_since_heading = False

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.rstrip()
        match = HEADING_RE.match(line)
        if match:
            level = len(match.group(1))
            title = match.group(2).strip()
            if previous_level and level > previous_level + 1:
                issues.append(
                    {
                        "line": line_number,
                        "type": "heading-jump",
                        "message": f"Heading jumps from H{previous_level} to H{level}.",
                    }
                )
            if last_heading_line and not content_since_heading:
                issues.append(
                    {
                        "line": last_heading_line,
                        "type": "empty-section",
                        "message": "Heading has no body content before the next heading.",
                    }
                )
            headings.append({"line": line_number, "level": level, "title": title})
            previous_level = level
            last_heading_line = line_number
            content_since_heading = False
        elif line.strip():
            content_since_heading = True

    if last_heading_line and not content_since_heading:
        issues.append(
            {
                "line": last_heading_line,
                "type": "empty-section",
                "message": "Heading has no body content.",
            }
        )

    slide_candidates = [heading for heading in headings if heading["level"] <= 2]
    estimated_slides = max(1, len(slide_candidates))
    if min_slides is not None and estimated_slides < min_slides:
        issues.append(
            {
                "line": 1,
                "type": "too-few-slides",
                "message": f"Estimated slides {estimated_slides} is below minimum {min_slides}.",
            }
        )
    if max_slides is not None and estimated_slides > max_slides:
        issues.append(
            {
                "line": 1,
                "type": "too-many-slides",
                "message": f"Estimated slides {estimated_slides} is above maximum {max_slides}.",
            }
        )
    return {
        "path": str(source),
        "heading_count": len(headings),
        "estimated_slides": estimated_slides,
        "min_slides": min_slides,
        "max_slides": max_slides,
        "issues": issues,
        "ok": len(issues) == 0,
    }


def render_text(report: dict) -> str:
    lines = [
        f"Outline: {report['path']}",
        f"Headings: {report['heading_count']}",
        f"Estimated slides: {report['estimated_slides']}",
    ]
    if report["ok"]:
        lines.append("OK: outline is ready for PPT generation.")
    else:
        lines.append(f"Found {len(report['issues'])} issue(s):")
        for issue in report["issues"]:
            lines.append(f"- line {issue['line']} {issue['type']}: {issue['message']}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate a Markdown outline before PPT generation.")
    parser.add_argument("outline", help="Path to a UTF-8 Markdown outline.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    parser.add_argument("--min-slides", type=int, help="Require at least this many estimated slides.")
    parser.add_argument("--max-slides", type=int, help="Require no more than this many estimated slides.")
    args = parser.parse_args(argv)

    report = lint_outline(args.outline, min_slides=args.min_slides, max_slides=args.max_slides)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
