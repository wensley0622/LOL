#!/usr/bin/env python3
"""Clean, anonymize, and export collected Hacker News records as Markdown."""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_SALT = "course-assignment-demo-salt"
EMAIL_RE = re.compile(r"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b")
PHONE_RE = re.compile(r"(?<!\d)(?:\+?\d[\d\s().-]{7,}\d)(?!\d)")
URL_RE = re.compile(r"https?://[^\s<>'\"]+")
TAG_AUTHOR_RE = re.compile(r"^author_")
ENTITY_RE = re.compile(
    r"\b(?:[A-Z][A-Za-z0-9&.-]*(?:\s+|$)){1,5}"
)


def pseudonymize(value: str | None, salt: str) -> str | None:
    if not value:
        return None
    digest = hashlib.sha256(f"{salt}:{value}".encode("utf-8")).hexdigest()
    return f"user_{digest[:12]}"


def clean_text(value: str | None) -> str | None:
    if not value:
        return None
    text = html.unescape(value)
    text = re.sub(r"<[^>]+>", " ", text)
    text = EMAIL_RE.sub("[EMAIL]", text)
    text = PHONE_RE.sub("[PHONE]", text)
    text = URL_RE.sub("[URL]", text)
    text = ENTITY_RE.sub("[ENTITY] ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or None


def domain_hash(value: str | None, salt: str) -> str | None:
    if not value:
        return None
    parsed = urlparse(value)
    domain = parsed.netloc.lower()
    return pseudonymize(domain, salt) if domain else None


def yaml_scalar(value: object) -> str:
    if value is None:
        return "null"
    return json.dumps(value, ensure_ascii=False)


def anonymize_record(raw: dict, salt: str) -> dict:
    record = raw.get("record", {})
    tags = record.get("_tags") or []
    public_tags = [tag for tag in tags if not TAG_AUTHOR_RE.match(tag)]

    return {
        "source": raw.get("source"),
        "source_url": raw.get("source_url"),
        "collected_at": raw.get("collected_at"),
        "license_note": "Cleaned sample for coursework; contains public HN metadata with usernames pseudonymized.",
        "record": {
            "object_id": record.get("objectID"),
            "story_id": record.get("story_id"),
            "created_at": record.get("created_at"),
            "updated_at": record.get("updated_at"),
            "author_hash": pseudonymize(record.get("author"), salt),
            "title": clean_text(record.get("title")),
            "story_text": clean_text(record.get("story_text")),
            "url_domain_hash": domain_hash(record.get("url"), salt),
            "points": record.get("points"),
            "num_comments": record.get("num_comments"),
            "tags": public_tags,
        },
    }


def to_markdown(cleaned: dict) -> str:
    record = cleaned["record"]
    front_matter = {
        "source": cleaned.get("source"),
        "source_url": cleaned.get("source_url"),
        "collected_at": cleaned.get("collected_at"),
        "record_id": record.get("object_id"),
        "story_id": record.get("story_id"),
        "record_created_at": record.get("created_at"),
        "anonymized_author": record.get("author_hash"),
        "anonymized_url_domain": record.get("url_domain_hash"),
        "cleaning_version": "1.1",
        "contains_raw_subject_names": False,
    }
    yaml_lines = ["---"]
    yaml_lines.extend(f"{key}: {yaml_scalar(value)}" for key, value in front_matter.items())
    yaml_lines.append("---")
    body = json.dumps(cleaned, ensure_ascii=False, indent=2)
    return "\n".join(yaml_lines) + "\n\n```json\n" + body + "\n```\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean, anonymize, and export raw HN JSON files.")
    parser.add_argument("--input-dir", default="data/raw", help="directory containing raw JSON files")
    parser.add_argument("--output-dir", default="data/clean", help="directory for cleaned Markdown files")
    parser.add_argument("--salt", default=DEFAULT_SALT, help="salt used for author pseudonymization")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for raw_file in sorted(input_dir.glob("*.json")):
        raw = json.loads(raw_file.read_text(encoding="utf-8"))
        cleaned = anonymize_record(raw, args.salt)
        target = output_dir / f"{raw_file.stem}.md"
        target.write_text(to_markdown(cleaned), encoding="utf-8")
        print(target)


if __name__ == "__main__":
    main()
