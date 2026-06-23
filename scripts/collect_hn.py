#!/usr/bin/env python3
"""Collect Hacker News story metadata from the public Algolia API."""

from __future__ import annotations

import argparse
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path


API_URL = "https://hn.algolia.com/api/v1/search_by_date"


def fetch_stories(limit: int) -> list[dict]:
    params = urllib.parse.urlencode({"tags": "story", "hitsPerPage": limit})
    request = urllib.request.Request(
        f"{API_URL}?{params}",
        headers={"User-Agent": "hn-open-data-pipeline/1.0"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return payload.get("hits", [])[:limit]


def main() -> None:
    parser = argparse.ArgumentParser(description="Collect HN story records.")
    parser.add_argument("--limit", type=int, default=10, help="number of stories to collect")
    parser.add_argument("--output-dir", default="data/raw", help="directory for raw JSON files")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    collected_at = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    for index, story in enumerate(fetch_stories(args.limit), start=1):
        story_id = story.get("objectID") or story.get("story_id") or f"story-{index}"
        record = {
            "source": "Hacker News Algolia Search API",
            "source_url": API_URL,
            "collected_at": collected_at,
            "record": story,
        }
        target = output_dir / f"{index:02d}_{story_id}.json"
        target.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        print(target)


if __name__ == "__main__":
    main()
