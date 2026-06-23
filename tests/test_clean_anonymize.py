import json
import subprocess
import sys
import tempfile
import unittest


class CleanAnonymizeTest(unittest.TestCase):
    def test_clean_script_anonymizes_author_and_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            from pathlib import Path

            tmp_path = Path(tmp)
            raw_dir = tmp_path / "raw"
            clean_dir = tmp_path / "clean"
            raw_dir.mkdir()
            sample = {
                "source": "Hacker News Algolia Search API",
                "source_url": "https://hn.algolia.com/api/v1/search_by_date",
                "collected_at": "2026-06-22T00:00:00Z",
                "record": {
                    "objectID": "1",
                    "story_id": 1,
                    "author": "alice",
                    "title": "Contact me at alice@example.com",
                    "story_text": "Call +1 555-010-9999 or visit https://example.com/a",
                    "url": "https://news.ycombinator.com/item?id=1",
                    "_tags": ["story", "author_alice"],
                },
            }
            (raw_dir / "01_1.json").write_text(json.dumps(sample), encoding="utf-8")

            subprocess.run(
                [
                    sys.executable,
                    "scripts/clean_anonymize.py",
                    "--input-dir",
                    str(raw_dir),
                    "--output-dir",
                    str(clean_dir),
                    "--salt",
                    "test",
                ],
                check=True,
            )

            output = (clean_dir / "01_1.md").read_text(encoding="utf-8")
            self.assertTrue(output.startswith("---\n"))
            self.assertIn("contains_raw_subject_names: false", output)
            json_body = output.split("```json\n", 1)[1].rsplit("\n```", 1)[0]
            cleaned = json.loads(json_body)
            record = cleaned["record"]
            self.assertTrue(record["author_hash"].startswith("user_"))
            self.assertNotIn("alice", json.dumps(cleaned))
            self.assertIn("[EMAIL]", record["title"])
            self.assertIn("[PHONE]", record["story_text"])
            self.assertIn("[URL]", record["story_text"])
            self.assertTrue(record["url_domain_hash"].startswith("user_"))
            self.assertNotIn("news.ycombinator.com", json.dumps(cleaned))
            self.assertNotIn("author_alice", record["tags"])


if __name__ == "__main__":
    unittest.main()
