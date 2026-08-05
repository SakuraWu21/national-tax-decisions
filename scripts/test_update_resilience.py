#!/usr/bin/env python3
"""持续更新关键保障的可重复故障与幂等测试。"""

from __future__ import annotations

import copy
import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

import update_tax_decisions as updater


class ContinuousUpdateTests(unittest.TestCase):
    def test_http_retries_three_times_with_backoff(self) -> None:
        retries = updater.new_session().get_adapter("https://").max_retries
        self.assertEqual(retries.total, 3)
        self.assertEqual(retries.connect, 3)
        self.assertEqual(retries.read, 3)
        self.assertEqual(retries.backoff_factor, 1.0)

    def test_historical_failure_priority_cannot_be_overwritten(self) -> None:
        url = "https://fujian.chinatax.gov.cn/example.html"
        merged = updater.merge_search_hits(
            [updater.SearchHit(url, "搜索结果", "Bing HTML", "query")],
            [updater.SearchHit(url, "历史记录", "历史记录复核", "pending-or-invalid")],
        )
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].provider, "历史记录复核")

    def test_retry_queue_persists_then_clears_after_recovery(self) -> None:
        original_path = updater.RETRY_QUEUE_PATH
        with tempfile.TemporaryDirectory() as directory:
            updater.RETRY_QUEUE_PATH = Path(directory) / "retry_queue.json"
            url = "https://henan.chinatax.gov.cn/example.html"
            failed = {
                "url": url,
                "provider": "税务机关公告栏目",
                "query": url,
                "title": "测试栏目",
                "status_code": 503,
                "page_state": "暂时无法访问",
                "error": "模拟超时",
            }
            queued = updater.update_retry_queue({}, [failed])
            self.assertIn(url, queued)
            saved = json.loads(updater.RETRY_QUEUE_PATH.read_text(encoding="utf-8"))
            self.assertEqual(saved["items"][0]["failureCount"], 1)

            recovered = {**failed, "status_code": 200, "page_state": "正常", "error": ""}
            self.assertEqual(updater.update_retry_queue(queued, [recovered]), {})
        updater.RETRY_QUEUE_PATH = original_path

    def test_one_official_timeout_does_not_block_other_result(self) -> None:
        failed_hit = updater.SearchHit(
            "https://guizhou.chinatax.gov.cn/timeout.html", provider="税务机关公告栏目"
        )
        good_hit = updater.SearchHit(
            "https://jiangsu.chinatax.gov.cn/good.html", provider="税务机关公告栏目"
        )

        def fake_fetch(hit: updater.SearchHit) -> updater.FetchResult:
            if hit is failed_hit:
                raise TimeoutError("模拟官网超时")
            return updater.FetchResult(
                hit.url, hit.url, True, 200, "text/html", b"", "有效正文", "有效文书", "正常"
            )

        def fake_parse(hit: updater.SearchHit, fetched: updater.FetchResult, *_: object) -> tuple[list[dict], dict]:
            return ([{"文书唯一ID": "测试文书"}], {
                "url": hit.url,
                "provider": hit.provider,
                "status_code": fetched.status_code,
                "page_state": fetched.state,
                "error": "",
                "records": 1,
            })

        with patch.object(updater, "fetch_candidate", side_effect=fake_fetch), patch.object(
            updater, "parse_fetch", side_effect=fake_parse
        ):
            records, audits = updater.process_hits([failed_hit, good_hit], date(2026, 1, 1), False)
        self.assertEqual(len(records), 1)
        self.assertEqual(len(audits), 2)
        self.assertTrue(any(audit.get("error") for audit in audits))

    def test_duplicate_trigger_merge_is_idempotent(self) -> None:
        existing = updater.read_csv_records(updater.CSV_PATH)
        once, _ = updater.merge_all(copy.deepcopy(existing), [], copy.deepcopy(existing))
        twice, stats = updater.merge_all(copy.deepcopy(once), [], copy.deepcopy(existing))
        self.assertEqual(len(once), len(twice))
        self.assertEqual(len({record["文书唯一ID"] for record in twice}), len(twice))
        self.assertEqual(stats["new_full"] + stats["new_clue"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
