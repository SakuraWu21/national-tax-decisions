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
    def test_legal_representative_rejects_notice_template_text(self) -> None:
        self.assertEqual(updater.find_legal_rep("法定代表人：张三"), "张三")
        self.assertEqual(
            updater.find_legal_rep("法定代表人或委托代理人持单位公章等材料办理"),
            "",
        )
        self.assertEqual(updater.find_legal_rep("法定代表人或者股东的个人账户"), "")

    def test_subject_parser_prefers_labeled_party_over_contextual_company(self) -> None:
        text = """广西苒炙机械设备有限公司：（纳税人识别号：91451202MAK718N46N）
        国家税务总局河池市税务局第二稽查局税务处理决定书 河市税二稽处〔2026〕12号
        同时，该地址也是广西槟咪科技有限公司（纳税人识别号：91451202MAK5F6YF55）登记注册地址。"""
        self.assertEqual(updater.candidate_subjects("税务文书送达公告", text), ["广西苒炙机械设备有限公司"])

    def test_suspicious_subject_is_revalidated_and_replaced(self) -> None:
        old = {
            "当事人名称": "同时,该地址也是广西槟咪科技有限公司",
            "官方原文链接": "https://guangxi.chinatax.gov.cn/example.html",
            "核验状态": "已核验",
            "页面当前状态": "正常",
        }
        hits = updater.historical_revalidation_hits([old])
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0].provider, "历史记录复核")

        existing = {field: "" for field in updater.FIELDS}
        existing.update(old)
        incoming = dict(existing)
        incoming["当事人名称"] = "广西苒炙机械设备有限公司"
        merged, changed = updater.merge_record(existing, incoming)
        self.assertTrue(changed)
        self.assertEqual(merged["当事人名称"], "广西苒炙机械设备有限公司")

    def test_undated_old_document_is_outside_collection_period(self) -> None:
        record = {field: "" for field in updater.FIELDS}
        record.update({
            "决定书文号": "常税稽一处〔2025〕35号",
            "首次发现日期": updater.TODAY.isoformat(),
        })
        self.assertEqual(updater.document_number_year(record["决定书文号"]), 2025)
        self.assertTrue(updater.provisional_out_of_scope(record))
        record["官方发布日期"] = "2026-08-16"
        self.assertFalse(updater.provisional_out_of_scope(record))

    def test_fine_amount_after_action_verb_is_extracted(self) -> None:
        amounts = updater.extract_amounts("对你公司虚开发票的行为处以200,000.00元的罚款。")
        self.assertEqual(amounts["罚款金额"], 200000.0)

    def test_http_retries_three_times_with_backoff(self) -> None:
        session = updater.new_session()
        retries = session.get_adapter("https://").max_retries
        self.assertEqual(retries.total, 3)
        self.assertEqual(retries.connect, 3)
        self.assertEqual(retries.read, 3)
        self.assertEqual(retries.backoff_factor, 1.0)
        self.assertNotIn("TaxDecisionResearchBot", session.headers["User-Agent"])

    def test_waf_responses_are_persisted_as_temporary_failures(self) -> None:
        for status_code in (403, 408, 412, 425, 429):
            with self.subTest(status_code=status_code):
                self.assertEqual(updater.detect_page_state(status_code, "访问被拦截", ""), "暂时无法访问")
                self.assertTrue(updater.transient_access_failure({
                    "status_code": status_code,
                    "page_state": "暂时无法访问",
                    "error": "",
                }))
        self.assertEqual(updater.detect_page_state(404, "", ""), "页面已删除")

    def test_attachment_identifier_supplies_publication_date(self) -> None:
        url = "https://fujian.chinatax.gov.cn/example/202509/P020250901344584682810.pdf"
        self.assertEqual(updater.extract_publication_date("", url), "2025-09-01")

    def test_official_delivery_table_extracts_multiple_parties(self) -> None:
        text = """关于送达税务处理决定书的公告（无锡奋西利科技有限公司、无锡强强运输有限公司、无锡市一川钢业有限公司）
        国家税务总局无锡市税务局稽查局《税务处理决定书》公告送达名单
        序号 纳税人名称 纳税人识别号 税务处理决定书文号
        1 无锡奋西利科技有限公司 91320205MAK1LYB36M 锡税稽处〔2026〕88号
        2 无锡强强运输有限公司 91320206089317767Y 锡税稽处〔2026〕89号
        3 无锡市一川钢业有限公司 320200590045750 锡税稽处〔2026〕92号
        """
        self.assertEqual(
            updater.candidate_subjects("税务处理决定书公告", text),
            ["无锡奋西利科技有限公司", "无锡强强运输有限公司", "无锡市一川钢业有限公司"],
        )
        block = updater.subject_block(text, "无锡奋西利科技有限公司", updater.candidate_subjects("", text))
        self.assertEqual(updater.find_uscc(block, "无锡奋西利科技有限公司"), "91320205MAK1LYB36M")
        self.assertEqual(updater.extract_doc_numbers(block), [("税务处理决定书", "锡税稽处〔2026〕88号")])

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
