#!/usr/bin/env python3
"""检查定向补录前后历史保留、连续运行幂等性，并可验收生产JSON与新增附件。"""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import uuid

import requests
import update_tax_decisions as updater


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-ref", required=True)
    parser.add_argument("--site", help="可选：验收生产站点，含提交SHA、完成时间、完整记录与新增附件")
    args = parser.parse_args()
    baseline = json.loads(subprocess.check_output(
        ["git", "show", f"{args.baseline_ref}:public/data/tax-decisions.json"], cwd=updater.ROOT,
    ))
    current = json.loads(updater.PUBLIC_JSON_PATH.read_text(encoding="utf-8"))
    old = {row["id"]: row for row in baseline}
    new = {row["id"]: row for row in current}
    assert len(new) == len(current), "文书ID重复"
    assert set(old) <= set(new), "历史记录丢失"
    assert all(new[key]["firstDiscoveredDate"] == row["firstDiscoveredDate"] for key, row in old.items()), "首次发现日期被改写"
    _, logs = updater.read_xlsx_records(updater.XLSX_PATH)
    assert len(logs) >= 2 and all(log["运行是否成功"] == "成功" for log in logs[-2:]), "连续运行日志缺失"
    assert logs[-1]["新增完整文书数量"] == 0 and logs[-1]["新增文号线索数量"] == 0, "第二次运行非幂等"
    added = [row for key, row in new.items() if key not in old]
    assert sum(log["新增完整文书数量"] + log["新增文号线索数量"] for log in logs[-2:]) == len(added), "两次补录日志数量不匹配"
    report = {"baselineRecords": len(old), "currentRecords": len(new), "added": len(added),
              "historyPreserved": True, "firstDiscoveryPreserved": True,
              "secondRunNew": 0, "runLogs": len(logs), "productionVerified": False}
    if args.site:
        base = args.site.rstrip("/")
        version = uuid.uuid4().hex
        def get(path: str) -> requests.Response:
            response = requests.get(f"{base}{path}?acceptance={version}", timeout=(10, 40),
                                    headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
            response.raise_for_status()
            assert response.status_code == 200
            return response
        status = get("/data/update-status.json").json()
        local_status = json.loads(updater.PUBLIC_STATUS_PATH.read_text(encoding="utf-8"))
        commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=updater.ROOT, text=True).strip()
        assert status["sourceCommit"] == commit, "生产SHA不一致"
        assert status["lastRunCompletedAt"] == local_status["lastRunCompletedAt"], "生产运行时间不一致"
        assert status["deploymentStatus"] == "deployed", "生产状态非deployed"
        assert get("/data/tax-decisions.json").json() == current, "生产JSON与本地记录不一致"
        manifest = get("/data/link-fallbacks.json").json()["attachments"]
        for row in added:
            entry = manifest[row["id"]]
            assert entry["cachedUrl"].startswith("/official-attachments/")
            content = get(entry["cachedUrl"]).content
            assert updater.valid_attachment_content(content), "生产附件并非有效文书文件"
            assert hashlib.sha256(content).hexdigest() == entry["sha256"], "生产附件哈希不一致"
        report.update(productionVerified=True, sourceCommit=commit,
                      lastSuccessfulRunAt=status["lastSuccessfulRunAt"],
                      productionAttachments200=len(added))
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
