"use client";

import { ExternalLink, FileText, Link2, Scale } from "lucide-react";
import { DocumentTypeBadge, VerificationBadge } from "@/components/decision-badges";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from "@/components/ui/sheet";
import {
  formatDate,
  formatMoney,
  isOfficialGovernmentUrl,
  moneyFields,
  type TaxDecision,
} from "@/lib/tax-decisions";

const fieldRows: Array<[keyof TaxDecision, string]> = [
  ["province", "省份"],
  ["city", "城市"],
  ["district", "区县"],
  ["issuingAuthority", "发布机关"],
  ["inspectionAuthority", "稽查机构"],
  ["unifiedSocialCreditCode", "统一社会信用代码"],
  ["legalRepresentative", "法定代表人"],
  ["decisionDate", "决定书作出日期"],
  ["officialPublishDate", "官方发布日期"],
  ["thirdPartyPublishDate", "第三方收录日期"],
  ["firstDiscoveredDate", "首次发现日期"],
  ["lastVerifiedDate", "最后核验日期"],
  ["completeness", "公开完整度"],
  ["sourceLevel", "来源级别"],
  ["pageStatus", "页面状态"],
  ["caseGroupId", "案件组ID"],
];

export function DecisionDetailSheet({
  decision,
  open,
  onOpenChange,
}: {
  decision: TaxDecision | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  if (!decision) return null;
  const official = isOfficialGovernmentUrl(decision.officialUrl);
  const sourceUrl = decision.officialUrl ?? decision.backupUrl;

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent className="w-full overflow-y-auto p-0 sm:max-w-2xl">
        <SheetHeader className="border-b bg-slate-50 px-6 py-5 text-left">
          <div className="mb-2 flex flex-wrap items-center gap-2">
            <DocumentTypeBadge type={decision.documentType} />
            <VerificationBadge status={decision.verificationStatus} pageStatus={decision.pageStatus} />
          </div>
          <SheetTitle className="pr-8 text-xl leading-8 text-slate-950">
            {decision.partyName ?? "未公开当事人"}
          </SheetTitle>
          <SheetDescription className="font-mono text-sm text-slate-600">
            {decision.documentNumber ?? "文号待核验"}
          </SheetDescription>
        </SheetHeader>

        <div className="space-y-7 px-6 py-6">
          <section>
            <h3 className="mb-3 flex items-center gap-2 text-sm font-semibold text-slate-950">
              <FileText className="size-4 text-blue-600" />
              文书信息
            </h3>
            <dl className="grid grid-cols-1 gap-x-6 gap-y-4 rounded-xl border bg-white p-4 sm:grid-cols-2">
              {fieldRows.map(([key, label]) => {
                const rawValue = decision[key];
                const value = key.toLowerCase().includes("date")
                  ? formatDate(rawValue as string | null)
                  : (rawValue ?? "—");
                return (
                  <div key={key}>
                    <dt className="text-xs text-slate-500">{label}</dt>
                    <dd className="mt-1 break-words text-sm text-slate-800">{String(value)}</dd>
                  </div>
                );
              })}
            </dl>
          </section>

          <section>
            <h3 className="mb-3 flex items-center gap-2 text-sm font-semibold text-slate-950">
              <Scale className="size-4 text-blue-600" />
              金额与处理结果
            </h3>
            <div className="grid grid-cols-2 gap-3">
              {moneyFields.map(([key, label]) => (
                <div key={key} className="rounded-xl border bg-slate-50 p-4">
                  <p className="text-xs text-slate-500">{label}</p>
                  <p className="mt-1 font-mono text-base font-semibold text-slate-900">
                    {formatMoney(decision[key])}
                  </p>
                </div>
              ))}
            </div>
            <div className="mt-4 space-y-4">
              <DetailText label="主要违法事实" value={decision.violationFacts} />
              <DetailText label="处理或处罚结果" value={decision.result} />
              <DetailText label="涉及税种" value={decision.taxTypes} />
            </div>
          </section>

          <section>
            <h3 className="mb-3 flex items-center gap-2 text-sm font-semibold text-slate-950">
              <Link2 className="size-4 text-blue-600" />
              关联文书与来源
            </h3>
            <div className="space-y-3 rounded-xl border p-4 text-sm">
              <MetaLine label="关联处理决定书" value={decision.relatedTreatmentDocumentNumber} />
              <MetaLine label="关联处罚决定书" value={decision.relatedPenaltyDocumentNumber} />
              <Separator />
              <MetaLine label="页面标题" value={decision.pageTitle} />
              <MetaLine label="备注" value={decision.notes} />
            </div>
            {sourceUrl ? (
              <Button asChild className="mt-4 w-full bg-blue-700 hover:bg-blue-800">
                <a href={sourceUrl} target="_blank" rel="noopener noreferrer">
                  {official ? "打开官方原文" : "打开来源页面"}
                  <ExternalLink className="size-4" />
                </a>
              </Button>
            ) : (
              <div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
                当前记录暂无可访问的原文链接，已保留历史信息并等待再次核验。
              </div>
            )}
          </section>
        </div>
      </SheetContent>
    </Sheet>
  );
}

function DetailText({ label, value }: { label: string; value: string | null }) {
  return (
    <div>
      <p className="text-xs font-medium text-slate-500">{label}</p>
      <p className="mt-1 whitespace-pre-wrap text-sm leading-6 text-slate-700">{value ?? "未公开"}</p>
    </div>
  );
}

function MetaLine({ label, value }: { label: string; value: string | null }) {
  return (
    <div className="grid gap-1 sm:grid-cols-[8rem_1fr]">
      <span className="text-slate-500">{label}</span>
      <span className="break-words text-slate-800">{value ?? "—"}</span>
    </div>
  );
}
