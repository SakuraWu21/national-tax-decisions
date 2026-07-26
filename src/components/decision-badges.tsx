import { Badge } from "@/components/ui/badge";
import type { TaxDecision } from "@/lib/tax-decisions";

export function DocumentTypeBadge({ type }: { type: TaxDecision["documentType"] }) {
  const treatment = type === "税务处理决定书";
  return (
    <Badge
      className={
        treatment
          ? "border-blue-200 bg-blue-50 text-blue-700"
          : "border-orange-200 bg-orange-50 text-orange-700"
      }
      variant="outline"
    >
      {treatment ? "处理决定书" : "行政处罚决定书"}
    </Badge>
  );
}

export function VerificationBadge({
  status,
  pageStatus,
}: {
  status: TaxDecision["verificationStatus"];
  pageStatus: TaxDecision["pageStatus"];
}) {
  const invalid = pageStatus && !["正常", "附件正常"].includes(pageStatus);
  if (invalid) {
    return (
      <Badge className="border-red-200 bg-red-50 text-red-700" variant="outline">
        {pageStatus}
      </Badge>
    );
  }
  if (status === "已核验") {
    return (
      <Badge className="border-emerald-200 bg-emerald-50 text-emerald-700" variant="outline">
        已核验
      </Badge>
    );
  }
  return (
    <Badge className="border-amber-200 bg-amber-50 text-amber-700" variant="outline">
      待核验
    </Badge>
  );
}
