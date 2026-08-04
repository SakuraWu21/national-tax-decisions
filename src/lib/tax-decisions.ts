export type TaxDecision = {
  id: string;
  caseGroupId: string | null;
  province: string | null;
  city: string | null;
  district: string | null;
  issuingAuthority: string | null;
  inspectionAuthority: string | null;
  partyName: string | null;
  unifiedSocialCreditCode: string | null;
  legalRepresentative: string | null;
  documentType: "税务处理决定书" | "税务行政处罚决定书";
  documentNumber: string | null;
  relatedTreatmentDocumentNumber: string | null;
  relatedPenaltyDocumentNumber: string | null;
  violationFacts: string | null;
  taxTypes: string | null;
  recoveredTaxAmount: number | null;
  lateFeeAmount: number | null;
  fineAmount: number | null;
  confiscatedIncomeAmount: number | null;
  result: string | null;
  decisionDate: string | null;
  officialPublishDate: string | null;
  thirdPartyPublishDate: string | null;
  firstDiscoveredDate: string | null;
  lastVerifiedDate: string | null;
  completeness: string | null;
  sourceLevel: string | null;
  verificationStatus: "已核验" | "待核验" | null;
  officialUrl: string | null;
  attachmentUrl: string | null;
  backupUrl: string | null;
  pageTitle: string | null;
  pageStatus:
    | "正常"
    | "附件正常"
    | "暂时无法访问"
    | "页面已删除"
    | "内容不匹配"
    | "需要人工核验"
    | null;
  notes: string | null;
};

export type UpdateStatus = {
  status: "normal" | "degraded";
  lastUpdated: string;
  timezone: string;
  nextScheduledUpdate: string;
  searchRange: string;
  searchedPages: number;
  total: number;
  todayNew: number;
  completeDocuments: number;
  clues: number;
  pending: number;
  invalidLinks: number;
  newCompleteDocuments: number;
  newClues: number;
  updatedRecords: number;
  duplicateRecords: number;
  runSuccess: boolean;
  message: string;
};

export type LinkFallbackEntry = {
  originalUrl: string;
  sourcePageUrl: string | null;
  linkType: "attachment" | "official_document";
  status: "cached" | "cached_previous" | "unavailable";
  checkedAt: string;
  statusCode: number | null;
  contentType: string | null;
  cachedUrl: string | null;
  sha256: string | null;
  bytes: number | null;
  message: string;
};

export type LinkFallbackManifest = {
  generatedAt: string;
  attachments: Record<string, LinkFallbackEntry>;
};

export type Filters = {
  query: string;
  documentType: string;
  province: string;
  city: string;
  dateFrom: string;
  dateTo: string;
  completeness: string;
  sourceLevel: string;
  verificationStatus: string;
  pageStatus: string;
  sort: string;
};

export const defaultFilters: Filters = {
  query: "",
  documentType: "",
  province: "",
  city: "",
  dateFrom: "",
  dateTo: "",
  completeness: "",
  sourceLevel: "",
  verificationStatus: "",
  pageStatus: "",
  sort: "publish-desc",
};

export const moneyFields = [
  ["recoveredTaxAmount", "追缴税款"],
  ["lateFeeAmount", "滞纳金"],
  ["fineAmount", "罚款"],
  ["confiscatedIncomeAmount", "没收违法所得"],
] as const;

export function formatDate(value: string | null): string {
  return value ?? "—";
}

export function formatDateTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value.replace("T", " ").slice(0, 16);
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  })
    .format(date)
    .replace(/\//g, "-");
}

export function formatMoney(value: number | null): string {
  if (value === null) return "—";
  return `¥${value.toLocaleString("zh-CN", { maximumFractionDigits: 2 })}`;
}

export function majorAmount(record: TaxDecision): number | null {
  const values = moneyFields
    .map(([key]) => record[key])
    .filter((value): value is number => value !== null);
  return values.length ? Math.max(...values) : null;
}

export function isOfficialGovernmentUrl(value: string | null): boolean {
  if (!value) return false;
  try {
    const hostname = new URL(value).hostname.toLowerCase();
    return hostname === "gov.cn" || hostname.endsWith(".gov.cn") || hostname.endsWith(".chinatax.gov.cn");
  } catch {
    return false;
  }
}
