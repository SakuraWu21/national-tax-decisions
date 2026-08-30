"use client";

import { useEffect, useMemo, useState } from "react";
import { usePathname, useRouter } from "next/navigation";
import {
  AlertCircle,
  CalendarClock,
  ChevronLeft,
  ChevronRight,
  CircleCheck,
  Download,
  ExternalLink,
  FileCheck2,
  FileSearch,
  Files,
  Filter,
  Landmark,
  RefreshCw,
  RotateCcw,
  Search,
} from "lucide-react";
import { DecisionDetailSheet } from "@/components/decision-detail-sheet";
import { DocumentTypeBadge, VerificationBadge } from "@/components/decision-badges";
import { DashboardSkeleton } from "@/components/dashboard-skeleton";
import { FilterSelect } from "@/components/filter-select";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Separator } from "@/components/ui/separator";
import { Sheet, SheetContent, SheetHeader, SheetTitle, SheetTrigger } from "@/components/ui/sheet";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import {
  defaultFilters,
  formatDate,
  formatDateTime,
  formatMoney,
  isOfficialGovernmentUrl,
  majorAmount,
  type Filters,
  type LinkFallbackEntry,
  type LinkFallbackManifest,
  type TaxDecision,
  type UpdateStatus,
} from "@/lib/tax-decisions";

const filterKeys = Object.keys(defaultFilters) as Array<keyof Filters>;
const emptyLinkFallbacks: LinkFallbackManifest = { generatedAt: "", attachments: {} };

function filtersFromParams(params: URLSearchParams): Filters {
  return filterKeys.reduce(
    (result, key) => ({ ...result, [key]: params.get(key) ?? defaultFilters[key] }),
    defaultFilters,
  );
}

function unique(values: Array<string | null>): string[] {
  return [...new Set(values.filter((value): value is string => Boolean(value)))].sort((a, b) =>
    a.localeCompare(b, "zh-CN"),
  );
}

function normalizedText(record: TaxDecision): string {
  return [
    record.partyName,
    record.documentNumber,
    record.issuingAuthority,
    record.inspectionAuthority,
    record.violationFacts,
    record.pageTitle,
  ]
    .filter(Boolean)
    .join(" ")
    .toLocaleLowerCase("zh-CN");
}

export function TaxDecisionsDashboard({ initialParams }: { initialParams: Record<string, string> }) {
  const router = useRouter();
  const pathname = usePathname();
  const initialSearchParams =
    typeof window === "undefined"
      ? new URLSearchParams(initialParams)
      : new URLSearchParams(window.location.search);
  const [records, setRecords] = useState<TaxDecision[]>([]);
  const [status, setStatus] = useState<UpdateStatus | null>(null);
  const [linkFallbacks, setLinkFallbacks] = useState<LinkFallbackManifest>(emptyLinkFallbacks);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(true);
  const [error, setError] = useState("");
  const [retry, setRetry] = useState(0);
  const [freshnessCheckedAt, setFreshnessCheckedAt] = useState(() => Date.now());
  const [filters, setFilters] = useState<Filters>(() => filtersFromParams(initialSearchParams));
  const [queryInput, setQueryInput] = useState(() => initialSearchParams.get("query") ?? "");
  const [page, setPage] = useState(() => Math.max(1, Number(initialSearchParams.get("page") ?? 1) || 1));
  const [pageSize, setPageSize] = useState(() => {
    const value = Number(initialSearchParams.get("pageSize") ?? 20);
    return [20, 50, 100].includes(value) ? value : 20;
  });
  const [selected, setSelected] = useState<TaxDecision | null>(null);

  useEffect(() => {
    const controller = new AbortController();
    const requestedAt = Date.now();
    fetch(`/data/update-status.json?v=${requestedAt}`, {
      cache: "no-store",
      headers: { Pragma: "no-cache" },
      signal: controller.signal,
    })
      .then((response) => {
        if (!response.ok) throw new Error("无法读取更新状态");
        return response.json() as Promise<UpdateStatus>;
      })
      .then((nextStatus) => {
        const version = encodeURIComponent(`${nextStatus.sourceCommit}-${requestedAt}`);
        return Promise.all([
          fetch(`/data/tax-decisions.json?v=${version}`, {
            cache: "no-store",
            headers: { Pragma: "no-cache" },
            signal: controller.signal,
          }).then((response) => {
            if (!response.ok) throw new Error("无法读取文书数据");
            return response.json() as Promise<TaxDecision[]>;
          }),
          Promise.resolve(nextStatus),
          fetch(`/data/link-fallbacks.json?v=${version}`, {
            cache: "no-store",
            headers: { Pragma: "no-cache" },
            signal: controller.signal,
          })
            .then((response) =>
              response.ok ? (response.json() as Promise<LinkFallbackManifest>) : emptyLinkFallbacks,
            )
            .catch(() => emptyLinkFallbacks),
        ]);
      })
      .then(([nextRecords, nextStatus, nextLinkFallbacks]) => {
        setRecords(nextRecords);
        setStatus(nextStatus);
        setLinkFallbacks(nextLinkFallbacks);
        setFreshnessCheckedAt(Date.now());
        setError("");
      })
      .catch((reason: unknown) => {
        if (reason instanceof DOMException && reason.name === "AbortError") return;
        setError(reason instanceof Error ? reason.message : "数据读取失败，请稍后重试");
      })
      .finally(() => {
        if (!controller.signal.aborted) {
          setLoading(false);
          setRefreshing(false);
        }
      });
    return () => controller.abort();
  }, [retry]);

  useEffect(() => {
    const refresh = () => {
      setRefreshing(true);
      setRetry((value) => value + 1);
    };
    const interval = window.setInterval(refresh, 5 * 60 * 1000);
    const handleVisibility = () => {
      if (document.visibilityState === "visible") refresh();
    };
    const handleFocus = () => refresh();
    document.addEventListener("visibilitychange", handleVisibility);
    window.addEventListener("focus", handleFocus);
    return () => {
      window.clearInterval(interval);
      document.removeEventListener("visibilitychange", handleVisibility);
      window.removeEventListener("focus", handleFocus);
    };
  }, []);

  useEffect(() => {
    if (queryInput === filters.query) return;
    const timeout = window.setTimeout(() => {
      setFilters((current) => ({ ...current, query: queryInput }));
      setPage(1);
    }, 300);
    return () => window.clearTimeout(timeout);
  }, [queryInput, filters.query]);

  useEffect(() => {
    const params = new URLSearchParams();
    filterKeys.forEach((key) => {
      const value = filters[key];
      if (value && value !== defaultFilters[key]) params.set(key, value);
    });
    if (page > 1) params.set("page", String(page));
    if (pageSize !== 20) params.set("pageSize", String(pageSize));
    const next = params.toString();
    router.replace(next ? `${pathname}?${next}` : pathname, { scroll: false });
  }, [filters, page, pageSize, pathname, router]);

  const filterOptions = useMemo(() => {
    const cities = records
      .filter((record) => !filters.province || record.province === filters.province)
      .map((record) => record.city);
    return {
      documentTypes: unique(records.map((record) => record.documentType)),
      provinces: unique(records.map((record) => record.province)),
      cities: unique(cities),
      completeness: unique(records.map((record) => record.completeness)),
      sourceLevels: unique(records.map((record) => record.sourceLevel)),
      verificationStatuses: unique(records.map((record) => record.verificationStatus)),
      pageStatuses: unique(records.map((record) => record.pageStatus)),
    };
  }, [records, filters.province]);

  const filtered = useMemo(() => {
    const query = filters.query.trim().toLocaleLowerCase("zh-CN");
    const matches = records.filter((record) => {
      if (query && !normalizedText(record).includes(query)) return false;
      if (filters.documentType && record.documentType !== filters.documentType) return false;
      if (filters.province && record.province !== filters.province) return false;
      if (filters.city && record.city !== filters.city) return false;
      if (filters.dateFrom && (!record.officialPublishDate || record.officialPublishDate < filters.dateFrom)) {
        return false;
      }
      if (filters.dateTo && (!record.officialPublishDate || record.officialPublishDate > filters.dateTo)) {
        return false;
      }
      if (filters.completeness && record.completeness !== filters.completeness) return false;
      if (filters.sourceLevel && record.sourceLevel !== filters.sourceLevel) return false;
      if (filters.verificationStatus && record.verificationStatus !== filters.verificationStatus) return false;
      if (filters.pageStatus && record.pageStatus !== filters.pageStatus) return false;
      return true;
    });
    return matches.toSorted((a, b) => {
      if (filters.sort === "decision-desc") return (b.decisionDate ?? "").localeCompare(a.decisionDate ?? "");
      if (filters.sort === "amount-desc") return (majorAmount(b) ?? -1) - (majorAmount(a) ?? -1);
      if (filters.sort === "party-asc") return (a.partyName ?? "").localeCompare(b.partyName ?? "", "zh-CN");
      return (b.officialPublishDate ?? b.thirdPartyPublishDate ?? "").localeCompare(
        a.officialPublishDate ?? a.thirdPartyPublishDate ?? "",
      );
    });
  }, [records, filters]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const visibleRecords = filtered.slice((safePage - 1) * pageSize, safePage * pageSize);

  function updateFilter<Key extends keyof Filters>(key: Key, value: Filters[Key]) {
    setFilters((current) => {
      const next = { ...current, [key]: value };
      if (key === "province") next.city = "";
      return next;
    });
    setPage(1);
  }

  function resetFilters() {
    setFilters(defaultFilters);
    setQueryInput("");
    setPage(1);
    setPageSize(20);
  }

  if (loading) return <DashboardSkeleton />;

  if (error || !status) {
    return (
      <main className="grid min-h-screen place-items-center bg-[#f5f7fb] p-6">
        <Alert className="max-w-lg border-red-200 bg-white">
          <AlertCircle className="size-4 text-red-600" />
          <AlertTitle>数据读取失败</AlertTitle>
          <AlertDescription className="mt-2">
            {error || "更新状态暂不可用"}。历史下载文件仍保留在仓库中。
          </AlertDescription>
          <Button
            className="mt-4"
            onClick={() => {
              setLoading(true);
              setError("");
              setRetry((value) => value + 1);
            }}
          >
            <RefreshCw className="size-4" />
            重新加载
          </Button>
        </Alert>
      </main>
    );
  }

  const statCards = [
    { label: "累计收录", value: status.total, icon: Files, note: "当前公开记录" },
    { label: "今日新增", value: status.todayNew, icon: CalendarClock, note: status.todayNew ? "今日首次发现" : "今日暂无新增" },
    { label: "完整文书", value: status.completeDocuments, icon: FileCheck2, note: "含案情或金额" },
    { label: "公告送达及文号线索", value: status.clues, icon: FileSearch, note: "保留可核验线索" },
    { label: "待核验", value: status.pending, icon: AlertCircle, note: "持续复核来源" },
  ];
  const lastSuccessfulTimestamp = Date.parse(status.lastSuccessfulRunAt);
  const updateDelayed =
    Number.isFinite(lastSuccessfulTimestamp)
    && freshnessCheckedAt - lastSuccessfulTimestamp > 26 * 60 * 60 * 1000;
  const coverageIncomplete = status.coverageStatus
    ? status.coverageStatus !== "complete"
    : status.failedSources.length > 0 || !status.runSuccess;
  const statusDegraded = status.status === "degraded" || updateDelayed || coverageIncomplete;
  const dataVersion = status.sourceCommit || status.lastUpdated;

  return (
    <main className="min-h-screen bg-[#f5f7fb]">
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-[1520px] flex-col gap-5 px-4 py-5 sm:px-6 lg:px-8 xl:flex-row xl:items-center xl:justify-between">
          <div className="flex items-start gap-3">
            <div className="grid size-11 shrink-0 place-items-center rounded-xl bg-blue-700 text-white shadow-sm">
              <Landmark className="size-6" strokeWidth={1.8} />
            </div>
            <div>
              <h1 className="text-xl font-semibold tracking-tight text-slate-950 sm:text-2xl">
                全国税务决定书查询平台
              </h1>
              <p className="mt-1 text-sm text-slate-500">
                汇集全国公开的税务处理决定书与税务行政处罚决定书
              </p>
              <div className="mt-2 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-500">
                <span
                  className={`inline-flex items-center gap-1.5 font-medium ${
                    statusDegraded ? "text-amber-700" : "text-emerald-700"
                  }`}
                >
                  <span
                    className={`size-2 rounded-full ring-4 ${
                      statusDegraded
                        ? "bg-amber-500 ring-amber-100"
                        : "bg-emerald-500 ring-emerald-100"
                    }`}
                  />
                  {updateDelayed ? "自动更新延迟" : status.status === "normal" ? "数据正常" : "部分来源待恢复"}
                </span>
                <span>最近任务完成：{formatDateTime(status.lastRunCompletedAt)}</span>
                <span>
                  最近生产部署：
                  {status.lastProductionDeploymentAt ? formatDateTime(status.lastProductionDeploymentAt) : "等待首次部署"}
                </span>
                <span>数据版本：{status.sourceCommit.slice(0, 8)}</span>
                <span>下次计划：{status.nextScheduledUpdate}</span>
                <Button
                  aria-label="立即刷新数据"
                  className="h-7 gap-1 px-2 text-xs"
                  disabled={refreshing}
                  onClick={() => {
                    setRefreshing(true);
                    setError("");
                    setRetry((value) => value + 1);
                  }}
                  size="sm"
                  title="绕过缓存并重新读取最新数据"
                  variant="ghost"
                >
                  <RefreshCw className={`size-3.5 ${refreshing ? "animate-spin" : ""}`} />
                  刷新数据
                </Button>
              </div>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            {status.runSuccess ? (
              <>
                <Button asChild className="bg-blue-700 hover:bg-blue-800">
                  <a
                    href={`/downloads/全国税务处理及行政处罚决定书汇总.xlsx?v=${encodeURIComponent(dataVersion)}`}
                    download
                  >
                    <Download className="size-4" />
                    下载Excel
                  </a>
                </Button>
                <Button asChild variant="outline">
                  <a
                    href={`/downloads/全国税务处理及行政处罚决定书汇总.csv?v=${encodeURIComponent(dataVersion)}`}
                    download
                  >
                    <Download className="size-4" />
                    下载CSV
                  </a>
                </Button>
              </>
            ) : (
              <Button disabled title="本次更新未完整成功，Excel暂不可下载">
                <Download className="size-4" />
                Excel暂不可下载
              </Button>
            )}
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1520px] space-y-5 px-4 py-5 sm:px-6 lg:px-8">
        {updateDelayed && (
          <Alert className="border-amber-300 bg-amber-50 text-amber-950">
            <AlertCircle className="size-4 text-amber-700" />
            <AlertTitle>自动更新已延迟，系统正在重试。</AlertTitle>
            <AlertDescription>
              最近成功检索时间：{formatDateTime(status.lastSuccessfulRunAt)}。历史数据仍可正常查询与下载。
            </AlertDescription>
          </Alert>
        )}
        <section className="grid grid-cols-2 gap-3 lg:grid-cols-5">
          {statCards.map((item) => (
            <Card key={item.label} className="border-slate-200 shadow-[0_1px_2px_rgba(15,23,42,0.03)]">
              <CardContent className="p-4 sm:p-5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-medium text-slate-500 sm:text-sm">{item.label}</span>
                  <item.icon className="size-4 text-blue-600" />
                </div>
                <p className="mt-3 font-mono text-2xl font-semibold tracking-tight text-slate-950 sm:text-3xl">
                  {item.value.toLocaleString("zh-CN")}
                </p>
                <p className="mt-1 text-[11px] text-slate-500 sm:text-xs">{item.note}</p>
              </CardContent>
            </Card>
          ))}
        </section>

        {coverageIncomplete && (
          <Alert className="border-amber-300 bg-amber-50 text-amber-950">
            <AlertCircle className="size-4 text-amber-700" />
            <AlertTitle>检索覆盖不完整，不能确认其他日期没有新文书</AlertTitle>
            <AlertDescription className="block space-y-2">
              <p>
                {status.runMode === "targeted" ? "最近一次为定向补录。" : "部分官网未能完成访问或解析。"}
                本轮可访问 {status.accessiblePages ?? "—"} 页、访问失败 {status.failedPages ?? "—"} 页；
                待重试 {status.retryQueueSize ?? "—"} 个链接。已核实的文书照常展示，未核实候选不计入收录数量。
              </p>
              {!!status.pendingCandidates?.length && (
                <details>
                  <summary className="cursor-pointer font-medium">
                    查看近14天待复核的官方候选链接（{status.pendingCandidates.length}个）
                  </summary>
                  <p className="my-2 text-xs">下列日期仅来自链接路径，不是已核实的官方发布日期；这些候选尚不能认定为符合收录标准的决定书。</p>
                  <ul className="max-h-64 space-y-2 overflow-y-auto pr-2 text-xs">
                    {status.pendingCandidates.map((candidate) => (
                      <li key={candidate.url} className="break-words">
                        <span className="mr-2 font-mono">{candidate.dateHint}</span>
                        <a className="underline underline-offset-2" href={candidate.url} target="_blank" rel="noopener noreferrer">
                          {candidate.title}
                        </a>
                        <span className="ml-2">（待复核，暂未计入）</span>
                      </li>
                    ))}
                  </ul>
                </details>
              )}
            </AlertDescription>
          </Alert>
        )}

        {status.todayNew === 0 && !coverageIncomplete && (
          <Alert className="border-blue-100 bg-blue-50/70 text-blue-900">
            <CircleCheck className="size-4 text-blue-600" />
            <AlertTitle>今日暂无新增文书</AlertTitle>
            <AlertDescription>本轮已检查来源未发现符合收录标准的新文书，历史记录与运行日志均已保留。</AlertDescription>
          </Alert>
        )}

        <Card className="overflow-hidden border-slate-200 shadow-[0_1px_3px_rgba(15,23,42,0.04)]">
          <CardContent className="p-0">
            <div className="border-b bg-white p-4 sm:p-5">
              <div className="relative">
                <Search className="pointer-events-none absolute left-4 top-1/2 size-5 -translate-y-1/2 text-slate-400" />
                <Input
                  className="h-12 border-slate-300 bg-slate-50 pl-12 pr-4 text-base shadow-none focus-visible:bg-white"
                  value={queryInput}
                  onChange={(event) => setQueryInput(event.target.value)}
                  placeholder="搜索当事人、决定书文号、发布机关、违法事实或页面标题"
                  aria-label="搜索税务决定书"
                />
              </div>

              <div className="mt-4 hidden grid-cols-2 gap-3 md:grid xl:grid-cols-5">
                <FilterControls
                  filters={filters}
                  options={filterOptions}
                  updateFilter={updateFilter}
                />
              </div>

              <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <Sheet>
                    <SheetTrigger asChild>
                      <Button variant="outline" className="md:hidden">
                        <Filter className="size-4" />
                        筛选
                      </Button>
                    </SheetTrigger>
                    <SheetContent side="bottom" className="max-h-[88vh] overflow-y-auto rounded-t-2xl">
                      <SheetHeader className="text-left">
                        <SheetTitle>筛选决定书</SheetTitle>
                      </SheetHeader>
                      <div className="grid gap-3 px-4 pb-5">
                        <FilterControls
                          filters={filters}
                          options={filterOptions}
                          updateFilter={updateFilter}
                        />
                        <Button variant="outline" onClick={resetFilters}>
                          <RotateCcw className="size-4" />
                          重置筛选
                        </Button>
                      </div>
                    </SheetContent>
                  </Sheet>
                  <Button variant="ghost" className="hidden text-slate-600 md:inline-flex" onClick={resetFilters}>
                    <RotateCcw className="size-4" />
                    重置筛选
                  </Button>
                </div>
                <div className="flex items-center gap-2 text-sm text-slate-500">
                  当前结果
                  <Badge variant="secondary" className="font-mono text-slate-800">
                    {filtered.length}
                  </Badge>
                </div>
              </div>
            </div>

            {visibleRecords.length === 0 ? (
              <div className="grid min-h-72 place-items-center px-6 py-12 text-center">
                <div>
                  <div className="mx-auto grid size-12 place-items-center rounded-full bg-slate-100">
                    <FileSearch className="size-5 text-slate-500" />
                  </div>
                  <h2 className="mt-4 font-medium text-slate-900">没有符合条件的结果</h2>
                  <p className="mt-1 text-sm text-slate-500">尝试减少筛选条件或更换搜索关键词。</p>
                  <Button className="mt-4" variant="outline" onClick={resetFilters}>
                    清除全部筛选
                  </Button>
                </div>
              </div>
            ) : (
              <>
                <div className="hidden xl:block">
                  <Table>
                    <TableHeader>
                      <TableRow className="bg-slate-50 hover:bg-slate-50">
                        <TableHead className="w-28">发布日期</TableHead>
                        <TableHead className="min-w-40">当事人</TableHead>
                        <TableHead className="w-28">地区</TableHead>
                        <TableHead className="w-36">文书类型</TableHead>
                        <TableHead className="min-w-48">决定书文号</TableHead>
                        <TableHead className="min-w-48">发布机关</TableHead>
                        <TableHead className="w-32 text-right">主要金额</TableHead>
                        <TableHead className="w-32">公开完整度</TableHead>
                        <TableHead className="w-32">核验状态</TableHead>
                        <TableHead className="w-48 text-right">来源与附件</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {visibleRecords.map((record) => (
                        <TableRow
                          key={record.id}
                          className="cursor-pointer border-slate-100 hover:bg-blue-50/50 focus-visible:bg-blue-50 focus-visible:outline-none"
                          tabIndex={0}
                          onClick={() => setSelected(record)}
                          onKeyDown={(event) => {
                            if (event.key === "Enter" || event.key === " ") setSelected(record);
                          }}
                        >
                          <TableCell className="font-mono text-xs text-slate-600">
                            {formatDate(record.officialPublishDate ?? record.thirdPartyPublishDate)}
                          </TableCell>
                          <TableCell className="font-medium text-slate-900">{record.partyName ?? "—"}</TableCell>
                          <TableCell className="text-slate-600">
                            {[record.province, record.city].filter(Boolean).join(" · ") || "—"}
                          </TableCell>
                          <TableCell>
                            <DocumentTypeBadge type={record.documentType} />
                          </TableCell>
                          <TableCell className="font-mono text-xs text-slate-700">
                            {record.documentNumber ?? "—"}
                          </TableCell>
                          <TableCell className="max-w-56 truncate text-slate-600" title={record.issuingAuthority ?? ""}>
                            {record.issuingAuthority ?? "—"}
                          </TableCell>
                          <TableCell className="text-right font-mono text-xs font-medium text-slate-800">
                            {formatMoney(majorAmount(record))}
                          </TableCell>
                          <TableCell>
                            <span className="text-xs text-slate-600">{record.completeness ?? "—"}</span>
                          </TableCell>
                          <TableCell>
                            <VerificationBadge status={record.verificationStatus} pageStatus={record.pageStatus} />
                          </TableCell>
                          <TableCell className="text-right">
                            <DecisionSourceLinks
                              fallback={linkFallbacks.attachments[record.id]}
                              record={record}
                            />
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>

                <div className="divide-y divide-slate-100 xl:hidden">
                  {visibleRecords.map((record) => (
                    <article key={record.id} className="space-y-3 p-4 sm:p-5">
                      <div className="flex items-start justify-between gap-3">
                        <div>
                          <h2 className="font-medium leading-6 text-slate-950">{record.partyName ?? "未公开当事人"}</h2>
                          <p className="mt-1 break-all font-mono text-xs text-slate-500">
                            {record.documentNumber ?? "文号待核验"}
                          </p>
                        </div>
                        <DocumentTypeBadge type={record.documentType} />
                      </div>
                      <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-xs">
                        <div>
                          <p className="text-slate-600">发布日期</p>
                          <p className="mt-0.5 font-mono text-slate-700">
                            {formatDate(record.officialPublishDate ?? record.thirdPartyPublishDate)}
                          </p>
                        </div>
                        <div>
                          <p className="text-slate-600">地区</p>
                          <p className="mt-0.5 text-slate-700">
                            {[record.province, record.city].filter(Boolean).join(" · ") || "—"}
                          </p>
                        </div>
                        <div className="col-span-2">
                          <p className="text-slate-600">发布机关</p>
                          <p className="mt-0.5 text-slate-700">{record.issuingAuthority ?? "—"}</p>
                        </div>
                      </div>
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <VerificationBadge status={record.verificationStatus} pageStatus={record.pageStatus} />
                        <div className="flex gap-1">
                          <Button size="sm" variant="ghost" onClick={() => setSelected(record)}>
                            查看详情
                          </Button>
                          <DecisionSourceLinks
                            fallback={linkFallbacks.attachments[record.id]}
                            record={record}
                          />
                        </div>
                      </div>
                    </article>
                  ))}
                </div>
              </>
            )}

            <Separator />
            <div className="flex flex-col gap-3 bg-slate-50/70 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-5">
              <div className="text-sm text-slate-500">
                第 <span className="font-mono text-slate-800">{safePage}</span> /{" "}
                <span className="font-mono text-slate-800">{totalPages}</span> 页
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <FilterSelect
                  label="每页数量"
                  value={String(pageSize)}
                  options={["20", "50", "100"]}
                  onChange={(value) => {
                    setPageSize(Number(value) || 20);
                    setPage(1);
                  }}
                />
                <Button
                  size="icon"
                  variant="outline"
                  disabled={safePage <= 1}
                  onClick={() => setPage((value) => Math.max(1, value - 1))}
                  aria-label="上一页"
                >
                  <ChevronLeft className="size-4" />
                </Button>
                <Button
                  size="icon"
                  variant="outline"
                  disabled={safePage >= totalPages}
                  onClick={() => setPage((value) => Math.min(totalPages, value + 1))}
                  aria-label="下一页"
                >
                  <ChevronRight className="size-4" />
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>

        <footer className="pb-4 text-center text-xs leading-5 text-slate-600">
          在当前已检索并核验的公开来源范围内 · 数据仅供查询参考，请以发布机关原文为准
        </footer>
      </div>

      <DecisionDetailSheet
        attachmentFallback={selected ? linkFallbacks.attachments[selected.id] : undefined}
        decision={selected}
        open={Boolean(selected)}
        onOpenChange={(open) => {
          if (!open) setSelected(null);
        }}
      />
    </main>
  );
}

function DecisionSourceLinks({
  fallback,
  record,
}: {
  fallback?: LinkFallbackEntry;
  record: TaxDecision;
}) {
  const documentUrl = fallback?.cachedUrl ?? (fallback?.status === "unavailable" ? null : record.attachmentUrl);
  if (!record.officialUrl && !documentUrl) {
    return <span className="text-xs text-slate-400">暂不可用</span>;
  }
  return (
    <div className="flex flex-wrap justify-end gap-1" onClick={(event) => event.stopPropagation()}>
      {documentUrl ? (
        <Button asChild size="sm" variant={fallback?.cachedUrl ? "secondary" : "outline"}>
          <a href={documentUrl} target="_blank" rel="noopener noreferrer">
            {fallback?.cachedUrl ? "稳定附件" : "文书附件"}
            <Download className="size-3.5" />
          </a>
        </Button>
      ) : null}
      {record.officialUrl ? (
        <Button asChild size="sm" variant="ghost">
          <a href={record.officialUrl} target="_blank" rel="noopener noreferrer">
            {isOfficialGovernmentUrl(record.officialUrl) ? "官网" : "来源"}
            <ExternalLink className="size-3.5" />
          </a>
        </Button>
      ) : null}
    </div>
  );
}

type FilterOptions = {
  documentTypes: string[];
  provinces: string[];
  cities: string[];
  completeness: string[];
  sourceLevels: string[];
  verificationStatuses: string[];
  pageStatuses: string[];
};

function FilterControls({
  filters,
  options,
  updateFilter,
}: {
  filters: Filters;
  options: FilterOptions;
  updateFilter: <Key extends keyof Filters>(key: Key, value: Filters[Key]) => void;
}) {
  return (
    <>
      <FilterSelect
        label="文书类型"
        value={filters.documentType}
        options={options.documentTypes}
        onChange={(value) => updateFilter("documentType", value)}
      />
      <FilterSelect
        label="省份"
        value={filters.province}
        options={options.provinces}
        onChange={(value) => updateFilter("province", value)}
      />
      <FilterSelect
        label="城市"
        value={filters.city}
        options={options.cities}
        onChange={(value) => updateFilter("city", value)}
      />
      <label className="grid grid-cols-[1fr_auto_1fr] items-center gap-2 text-xs text-slate-600">
        <Input
          type="date"
          className="h-10 bg-white text-slate-700"
          value={filters.dateFrom}
          onChange={(event) => updateFilter("dateFrom", event.target.value)}
          aria-label="官方发布日期起始日期"
        />
        至
        <Input
          type="date"
          className="h-10 bg-white text-slate-700"
          value={filters.dateTo}
          onChange={(event) => updateFilter("dateTo", event.target.value)}
          aria-label="官方发布日期结束日期"
        />
      </label>
      <FilterSelect
        label="公开完整度"
        value={filters.completeness}
        options={options.completeness}
        onChange={(value) => updateFilter("completeness", value)}
      />
      <FilterSelect
        label="来源级别"
        value={filters.sourceLevel}
        options={options.sourceLevels}
        onChange={(value) => updateFilter("sourceLevel", value)}
      />
      <FilterSelect
        label="核验状态"
        value={filters.verificationStatus}
        options={options.verificationStatuses}
        onChange={(value) => updateFilter("verificationStatus", value)}
      />
      <FilterSelect
        label="页面状态"
        value={filters.pageStatus}
        options={options.pageStatuses}
        onChange={(value) => updateFilter("pageStatus", value)}
      />
      <FilterSelect
        label="排序"
        value={filters.sort}
        options={["publish-desc", "decision-desc", "amount-desc", "party-asc"]}
        onChange={(value) => updateFilter("sort", value)}
      />
    </>
  );
}
