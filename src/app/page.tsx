import { Suspense } from "react";
import { DashboardSkeleton } from "@/components/dashboard-skeleton";
import { TaxDecisionsDashboard } from "@/components/tax-decisions-dashboard";

export default async function Home({
  searchParams,
}: {
  searchParams: Promise<Record<string, string | string[] | undefined>>;
}) {
  const resolvedParams = await searchParams;
  const initialParams = Object.fromEntries(
    Object.entries(resolvedParams).flatMap(([key, value]) => {
      const firstValue = Array.isArray(value) ? value[0] : value;
      return firstValue ? [[key, firstValue]] : [];
    }),
  );
  return (
    <Suspense fallback={<DashboardSkeleton />}>
      <TaxDecisionsDashboard initialParams={initialParams} />
    </Suspense>
  );
}
