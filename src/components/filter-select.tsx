"use client";

import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const displayLabels: Record<string, string> = {
  "publish-desc": "官方发布日期（新到旧）",
  "decision-desc": "决定日期（新到旧）",
  "amount-desc": "主要金额（高到低）",
  "party-asc": "当事人名称（A-Z）",
};

export function FilterSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}) {
  return (
    <Select value={value || "__all"} onValueChange={(next) => onChange(next === "__all" ? "" : next)}>
      <SelectTrigger className="h-10 w-full bg-white" aria-label={label}>
        <SelectValue placeholder={label} />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value="__all">全部{label}</SelectItem>
        {options.map((option) => (
          <SelectItem key={option} value={option}>
            {displayLabels[option] ?? option}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}
