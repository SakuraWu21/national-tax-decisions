# 全国税务决定书查询平台

公开查询全国已检索并核验的《税务处理决定书》和《税务行政处罚决定书》。网站使用 Next.js App Router，数据由 Python 3.12 采集程序生成，并由 GitHub Actions 在北京时间每天 12:07 增量更新。定时任务避开 GitHub Actions 整点拥堵，提交数据后会继续核验 Vercel 生产站点是否已经显示本次版本。

## 本地网站

```powershell
npm install
npm run dev
```

浏览器打开 `http://localhost:3000`。生产页面读取：

- `public/data/tax-decisions.json`
- `public/data/update-status.json`
- `public/downloads/全国税务处理及行政处罚决定书汇总.xlsx`
- `public/downloads/全国税务处理及行政处罚决定书汇总.csv`

## 本地更新数据

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe scripts\update_tax_decisions.py
.venv\Scripts\python.exe scripts\verify_system.py
```

首次没有历史文件时自动从 2026-01-01 回溯；后续默认检索过去 7 天，并重新核验历史待核验记录和失效链接。也可手动强制全量回溯：

```powershell
python scripts/update_tax_decisions.py --force-full
```

检索采用 Bing RSS、Bing HTML、百度 HTML、固定官方入口和已知税务机关公告栏目直扫，并逐条访问原页面及附件核验。程序会从历史官方原文自动扩展可直扫栏目，降低搜索引擎延迟收录造成的遗漏。第三方页面只作线索，无法找到官方原文时进入“待核验记录”。历史 Excel、CSV 与 JSON 始终增量合并，不清空旧记录。

## 验证

```powershell
npm run lint
npm run typecheck
npm run build
.venv\Scripts\python.exe scripts\verify_system.py
```

验证脚本检查 Excel 可打开、CSV/Excel/JSON 数量一致、唯一 ID 无重复、JSON Schema 有效以及公开下载文件与数据源一致。
