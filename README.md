# 快手数据库落库工具

这个目录用于把 `H:\888-CODEX\2-快手自动创编` 里已经打通的快手接口沉淀到 SQLite 数据库。

已接入的快手接口：

- `POST /rest/openapi/gw/uc/v1/advertisers`：授权广告账户列表
- `POST /rest/openapi/gw/dsp/campaign/list`：广告计划列表
- `POST /rest/openapi/gw/dsp/unit/list`：广告组列表
- `POST /rest/openapi/gw/dsp/creative/list`：创意列表
- `POST /rest/openapi/v1/report/account_report`：广告主报表
- `POST /rest/openapi/oauth2/authorize/access_token`：auth code 换 token
- `POST /rest/openapi/oauth2/authorize/refresh_token`：刷新 token

## 快速开始

复制 `.env.example` 为 `.env`，填入旧项目同一套快手配置：

```powershell
Copy-Item .env.example .env
```

初始化数据库：

```powershell
python .\sync_kuaishou.py init-db
```

查看鉴权状态：

```powershell
python .\sync_kuaishou.py status
```

同步单个广告主：

```powershell
python .\sync_kuaishou.py sync --advertiser-id 39059876
```

只同步某个计划：

```powershell
python .\sync_kuaishou.py snapshot --advertiser-id 39059876 --campaign-id 9295250964
```

同步授权账户列表：

```powershell
python .\sync_kuaishou.py accounts --auth-user-id 39059876
```

同步日报到中文表：

```powershell
python .\sync_kuaishou.py report --advertiser-id 39059876 --start-date 2026-06-23 --end-date 2026-06-23
```

导入旧项目导出的 CSV 报表到中文表：

```powershell
python .\sync_kuaishou.py import-report-csv --advertiser-id 39059876 --file "H:\888-CODEX\2-快手自动创编\data\acct39059876_20260519_20260617.csv"
```

筛选指定日期有消耗账户：

```powershell
python .\sync_kuaishou.py spenders --start-date 2026-06-23 --end-date 2026-06-23 --report-only
```

同步已筛出的有消耗账户配置数据：

```powershell
python .\sync_kuaishou.py sync-spenders
```

创建需求账号 Excel：

```powershell
python .\sync_kuaishou.py init-demand-excel
```

手动执行一次 30 分钟拉数逻辑：

```powershell
python .\sync_kuaishou.py scheduled-pull
```

安装 Windows 计划任务，每 30 分钟自动拉取：

```powershell
python .\sync_kuaishou.py install-schedule
```

## 数据库结构

默认数据库文件是 `data/kuaishou.db`。

- `advertisers`：广告主账户最新状态
- `campaigns`：广告计划最新状态
- `units`：广告组最新状态
- `creatives`：创意最新状态
- `entity_history`：实体 JSON 变化历史，相同内容不会重复写入
- `raw_api_responses`：每页接口原始返回
- `sync_runs`：每次同步任务的开始、结束、状态和数量统计
- `快手日报`：中文字段的日报指标表
- `快手字段映射`：中文字段和快手原始字段的对应关系
- `快手广告主`、`快手广告计划`、`快手广告组`、`快手创意`、`快手同步任务`：中文字段查询视图
- `快手有消耗账户`：`花费 > 0` 的广告主视图

每个业务表都有 `raw_json`，先保证快手返回的完整字段落地；常用字段再单独拆列，方便后续做监控和查询。

## 常用查询

最近同步任务：

```powershell
python .\query_db.py "select run_id, scope, advertiser_id, status, started_at, finished_at, stats_json from sync_runs order by run_id desc limit 10;"
```

查看未投放或异常状态的计划：

```powershell
python .\query_db.py "select advertiser_id, campaign_id, campaign_name, put_status, status, update_time from campaigns order by last_seen_at desc limit 50;"
```

查看广告组和 ROI：

```powershell
python .\query_db.py "select campaign_id, unit_id, unit_name, roi_ratio, put_status, status from units order by last_seen_at desc limit 50;"
```

执行内置查询集合：

```powershell
python .\query_db.py --file .\queries.sql
```

查询中文日报：

```powershell
python .\query_db.py 'select "日期", "投放场景", "花费", "当日广告变现ROI", "激活七日广告变现ROI", "IAA广告变现ROI" from "快手日报" order by "日期" desc limit 20;'
```

## 当前验证结果

已用旧项目 `.env` 的授权刷新 token，并完成真实接口落库验证：

- 授权账户列表：960 个广告主
- 广告主 `39059876`：12 个计划、39 个广告组、611 条创意
- 当前只调用查询类接口，不会创建、暂停或修改快手侧数据
- `exchange-token` / `refresh-token` 的终端输出已做 token 脱敏

## 从旧项目学到的接口规则

旧项目没有额外请求签名，普通 JSON 请求使用 `Access-Token` 请求头。分页列表接口统一用 `page` 和 `page_size`，返回结构主要是 `data.details` 和 `data.total_count`。

旧项目创建链路是：先查计划、广告组、创意，再用这些字段构造创建 payload。当前目录先做数据落库，不主动创建或修改快手侧数据，避免监控项目误触发投放变更。
