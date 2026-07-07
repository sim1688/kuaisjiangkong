-- 最近同步任务
select run_id, scope, advertiser_id, status, started_at, finished_at, stats_json
from sync_runs
order by run_id desc
limit 20;

-- 计划最新状态
select advertiser_id, campaign_id, campaign_name, put_status, status, day_budget, update_time, last_seen_at
from campaigns
order by last_seen_at desc
limit 100;

-- 广告组 ROI / 状态
select advertiser_id, campaign_id, unit_id, unit_name, roi_ratio, put_status, status, day_budget, update_time
from units
order by last_seen_at desc
limit 100;

-- 创意最新状态
select advertiser_id, campaign_id, unit_id, creative_id, creative_name, photo_id, put_status, status, update_time
from creatives
order by last_seen_at desc
limit 100;

-- 最近发生变化的实体
select entity_type, advertiser_id, entity_id, parent_id, captured_at
from entity_history
order by captured_at desc, history_id desc
limit 100;

-- 中文日报核心指标
select "广告主ID", "日期", "投放场景", "花费", "当日广告变现ROI", "激活七日广告变现ROI", "IAA广告变现ROI"
from "快手日报"
order by "日期" desc, "小时" desc
limit 100;

-- 中文广告组视图
select "广告主ID", "计划ID", "广告组ID", "广告组名称", "ROI系数", "投放状态", "平台状态"
from "快手广告组"
order by "最后入库时间" desc
limit 50;

-- 中文字段映射
select "中文表名", "中文字段名", "快手原始字段", "说明"
from "快手字段映射"
order by "中文表名", "中文字段名";
