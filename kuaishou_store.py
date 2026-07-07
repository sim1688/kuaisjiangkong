import hashlib
import json
import os
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

from kuaishou_client import load_runtime_env


ROOT_DIR = Path(__file__).resolve().parent
SHANGHAI_TZ = timezone(timedelta(hours=8), name="Asia/Shanghai")

REPORT_FIELD_ALIASES = {
    "花费": ["charge", "花费"],
    "当日广告变现ROI": [
        "event_pay_first_day_overall_roi",
        "event_pay_first_day_roi",
        "event_pay_purchase_amount_one_day_roi",
        "event_pay_purchase_amount_one_day_by_conversion_roi",
    ],
    "激活七日广告变现ROI": [
        "event_pay_purchase_amount_week_by_conversion_roi",
        "event_pay_week_overall_roi",
    ],
    "IAA广告变现ROI": [
        "minigame_iaa_purchase_roi",
        "minigame_iaa_purchase_amount_week_by_conversion_roi",
    ],
}


def local_now():
    return datetime.now(SHANGHAI_TZ).replace(microsecond=0).strftime("%Y-%m-%d %H:%M:%S")


def json_text(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def raw_hash(value):
    return hashlib.sha256(json_text(value).encode("utf-8")).hexdigest()


def first_value(source, *keys):
    for key in keys:
        value = source.get(key)
        if value not in (None, ""):
            return value
    return None


class KuaishouStore:
    def __init__(self, db_path=None, timeout=60, busy_timeout_ms=60000, set_wal=True):
        load_runtime_env()
        configured = db_path or os.getenv("KS_DB_PATH", "data/kuaishou.db")
        self.db_path = Path(configured)
        if not self.db_path.is_absolute():
            self.db_path = ROOT_DIR / self.db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path), timeout=timeout)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute(f"PRAGMA busy_timeout={int(busy_timeout_ms)}")
        if set_wal:
            self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA foreign_keys=ON")

    def close(self):
        self.conn.close()

    def init_schema(self):
        self.conn.executescript(
            """
            DROP VIEW IF EXISTS "快手广告主";
            DROP VIEW IF EXISTS "快手广告计划";
            DROP VIEW IF EXISTS "快手广告组";
            DROP VIEW IF EXISTS "快手创意";
            DROP VIEW IF EXISTS "快手同步任务";
            DROP VIEW IF EXISTS "快手报表汇总";
            DROP VIEW IF EXISTS "快手广告组报表汇总";
            DROP VIEW IF EXISTS "快手有消耗账户";

            CREATE TABLE IF NOT EXISTS sync_runs (
              run_id INTEGER PRIMARY KEY AUTOINCREMENT,
              scope TEXT NOT NULL,
              advertiser_id INTEGER,
              started_at TEXT NOT NULL,
              finished_at TEXT,
              status TEXT NOT NULL DEFAULT 'running',
              message TEXT,
              stats_json TEXT NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS raw_api_responses (
              response_id INTEGER PRIMARY KEY AUTOINCREMENT,
              run_id INTEGER,
              endpoint TEXT NOT NULL,
              request_json TEXT NOT NULL,
              response_json TEXT NOT NULL,
              code INTEGER,
              message TEXT,
              created_at TEXT NOT NULL,
              FOREIGN KEY(run_id) REFERENCES sync_runs(run_id)
            );

            CREATE TABLE IF NOT EXISTS advertisers (
              advertiser_id INTEGER PRIMARY KEY,
              advertiser_name TEXT,
              product_name TEXT,
              corporation_name TEXT,
              agent_id INTEGER,
              user_id INTEGER,
              auth_status INTEGER,
              frozen_status INTEGER,
              raw_json TEXT NOT NULL,
              first_seen_at TEXT NOT NULL,
              last_seen_at TEXT NOT NULL,
              last_sync_run_id INTEGER
            );

            CREATE TABLE IF NOT EXISTS campaigns (
              advertiser_id INTEGER NOT NULL,
              campaign_id INTEGER NOT NULL,
              campaign_name TEXT,
              ad_type INTEGER,
              campaign_type INTEGER,
              bid_type INTEGER,
              put_status INTEGER,
              status INTEGER,
              day_budget REAL,
              create_time TEXT,
              update_time TEXT,
              raw_json TEXT NOT NULL,
              first_seen_at TEXT NOT NULL,
              last_seen_at TEXT NOT NULL,
              last_sync_run_id INTEGER,
              PRIMARY KEY(advertiser_id, campaign_id)
            );

            CREATE TABLE IF NOT EXISTS units (
              advertiser_id INTEGER NOT NULL,
              campaign_id INTEGER,
              unit_id INTEGER NOT NULL,
              unit_name TEXT,
              bid_type INTEGER,
              ocpx_action_type INTEGER,
              roi_ratio REAL,
              put_status INTEGER,
              status INTEGER,
              day_budget REAL,
              begin_time TEXT,
              create_time TEXT,
              update_time TEXT,
              raw_json TEXT NOT NULL,
              first_seen_at TEXT NOT NULL,
              last_seen_at TEXT NOT NULL,
              last_sync_run_id INTEGER,
              PRIMARY KEY(advertiser_id, unit_id)
            );

            CREATE TABLE IF NOT EXISTS creatives (
              advertiser_id INTEGER NOT NULL,
              campaign_id INTEGER,
              unit_id INTEGER,
              creative_id INTEGER NOT NULL,
              creative_name TEXT,
              photo_id TEXT,
              creative_material_type INTEGER,
              put_status INTEGER,
              status INTEGER,
              action_bar_text TEXT,
              description TEXT,
              create_time TEXT,
              update_time TEXT,
              raw_json TEXT NOT NULL,
              first_seen_at TEXT NOT NULL,
              last_seen_at TEXT NOT NULL,
              last_sync_run_id INTEGER,
              PRIMARY KEY(advertiser_id, creative_id)
            );

            CREATE TABLE IF NOT EXISTS entity_history (
              history_id INTEGER PRIMARY KEY AUTOINCREMENT,
              entity_type TEXT NOT NULL,
              advertiser_id INTEGER NOT NULL,
              entity_id INTEGER NOT NULL,
              parent_id INTEGER,
              captured_at TEXT NOT NULL,
              raw_hash TEXT NOT NULL,
              raw_json TEXT NOT NULL,
              run_id INTEGER,
              UNIQUE(entity_type, advertiser_id, entity_id, raw_hash)
            );

            CREATE INDEX IF NOT EXISTS idx_campaigns_status ON campaigns(advertiser_id, put_status, status);
            CREATE INDEX IF NOT EXISTS idx_units_campaign ON units(advertiser_id, campaign_id);
            CREATE INDEX IF NOT EXISTS idx_units_status ON units(advertiser_id, put_status, status);
            CREATE INDEX IF NOT EXISTS idx_creatives_unit ON creatives(advertiser_id, unit_id);
            CREATE INDEX IF NOT EXISTS idx_creatives_status ON creatives(advertiser_id, put_status, status);
            CREATE INDEX IF NOT EXISTS idx_history_entity ON entity_history(entity_type, advertiser_id, entity_id, captured_at);

            CREATE TABLE IF NOT EXISTS "快手日报" (
              "广告主ID" INTEGER NOT NULL,
              "日期" TEXT NOT NULL,
              "小时" INTEGER NOT NULL DEFAULT 0,
              "投放场景" TEXT NOT NULL DEFAULT '',
              "花费" REAL,
              "当日广告变现ROI" REAL,
              "激活七日广告变现ROI" REAL,
              "IAA广告变现ROI" REAL,
              "来源" TEXT,
              "来源文件" TEXT,
              "原始数据" TEXT NOT NULL,
              "首次入库时间" TEXT NOT NULL,
              "最后入库时间" TEXT NOT NULL,
              "同步任务ID" INTEGER,
              PRIMARY KEY("广告主ID", "日期", "小时", "投放场景")
            );

            CREATE INDEX IF NOT EXISTS "idx_快手日报_日期" ON "快手日报"("广告主ID", "日期");

            CREATE TABLE IF NOT EXISTS "快手广告组日报" (
              "广告主ID" INTEGER NOT NULL,
              "日期" TEXT NOT NULL,
              "小时" INTEGER NOT NULL DEFAULT 0,
              "投放场景" TEXT NOT NULL DEFAULT '',
              "计划ID" INTEGER,
              "计划名称" TEXT,
              "广告组ID" INTEGER NOT NULL,
              "广告组名称" TEXT,
              "花费" REAL,
              "当日广告变现ROI" REAL,
              "激活七日广告变现ROI" REAL,
              "IAA广告变现ROI" REAL,
              "来源" TEXT,
              "原始数据" TEXT NOT NULL,
              "首次入库时间" TEXT NOT NULL,
              "最后入库时间" TEXT NOT NULL,
              "同步任务ID" INTEGER,
              PRIMARY KEY("广告主ID", "日期", "小时", "投放场景", "广告组ID")
            );

            CREATE INDEX IF NOT EXISTS "idx_快手广告组日报_日期" ON "快手广告组日报"("广告主ID", "日期");
            CREATE INDEX IF NOT EXISTS "idx_快手广告组日报_广告组" ON "快手广告组日报"("广告主ID", "广告组ID");

            CREATE TABLE IF NOT EXISTS "快手字段映射" (
              "中文表名" TEXT NOT NULL,
              "中文字段名" TEXT NOT NULL,
              "快手原始字段" TEXT NOT NULL,
              "说明" TEXT,
              PRIMARY KEY("中文表名", "中文字段名")
            );

            CREATE VIEW IF NOT EXISTS "快手广告主" AS
            WITH advertiser_ids AS (
              SELECT advertiser_id
              FROM advertisers
              WHERE date(last_seen_at) = date('now', 'localtime')
              UNION
              SELECT "广告主ID" AS advertiser_id FROM "快手报表汇总"
            )
            SELECT
              ids.advertiser_id AS "广告主ID",
              a.advertiser_name AS "广告主名称",
              a.product_name AS "产品名称",
              a.corporation_name AS "公司主体",
              a.agent_id AS "代理商ID",
              a.user_id AS "用户ID",
              a.auth_status AS "授权状态",
              a.frozen_status AS "冻结状态",
              r."报表开始日期",
              r."报表结束日期",
              r."花费",
              r."当日广告变现ROI",
              r."激活七日广告变现ROI",
              r."IAA广告变现ROI",
              COALESCE(a.first_seen_at, r."报表首次入库时间") AS "首次入库时间",
              CASE
                WHEN a.last_seen_at IS NULL THEN r."报表最后入库时间"
                WHEN r."报表最后入库时间" IS NULL THEN a.last_seen_at
                WHEN r."报表最后入库时间" > a.last_seen_at THEN r."报表最后入库时间"
                ELSE a.last_seen_at
              END AS "最后入库时间",
              r."报表最后入库时间"
            FROM advertiser_ids ids
            LEFT JOIN advertisers a ON a.advertiser_id = ids.advertiser_id
            LEFT JOIN "快手报表汇总" r ON r."广告主ID" = ids.advertiser_id;

            CREATE VIEW IF NOT EXISTS "快手广告计划" AS
            SELECT
              c.advertiser_id AS "广告主ID",
              c.campaign_id AS "计划ID",
              c.campaign_name AS "计划名称",
              c.ad_type AS "广告类型",
              c.campaign_type AS "计划类型",
              c.bid_type AS "出价类型",
              c.put_status AS "投放状态",
              c.status AS "平台状态",
              c.day_budget AS "日预算",
              NULL AS "报表开始日期",
              NULL AS "报表结束日期",
              NULL AS "花费",
              NULL AS "当日广告变现ROI",
              NULL AS "激活七日广告变现ROI",
              NULL AS "IAA广告变现ROI",
              c.create_time AS "创建时间",
              c.update_time AS "更新时间",
              c.first_seen_at AS "首次入库时间",
              c.last_seen_at AS "最后入库时间"
            FROM campaigns c;

            CREATE VIEW IF NOT EXISTS "快手广告组报表汇总" AS
            SELECT
              "广告主ID",
              "广告组ID",
              MIN("日期") AS "报表开始日期",
              MAX("日期") AS "报表结束日期",
              MIN("首次入库时间") AS "报表首次入库时间",
              MAX("最后入库时间") AS "报表最后入库时间",
              SUM(COALESCE("花费", 0)) AS "花费",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("当日广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("当日广告变现ROI"), 2)
              END AS "当日广告变现ROI",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("激活七日广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("激活七日广告变现ROI"), 2)
              END AS "激活七日广告变现ROI",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("IAA广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("IAA广告变现ROI"), 2)
              END AS "IAA广告变现ROI"
            FROM "快手广告组日报"
            GROUP BY "广告主ID", "广告组ID";

            CREATE VIEW IF NOT EXISTS "快手广告组" AS
            SELECT
              u.advertiser_id AS "广告主ID",
              u.campaign_id AS "计划ID",
              u.unit_id AS "广告组ID",
              u.unit_name AS "广告组名称",
              u.bid_type AS "出价类型",
              u.ocpx_action_type AS "优化目标",
              ROUND(u.roi_ratio, 2) AS "ROI系数",
              u.put_status AS "投放状态",
              u.status AS "平台状态",
              u.day_budget AS "日预算",
              r."报表开始日期",
              r."报表结束日期",
              r."花费",
              r."当日广告变现ROI",
              r."激活七日广告变现ROI",
              r."IAA广告变现ROI",
              u.begin_time AS "开始时间",
              u.create_time AS "创建时间",
              u.update_time AS "更新时间",
              u.first_seen_at AS "首次入库时间",
              CASE
                WHEN r."报表最后入库时间" IS NULL THEN u.last_seen_at
                WHEN r."报表最后入库时间" > u.last_seen_at THEN r."报表最后入库时间"
                ELSE u.last_seen_at
              END AS "最后入库时间"
            FROM units u
            LEFT JOIN "快手广告组报表汇总" r
              ON r."广告主ID" = u.advertiser_id
             AND r."广告组ID" = u.unit_id;

            CREATE VIEW IF NOT EXISTS "快手创意" AS
            SELECT
              cr.advertiser_id AS "广告主ID",
              cr.campaign_id AS "计划ID",
              cr.unit_id AS "广告组ID",
              cr.creative_id AS "创意ID",
              cr.creative_name AS "创意名称",
              cr.photo_id AS "视频ID",
              cr.creative_material_type AS "创意素材类型",
              cr.put_status AS "投放状态",
              cr.status AS "平台状态",
              cr.action_bar_text AS "行动号召",
              cr.description AS "创意文案",
              NULL AS "报表开始日期",
              NULL AS "报表结束日期",
              NULL AS "花费",
              NULL AS "当日广告变现ROI",
              NULL AS "激活七日广告变现ROI",
              NULL AS "IAA广告变现ROI",
              cr.create_time AS "创建时间",
              cr.update_time AS "更新时间",
              cr.first_seen_at AS "首次入库时间",
              cr.last_seen_at AS "最后入库时间"
            FROM creatives cr;

            CREATE VIEW IF NOT EXISTS "快手同步任务" AS
            SELECT
              run_id AS "任务ID",
              scope AS "任务类型",
              advertiser_id AS "广告主ID",
              started_at AS "开始时间",
              finished_at AS "结束时间",
              status AS "任务状态",
              message AS "消息",
              stats_json AS "统计"
            FROM sync_runs;

            CREATE VIEW IF NOT EXISTS "快手报表汇总" AS
            SELECT
              "广告主ID",
              MIN("日期") AS "报表开始日期",
              MAX("日期") AS "报表结束日期",
              MIN("首次入库时间") AS "报表首次入库时间",
              MAX("最后入库时间") AS "报表最后入库时间",
              SUM(COALESCE("花费", 0)) AS "花费",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("当日广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("当日广告变现ROI"), 2)
              END AS "当日广告变现ROI",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("激活七日广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("激活七日广告变现ROI"), 2)
              END AS "激活七日广告变现ROI",
              CASE
                WHEN SUM(COALESCE("花费", 0)) > 0
                THEN ROUND(SUM(COALESCE("IAA广告变现ROI", 0) * COALESCE("花费", 0)) / SUM(COALESCE("花费", 0)), 2)
                ELSE ROUND(AVG("IAA广告变现ROI"), 2)
              END AS "IAA广告变现ROI"
            FROM "快手日报"
            GROUP BY "广告主ID";

            CREATE VIEW IF NOT EXISTS "快手有消耗账户" AS
            SELECT
              a."广告主ID",
              a."广告主名称",
              a."产品名称",
              a."公司主体",
              a."报表开始日期",
              a."报表结束日期",
              a."花费",
              a."当日广告变现ROI",
              a."激活七日广告变现ROI",
              a."IAA广告变现ROI",
              a."最后入库时间"
            FROM "快手广告主" a
            WHERE COALESCE(a."花费", 0) > 0
            ORDER BY a."花费" DESC;
            """
        )
        self._upsert_field_mappings()
        self.conn.commit()

    def start_run(self, scope, advertiser_id=None):
        now = local_now()
        cursor = self.conn.execute(
            "INSERT INTO sync_runs(scope, advertiser_id, started_at, status) VALUES (?, ?, ?, 'running')",
            (scope, advertiser_id, now),
        )
        self.conn.commit()
        return cursor.lastrowid

    def finish_run(self, run_id, status, stats=None, message=None):
        self.conn.execute(
            """
            UPDATE sync_runs
            SET finished_at = ?, status = ?, message = ?, stats_json = ?
            WHERE run_id = ?
            """,
            (local_now(), status, message, json_text(stats or {}), run_id),
        )
        self.conn.commit()

    def mark_stale_running_runs(self, before_time, message):
        cursor = self.conn.execute(
            """
            UPDATE sync_runs
            SET finished_at = ?, status = 'failed', message = ?
            WHERE status = 'running' AND started_at < ?
            """,
            (local_now(), message, before_time),
        )
        self.conn.commit()
        return cursor.rowcount

    def save_raw_response(self, run_id, endpoint, request_body, response):
        self.conn.execute(
            """
            INSERT INTO raw_api_responses(run_id, endpoint, request_json, response_json, code, message, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                endpoint,
                json_text(request_body or {}),
                json_text(response or {}),
                response.get("code") if isinstance(response, dict) else None,
                response.get("message") if isinstance(response, dict) else None,
                local_now(),
            ),
        )

    def keep_only_today(self, today):
        keep_dates = {
            today,
            (datetime.strptime(today, "%Y-%m-%d").date() - timedelta(days=1)).isoformat(),
        }
        placeholders = ",".join("?" for _ in keep_dates)
        params = tuple(sorted(keep_dates))

        cursor = self.conn.execute(f'DELETE FROM "快手日报" WHERE "日期" NOT IN ({placeholders})', params)
        stats = {"快手日报": cursor.rowcount}

        cursor = self.conn.execute(f'DELETE FROM "快手广告组日报" WHERE "日期" NOT IN ({placeholders})', params)
        stats["快手广告组日报"] = cursor.rowcount

        for table in ("campaigns", "units", "creatives"):
            cursor = self.conn.execute(
                f"DELETE FROM {table} WHERE substr(last_seen_at, 1, 10) <> ?",
                (today,),
            )
            stats[table] = cursor.rowcount

        cursor = self.conn.execute(
            "DELETE FROM entity_history WHERE substr(captured_at, 1, 10) <> ?",
            (today,),
        )
        stats["entity_history"] = cursor.rowcount

        cursor = self.conn.execute(
            "DELETE FROM raw_api_responses WHERE substr(created_at, 1, 10) <> ?",
            (today,),
        )
        stats["raw_api_responses"] = cursor.rowcount

        cursor = self.conn.execute(
            """
            DELETE FROM sync_runs
            WHERE substr(started_at, 1, 10) <> ?
            """,
            (today,),
        )
        stats["sync_runs"] = cursor.rowcount

        self.conn.commit()
        return stats

    def upsert_advertisers(self, details, run_id=None):
        count = 0
        now = local_now()
        for item in details:
            advertiser_id = first_value(item, "advertiser_id", "advertiserId")
            if not advertiser_id:
                continue
            raw = json_text(item)
            self.conn.execute(
                """
                INSERT INTO advertisers (
                  advertiser_id, advertiser_name, product_name, corporation_name,
                  agent_id, user_id, auth_status, frozen_status,
                  raw_json, first_seen_at, last_seen_at, last_sync_run_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(advertiser_id) DO UPDATE SET
                  advertiser_name = excluded.advertiser_name,
                  product_name = excluded.product_name,
                  corporation_name = excluded.corporation_name,
                  agent_id = excluded.agent_id,
                  user_id = excluded.user_id,
                  auth_status = excluded.auth_status,
                  frozen_status = excluded.frozen_status,
                  raw_json = excluded.raw_json,
                  last_seen_at = excluded.last_seen_at,
                  last_sync_run_id = excluded.last_sync_run_id
                """,
                (
                    advertiser_id,
                    first_value(item, "advertiser_name", "advertiserName"),
                    first_value(item, "product_name", "productName"),
                    first_value(item, "corporation_name", "corporationName"),
                    first_value(item, "agent_id", "agentId"),
                    first_value(item, "user_id", "userId"),
                    first_value(item, "auth_status", "authStatus"),
                    first_value(item, "frozen_status", "frozenStatus"),
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            self._insert_history("advertiser", int(advertiser_id), int(advertiser_id), None, item, run_id, now)
            count += 1
        self.conn.commit()
        return count

    def upsert_campaigns(self, advertiser_id, details, run_id=None):
        count = 0
        now = local_now()
        for item in details:
            campaign_id = first_value(item, "campaign_id", "campaignId")
            if not campaign_id:
                continue
            raw = json_text(item)
            self.conn.execute(
                """
                INSERT INTO campaigns (
                  advertiser_id, campaign_id, campaign_name, ad_type, campaign_type,
                  bid_type, put_status, status, day_budget, create_time, update_time,
                  raw_json, first_seen_at, last_seen_at, last_sync_run_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(advertiser_id, campaign_id) DO UPDATE SET
                  campaign_name = excluded.campaign_name,
                  ad_type = excluded.ad_type,
                  campaign_type = excluded.campaign_type,
                  bid_type = excluded.bid_type,
                  put_status = excluded.put_status,
                  status = excluded.status,
                  day_budget = excluded.day_budget,
                  create_time = excluded.create_time,
                  update_time = excluded.update_time,
                  raw_json = excluded.raw_json,
                  last_seen_at = excluded.last_seen_at,
                  last_sync_run_id = excluded.last_sync_run_id
                """,
                (
                    advertiser_id,
                    campaign_id,
                    first_value(item, "campaign_name", "campaignName"),
                    first_value(item, "ad_type", "adType"),
                    first_value(item, "campaign_type", "campaignType", "type"),
                    first_value(item, "bid_type", "bidType"),
                    first_value(item, "put_status", "putStatus"),
                    first_value(item, "status"),
                    first_value(item, "day_budget", "dayBudget"),
                    first_value(item, "create_time", "createTime"),
                    first_value(item, "update_time", "updateTime"),
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            self._insert_history("campaign", int(advertiser_id), int(campaign_id), None, item, run_id, now)
            count += 1
        self.conn.commit()
        return count

    def upsert_units(self, advertiser_id, details, run_id=None):
        count = 0
        now = local_now()
        for item in details:
            unit_id = first_value(item, "unit_id", "unitId")
            if not unit_id:
                continue
            campaign_id = first_value(item, "campaign_id", "campaignId")
            raw = json_text(item)
            self.conn.execute(
                """
                INSERT INTO units (
                  advertiser_id, campaign_id, unit_id, unit_name, bid_type,
                  ocpx_action_type, roi_ratio, put_status, status, day_budget,
                  begin_time, create_time, update_time,
                  raw_json, first_seen_at, last_seen_at, last_sync_run_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(advertiser_id, unit_id) DO UPDATE SET
                  campaign_id = excluded.campaign_id,
                  unit_name = excluded.unit_name,
                  bid_type = excluded.bid_type,
                  ocpx_action_type = excluded.ocpx_action_type,
                  roi_ratio = excluded.roi_ratio,
                  put_status = excluded.put_status,
                  status = excluded.status,
                  day_budget = excluded.day_budget,
                  begin_time = excluded.begin_time,
                  create_time = excluded.create_time,
                  update_time = excluded.update_time,
                  raw_json = excluded.raw_json,
                  last_seen_at = excluded.last_seen_at,
                  last_sync_run_id = excluded.last_sync_run_id
                """,
                (
                    advertiser_id,
                    campaign_id,
                    unit_id,
                    first_value(item, "unit_name", "unitName"),
                    first_value(item, "bid_type", "bidType"),
                    first_value(item, "ocpx_action_type", "ocpxActionType"),
                    first_value(item, "roi_ratio", "roiRatio"),
                    first_value(item, "put_status", "putStatus"),
                    first_value(item, "status"),
                    first_value(item, "day_budget", "dayBudget"),
                    first_value(item, "begin_time", "beginTime"),
                    first_value(item, "create_time", "createTime"),
                    first_value(item, "update_time", "updateTime"),
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            self._insert_history("unit", int(advertiser_id), int(unit_id), _to_int(campaign_id), item, run_id, now)
            count += 1
        self.conn.commit()
        return count

    def upsert_creatives(self, advertiser_id, details, run_id=None):
        count = 0
        now = local_now()
        for item in details:
            creative_id = first_value(item, "creative_id", "creativeId")
            if not creative_id:
                continue
            campaign_id = first_value(item, "campaign_id", "campaignId")
            unit_id = first_value(item, "unit_id", "unitId")
            display_info = item.get("display_info") if isinstance(item.get("display_info"), dict) else {}
            raw = json_text(item)
            self.conn.execute(
                """
                INSERT INTO creatives (
                  advertiser_id, campaign_id, unit_id, creative_id, creative_name,
                  photo_id, creative_material_type, put_status, status,
                  action_bar_text, description, create_time, update_time,
                  raw_json, first_seen_at, last_seen_at, last_sync_run_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(advertiser_id, creative_id) DO UPDATE SET
                  campaign_id = excluded.campaign_id,
                  unit_id = excluded.unit_id,
                  creative_name = excluded.creative_name,
                  photo_id = excluded.photo_id,
                  creative_material_type = excluded.creative_material_type,
                  put_status = excluded.put_status,
                  status = excluded.status,
                  action_bar_text = excluded.action_bar_text,
                  description = excluded.description,
                  create_time = excluded.create_time,
                  update_time = excluded.update_time,
                  raw_json = excluded.raw_json,
                  last_seen_at = excluded.last_seen_at,
                  last_sync_run_id = excluded.last_sync_run_id
                """,
                (
                    advertiser_id,
                    campaign_id,
                    unit_id,
                    creative_id,
                    first_value(item, "creative_name", "creativeName"),
                    first_value(item, "photo_id", "photoId"),
                    first_value(item, "creative_material_type", "creativeMaterialType"),
                    first_value(item, "put_status", "putStatus"),
                    first_value(item, "status"),
                    first_value(item, "action_bar_text", "actionBarText") or display_info.get("action_bar_text"),
                    first_value(item, "description") or display_info.get("description"),
                    first_value(item, "create_time", "createTime"),
                    first_value(item, "update_time", "updateTime"),
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            self._insert_history("creative", int(advertiser_id), int(creative_id), _to_int(unit_id), item, run_id, now)
            count += 1
        self.conn.commit()
        return count

    def upsert_daily_reports(self, advertiser_id, rows, run_id=None, source="", source_file=""):
        count = 0
        now = local_now()
        for row in rows:
            stat_date = first_value(row, "stat_date", "日期", "date")
            if not stat_date:
                continue
            stat_hour = _to_int(first_value(row, "stat_hour", "小时", "hour"))
            ad_scene = str(first_value(row, "ad_scene", "投放场景") or "")
            raw = json_text(row)
            self.conn.execute(
                """
                INSERT INTO "快手日报" (
                  "广告主ID", "日期", "小时", "投放场景",
                  "花费", "当日广告变现ROI", "激活七日广告变现ROI", "IAA广告变现ROI",
                  "来源", "来源文件", "原始数据", "首次入库时间", "最后入库时间", "同步任务ID"
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT("广告主ID", "日期", "小时", "投放场景") DO UPDATE SET
                  "花费" = excluded."花费",
                  "当日广告变现ROI" = excluded."当日广告变现ROI",
                  "激活七日广告变现ROI" = excluded."激活七日广告变现ROI",
                  "IAA广告变现ROI" = excluded."IAA广告变现ROI",
                  "来源" = excluded."来源",
                  "来源文件" = excluded."来源文件",
                  "原始数据" = excluded."原始数据",
                  "最后入库时间" = excluded."最后入库时间",
                  "同步任务ID" = excluded."同步任务ID"
                """,
                (
                    advertiser_id,
                    stat_date,
                    stat_hour if stat_hour is not None else 0,
                    ad_scene,
                    _to_float(_first_alias(row, REPORT_FIELD_ALIASES["花费"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["当日广告变现ROI"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["激活七日广告变现ROI"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["IAA广告变现ROI"])),
                    source,
                    source_file,
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            count += 1
        self.conn.commit()
        return count

    def upsert_unit_daily_reports(self, advertiser_id, rows, run_id=None, source=""):
        count = 0
        now = local_now()
        for row in rows:
            stat_date = first_value(row, "stat_date", "日期", "date")
            unit_id = first_value(row, "unit_id", "unitId", "广告组ID")
            if not stat_date or not unit_id:
                continue
            stat_hour = _to_int(first_value(row, "stat_hour", "小时", "hour"))
            ad_scene = str(first_value(row, "ad_scene", "投放场景") or "")
            raw = json_text(row)
            self.conn.execute(
                """
                INSERT INTO "快手广告组日报" (
                  "广告主ID", "日期", "小时", "投放场景",
                  "计划ID", "计划名称", "广告组ID", "广告组名称",
                  "花费", "当日广告变现ROI", "激活七日广告变现ROI", "IAA广告变现ROI",
                  "来源", "原始数据", "首次入库时间", "最后入库时间", "同步任务ID"
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT("广告主ID", "日期", "小时", "投放场景", "广告组ID") DO UPDATE SET
                  "计划ID" = excluded."计划ID",
                  "计划名称" = excluded."计划名称",
                  "广告组名称" = excluded."广告组名称",
                  "花费" = excluded."花费",
                  "当日广告变现ROI" = excluded."当日广告变现ROI",
                  "激活七日广告变现ROI" = excluded."激活七日广告变现ROI",
                  "IAA广告变现ROI" = excluded."IAA广告变现ROI",
                  "来源" = excluded."来源",
                  "原始数据" = excluded."原始数据",
                  "最后入库时间" = excluded."最后入库时间",
                  "同步任务ID" = excluded."同步任务ID"
                """,
                (
                    advertiser_id,
                    stat_date,
                    stat_hour if stat_hour is not None else 0,
                    ad_scene,
                    _to_int(first_value(row, "campaign_id", "campaignId", "计划ID")),
                    first_value(row, "campaign_name", "campaignName", "计划名称"),
                    int(unit_id),
                    first_value(row, "unit_name", "unitName", "广告组名称"),
                    _to_float(_first_alias(row, REPORT_FIELD_ALIASES["花费"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["当日广告变现ROI"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["激活七日广告变现ROI"])),
                    _to_roi(_first_alias(row, REPORT_FIELD_ALIASES["IAA广告变现ROI"])),
                    source,
                    raw,
                    now,
                    now,
                    run_id,
                ),
            )
            count += 1
        self.conn.commit()
        return count

    def _upsert_field_mappings(self):
        rows = [
            ("快手日报", "花费", ",".join(REPORT_FIELD_ALIASES["花费"]), "快手报表消耗字段"),
            (
                "快手日报",
                "当日广告变现ROI",
                ",".join(REPORT_FIELD_ALIASES["当日广告变现ROI"]),
                "优先取当日整体广告变现 ROI，兼容旧字段",
            ),
            (
                "快手日报",
                "激活七日广告变现ROI",
                ",".join(REPORT_FIELD_ALIASES["激活七日广告变现ROI"]),
                "优先取激活归因七日广告变现金额 ROI",
            ),
            (
                "快手日报",
                "IAA广告变现ROI",
                ",".join(REPORT_FIELD_ALIASES["IAA广告变现ROI"]),
                "优先取小游戏 IAA 广告变现 ROI",
            ),
        ]
        self.conn.executemany(
            """
            INSERT INTO "快手字段映射"("中文表名", "中文字段名", "快手原始字段", "说明")
            VALUES (?, ?, ?, ?)
            ON CONFLICT("中文表名", "中文字段名") DO UPDATE SET
              "快手原始字段" = excluded."快手原始字段",
              "说明" = excluded."说明"
            """,
            rows,
        )

    def _insert_history(self, entity_type, advertiser_id, entity_id, parent_id, item, run_id, captured_at):
        self.conn.execute(
            """
            INSERT OR IGNORE INTO entity_history(
              entity_type, advertiser_id, entity_id, parent_id, captured_at, raw_hash, raw_json, run_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                entity_type,
                advertiser_id,
                entity_id,
                parent_id,
                captured_at,
                raw_hash(item),
                json_text(item),
                run_id,
            ),
        )


def _to_int(value):
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _to_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_roi(value):
    number = _to_float(value)
    return round(number, 2) if number is not None else None


def _first_alias(row, aliases):
    for key in aliases:
        value = row.get(key)
        if value not in (None, ""):
            return value
    return None
