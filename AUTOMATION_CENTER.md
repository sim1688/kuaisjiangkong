# 快手自动投放中台

本中台基于当前 SQLite 数据库和 `需求账号.xlsx`，只处理 Excel 中启用的广告主账号。规则按“今天累计”的广告组报表和广告组最新状态判断，命中后可自动调用快手接口操作广告组。

## 启动页面

```powershell
python .\sync_kuaishou.py web
```

默认地址：

```text
http://127.0.0.1:8787
```

默认设置是 `自动执行关闭`、`试运行开启`。配置好规则后，先点页面右上角“试运行”，确认日志符合预期，再到“设置”里开启真实自动执行。

## 支持动作

- 暂停广告组
- 开启广告组
- 修改 ROI 系数
- 修改广告组预算

页面里的预算按“元”填写，后端调用快手接口时自动换算成“厘”。

## 支持条件字段

- 广告主ID
- 计划ID / 计划名称 / 计划投放状态
- 广告组ID / 广告组名称
- 账号余额
- 今日花费
- 昨天花费 / 昨天 IAA广告变现ROI
- IAA广告变现ROI / 当日广告变现ROI / 激活七日 ROI
- 当前 ROI 系数
- 广告组投放状态 / 平台状态
- 广告组日预算

多条件支持“全部条件”或“任一条件”。多条规则命中同一个广告组时，按优先级从小到大，只执行第一条命中的规则。

## 安全限制

在页面“设置”中可以自定义：

- 单轮最多操作数量
- 单账号单轮最多操作数量
- 单广告组每天最多成功操作次数
- 接口调用间隔秒数
- 广告主 / 计划 / 广告组黑名单

定时任务每 30 分钟拉数后会检查自动执行开关。只有 `自动执行` 开启时，才会在拉数后自动跑规则。

## 命令行

查看中台状态：

```powershell
python .\sync_kuaishou.py automation-status
```

只试运行一次规则：

```powershell
python .\sync_kuaishou.py automation-run --ignore-disabled
```

真实执行一次规则：

```powershell
python .\sync_kuaishou.py automation-run --ignore-disabled --execute
```

不加 `--ignore-disabled` 时，会遵守页面里的“自动执行”开关。

## 数据表

自动投放相关数据保存在同一个 SQLite 数据库：

- `automation_settings`：自动执行开关和安全限制
- `automation_rules`：规则配置
- `automation_logs`：每次命中、跳过、成功、失败的执行日志
