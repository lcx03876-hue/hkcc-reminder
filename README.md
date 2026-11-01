# hkcc-reminder

每天在设定的提醒时间，把「今天起 7 天」的课表、作业/考试截止日期整理成一封邮件，
同时生成可导入 Outlook 的 `.ics` 日历。纯 Windows PowerShell 5.1，
零第三方模块，数据全部留在本机。

> 本项目按「个人工具」设计：课程名、邮箱、日历源都放在本地配置里，
> 仓库本身只包含代码和示例配置，**不含任何个人信息**。
> 发布前请按 [隐私与安全](#隐私与安全) 一节核对。

> ⚠️ **运行前提：到了你设定的提醒时间，电脑必须开机、联网，并且当前 Windows 用户已登录。**
> 计划任务使用 `Interactive` 登录类型（DPAPI 凭据要用户会话才能解密），
> 且默认 **不会唤醒睡眠中的电脑**（`WakeToRun = $false`）。关机、睡眠、
> 注销期间都不会发送；开机或唤醒后 `-StartWhenAvailable` 会补跑一次，
> 但邮件会晚于设定时间，**不是准点**。详见「注册计划任务」和「已知限制」。

---

## 功能

| 功能 | 说明 |
|---|---|
| 课表 → 日历 | 支持单双周、跨年时区、单次课；输出标准 `.ics` |
| Blackboard → 日历 | 抓取个人日历源（iCal），作业/测验/考试截止自动进日历 |
| 每日邮件 | 每天在设定时间（默认 08:00，`-Time` 可改）发送 7 天窗口摘要 + 窗口外的后续截止预览（**需电脑开机且已登录**） |
| 日历源变动监测 | 对比事件指纹，发现新增/改时间/换教室/停课并在邮件顶部告警 |
| 科目标注 | Blackboard 源不含课程信息，用 UID 映射或标题规则补上 |
| 调课覆盖 | 公告里的调课/补课手工登记，生成日历时应用（EXDATE / 新增事件） |
| 多语言邮件 | `zh-CN` / `zh-TW` / `en`，只影响邮件措辞 |

**不包含**：云服务、后台常驻进程、任何自动改课表的逻辑。

---

## 一键安装（整合包）

不想手动敲命令：下载 [Releases](../../releases/latest) 里最新的
`hkcc-reminder-v*.zip`，解压到固定目录，双击 `install.bat`。
向导会完成邮箱、日历源、提醒时间配置，并注册计划任务；包内
[INSTALL.txt](INSTALL.txt) 有完整步骤。

> 需要 Windows 10/11、PowerShell 5.1 和管理员权限（UAC）。
> 安装包只含代码与示例配置，不含任何个人信息。

---

## 数据流

```
Blackboard 日历源 ──sync-feed.ps1──┐
                                   ├─ build-merged.ps1 ─┬─ HKCC_merged.ics
data/timetable.ics ────────────────┤                     ├─ HKCC_classes.ics
config/overrides.psd1 ─────────────┘                     └─ HKCC_deadlines.ics
                                                              │
                                          daily-reminder.ps1 ─┴─ 邮件 + 摘要
```

邮件摘要**不直接读**原始 `blackboard.ics`，而是读合并产物
（`HKCC_deadlines.ics` / `HKCC_classes.ics`）。这样「旧学期过滤」「科目
标注」只实现一次，邮件、日历、CSV 报表不会互相矛盾。

---

## 快速开始

### 0. 前置

- Windows 10/11
- Windows PowerShell 5.1（系统自带）
- 一个支持 SMTP 的邮箱（Office 365 / Gmail / QQ / 163 均可）
- 到了设定的提醒时间，电脑开机、联网，并且当前用户已登录（见上方运行前提）

### 1. 复制配置

```powershell
cd <repo>
Copy-Item .\config\config.example.psd1        .\config\config.psd1
Copy-Item .\config\courses.example.psd1       .\config\courses.psd1
Copy-Item .\config\extra-deadlines.example.psd1 .\config\extra-deadlines.psd1
Copy-Item .\config\overrides.example.psd1     .\config\overrides.psd1
Copy-Item .\config\uid-map.example.psd1       .\config\uid-map.psd1
```

`config.psd1` 里至少填 `From`、`To`、`FeedUrl`。
`FeedUrl` 是 Blackboard 的个人日历源（Blackboard → Calendar →
Get external calendar link），**带个人 token，等同密码**，只放在本地配置。

### 2. 准备课表

```powershell
Copy-Item .\data\timetable.sample.ics .\data\timetable.ics
```

编辑 `data/timetable.ics`：`SUMMARY` 用「课程代号 + 名称」，`DESCRIPTION`
写课型（`Lecture` / `Tutorial` / `Laboratory`），`LOCATION` 写教室。
格式说明见 [`data/README.md`](data/README.md)。

### 3. 配置邮件与凭据

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup.ps1
```

向导会问发件邮箱、SMTP、收件人、日历源 URL，然后弹出 Windows 凭据窗口
输入**应用专用密码/授权码**（不是网页登录密码）。密码经 DPAPI 加密后写到
`config\smtp.cred.xml`，只有本机、当前 Windows 账户能解密。

也可以手动改 `config\config.psd1` 后，用
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\init-credential.ps1
```
只更新密码。

### 4. 试跑

```powershell
# 只预览，不发邮件
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\daily-reminder.ps1 -Preview

# 同步日历、生成 .ics，不发邮件
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\daily-reminder.ps1 -NoMail

# 发一封测试邮件
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-mail.ps1
```

生成的日历在 `out\`：

| 文件 | 内容 |
|---|---|
| `HKCC_merged.ics` | 课表 + 截止日期（导入 Outlook 用这个） |
| `HKCC_classes.ics` | 只有课 |
| `HKCC_deadlines.ics` | 只有作业/考试 |

### 5. 注册计划任务

```powershell
# -Time 改成你自己的提醒时间；不填默认 08:00
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-task.ps1 -Time "08:00"
```

任务名 `HKCC Daily Reminder`，`LogonType Interactive`（DPAPI 需要用户登录
才能解密），启用了 `-StartWhenAvailable`。

> ⚠️ **这不是云端定时任务，它依赖你的电脑。** 到了你设置的提醒时间（默认
> 08:00），电脑必须开机，并且你已经登录 Windows（锁屏可以，注销不行）。
>
> - **睡眠 / 休眠**：默认不会自动唤醒（`WakeToRun = $false`），邮件会等你唤醒后补跑。
> - **关机**：开机并登录后补跑。
> - **注销**：不运行；重新登录后补跑。
>
> 补跑靠 `-StartWhenAvailable`，所以邮件时间会晚于设定时间。想让它唤醒睡眠中的
> 电脑，可在「任务计划程序 → 属性 → 条件」勾选“唤醒计算机运行此任务”
> ——只对睡眠有效，关机无效。

```powershell
# 立即跑一次
Start-ScheduledTask -TaskName "HKCC Daily Reminder"

# 看结果（0 = 成功，1 = 出错，2 = 缺凭据/配置）
Get-ScheduledTaskInfo -TaskName "HKCC Daily Reminder" | Format-List
```

---

## 配置说明

### `config.psd1` 主要字段

| 字段 | 说明 |
|---|---|
| `From` / `To` | 发件邮箱 / 收件人（数组） |
| `SmtpServer` / `SmtpPort` / `UseSsl` | SMTP 服务器；Office 365 用 `smtp.office365.com:587` |
| `FeedUrl` | Blackboard 个人日历源；留空则不同步，沿用 `data\blackboard.ics` |
| `SubjectPrefix` | 邮件标题前缀（纯 ASCII） |
| `Days` | 摘要窗口天数，默认 7（含今天） |
| `Language` | 邮件语言：`zh-CN` / `zh-TW` / `en` |
| `HideBeforeDate` | 早于该日期的 Blackboard 事件全部隐藏（过滤旧学期） |
| `HideTitles` | 按标题通配符永久隐藏，如 `'Application for LC0036*'` |
| `ClientId` / `Tenant` | 可选，读邮件功能用；留空即禁用 |

> `config.psd1` 必须保持**纯 ASCII**：Windows PowerShell 5.1 解析无 BOM 的
> 非 ASCII `.psd1` 会出错。中文说明写在这里和 README 里。

### 科目标注

Blackboard 的 `.ics` 不带课程名（实测 `CATEGORIES` / `URL` 全为空），所以
用两种方式补：

1. **UID 映射**（`config\uid-map.psd1`）：UID 稳定，标一次即可。
   交互式标注：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\label-events.ps1
   ```
   或先把 `out\label-report.csv` 用 Excel 填好 `CourseCode` 列，再：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\import-label-csv.ps1
   ```
2. **标题规则**（`config\courses.psd1` 的 `Rules`）：按通配符/正则批量匹配，
   适合 `Lab assignment*`、`Quiz \d+` 这类固定命名。

优先级：**UID 映射 > 标题规则 > `UNMAPPED-COURSE`**。未标注的事件会明确
显示为未标注，不会静默漏掉。

### 调课 / 补课

公告不会进日历源，登记到 `config\overrides.psd1`：

- `Exclude`：跳过某一次课（注入精确到开始时刻的 `EXDATE`）
- `Add`：新增一次性事件（补课、换教室、考试）
- `Enabled = $false`：保留记录但不生效
- `AllowDuplicate = $true`：与常规课同时间也强制加入

改完重新生成：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\lib\build-merged.ps1
```

### 额外截止

不在 Blackboard 日历源里的作业（只出现在 Course Content / 公告）写到
`config\extra-deadlines.psd1`。不要手改 `data\blackboard.ics`，它每次同步
都会被整体覆盖。

---

## 目录结构

```
hkcc-reminder/
├─ config/
│  ├─ config.psd1            本地配置（不入库）
│  ├─ *.example.psd1         示例配置（入库）
│  ├─ smtp.cred.xml          DPAPI 凭据（不入库）
│  └─ courses / overrides / extra-deadlines / uid-map
├─ data/
│  ├─ timetable.ics          课表（不入库）
│  ├─ timetable.sample.ics   示例课表
│  ├─ blackboard.ics         每天抓取的源（不入库）
│  └─ feed-state.json        事件指纹快照（不入库）
├─ lib/
│  ├─ ics.ps1                iCalendar 解析 / RRULE 展开 / 折行
│  ├─ sync-feed.ps1          抓取 Blackboard 日历源
│  ├─ build-merged.ps1       合并导出 .ics
│  ├─ feed-monitor.ps1       源变动监测
│  ├─ graph-auth.ps1         OAuth2 设备码登录（可选，默认不启用）
│  ├─ mail-core.ps1          读收件箱 / 摘要（可选）
│  └─ i18n.ps1               邮件三语词条
├─ scripts/
│  ├─ setup.ps1              配置向导
│  ├─ daily-reminder.ps1     每日主程序
│  ├─ test-mail.ps1          测试邮件
│  ├─ install-task.ps1       注册/删除计划任务
│  ├─ label-events.ps1       交互式标注
│  ├─ import-label-csv.ps1   CSV 批量标注
│  ├─ init-credential.ps1    只更新凭据
│  ├─ fix-bom.ps1            给 .ps1 补 UTF-8 BOM
│  └─ mail-login.ps1 / mail-dump.ps1  读邮件（可选）
├─ out/                      生成的 .ics / 摘要（不入库）
└─ logs/                     运行日志（不入库）
```

---

## 设计要点（踩过的坑）

- **时区**：所有时间先转 UTC（`Z` 结尾）再输出。香港恒为 UTC+8、无夏令时，
  这样 `.ics` 不依赖客户端时区库，Outlook 不会提示「未知时区」。
- **RRULE**：常规课用 `FREQ=WEEKLY;COUNT=13;BYDAY=xx`；单双周**不用
  `INTERVAL`**，直接展开成独立事件，避免客户端算错。
- **EXDATE 必须精确到开始时刻**，只写当天 `00:00` 会静默失效。
- **`Add` 查重基于展开后的实例**，不是静态 `DTSTART`；否则重复课永远查不到。
- **补课要同时写进 `HKCC_classes.ics`**，因为邮件摘要读的是 classes/deadlines，
  只写 merged 会出现「日历有、邮件没有」。
- **UTF-8 BOM**：PowerShell 5.1 把无 BOM 的 UTF-8 当 GBK 读，中文脚本会语法
  错误。规则：`.ps1` 带 BOM；`.psd1` 纯 ASCII。改完 `.ps1` 跑一次
  `scripts\fix-bom.ps1`。
- **7 天窗口从「今天 00:00」起算**，不是「此刻」，否则当天已上完的课会被
  排除；当天已结束的条目显示 `[已结束]`。
- **窗口外的后续截止**：末尾额外列出未来最近 8 条截止/一次性日程，避免
  「下周一要交的作业落在窗口外」。
- **抓取失败保留旧副本**：Blackboard 源失效时提醒照常发送，退出码区分
  `0` 成功 / `1` 出错 / `2` 缺配置或凭据。

---

## 隐私与安全

- `config\config.psd1`、`smtp.cred.xml`、`graph.cred.xml`、`data\*.ics`、
  `out\`、`logs\`、`acl-recovery\` 都在 `.gitignore` 里，**不会入库**。
- SMTP / Graph 凭据用 Windows DPAPI 加密，只有本机当前账户可解密；仓库里
  只有示例配置，没有任何真实密码或 token。
- Blackboard 的 `FeedUrl` 是个人日历 token，等同于密码，只放本地配置。
- 第一次 `git add` 前建议再扫一遍：
  ```powershell
  git grep -nE "[\w.+-]+@[\w.-]+\.\w+|[0-9]{8}[A-Z]" -- . ':(exclude)*.example.psd1'
  ```
- 仓库已公开；如果 fork 或二次发布，仍建议确认 `README`、示例配置、提交历史
  里没有真实邮箱、学号、课程表和日历源 token。

---

## 已知限制

1. 学校邮箱若强制 OAuth，`Send-MailMessage` 的简单认证会失败；改用应用密码，
   或换 Gmail/QQ 等支持基本认证的邮箱发件。
2. Blackboard 源不含课程名，未标注的条目只能显示标题。
3. **电脑必须在设定的提醒时间开机且已登录**：`LogonType Interactive` 需要用户会话
   （DPAPI 解密），默认 `WakeToRun = $false`，睡眠不会自动唤醒；关机、睡眠、
   注销期间不发送。`-StartWhenAvailable` 会在下次可用时补跑，但邮件会迟到，
   不是设定时间准点。需要准点，可把电脑设为提醒时间前自动唤醒，或把任务改到
   你确定会开机的时段。
4. Blackboard feed URL 在改密码或课程结束后可能失效，需重新获取并更新
   `config.psd1` 的 `FeedUrl`。
5. 没有自动化测试；改动后用 `-Preview` 和 `-NoMail` 验证。
6. 本项目不会替你常驻运行任何东西，调度完全交给 Windows 计划任务。

---

## 状态

个人自用项目，已在 HKCC/CPCE 的真实课表上跑通。读邮件功能
（`graph-auth.ps1` / `mail-core.ps1`）已实现但因租户权限限制默认不启用，
不影响课表、日历和每日邮件。
