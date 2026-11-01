# 手动安装与详细配置

> 本文用于**从源码运行 / 手动配置 / 排错**。普通用户优先用
> [一键安装](../INSTALL.txt)：下载 zip、解压、双击 `install.bat` 即可。

## 1. 前置

- Windows 10/11
- Windows PowerShell 5.1（系统自带）
- 一个支持 SMTP 的邮箱（Office 365 / Gmail / QQ / 163 均可）
- 到设定时间时，电脑开机、联网、当前用户已登录

## 2. 从源码安装

### 2.1 复制示例配置

```powershell
cd <repo>
Copy-Item .\config\config.example.psd1          .\config\config.psd1
Copy-Item .\config\courses.example.psd1         .\config\courses.psd1
Copy-Item .\config\extra-deadlines.example.psd1 .\config\extra-deadlines.psd1
Copy-Item .\config\overrides.example.psd1       .\config\overrides.psd1
Copy-Item .\config\uid-map.example.psd1         .\config\uid-map.psd1
```

`config.psd1` 里至少填 `From`、`To`、`FeedUrl`。

`FeedUrl` 是 Blackboard 的个人日历源（Blackboard → Calendar →
Get external calendar link），**带个人 token，等同密码**，只放在本地配置。

### 2.2 准备课表

```powershell
Copy-Item .\data\timetable.sample.ics .\data\timetable.ics
```

编辑 `data/timetable.ics`：

- `SUMMARY` 用「课程代号 + 名称」
- `DESCRIPTION` 写课型（`Lecture` / `Tutorial` / `Laboratory`）
- `LOCATION` 写教室

格式说明见 [data/README.md](../data/README.md)。

### 2.3 配置邮件与凭据

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup.ps1
```

向导会问发件邮箱、SMTP、收件人、日历源 URL，然后弹出 Windows 凭据窗口
输入**应用专用密码/授权码**（不是网页登录密码）。密码经 DPAPI 加密后写到
`config\smtp.cred.xml`，只有本机、当前 Windows 账户能解密。

也可以手动改 `config\config.psd1` 后，只更新密码：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\init-credential.ps1
```

### 2.4 试跑

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

### 2.5 注册计划任务

```powershell
# -Time 改成你自己的提醒时间；不填默认 08:00
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-task.ps1 -Time "08:00"
```

任务名 `HKCC Daily Reminder`，`LogonType Interactive`（DPAPI 需要用户登录才
能解密），启用了 `-StartWhenAvailable`。

> **这不是云端定时任务，它依赖你的电脑。** 到设定时间，电脑必须开机，并且
> 你已经登录 Windows（锁屏可以，注销不行）。
>
> - **睡眠 / 休眠**：默认不会自动唤醒（`WakeToRun = $false`），唤醒后补跑。
> - **关机**：开机并登录后补跑。
> - **注销**：不运行；重新登录后补跑。
> - 补跑会晚于设定时间；想唤醒睡眠中的电脑，可在「任务计划程序 → 属性 →
>   条件」勾选“唤醒计算机运行此任务”（只对睡眠有效，关机无效）。

```powershell
# 立即跑一次
Start-ScheduledTask -TaskName "HKCC Daily Reminder"

# 看结果（0 = 成功，1 = 出错，2 = 缺凭据/配置）
Get-ScheduledTaskInfo -TaskName "HKCC Daily Reminder" | Format-List
```

## 3. 配置说明

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
| `HideTitles` | 按标题通配符永久隐藏，如 `'Application for XYZ*'` |
| `ClientId` / `Tenant` | 可选，读邮件功能用；留空即禁用 |

> `config.psd1` 必须保持**纯 ASCII**：Windows PowerShell 5.1 解析无 BOM 的
> 非 ASCII `.psd1` 会出错。

### 科目标注

Blackboard 的 `.ics` 不带课程名（`CATEGORIES` / `URL` 为空），所以用两种方式补：

1. **UID 映射**（`config\uid-map.psd1`）：UID 稳定，标一次即可。
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\label-events.ps1
   ```
   或先把 `out\label-report.csv` 用 Excel 填好 `CourseCode` 列，再：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\import-label-csv.ps1
   ```
2. **标题规则**（`config\courses.psd1` 的 `Rules`）：按通配符/正则批量匹配，
   适合 `Lab assignment*`、`Quiz \d+` 这类固定命名。

优先级：**UID 映射 > 标题规则 > `UNMAPPED-COURSE`**。未标注的事件会明确显示为
未标注，不会静默漏掉。

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

## 4. 目录结构

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

## 5. 设计要点（踩过的坑）

- **时区**：所有时间先转 UTC（`Z` 结尾）再输出。香港恒为 UTC+8、无夏令时，
  这样 `.ics` 不依赖客户端时区库，Outlook 不会提示「未知时区」。
- **RRULE**：常规课用 `FREQ=WEEKLY;COUNT=13;BYDAY=xx`；单双周**不用
  `INTERVAL`**，直接展开成独立事件，避免客户端算错。
- **EXDATE 必须精确到开始时刻**，只写当天 `00:00` 会静默失效。
- **`Add` 查重基于展开后的实例**，不是静态 `DTSTART`；否则重复课永远查不到。
- **补课要同时写进 `HKCC_classes.ics`**，因为邮件摘要读的是 classes/deadlines，
  只写 merged 会出现「日历有、邮件没有」。
- **UTF-8 BOM**：PowerShell 5.1 把无 BOM 的 UTF-8 当 GBK 读。规则：`.ps1` 带
  BOM；`.psd1` 纯 ASCII。改完 `.ps1` 跑一次 `scripts\fix-bom.ps1`。
- **7 天窗口从「今天 00:00」起算**，不是「此刻」；当天已结束的条目显示 `[已结束]`。
- **窗口外的后续截止**：末尾额外列出未来最近 8 条截止/一次性日程。
- **抓取失败保留旧副本**：Blackboard 源失效时提醒照常发送，退出码区分
  `0` 成功 / `1` 出错 / `2` 缺配置或凭据。

## 6. 排错

- **退出码 2 / 缺凭据**：先跑 `scripts\init-credential.ps1` 或 `setup.ps1`。
- **测试邮件失败**：检查 SMTP 端口、是否用了应用专用密码/授权码；Office 365
  若强制 OAuth，换 Gmail/QQ 或使用应用密码。
- **没收到每日邮件**：看 `logs\reminder-YYYY-MM.log`；确认到点电脑开机、已登录；
  在任务计划程序里看 `HKCC Daily Reminder` 的最近运行结果。
- **Blackboard 源失效**：重新获取 feed URL，更新 `config.psd1` 的 `FeedUrl`。
- **中文脚本语法错**：`.ps1` 可能丢了 UTF-8 BOM，跑 `scripts\fix-bom.ps1`。
