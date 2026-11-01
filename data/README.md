# data/ — 本地数据（不入库）

这个目录存放运行时数据。除本说明和示例外，`data/*` 都已在 `.gitignore` 里，
因为课表和 Blackboard 源属于个人信息。

| 文件 | 来源 | 说明 |
|---|---|---|
| `timetable.ics` | 你自己维护 | 课表（单双周、单次课都写在这里）。可从 `timetable.sample.ics` 复制改名后修改 |
| `blackboard.ics` | `lib/sync-feed.ps1` 自动抓取 | Blackboard 日历源。每次同步会整体覆盖，**不要手改** |
| `feed-state.json` | `lib/feed-monitor.ps1` 自动生成 | 事件指纹快照，用于发现调课/停课/换教室 |

## 开始使用

```powershell
Copy-Item .\data\timetable.sample.ics .\data\timetable.ics
```

然后编辑 `timetable.ics`，把示例事件换成你的课程。`SUMMARY` 建议
`课程代号 + 名称`，`DESCRIPTION` 写明课型（`Lecture` / `Tutorial` /
`Laboratory`），`LOCATION` 写教室。

## 格式提醒

- 时间用 `DTSTART;TZID=Asia/Hong_Kong:yyyyMMddTHHmmss`
- 每周重复用 `RRULE:FREQ=WEEKLY;COUNT=13;BYDAY=MO`
- 单双周不要用 `INTERVAL`，直接展开成独立事件更稳（见 `README.md`）
- 行尾用 CRLF 或 LF 均可；文件需为 UTF-8
