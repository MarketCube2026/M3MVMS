# M3MVMS

市场部会议可视化管理系统。系统以本地 Excel 汇总表为主数据源，自动关联每场会议资料文件夹，并提供月历、年度视图、会议详情、附件入口和自动同步。

## 目录约定

本地运行时需要在项目根目录创建 `SM/` 目录：

- `SM/汇总表-市场部市场活动计划汇总.xlsx`
- `SM/每场会议文件夹/`

`SM/` 包含业务资料，已通过 `.gitignore` 排除，不会上传到 GitHub。

## 启动

```powershell
python .\server.py
```

或双击：

```text
start_dashboard.vbs
```

启动后打开：

```text
http://127.0.0.1:8765
```

## 自动同步

- 页面每 30 秒自动读取 `/api/events`。
- 修改并保存 Excel 总表后，页面会自动刷新。
- 新增或修改会议文件夹、日程、课件、参考材料后，详情页会自动显示最新附件。
- 文件访问被限制在 `SM/` 目录内。

## 开机自启动

运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install_startup_shortcut.ps1
```

该脚本会在当前用户的 Windows 启动文件夹中创建 `MeetingDashboard.lnk`，并通过 `watchdog_dashboard.vbs` 守护本地服务。
