# M3MVMS

市场部会议可视化管理系统。系统以会议汇总 Excel 为主数据源，提供月历、年度、会议汇总分析、会议详情、资料文件上传下载、会议 Excel 导入导出、会议新增和删除。

## 本地运行

本地运行时，项目根目录需要有 `SM/` 目录：

- `SM/汇总表-市场部市场活动计划汇总.xlsx`
- `SM/每场会议资料文件夹`

启动：

```powershell
python .\server.py
```

或双击：

```text
start_dashboard.vbs
```

访问：

```text
http://127.0.0.1:8765
```

## 会议 Excel 标准字段

导入/导出会议 Excel 使用以下表头：

`序号`、`项目编号`、`类型`、`项目类型`、`项目名称`、`区域`、`活动/会议名称`、`日期`、`地理位置`、`项目负责人`、`参与环节`、`讲题/内容`、`参与嘉宾`、`项目状态`、`关键进度/问题`、`备注`

导入规则：

- `项目编号` 已存在：更新该会议。
- `项目编号` 为空或不存在：新增会议。
- 建议先点击“导出会议Excel”获得标准模板，再编辑后导入。

## Vercel + Supabase Storage

Vercel 函数没有长期可写本地磁盘。云端部署时，主 Excel 和附件应存放在 Supabase Storage。

在 Supabase Storage 创建 bucket，例如：

```text
meeting-files
```

上传主 Excel 到：

```text
data/汇总表-市场部市场活动计划汇总.xlsx
```

Vercel 环境变量建议配置：

```text
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_ANON_KEY=your-anon-key
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
SUPABASE_BUCKET=meeting-files
SUPABASE_MASTER_EXCEL_PATH=data/汇总表-市场部市场活动计划汇总.xlsx
MEETING_DATA_SOURCE=supabase
```

说明：

- `SUPABASE_SERVICE_ROLE_KEY` 只放在 Vercel 服务端环境变量中，不要写入前端。
- 本地优先使用 `SM/`；Vercel 建议设置 `MEETING_DATA_SOURCE=supabase`。
- 附件上传支持 `PDF / Word / PPT / Excel / JPG / PNG`。

## 自动同步

- 页面每 30 秒自动读取 `/api/events`。
- 本地修改并保存 Excel 后，页面会自动刷新。
- 云端新增、删除、导入会议 Excel 后，会写回 Supabase Storage 中的主 Excel。
