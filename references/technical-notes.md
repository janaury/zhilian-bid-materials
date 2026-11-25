# 技术手法与踩坑记录

## 环境依赖

Python 库：`python-docx`、`python-pptx`、`xlrd`（读 .xls）、`openpyxl`（读 .xlsx）、`pywin32`（COM，读 .doc / 操作 WPS）。按需 `python -m pip install python-docx python-pptx openpyxl xlrd pywin32`。

本地文档编辑通道：WorkBuddy 内置 `workbuddy-local-office-edit` skill（`edsdk.py`），路径：
`C:/Users/<用户>/AppData/Local/Programs/WorkBuddy/resources/app.asar.unpacked/resources/plugins/workbuddy-builtin/skills/workbuddy-local-office-edit`
调用前先 `python edsdk.py schema <tool>` 查参数，不要凭记忆猜。

## 1. .doc 合同读取（老式二进制）

`python-docx` 只能读 .docx，.doc 需用以下任一方式：
- **编辑器 SDK**：`edsdk.py call open_file` → `doc_convert_to_html`（output_path 用纯 ASCII 路径，如 `C:/Users/<用户>/contract_out.html`，中文路径会失败）→ 用 `re.sub` 去 style/head 标签、把 `</td>` 转 ` | `、`</p|div|tr>` 转换行，清洗成纯文本。
- **COM**：`win32com` 调 Word/WPS 打开导出文本。
- 也可 `textract`、`antiword`（若可用）。

## 2. 保留格式的 run 级替换（最重要）

模板段落被切成多个 run（字体/加粗各异）。**整段 `p.text = xxx` 会丢失格式**。正确做法：

```python
def set_run(p, i, t):          # 替换第 i 个 run
    if i < len(p.runs): p.runs[i].text = t
    else: p.add_run().text = t

def set_para(p, t):            # 整段替换：内容塞进 run0，其余 run 清空
    if p.runs:
        p.runs[0].text = t
        for r in p.runs[1:]: r.text = ""
    else:
        p.add_run().text = t
```

替换前先 dump `[r.text for r in p.runs]` 确认索引，例如第11页 run 是 `['一、采购内容：','XX','项目','XX','个铁塔的维护费，','。','服务','期为X年。']` 这种切法，B/A 类字母常单独成 run。

## 3. PPT 被 WPS 占用无法写盘（历史踩坑）

现象：`prs.save(f)` 抛 PermissionError；目录里有 `~$xxx.pptx` 锁文件；`Get-Process` 显示 `wpp`（WPS 演示）占用。**不要强关用户窗口**。改用 COM 连接已打开文稿：

```python
import win32com.client as win32
app = win32.GetActiveObject("KWpp.Application")   # WPS 演示
pres = next(p for p in app.Presentations if "项目名" in p.Name)
slide = pres.Slides.Item(11)
for sh in slide.Shapes:
    if sh.HasTextFrame and "分项三" in sh.TextFrame.TextRange.Text:
        tr = sh.TextFrame.TextRange
        for old, new in [("XX项目维护服务","杆塔资源综合服务项目维护服务"), ...]:
            # 逐 run 替换用 Paragraphs(i).Runs(j).Text 或 Find
            tr.Find(old).Text = new   # 或更精细的 run 定位
pres.Save()
```

要点：WPS COM 的集合不是标准 Python 序列，`Paragraphs(i)`/`Runs(i)` 用调用或索引形式需实测；最稳是用 `TextRange.Find(old).Text = new` 直接替换唯一字符串。保存后用 python-pptx 从磁盘独立校验，并逐页 dump 比对确认只改目标页。

**删除被占用旧文件**：安全删除机制/系统 trash 会失败（sharing violation error 32），`os.remove`/`Remove-Item`/Win32 `DeleteFileW` 都不行。不要强删——在回复里告知用户关闭 WPS 后自行删除旧模板。

## 4. docx 表格合并单元格（附件2 配置与成本测算表）

用 `edsdk.py call doc_get_table_info` 拿 `cells` 数组，每个 cell 有 `row`/`col`/`mc`/`v_merge` 字段：
- `mc` 真 = 横向合并的占位格（gridSpan 延续），不要写。
- `v_merge` 非空 = 纵向合并格，判断是不是起始格再写。
- 只在真实起始格写值，否则内容写进不可见格。

写入用 `doc_set_table_cells --json-file <args.json>`，参数形如 `{"file_id":..., "table_id":..., "cells":[{"row":13,"col":4,"text":"..."}, ...]}`（行号 1-based，先 `schema` 确认）。

## 5. docx 中文自动编号标题后插入正文（附件2 九十十一节）

三个标题用 numId=1 中文编号（一…十一），正文原本为空。步骤：
1. 检查 `w:numPr` / numbering.xml 确认是自动编号，不要手打"九、"。
2. `doc_find` 用**完整标题串**（含"（按照实际情况填报…）"后缀）定位，取 `locations[0].end + 1` 作插入索引。
3. 从后往前插入（`doc_insert_paragraph_with_text`），避免索引漂移；插入的正文不带 numPr，标题编号保持不变。
4. 存盘后校验编号属性：标题段 num=1、正文段 num=None。

## 6. docx 文件被占用 vs 可写

写前先测：
```python
try:
    with open(f, "r+b"): print("WRITABLE")
except Exception as e: print("LOCKED", e)
```
可写时优先用 editor_sdk 编辑并 `save_file` 回原路径（或 python-docx 直接存）。被占用时同上走 COM。

## 7. 校验脚本的正确路径写法

在 Git Bash 里 Python 打开 Windows 文件用 `C:/Users/...` 正斜杠，不要用 `/c/Users/...`（python 不认），也不要相对路径依赖 cwd。

## 8. 填写范围原则

用户要求"只填空缺信息、其余保持模板"时：只替换 XX/____ 占位、矛盾项、残留的上一项目示例；保留示例说明、编写要求、10不准自查话术等模板原文。空的会议时间/地点/参会人等按约定留空，不要自作主张补写。
