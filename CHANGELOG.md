# 变更记录（CHANGELOG）

> **本文件顶部条目 = 当前发布版本的唯一事实源。** 版本号须与 `SKILL.md` frontmatter 的 `version`、GitHub tag、团队技能面板上传时填写的 version 保持一致。
>
> 格式：倒序（最新在上）。类型：`feat` 新功能 / `fix` 修复 / `docs` 文档 / `test` 测试 / `refactor` 重构。

---

## v1.0.0 — 2026-11-23

- **类型**：feat（基线版本）
- **发布人**：技能作者（项目负责人）
- **变更摘要**：
  - 从零多轮实战中固化"智联商机报备材料生成"技能：三份输入 → 五份材料，含防漏填强制校验。
  - 沉淀三处历史漏填/填错点：PPT 第11页维护服务、附件2 成本3行明细与九/十/十一节、PPT 第15页项目类型。
  - **本次改造**：源码脱敏（去除所有具体金额/客户名/站址），校验器改为从 `project_data.json` 动态读取数值基准，换项目可复用；旧项目残留词外置到 `residue-patterns.json`。
  - 新增版本号、CHANGELOG、维护规范、团队使用一页纸、`.gitignore`，转入项目级技能目录，准备纳入 GitHub 私有库版本管理。
- **影响文件**：SKILL.md、scripts/verify_outputs.py（重写）、scripts/extract_data.py（注释）、scripts/fill_utils.py（示例）、references/data-sources.md、references/fill-map.md、新增 5 个文件。
- **本地验证结果**（2026-11-23 实测，具体数值属项目数据，此处只记结论）：
  - `extract_data.py` 基准值核对：通过——含税/不含税收入、税率、成本三项、维护费（元与万元两口径）、毛利率均与真实测算表一致。
  - `verify_outputs.py` 完整模式：`合计 0 项问题`，退出码 0；数值基准来自 `project_data.json`，PPT p11 与附件2 三行明细均与基准一致。
  - 降级模式（无 JSON）：按预期打印 `[WARN] 未提供 project_data.json，数值基准校验已跳过`。
  - 脱敏 grep 扫描：无具体金额/客户名/站址残留；Python 语法检查通过；skill 结构校验通过；zip 共 12 个文件、根目录正确、无缓存、无 `project_data.json`。
- **发布状态**：待负责人推送 GitHub 私有库 `zhilian-bid-materials`（tag `v1.0.0`）+ 待管理员上传珠海智联团队技能面板。
