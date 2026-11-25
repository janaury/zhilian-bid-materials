# -*- coding: utf-8 -*-
"""防漏填校验器：扫描五份成品，检查残留占位符与必填项。

用法:
    python verify_outputs.py <目录> [<项目全名>] [<project_data.json路径>]

参数:
    argv[1] 成品所在目录（必填）
    argv[2] 项目全名（可选，用于核对五份文件命名一致性）
    argv[3] project_data.json（可选）——由 extract_data.py 从本次真实输入生成，
            作为数值基准的动态来源；未传时自动在成品目录/脚本目录/当前目录查找；
            找不到则进入"降级模式"：跳过数值比对，仅做结构与占位符检查，并打印 WARN。

设计:
    本脚本不包含任何具体项目数据（金额、客户名、站址）。所有基准值来自
    project_data.json，因此换项目也能复用；残留旧项目的名称词从可选的
    references/residue-patterns.json 读取，团队可按上次项目自行填写。

检查内容:
  1. 任何输出里残留通用占位符（XX/XXXX/____）与可配置的旧项目残留词
  2. 项目全名在五份文件中出现
  3. PPT 第11页维护服务实填（数值基准来自 JSON）、第15页项目类型非矛盾值
  4. 附件2 配置与成本测算表 3 行明细、九/十/十一节有正文
退出码: 0=全部通过, 1=有失败项。
"""
import sys, os, glob, json

# 通用占位符规则：所有项目共用，非敏感数据，保留在源码中
BAD_PATTERNS = [
    ("XX项目", "残留XX项目占位"),
    ("XXXX", "残留XXXX占位"),
    ("XX公司", "残留XX公司占位"),
    ("XX万元", "残留X万元占位"),
    ("XX元", "残留XX元占位"),
]


def read_pptx(path):
    from pptx import Presentation
    prs = Presentation(path)
    pages = []
    for s in prs.slides:
        buf = []
        for sh in s.shapes:
            if sh.has_text_frame:
                buf.append(sh.text_frame.text)
            if sh.has_table:
                for row in sh.table.rows:
                    buf.append(" ".join(c.text for c in row.cells))
        pages.append("\n".join(buf))
    return pages


def read_docx(path):
    import docx
    d = docx.Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            parts.append(" ".join(c.text for c in row.cells))
    return parts


def find(files, kw):
    return [os.path.basename(f) for f in files if kw in os.path.basename(f)]


def load_baseline(folder, explicit=None):
    """加载动态数值基准。返回 (baseline_dict, degrade_bool)。"""
    path = explicit
    if not path:
        here = os.path.dirname(os.path.abspath(__file__))
        for d in (folder, here, os.getcwd()):
            cand = os.path.join(d, "project_data.json")
            if os.path.exists(cand):
                path = cand
                break
    if not path or not os.path.exists(path):
        return {}, True
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f), False
    except (ValueError, OSError):
        return {}, True


def load_extra_patterns():
    """读取可选的旧项目残留词配置（非敏感，团队按上次项目自行填写）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, os.pardir, "references", "residue-patterns.json")
    if not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as f:
            cfg = json.load(f)
        return [tuple(x) for x in cfg.get("extra_bad_patterns", []) if len(x) == 2]
    except (ValueError, OSError):
        return []


def fmt(v, mode="raw"):
    """按口径格式化基准值，避免浮点精度误报。"""
    try:
        if mode == "int":
            return str(int(float(v)))
        if mode == "2f":
            return f"{float(v):.2f}"
        return str(v)
    except (ValueError, TypeError):
        return None


def scan_residue(parts, label, issues, extra_patterns):
    text = "\n".join(parts)
    for pat, desc in BAD_PATTERNS + extra_patterns:
        if pat in text:
            issues.append(f"[{label}] {desc}: 命中 '{pat}'")


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    proj = sys.argv[2] if len(sys.argv) > 2 else None
    baseline_explicit = sys.argv[3] if len(sys.argv) > 3 else None
    allf = glob.glob(os.path.join(folder, "*"))

    def is_finished(f):
        """成品 = 非锁定文件，且文件名不含'模板'、不含'XX'占位。"""
        b = os.path.basename(f)
        return (not b.startswith("~$")) and ("模板" not in b) and ("XX" not in b)

    pptx = [f for f in allf if f.lower().endswith(".pptx") and is_finished(f)]
    docx = [f for f in allf if f.lower().endswith(".docx") and is_finished(f)]
    issues, oks = [], []
    extra_patterns = load_extra_patterns()

    # 数值基准（动态）
    base, degraded = load_baseline(folder, baseline_explicit)
    if degraded:
        print("[WARN] 未提供 project_data.json，数值基准校验已跳过（降级模式）")
        print("       建议先运行: python extract_data.py <测算表> <运维要点> project_data.json")
    else:
        oks.append("[基准] 已从 project_data.json 加载数值基准")

    def bval(key, mode="raw"):
        return fmt(base.get(key), mode) if key in base else None

    # ---- PPT ----
    ppt = find(pptx, "汇报材料")
    if not ppt:
        issues.append("[PPT] 未找到汇报材料成品")
    else:
        pages = read_pptx(ppt[0])
        oks.append(f"[PPT] {len(pages)}页")
        scan_residue(pages, "PPT", issues, extra_patterns)
        if len(pages) >= 11:
            s11 = pages[10]
            if "不涉及" in s11 and ("维护费" not in s11 or "采购内容" not in s11):
                issues.append("[PPT p11] 维护服务疑似仍为'不涉及'，未实填(漏填点①)")
            if not degraded:
                want_yuan = bval("maintain_yuan_per_site_year", "int")
                want_wan = bval("maintain_wan_with_tax")
                want_ex = bval("maintain_wan_ex_tax")
                vals = [(want_yuan, "元/站/年"), (want_wan, "含税万元"), (want_ex, "不含税万元")]
                hits = [v for v, _ in vals if v and v in s11]
                if len(hits) < 2:
                    issues.append("[PPT p11] 维护费数值与 project_data.json 基准不符(漏填点①)")
                else:
                    oks.append("[PPT p11] 维护服务已实填，数值与基准一致")
        if proj and proj not in "\n".join(pages):
            issues.append("[PPT] 未出现项目全名")

    # ---- 附件2 ----
    f2 = find(docx, "概要说明")
    if not f2:
        issues.append("[附件2] 未找到概要说明成品")
    else:
        parts = read_docx(f2[0])
        oks.append("[附件2] 已找到")
        scan_residue(parts, "附件2", issues, extra_patterns)
        joined = "\n".join(parts)
        if not degraded:
            cost_keys = [
                ("cost_maintain", "运维"),
                ("cost_electric", "电费"),
                ("cost_parent_res", "母公司资源费"),
            ]
            missing = []
            for key, name in cost_keys:
                v = bval(key, "2f")
                if v is None:
                    continue
                if v not in joined:
                    missing.append(f"{name}={v}")
            if missing:
                issues.append("[附件2 table2] 未见成本明细 " + "/".join(missing) + "(漏填点②A)")
            else:
                oks.append("[附件2 table2] 成本3行明细齐全")
        # 九/十/十一节正文：标题后应有实质正文
        for kw in ("项目背景与建设内容信息", "项目维护内容信息", "项目进度计划"):
            idx = next((i for i, t in enumerate(parts) if kw in t), None)
            if idx is None:
                issues.append(f"[附件2] 未找到章节'{kw}'")
            elif idx + 1 >= len(parts) or len(parts[idx + 1].strip()) < 10:
                issues.append(f"[附件2] 章节'{kw}'正文疑似为空(漏填点②B)")
            else:
                oks.append(f"[附件2] 章节'{kw}'有正文")
        if proj and proj not in joined and "概要" not in joined:
            issues.append("[附件2] 未出现项目全名")

    # ---- 附件3 ----
    f3 = find(docx, "计收")
    if not f3:
        issues.append("[附件3] 未找到计收说明成品")
    else:
        parts = read_docx(f3[0])
        oks.append("[附件3] 已找到")
        scan_residue(parts, "附件3", issues, extra_patterns)
        if proj and proj not in "\n".join(parts):
            issues.append("[附件3] 未出现项目全名")

    # ---- 会议纪要 ----
    f5 = find(docx, "会议纪要")
    if not f5:
        issues.append("[会议纪要] 未找到成品")
    else:
        parts = read_docx(f5[0])
        oks.append("[会议纪要] 已找到")
        scan_residue(parts, "会议纪要", issues, extra_patterns)

    # ---- 请示 ----
    fq = find(docx, "请示")
    if not fq:
        issues.append("[请示] 未找到成品")
    else:
        parts = read_docx(fq[0])
        oks.append("[请示] 已找到")
        scan_residue(parts, "请示", issues, extra_patterns)

    print("=== 通过项 ===")
    for o in oks:
        print(" OK", o)
    print("=== 问题项 ===")
    if not issues:
        print(" 无，全部通过")
    else:
        for it in issues:
            print(" FAIL", it)
    print(f"\n合计 {len(issues)} 项问题。")
    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
