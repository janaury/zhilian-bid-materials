# -*- coding: utf-8 -*-
"""从 项目前评估测算表.xls 与 智联项目规范运维前置审核要点.xlsx 抽取结构化项目数据。

用法:
    python extract_data.py <测算表.xls> <运维要点.xlsx> [输出.json]

说明:
  - 中文路径/文件名在命令行传递可能乱码。若传参失败，可直接在 Python 内 import 本模块，
    调用 read_calc(path) / read_ops(path) / derive(d)，路径用 glob 在脚本内解析。
  - 采用"标签单元格右邻取值"策略，并用数值类型过滤表头行，避免把"合同年限→总支出"误判。
  - 合同 .doc 需另用编辑器 SDK/COM 读取补充站址字段（见 references/technical-notes.md）。
"""
import sys, json, re

def num(v):
    try:
        if v in ("", None): return None
        return float(v)
    except (ValueError, TypeError):
        return None

def clean(v):
    return re.sub(r"\s+", "", str(v)).strip()

def read_calc(path):
    """读取测算表，返回字段字典。

    优先使用"转置表"sheet：第0行是表头、第1行是对应数值（一列一指标），
    这是最干净的数据源；概况文本字段用"标签右邻"策略补充。
    """
    import xlrd
    wb = xlrd.open_workbook(path)
    rows = []
    for ws in wb.sheets():
        for r in range(ws.nrows):
            rows.append([ws.cell_value(r, c) for c in range(ws.ncols)])

    # 1) 转置表映射：表头单元格 -> 同列数值
    transposed = {}
    for ws in wb.sheets():
        for r in range(ws.nrows - 1):
            head = [clean(ws.cell_value(r, c)) for c in range(ws.ncols)]
            if head and head[0] == "模板ID":
                for c, h in enumerate(head):
                    n = num(ws.cell_value(r + 1, c))
                    if h and n is not None and h not in transposed:
                        transposed[h] = n
                break

    def tget(key):
        """按子串从转置表取数值。"""
        for h, v in transposed.items():
            if key in h:
                return v
        return None

    def find_right(label_key, numeric=True, exact=False, skip_labels=()):
        """找到含 label_key 的单元格，取其右侧第一个符合类型的值。"""
        target = clean(label_key)
        for row in rows:
            for i, cell in enumerate(row):
                c = clean(cell)
                hit = (c == target) if exact else (target in c)
                if not hit:
                    continue
                for j in range(i + 1, len(row)):
                    v, cv = row[j], clean(row[j])
                    if cv == "":
                        continue
                    if numeric:
                        n = num(v)
                        if n is not None:
                            return n
                    else:
                        if cv in skip_labels or cv == target:
                            continue
                        return str(v).strip()
        return None

    data = {}
    # 文本字段（概况区，label|value 横向）
    for key, label in [
        ("project_name", "项目名称"),
        ("client", "客户名称"),
        ("project_class", "项目大类"),
        ("project_subclass", "项目子项"),
    ]:
        v = find_right(label, numeric=False,
                       skip_labels=("项目大类", "项目子项", "客户名称", "项目名称"))
        if v is not None: data[key] = v

    # 数值字段：转置表优先，find_right 兜底
    def pick(tkey, flabel=None, exact=False):
        v = tget(tkey)
        if v is None and flabel:
            v = find_right(flabel, numeric=True, exact=exact)
        return v

    fields = {
        "contract_years":   ("合同年限", "合同年限", False),
        "construction_months": (None, "预计建设工期", False),
        "prepay_ratio":     (None, "首批预付款比例", False),
        "income_net":       ("净额收入", "净额收入", True),
        "income_with_tax":  ("含税收入合计", "含税收入合计", False),
        "cost_maintain":    ("运行维护成本", "运行维护成本", False),
        "cost_electric":    ("电力成本", "电费", True),
        "cost_parent_res":  ("占用母公司资源应结算费用", "占用母公司资源应结算费用", False),
        "cost_total":       ("运营成本支出", "运营成本合计", False),
        "profit":           ("项目毛利润", "项目毛利润", False),
        "npv":              (None, "财务净现值", False),
        "investment":       ("项目投资合计", "项目投资合计", False),
        "gross_margin":     ("毛利率（导入）", "含代垫资金回收", False),
        "gross_margin_ex_tower": ("剔除塔租", "剔除塔租", False),
    }
    for key, (tkey, flabel, exact) in fields.items():
        v = pick(tkey, flabel, exact) if tkey else find_right(flabel, numeric=True, exact=exact)
        if v is not None: data[key] = v
    return data

def read_ops(path):
    """读取运维前置审核要点，抽取运维模式与单价。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    parts.append(str(c.value))
    blob = "\n".join(parts)
    out = {}
    out["ops_mode"] = "一体代维" if "一体代维" in blob else "未知"
    # 共享巡检单价：按"元/次"匹配（D9）
    m = re.search(r"共享巡检[：:]?\s*(\d+)\s*元\s*/\s*次", blob)
    out["inspect_price"] = int(m.group(1)) if m else None
    # 不上塔维修单价：必须按"元/站"匹配，避开"元/次"的历史报价干扰（D14）
    m = re.search(r"不上塔\s*(\d+)\s*元\s*/\s*站", blob)
    out["no_climb_price"] = int(m.group(1)) if m else None
    # 上塔价：(?<!不) 排除"不上塔"
    m = re.search(r"(?<!不)上塔[，,]?\s*(\d+)\s*元", blob)
    out["climb_price"] = int(m.group(1)) if m else None
    out["inspect_times_per_year"] = 1 if ("一年一次" in blob or "每年1次" in blob or "1次/年" in blob) else None
    return out

def derive(d):
    """计算派生指标。"""
    inc = num(d.get("income_net")) or 0
    tax_inc = num(d.get("income_with_tax")) or 0
    if inc and tax_inc:
        d["tax_rate"] = round((tax_inc / inc - 1) * 100)
    ins = d.get("inspect_price") or 0
    times = d.get("inspect_times_per_year") or 0
    nocl = d.get("no_climb_price") or 0
    if ins and times and nocl:
        d["maintain_yuan_per_site_year"] = ins * times + nocl
        d["maintain_wan_ex_tax"] = round(d["maintain_yuan_per_site_year"] / 10000, 4)
        rate = 1 + (d.get("tax_rate", 6) / 100.0)
        d["maintain_wan_with_tax"] = round(d["maintain_yuan_per_site_year"] * rate / 10000, 4)
    for k in ("gross_margin", "gross_margin_ex_tower"):
        if d.get(k) is not None:
            d[k.replace("_ex_tower","") + "_pct"] = round(d[k] * 100, 2)
    return d

def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    d = read_calc(sys.argv[1])
    d.update(read_ops(sys.argv[2]))
    d = derive(d)
    name = str(d.get("project_name", ""))
    m = re.search(r"与(.+?)订立的(.+)", name)
    if m:
        d["client_short"] = m.group(1)
        d["project_short"] = m.group(2)
    out = json.dumps(d, ensure_ascii=False, indent=2)
    if len(sys.argv) >= 4:
        with open(sys.argv[3], "w", encoding="utf-8") as f:
            f.write(out)
        print("已写入", sys.argv[3])
    else:
        print(out)

if __name__ == "__main__":
    main()
