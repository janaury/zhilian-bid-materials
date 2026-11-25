# -*- coding: utf-8 -*-
"""可复用填写原语：run 级替换、WPS COM 替换、docx 表格/编号插入。

从本模块 import 函数使用，或直接运行看用法说明。
"""
USAGE = """
# 1) run 级替换（保留字体格式）——最常用
from fill_utils import set_run, set_para
from pptx import Presentation
prs = Presentation("模板.pptx")
p = prs.slides[10].shapes[0].text_frame.paragraphs[1]
print([r.text for r in p.runs])      # 先 dump 确认索引
set_run(p, 1, "<项目简称>")          # 具体值取自 project_data.json
set_run(p, 3, "<站址数>")
prs.save("成品.pptx")

# 2) 整段替换（内容塞 run0，其余清空）
set_para(p, "一、采购内容：……")

# 3) PPT 被 WPS 占用时，COM 改写已打开文稿并保存
from fill_utils import com_ppt_replace
com_ppt_replace(keyword="<项目名片段>", slide=11,
                replacements=[("XX项目维护服务", "<项目简称>维护服务")])

# 4) docx 章节标题后插入正文（用 editor_sdk 的 schema 先查参数）
#    见 references/technical-notes.md 第5节
"""

def set_run(p, i, t):
    """替换段落第 i 个 run 文本；越界则新建 run。"""
    if i < len(p.runs):
        p.runs[i].text = t
    else:
        p.add_run().text = t

def set_para(p, t):
    """整段替换：内容写入 run0（保留其格式），其余 run 清空。"""
    if p.runs:
        p.runs[0].text = t
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run().text = t

def dump_para(p):
    """打印段落的 run 结构，便于确定替换索引。"""
    return [(i, r.text) for i, r in enumerate(p.runs)]

def com_ppt_replace(keyword, slide, replacements):
    """通过 WPS/Office COM 改写已打开的 PPT 并保存，不关闭用户窗口。
    keyword: 用于在 Presentations 中定位目标文稿的文件名片段
    slide: 1-based 页码
    replacements: [(old, new), ...] 字符串替换对
    """
    import win32com.client as win32
    app = win32.GetActiveObject("KWpp.Application")
    pres = None
    for p in app.Presentations:
        if keyword in p.Name:
            pres = p
            break
    if pres is None:
        raise RuntimeError(f"未找到含 '{keyword}' 的已打开文稿")
    s = pres.Slides.Item(slide)
    for sh in s.Shapes:
        if sh.HasTextFrame:
            tr = sh.TextFrame.TextRange
            txt = tr.Text
            if any(old in txt for old, _ in replacements):
                for old, new in replacements:
                    if old in txt:
                        tr.Find(old).Text = new
                        txt = txt.replace(old, new)
    pres.Save()
    return pres.Name

def safe_open(path):
    """测试文件是否可写，返回 (writable, err)。"""
    try:
        with open(path, "r+b"):
            return True, None
    except Exception as e:
        return False, str(e)

if __name__ == "__main__":
    print(USAGE)
