
import os, re, io, json, base64, zipfile, sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from PIL import Image
sys.stdout.reconfigure(encoding='utf-8')

BASE = r"D:\AI\healthy"
HTMLF = os.path.join(BASE, "build", "体检参考工具.html")
FILES = {2:"1.健康体检自测问卷（2025年版）.docx",3:"2.基本体检项目（2025年版）.docx",
         4:"3.慢性病风险筛查项目（2025年版）.docx",5:"4.成人健康体检工作流程图（2025年版）.docx",
         6:"5.常见慢性病风险评估指标（2025年版）.docx"}
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def cell_text(c): return "<br>".join(esc(p.text.strip()) for p in c.paragraphs if p.text.strip())
def table_html(t):
    rows=[]
    for ri,row in enumerate(t.rows):
        cells=[]; seen=set()
        for c in row.cells:
            key=id(c._tc)
            if key in seen: continue
            seen.add(key)
            tag="th" if ri==0 else "td"
            cells.append("<%s>%s</%s>"%(tag,cell_text(c),tag))
        rows.append("<tr>"+"".join(cells)+"</tr>")
    return "<table>"+"".join(rows)+"</table>"
def iter_blocks(doc):
    for ch in doc.element.body.iterchildren():
        if ch.tag==qn('w:p'): yield Paragraph(ch,doc)
        elif ch.tag==qn('w:tbl'): yield Table(ch,doc)
def render(path):
    doc=Document(path); out=[]
    for blk in iter_blocks(doc):
        if isinstance(blk,Paragraph):
            t=blk.text.strip()
            if not t: continue
            out.append("<h4>"+esc(t)+"</h4>" if (len(t)<=20 and not t.endswith("。") and not t.endswith("：")) else "<p>"+esc(t)+"</p>")
        else:
            out.append(table_html(blk))
    z=zipfile.ZipFile(path)
    for n in z.namelist():
        if not n.startswith("word/media/"): continue
        if n.endswith("/") or z.getinfo(n).file_size==0: continue
        try:
            im=Image.open(io.BytesIO(z.read(n))).convert("RGB")
            if im.width>1000: im=im.resize((1000,int(im.height*1000/im.width)), Image.LANCZOS)
            buf=io.BytesIO()
            if n.lower().endswith(".png"):
                # 线稿图：转调色板，体积极小
                im.convert("P", palette=Image.ADAPTIVE, colors=48).save(buf,"PNG",optimize=True)
                mime="image/png"
            else:
                im.save(buf,"JPEG",quality=78,optimize=True,progressive=True)
                mime="image/jpeg"
            b64=base64.b64encode(buf.getvalue()).decode()
            out.append('<img src="data:'+mime+';base64,'+b64+'" alt="附件原图">')
            print("   内嵌图 %s %sx%s → %d 字节(base64)" % (n, im.width, im.height, len(b64)))
        except Exception as e:
            print("   图片跳过", n, e)
    return "".join(out)

DOCS_HTML={}; DOCFILES={}
for n,fn in FILES.items():
    p=os.path.join(BASE,fn)
    DOCS_HTML[str(n)]=render(p)
    DOCFILES[str(n)]={"name":fn,"data":"data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,"+base64.b64encode(open(p,"rb").read()).decode()}
    print("  %d 原文 %7d 字节 → 带排版 HTML %5d 字符" % (n, os.path.getsize(p), len(DOCS_HTML[str(n)])))

payload=("var DOCS_HTML="+json.dumps(DOCS_HTML,ensure_ascii=False)+";\nvar DOCFILES="+json.dumps(DOCFILES,ensure_ascii=False)+";\n").replace("</","<\\/")
h=open(HTMLF,encoding="utf-8").read(); before=len(h)
pat=re.compile(r"/\*==DOCPAYLOAD_START==\*/[\s\S]*?/\*==DOCPAYLOAD_END==\*/")
if pat.search(h):
    h=pat.sub(lambda m:"/*==DOCPAYLOAD_START==*/\n"+payload+"/*==DOCPAYLOAD_END==*/", h); mode="替换已有"
else:
    h=h.replace("/*==DOCPAYLOAD==*/","/*==DOCPAYLOAD_START==*/\n"+payload+"/*==DOCPAYLOAD_END==*/"); mode="首次注入"
open(HTMLF,"w",encoding="utf-8").write(h)
print("\n注入方式:",mode)
print("HTML: %.0f KB → %.0f KB" % (before/1024, len(h)/1024))
