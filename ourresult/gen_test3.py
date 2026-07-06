#!/usr/bin/env python3
"""验证：去掉 compound nodes，看节点是否能正常显示"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import read_excel, build_graph, _make_cytoscape_style

_DIR = os.path.dirname(os.path.abspath(__file__))
cy_path = os.path.join(_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape.min.js")

records  = read_excel(os.path.join(_DIR, "sample.xlsx"))
elements = build_graph(records)

# 去掉所有 parent 属性，移除 container 节点
flat_elements = []
for e in elements:
    if e.get("group") == "nodes" and e["data"].get("is_container"):
        continue  # 跳过 region/group 容器节点
    d = dict(e["data"])
    d.pop("parent", None)  # 移除 parent 关系
    flat_elements.append({"group": e["group"], "data": d})

print("flat elements:", len(flat_elements))

cy_js    = open(cy_path, encoding="utf-8").read()
elems_js = json.dumps(flat_elements, ensure_ascii=True)
style_js = json.dumps(_make_cytoscape_style(), ensure_ascii=True)

html = """<!DOCTYPE html>
<html><head><meta charset="UTF-8"/>
<style>
body{margin:0;height:100vh;display:flex;flex-direction:column}
#info{background:#ffe082;padding:6px 12px;font-size:13px;font-family:monospace;flex-shrink:0}
#cy{flex:1;min-height:0;background:#fff}
</style></head><body>
<div id="info">加载中...</div>
<div id="cy"></div>
<script>""" + cy_js + """</script>
<script>
var info = document.getElementById("info");
window.onerror = function(m,s,l){ info.style.background="#fcc"; info.textContent="JS错误:"+m; };
try {
  var cy = cytoscape({
    container: document.getElementById("cy"),
    elements: """ + elems_js + """,
    style: """ + style_js + """,
    layout: {name:"cose", animate:false, padding:50}
  });
  info.textContent = "节点=" + cy.nodes().length + " 边=" + cy.edges().length + " 高=" + document.getElementById("cy").offsetHeight + "px";
} catch(e) { info.style.background="#fcc"; info.textContent="失败:"+e; }
</script></body></html>"""

out = os.path.join(_DIR, "test_flat.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("生成:", out)
