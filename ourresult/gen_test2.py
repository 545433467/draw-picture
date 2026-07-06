#!/usr/bin/env python3
"""测试：只嵌入 cytoscape.min.js，不加任何插件，使用真实 Excel 数据"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import read_excel, build_graph, _make_cytoscape_style

_DIR = os.path.dirname(os.path.abspath(__file__))
cy_path = os.path.join(_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape.min.js")

records  = read_excel(os.path.join(_DIR, "sample.xlsx"))
elements = build_graph(records)
print("元素数:", len(elements))

cy_js    = open(cy_path, encoding="utf-8").read()
elems_js = json.dumps(elements, ensure_ascii=True)   # ASCII-safe
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
window.onerror = function(m,s,l){ info.style.background="#fcc"; info.textContent="JS错误: "+m+" line:"+l; return false; };
if (typeof cytoscape === "undefined") {
  info.textContent = "ERROR: cytoscape 未定义";
} else {
  try {
    var cy = cytoscape({
      container: document.getElementById("cy"),
      elements: """ + elems_js + """,
      style: """ + style_js + """,
      layout: {name:"cose", animate:false, padding:50}
    });
    var h = document.getElementById("cy").offsetHeight;
    info.textContent = "OK: 节点=" + cy.nodes().length + " 边=" + cy.edges().length + " 容器高=" + h + "px";
  } catch(e) {
    info.style.background = "#fcc";
    info.textContent = "cy初始化失败: " + e;
  }
}
</script></body></html>"""

out = os.path.join(_DIR, "test_noPlugin.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("生成:", out, len(html)//1024, "KB")
