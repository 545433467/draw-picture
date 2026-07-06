#!/usr/bin/env python3
import os
_DIR = os.path.dirname(os.path.abspath(__file__))
cy_path = os.path.join(_DIR, "..", "cloudmapper-main", "web", "js", "cytoscape.min.js")

with open(cy_path, "r", encoding="utf-8") as f:
    cy_js = f.read()

html = """<!DOCTYPE html>
<html><head><meta charset="UTF-8"/>
<style>
body{margin:0;height:100vh;display:flex;flex-direction:column}
#info{background:#ffe082;padding:8px;font-size:13px;font-family:monospace;flex-shrink:0}
#cy{flex:1;min-height:0;background:#f5f5f5}
</style></head><body>
<div id="info">加载中...</div>
<div id="cy"></div>
<script>""" + cy_js + """</script>
<script>
var info = document.getElementById("info");
if (typeof cytoscape === "undefined") {
  info.textContent = "ERROR: cytoscape 未定义";
} else {
  info.textContent = "cytoscape 已加载";
  try {
    var cy = cytoscape({
      container: document.getElementById("cy"),
      elements: [
        {data:{id:"a",label:"节点A"}},
        {data:{id:"b",label:"节点B"}},
        {data:{id:"ab",source:"a",target:"b"}}
      ],
      style:[
        {selector:"node",style:{"background-color":"#6FB1FC","label":"data(label)","color":"#000"}},
        {selector:"edge",style:{"line-color":"#999"}}
      ],
      layout:{name:"cose",animate:false}
    });
    var h = document.getElementById("cy").offsetHeight;
    info.textContent = "OK: cy节点=" + cy.nodes().length + " 容器高=" + h + "px";
  } catch(e) {
    info.textContent = "cy初始化失败: " + e;
  }
}
</script></body></html>"""

out = os.path.join(_DIR, "test_minimal.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("Generated:", out, len(html)//1024, "KB")
