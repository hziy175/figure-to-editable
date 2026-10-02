# 清单与审计

脚本仅依赖 Python 标准库，不依赖 OfficeCLI、PowerPoint 或第三方包。它检查可机器核对的内容，无法判断原图是否少列了一个节点。逐区域清单和视觉检查仍由执行者完成。

## 清单格式（schema_version: 1）

```json
{
  "schema_version": 1,
  "source_sha256": "源图的64位十六进制SHA-256",
  "slide_size_pt": [936, 574.5],
  "slide_count": 1,
  "regions": ["A", "B"],
  "inventory_complete": true,
  "items": [
    {
      "name": "FIG_A_title",
      "slide": 1,
      "region": "A",
      "kind": "shape",
      "text": "(A) ΔAUPRC",
      "bounds_pt": [95.25, 0, 367.5, 24]
    },
    {
      "name": "FIG_A_interval_0",
      "slide": 1,
      "region": "A",
      "kind": "shape",
      "line_endpoints_pt": [[227.25, 86.7], [321.45, 86.7]]
    }
  ]
}
```

示例只展示字段，不是完整图的清单。`kind` 使用 `shape`、`connector`、`picture`、`chart` 或 `group`。图片项另填 `raster_reason`。对象名称在同一幻灯片唯一；嵌套组及组内对象都须列出。`bounds_pt` 和 `line_endpoints_pt` 可省略，填写时是**已核对的目标坐标**，容差默认 0.2 pt。源图像素矩形、节点边列表和图表读数继续记录在项目 scene map，不能用 PPT 自动反推并冒充源图证据。

线端点核查支持未旋转、未翻转、未分组的原生直线几何及由 moveTo/lnTo 构成的单段自由路径。对不支持的变换会报错，执行者应使用原生应用或正确的变换解析器完成检查，不要填写虚假坐标让它通过。

## 使用

```text
python scripts/audit_pptx.py --manifest manifest.json --source source.png --planning
python scripts/audit_pptx.py --manifest manifest.json --source source.png --deck final.pptx --preview preview.png --evidence evidence.json --output audit.json
```

最终证据记录格式：

```json
{
  "deck_sha256": "最终PPTX指纹",
  "source_sha256": "源图指纹",
  "preview_sha256": "最终预览指纹",
  "renderer": "实际使用的渲染器名称和必要限制",
  "visual_review": {
    "regions": [{"id": "A", "status": "pass"}, {"id": "B", "status": "pass"}],
    "findings": [],
    "limitations": []
  }
}
```

证据文件须在真实渲染和逐区域对照之后写入。`findings` 保留尚未修复的问题，不能为了通过审计清空。`limitations` 记录实际工具/位图限制，不能替代失败项。

脚本退出码：0 表示所执行的清单/文件检查通过，1 表示检查失败，2 表示输入或文件解析错误。审计输出明确区分“已验证文件属性”与“执行者提供的视觉检查记录”，不将后者包装成计算出的相似度。
