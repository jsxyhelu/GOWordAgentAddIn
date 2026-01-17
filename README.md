# GOWordAgentAddIn

GOWordAgentAddIn 是一个用于 Microsoft Word 的 Add-In，集成了 AI 校对功能。它会将文档发送给语言模型，获取错别字诊断，并将诊断以审阅形式写回到文档中。

主要功能

- 将当前文档正文发送给 AI，获取错别字检测结果（要求 AI 返回严格的 JSON 数组）。
- 对于可定位的错别字条目，Add-In 会：
  - 添加批注，批注内容包含：原文、修改后、理由；
  - 使用不同颜色高亮（根据严重性），并在有建议时将替换后的文本设置为红色；
- 在无法定位的条目处，Add-In 会在文档开头插入汇总批注，说明原文、建议与理由。

使用

1. 在 Visual Studio 中打开解决方案并编译（项目目标: .NET Framework 4.8）。
2. 安装 NuGet 包 `Newtonsoft.Json`。
3. 启动 Word 并加载 Add-In。打开要校对的文档，点击面板中的 `纠错（审阅）` 按钮。

注意

- 当前示例中 API Key 被硬编码在 `GOWordAgentPaneControl.cs` 的 `GOWordAgentPaneControl_Load` 中，请在生产环境中改为从安全存储读取。
- AI 返回的 JSON 结构示例（数组元素）：

```json
[
  {
    "excerpt": "错别字原文",
    "suggestion": "建议替换为",
    "severity": "low",
    "context_before": "前文",
    "context_after": "后文",
    "reason": "修改理由"
  }
]
```

许可

本项目为示例代码，请根据需要调整并遵守使用的第三方库许可。
