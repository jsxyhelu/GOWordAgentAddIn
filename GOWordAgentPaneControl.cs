using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn
{

    public partial class GOWordAgentPaneControl : UserControl
    {

        private readonly List<(string Role, string Content)> _messages = new List<(string Role, string Content)>();
        private DeepSeekService _deepSeekService;
        private List<object> _messageHistory;
        private Button btnFetchDoc;

        public GOWordAgentPaneControl()
        {
            _messageHistory = new List<object>();
            InitializeComponent();

            // 动态创建“纠错”按钮，避免修改 Designer 文件
            btnFetchDoc = new Button
            {
                Name = "btnFetchDoc",
                Text = "纠错（审阅）",
                AutoSize = false,
                Size = new Size(90, 28),
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            // 工具提示
            try
            {
                var tt = new ToolTip();
                tt.SetToolTip(btnFetchDoc, "检测当前文档中的中文错别字并以批注/高亮写回。");
            }
            catch { }

            // 尝试将按钮放在发送按钮左侧
            try
            {
                if (btnSend != null)
                {
                    // 计算位置：与 btnSend 平行并在其左侧 6 像素
                    btnFetchDoc.Location = new Point(btnSend.Left - btnFetchDoc.Width - 6, btnSend.Top);
                }
                else
                {
                    // 若 btnSend 不存在，放在靠底部靠右的默认位置
                    btnFetchDoc.Location = new Point(this.Width - btnFetchDoc.Width - 10, this.Height - btnFetchDoc.Height - 10);
                }
            }
            catch
            {
                // 忽略布局错误，Windows 会自动布局
            }
            btnFetchDoc.Click += BtnFetchDoc_Click;
            // 添加到控件集合
            this.Controls.Add(btnFetchDoc);

            // UI 调整：隐藏聊天框边框
            if (txtChatHistory != null)
            {
                txtChatHistory.BorderStyle = BorderStyle.None;
                // 允许在容器变化时上下左右拉伸，占用尽可能多的空间
                txtChatHistory.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                // 设置最小高度，避免被压缩到看不见
                try
                {
                    txtChatHistory.MinimumSize = new Size(100, 120);
                }
                catch { }
                txtChatHistory.Margin = new Padding(4);
            } 

            // 状态标签布局
            if (lblStatus != null)
            {
                lblStatus.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
                try
                {
                    lblStatus.MinimumSize = new Size(50, lblStatus.Height);
                }
                catch { }
            }

            // 发送按钮布局
            if (btnSend != null)
            {
                btnSend.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                try
                {
                    btnSend.MinimumSize = new Size(btnSend.Width, btnSend.Height);
                }
                catch { }
            }

            // txtInput 布局与快捷键
            if (txtInput != null)
            {
                txtInput.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                txtInput.Margin = new Padding(0);
                try
                {
                    txtInput.MinimumSize = new Size(100, txtInput.Height);
                }
                catch { }

                // 添加键盘快捷键处理：Ctrl+Enter 发送
                txtInput.KeyDown += TxtInput_KeyDown;
            }

            // 控件加载事件
            this.Load += GOWordAgentPaneControl_Load;

            // 调整 btnFetchDoc 在 Resize 时的位置
            this.Resize += (s, e) =>
            {
                try
                {
                    if (btnSend != null && btnFetchDoc != null)
                    {
                        btnFetchDoc.Location = new Point(btnSend.Left - btnFetchDoc.Width - 6, btnSend.Top);
                    }
                    else if (btnFetchDoc != null)
                    {
                        btnFetchDoc.Location = new Point(this.Width - btnFetchDoc.Width - 10, this.Height - btnFetchDoc.Height - 10);
                    }
                }
                catch { }
            };
        }

        private void TxtInput_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+Enter 发送
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                // 阻止系统发出提示音或插入换行
                e.SuppressKeyPress = true;

                // 只有在发送按钮可用时触发发送
                if (btnSend != null && btnSend.Enabled)
                {
                    // 调用发送逻辑（按钮事件为 async void，直接调用即可）
                    btnSend_Click_1(btnSend, EventArgs.Empty);
                }
            }
        }

        private void GOWordAgentPaneControl_Load(object sender, EventArgs e)
        {
            try
            {
                // 注意：目前 API Key 为硬编码。生产环境请从安全存储读取并避免在源码中暴露。
                _deepSeekService = new DeepSeekService("sk-db6aab7933a2427497761018834fe5b1");
                SetStatus("状态: 自动完成大模型注册，API Key 已设置", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatus("状态: 大模型注册失败", Color.Red);
                AppendMessage("错误", ex.Message, Color.Red);
            }
        }

        /*private void btnConn_Click(object sender, EventArgs e)
        {
            _deepSeekService = new DeepSeekService("sk-db6aab7933a2427497761018834fe5b1");
            // 使用线程安全的 SetStatus 来更新 UI
            SetStatus("状态: API Key 已设置，可以开始对话", Color.Green);
        }*/

        private async void btnSend_Click_1(object sender, EventArgs e)
        {
            if (_deepSeekService == null)
            {
                MessageBox.Show("请先设置 API Key", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtInput.Text))
            {
                MessageBox.Show("请输入消息", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string userMessage = txtInput.Text.Trim();

            // 显示用户消息
            AppendMessage("用户", userMessage, Color.Blue);

            // 清空输入框（线程安全）
            UiInvoke(() => txtInput.Clear());

            // 禁用发送按钮并更新状态
            UiInvoke(() => btnSend.Enabled = false);
            SetStatus("状态: AI 正在思考...", Color.Orange);

            // 添加到历史记录
            _messageHistory.Add(new { role = "user", content = userMessage });

            try
            {
                // 调用 API
                string response = await _deepSeekService.SendMessagesWithHistoryAsync(_messageHistory.ToArray());

                // AppendMessage 已内部处理跨线程，这里直接调用
                AppendMessage("DeepSeek", response, Color.Green);

                // 添加到历史记录
                _messageHistory.Add(new { role = "assistant", content = response });

                // 更新状态（线程安全）
                SetStatus("状态: 就绪", Color.Green);
            }
            catch (Exception ex)
            {
                // AppendMessage 已处理跨线程
                AppendMessage("错误", ex.Message, Color.Red);
            }
            finally
            {
                // 恢复发送按钮（线程安全）
                UiInvoke(() => btnSend.Enabled = true);
            }
        }   

        // 点击按钮：读取文档、调用 AI、并将诊断写回（批注、高亮、必要时替换并标红）
        private async void BtnFetchDoc_Click(object sender, EventArgs e)
        {
            // 防止重复点击
            UiInvoke(() => btnFetchDoc.Enabled = false);
            SetStatus("状态: 正在读取文档正文...", Color.Orange);

            string docText = null;
            try
            {
                // 在主线程读取 Word 文档内容
                UiInvoke(() =>
                {
                    var app = Globals.ThisAddIn?.Application;
                    if (app == null)
                        throw new InvalidOperationException("无法访问 Word 应用。");

                    var doc = app.ActiveDocument;
                    if (doc == null)
                        throw new InvalidOperationException("当前未打开任何文档。");

                    var contentRange = doc.Content;
                    docText = contentRange?.Text ?? string.Empty;
                });
            }
            catch (Exception ex)
            {
                UiInvoke(() =>
                {
                    AppendMessage("错误", $"获取文档正文失败: {ex.Message}", Color.Red);
                    SetStatus("状态: 获取失败", Color.Red);
                    btnFetchDoc.Enabled = true;
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(docText))
            {
                UiInvoke(() =>
                {
                    AppendMessage("提示", "文档正文为空或仅包含不可见字符。", Color.Orange);
                    SetStatus("状态: 文档为空", Color.Orange);
                    btnFetchDoc.Enabled = true;
                });
                return;
            }

            // 构建提示词并调用 AI
            string prompt = "你是一个专业的中文校对助手，任务仅限于检测错别字（即用字或字形错误）。" +
                            "不要检查语法、逻辑、事实或风格；仅发现并列出文本中的错别字。" +
                            "请严格以纯 JSON 数组形式输出，数组每个元素为一个对象，字段如下：" +
                            "  - excerpt: 有错的原始片段（尽量简短，不超过200字符），" +
                            "  - suggestion: 正确写法或替换建议，" +
                            "  - severity: 严重性，可选值：low/medium/high，" +
                            "  - 可选字段 context_before/context_after：分别为错字前后各最多40字符的上下文片段，便于在文档中精确定位（可为空或省略）。" +
                            "不要输出任何额外说明或解释、也不要在 JSON 之外添加文本，严格只返回 JSON 数组。" +
                            "\n\n正文如下：\n\n" + docText;

            SetStatus("状态: 正在向 DeepSeek 请求诊断...", Color.Orange);
            AppendMessage("系统", "已发送诊断请求，等待 AI 返回...", Color.Gray);

            string aiResponse = null;
            try
            {
                aiResponse = await _deepSeekService.SendMessageAsync(prompt);
            }
            catch (Exception ex)
            {
                UiInvoke(() =>
                {
                    AppendMessage("错误", $"调用 DeepSeek 失败: {ex.Message}", Color.Red);
                    SetStatus("状态: AI 调用失败", Color.Red);
                    btnFetchDoc.Enabled = true;
                });
                return;
            }

            // 解析 AI 返回的 JSON
            JArray issues = null;
            try
            {
                issues = JArray.Parse(aiResponse);
            }
            catch (Exception)
            {
                UiInvoke(() =>
                {
                    AppendMessage("错误", "AI 返回内容无法解析为 JSON。原文如下：\n" + aiResponse, Color.Red);
                    SetStatus("状态: AI 返回非 JSON", Color.Red);
                    btnFetchDoc.Enabled = true;
                });
                return;
            }

            // 将诊断摘要追加到聊天历史
            try
            {
                var sb = new StringBuilder();
                int idx = 1;
                foreach (var item in issues)
                {
                    string excerpt = item["excerpt"]?.ToString();
                    string suggestion = item["suggestion"]?.ToString();
                    string severity = item["severity"]?.ToString();
                    string ctxBefore = item["context_before"]?.ToString();
                    string ctxAfter = item["context_after"]?.ToString();
                    string reason = item["reason"]?.ToString();

                    sb.AppendLine($"{idx}. {(string.IsNullOrWhiteSpace(excerpt) ? "(无 excerpt)" : excerpt)}");
                    if (!string.IsNullOrWhiteSpace(suggestion))
                    {
                        sb.AppendLine($"   建议: {suggestion}");
                    }
                    if (!string.IsNullOrWhiteSpace(severity))
                    {
                        sb.AppendLine($"   严重性: {severity}");
                    }
                    if (!string.IsNullOrWhiteSpace(reason))
                    {
                        sb.AppendLine($"   理由: {reason}");
                    }
                    if (!string.IsNullOrWhiteSpace(ctxBefore) || !string.IsNullOrWhiteSpace(ctxAfter))
                    {
                        sb.AppendLine($"   上下文: {ctxBefore ?? ""}[{excerpt ?? ""}]{ctxAfter ?? ""}");
                    }
                    sb.AppendLine();
                    idx++;
                }

                string summaryText = sb.Length == 0 ? "未检测到错别字。" : sb.ToString().TrimEnd();
                AppendMessage("纠错结果", summaryText, Color.Purple);
            }
            catch (Exception ex)
            {
                // 如果生成摘要失败，不影响后续写回，只记录错误
                AppendMessage("警告", $"生成可读结果时发生错误: {ex.Message}", Color.Orange);
            }

            // 在 UI 线程对文档进行修改并写回诊断
            UiInvoke(() =>
            {
                try
                {
                    var app = Globals.ThisAddIn?.Application;
                    if (app == null)
                        throw new InvalidOperationException("无法访问 Word 应用。");

                    var doc = app.ActiveDocument;
                    if (doc == null)
                        throw new InvalidOperationException("当前未打开任何文档。");

                    int processed = 0;
                    foreach (var item in issues)
                    {
                        string excerpt = item["excerpt"]?.ToString();
                        string suggestion = item["suggestion"]?.ToString();
                        string severity = item["severity"]?.ToString();
                        string ctxBefore = item["context_before"]?.ToString();
                        string ctxAfter = item["context_after"]?.ToString();
                        string reason = item["reason"]?.ToString();

                        if (string.IsNullOrWhiteSpace(excerpt))
                        {
                            continue;
                        }

                        // 优先尝试按完整 excerpt 精确查找
                        Word.Range searchRange = doc.Content;
                        bool found = searchRange.Find.Execute(FindText: excerpt,
                                                              MatchCase: false,
                                                              MatchWholeWord: false,
                                                              MatchWildcards: false,
                                                              Forward: true,
                                                              Wrap: Word.WdFindWrap.wdFindStop);

                        // 如果未找到且提供上下文，尝试基于上下文查找
                        if (!found && (!string.IsNullOrWhiteSpace(ctxBefore) || !string.IsNullOrWhiteSpace(ctxAfter)))
                        {
                            string combined = (ctxBefore ?? "") + excerpt + (ctxAfter ?? "");
                            // 截断过长的上下文以避免查找失败
                            if (combined.Length > 300)
                            {
                                combined = combined.Substring(0, 300);
                            }
                            searchRange = doc.Content;
                            found = searchRange.Find.Execute(FindText: combined,
                                                             MatchCase: false,
                                                             MatchWholeWord: false,
                                                             MatchWildcards: false,
                                                             Forward: true,
                                                             Wrap: Word.WdFindWrap.wdFindStop);
                        }

                        // 如果仍未找到且 excerpt 较长，尝试用前 80 字进行搜索
                        if (!found && excerpt.Length > 80)
                        {
                            string shortExcerpt = excerpt.Substring(0, 80);
                            searchRange = doc.Content;
                            found = searchRange.Find.Execute(FindText: shortExcerpt,
                                                             MatchCase: false,
                                                             MatchWholeWord: false,
                                                             MatchWildcards: false,
                                                             Forward: true,
                                                             Wrap: Word.WdFindWrap.wdFindStop);
                        }

                        // 再尝试使用正则去除多余空格或换行后的简短匹配（逐句匹配）
                        if (!found)
                        {
                            string normalized = Regex.Replace(excerpt, @"\s+", " ").Trim();
                            if (!string.IsNullOrWhiteSpace(normalized) && normalized.Length > 5)
                            {
                                searchRange = doc.Content;
                                found = searchRange.Find.Execute(FindText: normalized,
                                                                 MatchCase: false,
                                                                 MatchWholeWord: false,
                                                                 MatchWildcards: false,
                                                                 Forward: true,
                                                                 Wrap: Word.WdFindWrap.wdFindStop);
                            }
                        }

                        if (found)
                        {
                            string commentText = BuildCommentText(excerpt, suggestion, reason);
                            try
                            {
                                Word.WdColorIndex highlightIndex = GetHighlightColorForSeverity(severity);

                                if (!string.IsNullOrWhiteSpace(suggestion))
                                {
                                    int replaceStart = searchRange.Start;
                                    searchRange.Text = suggestion;
                                    Word.Range replacedRange = doc.Range(replaceStart, replaceStart + (suggestion.Length));
                                    try { replacedRange.Font.Color = Word.WdColor.wdColorRed; } catch { }
                                    try { replacedRange.HighlightColorIndex = highlightIndex; } catch { }
                                    try { doc.Comments.Add(replacedRange, commentText); } catch
                                    {
                                        try { doc.Comments.Add(doc.Range(0, 0), commentText); } catch { }
                                    }
                                }
                                else
                                {
                                    try { doc.Comments.Add(searchRange, commentText); } catch { }
                                    try { searchRange.HighlightColorIndex = highlightIndex; } catch { }
                                }
                            }
                            catch { }

                            processed++;
                        }
                        else
                        {
                            // 无法定位时，向文档开始处插入一条汇总批注（不包含 severity 标签）
                            Word.Range firstRange = doc.Range(0, 0);
                            string shortExcerpt = excerpt.Length > 60 ? excerpt.Substring(0, 60) + "..." : excerpt;
                            string summary = $"原文: \"{shortExcerpt}\"；建议: {(string.IsNullOrWhiteSpace(suggestion) ? "（无）" : suggestion)}；理由: {(string.IsNullOrWhiteSpace(reason) ? "（未提供）" : reason)}";
                            doc.Comments.Add(firstRange, summary);
                            processed++;
                        }
                    }

                    AppendMessage("DeepSeek", $"已将 {processed} 条诊断以批注形式写入文档（并对有建议的条目直接替换且标红）。", Color.Green);
                    SetStatus("状态: 审阅（批注+替换）已插入", Color.Green);
                }
                catch (Exception ex)
                {
                    AppendMessage("错误", $"将审阅写回文档时发生错误: {ex.Message}", Color.Red);
                    SetStatus("状态: 写回失败", Color.Red);
                }
                finally
                {
                    btnFetchDoc.Enabled = true;
                }
            });
        }

        // 辅助：构建批注文本，包含修改前、修改后与理由（不显示 severity 标签）
        private string BuildCommentText(string original, string suggestion, string reason)
        {
            var sb = new StringBuilder();
            sb.AppendLine("原文: " + (original ?? "（无）"));
            sb.AppendLine("修改后: " + (string.IsNullOrWhiteSpace(suggestion) ? "（无建议）" : suggestion));
            //sb.AppendLine("理由: " + (string.IsNullOrWhiteSpace(reason) ? "（未提供）" : reason));
            return sb.ToString().Trim();
        }

        // 根据 severity 返回适当的高亮颜色（尽量接近浅绿 / 浅黄 / 浅红）
        private Word.WdColorIndex GetHighlightColorForSeverity(string severity)
        {
            if (string.IsNullOrWhiteSpace(severity))
                return Word.WdColorIndex.wdYellow;

            switch (severity.Trim().ToLowerInvariant())
            {
                case "low":
                    return Word.WdColorIndex.wdBrightGreen; // 浅绿
                case "medium":
                    return Word.WdColorIndex.wdYellow; // 浅黄
                case "high":
                    return Word.WdColorIndex.wdPink; // 浅红（pink 近似浅红）
                default:
                    return Word.WdColorIndex.wdYellow;
            }
        }

        private void AppendMessage(string sender, string message, Color color)
        {
            if (txtChatHistory.InvokeRequired)
            {
                txtChatHistory.Invoke(new Action(() => AppendMessage(sender, message, color)));
                return;
            }

            txtChatHistory.SelectionStart = txtChatHistory.TextLength;
            txtChatHistory.SelectionLength = 0;

            txtChatHistory.SelectionColor = color;
            txtChatHistory.SelectionFont = new Font("微软雅黑", 10, FontStyle.Bold);
            txtChatHistory.AppendText($"[{sender}] {DateTime.Now:HH:mm:ss}\n");

            txtChatHistory.SelectionColor = Color.Black;
            txtChatHistory.SelectionFont = new Font("微软雅黑", 10, FontStyle.Regular);
            txtChatHistory.AppendText($"{message}\n\n");

            txtChatHistory.ScrollToCaret();
        }

        private void SetStatus(string text, Color color)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => SetStatus(text, color)));
                return;
            }
            lblStatus.Text = text;
            lblStatus.ForeColor = color;
        }

        /// <summary>
        /// 将任意对 UI 的操作通过控件线程安全地执行。
        /// 使用 this.InvokeRequired 判断，保证在非 UI 线程也能安全更新控件。
        /// </summary>
        /// <param name="action">要在 UI 线程执行的操作</param>
        private void UiInvoke(Action action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
