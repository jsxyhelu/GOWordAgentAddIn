using System;
using System.IO;
using Microsoft.Office.Tools.Word;

namespace GOWordAgentAddIn
{
    public partial class ThisAddIn
    {
        // 保存当前实例和任务窗格引用，方便 Ribbon 中访问
        internal static ThisAddIn Current;
        internal Microsoft.Office.Tools.CustomTaskPane GOWordAgentPane;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Current = this;

            // 创建侧边栏控件实例
            var control = new GOWordAgentPaneControl();

            // 在 Word 中创建自定义任务窗格，标题也是 GOWordAgent
            GOWordAgentPane = this.CustomTaskPanes.Add(control, "GOWordAgent");
            GOWordAgentPane.Visible = true;      // 启动时默认显示

            // 尝试从上次保存的配置加载宽度，若无则使用默认 400
            int width = LoadSavedPaneWidth() ?? 400;
            GOWordAgentPane.Width = width;

            // 当用户在 UI 中调整任务窗格宽度时（控件 Resize），立即保存宽度
            // 使用控件的 SizeChanged/Resize 事件作为任务窗格宽度变化的代理
            control.SizeChanged += (s, args) =>
            {
                try
                {
                    SavePaneWidth(GOWordAgentPane.Width);
                }
                catch
                {
                    // 忽略持久化错误，避免影响 UI
                }
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 在关闭时再次保存一次当前宽度，双保险
            try
            {
                if (GOWordAgentPane != null)
                {
                    SavePaneWidth(GOWordAgentPane.Width);
                }
            }
            catch { }
        }

        #region VSTO 生成的代码
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        // ---------- 持久化实现 ----------
        private string GetSettingsFilePath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GOWordAgentAddIn");
            return Path.Combine(dir, "paneWidth.txt");
        }

        private int? LoadSavedPaneWidth()
        {
            try
            {
                string path = GetSettingsFilePath();
                if (File.Exists(path))
                {
                    string text = File.ReadAllText(path);
                    if (int.TryParse(text, out int w) && w > 0)
                    {
                        return w;
                    }
                }
            }
            catch
            {
                // 忽略读取错误
            }
            return null;
        }

        private void SavePaneWidth(int width)
        {
            try
            {
                string path = GetSettingsFilePath();
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.WriteAllText(path, width.ToString());
            }
            catch
            {
                // 忽略写入错误，避免抛出影响主流程
            }
        }
    }
}