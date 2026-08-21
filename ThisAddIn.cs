using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordMan
{
    public partial class ThisAddIn
    {
        // 每个 Word 窗口一个排版工具窗格，各自独立显示/隐藏，互不干涉
        private readonly Dictionary<Word.Window, Microsoft.Office.Tools.CustomTaskPane> _typesettingPanes =
            new Dictionary<Word.Window, Microsoft.Office.Tools.CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 监听选择变化，自动更新重复标题行按钮状态
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            // 监听窗口切换，用于清理已关闭窗口的排版窗格
            Application.WindowActivate += Application_WindowActivate;
        }

        /// <summary>
        /// 窗口激活时：清理已关闭窗口的排版窗格（窗格按窗口独立显示，互不干涉，不做任何隐藏）
        /// </summary>
        private void Application_WindowActivate(Word.Document doc, Word.Window wn)
        {
            try
            {
                CleanupClosedWindows();
            }
            catch
            {
                // 忽略错误，避免影响正常使用
            }
        }

        /// <summary>
        /// 切换当前窗口的排版工具窗格显示/隐藏（供功能区按钮调用）。
        /// 每个 Word 窗口的窗格独立控制，互不影响：A 窗口的操作只作用于 A 的窗格。
        /// </summary>
        public void ToggleTypesettingPane()
        {
            try
            {
                CleanupClosedWindows();

                var win = Application.ActiveWindow;
                if (win == null) return;

                var pane = GetOrCreateTypesettingPane(win);
                pane.Visible = !pane.Visible;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"切换排版窗格出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 隐藏当前活动窗口的排版窗格（供排版控件关闭按钮调用）
        /// </summary>
        public void HideActiveTypesettingPane()
        {
            try
            {
                var win = Application.ActiveWindow;
                if (win == null) return;

                Microsoft.Office.Tools.CustomTaskPane pane;
                if (_typesettingPanes.TryGetValue(win, out pane) && pane != null)
                {
                    pane.Visible = false;
                }
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 获取（或创建）指定窗口的排版窗格
        /// </summary>
        private Microsoft.Office.Tools.CustomTaskPane GetOrCreateTypesettingPane(Word.Window win)
        {
            Microsoft.Office.Tools.CustomTaskPane pane;
            if (_typesettingPanes.TryGetValue(win, out pane) && pane != null)
            {
                return pane;
            }

            pane = CustomTaskPanes.Add(
                control: new TypesettingTaskPane(),
                title: "排版工具"
            );
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width = 200;

            _typesettingPanes[win] = pane;
            return pane;
        }

        /// <summary>
        /// 清理已关闭窗口对应的窗格（惰性：每次操作前检查）
        /// </summary>
        private void CleanupClosedWindows()
        {
            var closedKeys = new List<Word.Window>();
            foreach (var kvp in _typesettingPanes)
            {
                bool closed = false;
                try
                {
                    if (kvp.Key == null || kvp.Key.Document == null)
                    {
                        closed = true;
                    }
                }
                catch
                {
                    closed = true;
                }

                if (closed)
                {
                    closedKeys.Add(kvp.Key);
                    try
                    {
                        if (kvp.Value != null)
                        {
                            CustomTaskPanes.Remove(kvp.Value);
                        }
                    }
                    catch { }
                }
            }

            foreach (var key in closedKeys)
            {
                _typesettingPanes.Remove(key);
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                // 无论是否在表格中，都更新重复标题行按钮状态
                // 当光标移出表格时，按钮状态会自动更新为未选中
                var ribbon = Globals.Ribbons.GetRibbon<MainRibbon>();
                if (ribbon != null)
                {
                    ribbon.UpdateRepeatHeaderRowsButtonState();
                }
            }
            catch
            {
                // 忽略错误，避免影响正常使用
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 清理资源
            try
            {
                // 移除事件监听
                if (Application != null)
                {
                    Application.WindowSelectionChange -= Application_WindowSelectionChange;
                    Application.WindowActivate -= Application_WindowActivate;
                }

                var ribbon = Globals.Ribbons.GetRibbon<MainRibbon>();
                if (ribbon != null)
                {
                    ribbon.Cleanup();
                }
            }
            catch
            {
                // 忽略清理时的错误，避免影响正常关闭
            }
        }

        #region 全局工具方法
        /// <summary>
        /// 执行操作并将其封装为一个撤销步骤
        /// </summary>
        /// <param name="undoRecordName">撤销记录的名称，将显示在撤销历史中</param>
        /// <param name="action">要执行的操作</param>
        public void ExecuteWithUndoRecord(string undoRecordName, System.Action action)
        {
            if (action == null) return;

            Word.UndoRecord undoRecord = null;
            try
            {
                var doc = Application.ActiveDocument;
                
                if (doc == null) return;
                
                // 开始自定义撤销记录
                undoRecord = doc.Application.UndoRecord;
                undoRecord.StartCustomRecord(undoRecordName);
                
                // 执行操作
                action();
            }
            finally
            {
                // 确保撤销记录已结束（无论是否出错）
                if (undoRecord != null)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                    }
                    catch { }
                }
            }
        }
        #endregion

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
