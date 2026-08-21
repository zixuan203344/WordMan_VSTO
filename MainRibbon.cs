using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordMan.SplitAndMerge;
using WordMan.MultiLevel;
using WordMan;
using static WordMan.CaptionManager;

namespace WordMan
{
    public partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private ImageProcessor imageProcessor = new ImageProcessor();
        private TextProcessor textProcessor = new TextProcessor();
        private TableProcessor tableProcessor = new TableProcessor();
        private CaptionManager captionManager = new CaptionManager();
        private DocumentProcessor documentProcessor = new DocumentProcessor();
        private SplitAndMerge.DocumentMerger documentMerger;
        private DocumentSplitter documentSplitter;
        private MultiLevelListForm multiLevelListForm;
        private StyleSettings styleSettingsForm; // 样式设置窗口（非模态单实例）

        #region 文本处理组
        // Word 内置功能
        private void 清除格式_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("清除格式", () =>
            {
                textProcessor.ClearFormatting();
            });
        }

        private void 格式刷_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            textProcessor.FormatPainter_Click(toggleButton);
        }

        private void 只留文本_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("只留文本粘贴", () =>
            {
                textProcessor.PasteTextOnly();
            });
        }

        private void 去除断行_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("去除断行", () =>
            {
                textProcessor.RemoveLineBreaks();
            });
        }

        private void 去除空格_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("去除空格", () =>
            {
                textProcessor.RemoveSpaces();
            });
        }

        private void 去除空行_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("去除空行", () =>
            {
                textProcessor.RemoveEmptyLines();
            });
        }

        private void 英标转中标_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("英标转中标", () =>
            {
                textProcessor.ConvertEnglishToChinesePunctuation();
            });
        }

        private void 中标转英标_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("中标转英标", () =>
            {
                textProcessor.ConvertChineseToEnglishPunctuation();
            });
        }

        private void 自动加空格_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("自动加空格", () =>
            {
                textProcessor.AutoAddSpaces();
            });
        }

        private void 缩进2字符_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("缩进2字符", () =>
            {
                textProcessor.IndentTwoCharacters();
            });
        }

        private void 去除缩进_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("去除缩进", () =>
            {
                textProcessor.RemoveIndent();
            });
        }

        private void 希腊字母_Click(object sender, RibbonControlEventArgs e)
        {
            ShowForm<GreekLetterForm>("希腊字母");
        }

        private void 常用符号_Click(object sender, RibbonControlEventArgs e)
        {
            ShowForm<CommonSymbolForm>("常用符号");
        }

        /// <summary>
        /// 统一显示表单的辅助方法
        /// </summary>
        private void ShowForm<T>(string formName) where T : Form, new()
        {
            try
            {
                T form = new T();
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开{formName}窗口失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 仿宋替换_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("字体替换：仿宋", () =>
            {
                textProcessor.ReplaceFangSongGB2312ToFangSong();
            });
        }

        private void 楷体替换_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("字体替换：楷体", () =>
            {
                textProcessor.ReplaceKaiTiGB2312ToKaiTi();
            });
        }

        private void 数字替换_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("字体替换：数字英文", () =>
            {
                textProcessor.ReplaceAllToTimesNewRoman();
            });
        }
        #endregion

        #region 表格处理组
        private void 创建表格_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            var gallery = sender as Microsoft.Office.Tools.Ribbon.RibbonGallery;
            if (gallery == null)
                return;

            int selectedIndex = gallery.SelectedItemIndex;
            if (selectedIndex < 0)
                return;

            Globals.ThisAddIn.ExecuteWithUndoRecord("创建表格", () =>
            {
                if (selectedIndex == 0)
                {
                    tableProcessor.CreateThreeLineTableStyle();
                }
                else if (selectedIndex == 1)
                {
                    tableProcessor.CreateGBTableStyle();
                }
                else if (selectedIndex == 2)
                {
                    tableProcessor.CreateNoBorderTableStyle();
                }
            });
        }

        private void 设置表格_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            var gallery = sender as Microsoft.Office.Tools.Ribbon.RibbonGallery;
            if (gallery == null)
                return;

            int selectedIndex = gallery.SelectedItemIndex;
            if (selectedIndex < 0)
                return;

            Globals.ThisAddIn.ExecuteWithUndoRecord("设置表格", () =>
            {
                if (selectedIndex == 0)
                {
                    tableProcessor.SetCurrentTableToThreeLineStyle();
                }
                else if (selectedIndex == 1)
                {
                    tableProcessor.SetCurrentTableToGBStyle();
                }
                else if (selectedIndex == 2)
                {
                    tableProcessor.SetCurrentTableToNoBorderStyle();
                }
            });
        }

        private void 插入N行_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("插入表格行", () =>
            {
                tableProcessor.InsertNRows();
            });
        }

        private void 插入N列_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("插入表格列", () =>
            {
                tableProcessor.InsertNColumns();
            });
        }

        private void 重复标题行_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            // 直接执行Word内置命令
            Globals.ThisAddIn.ExecuteWithUndoRecord("重复标题行", () =>
            {
                tableProcessor.RepeatHeaderRows();
            });
            // 执行命令后，延迟更新按钮状态以反映实际状态
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 50;
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                if (toggleButton != null)
                {
                    toggleButton.Checked = tableProcessor.GetRepeatHeaderRowsState();
                }
            };
            timer.Start();
        }

        /// <summary>
        /// 更新重复标题行按钮状态（供外部调用）
        /// </summary>
        public void UpdateRepeatHeaderRowsButtonState()
        {
            if (重复标题行 != null)
            {
                重复标题行.Checked = tableProcessor.GetRepeatHeaderRowsState();
            }
        }
        #endregion

        #region 题注与引用组
        private void 图注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置图注样式1", () =>
            {
                captionManager.SetPictureStyle(图注样式1, 图注样式2, 图注样式3, CaptionNumberStyle.Arabic);
            });
        }

        private void 图注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置图注样式2", () =>
            {
                captionManager.SetPictureStyle(图注样式2, 图注样式1, 图注样式3, CaptionNumberStyle.Dash);
            });
        }

        private void 图注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置图注样式3", () =>
            {
                captionManager.SetPictureStyle(图注样式3, 图注样式1, 图注样式2, CaptionNumberStyle.Dot);
            });
        }

        private void 图编号_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("插入图编号", () =>
            {
                captionManager.InsertPictureNumber();
            });
        }

        private void 表注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置表注样式1", () =>
            {
                captionManager.SetTableStyle(表注样式1, 表注样式2, 表注样式3, CaptionNumberStyle.Arabic);
            });
        }

        private void 表注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置表注样式2", () =>
            {
                captionManager.SetTableStyle(表注样式2, 表注样式1, 表注样式3, CaptionNumberStyle.Dash);
            });
        }

        private void 表注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置表注样式3", () =>
            {
                captionManager.SetTableStyle(表注样式3, 表注样式1, 表注样式2, CaptionNumberStyle.Dot);
            });
        }

        private void 表编号_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("插入表编号", () =>
            {
                captionManager.InsertTableNumber();
            });
        }

        private void 公式样式1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置公式样式1", () =>
            {
                captionManager.SetFormulaStyle(公式样式1, 公式样式2, 公式样式3, FormulaNumberStyle.Parenthesis1);
            });
        }

        private void 公式样式2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置公式样式2", () =>
            {
                captionManager.SetFormulaStyle(公式样式2, 公式样式1, 公式样式3, FormulaNumberStyle.Parenthesis1_1);
            });
        }

        private void 公式样式3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置公式样式3", () =>
            {
                captionManager.SetFormulaStyle(公式样式3, 公式样式1, 公式样式2, FormulaNumberStyle.Parenthesis1_1dot);
            });
        }

        private void 式编号_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("插入式编号", () =>
            {
                captionManager.InsertFormulaNumber();
            });
        }

        private void 交叉引用_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            captionManager.ToggleCrossReferenceMode(toggleButton);
        }
        #endregion

        #region 图片处理组
        private void 宽度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.WidthBrush_Click(sender, e, 宽度刷);
        }

        private void 高度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.HeightBrush_Click(sender, e, 高度刷);
        }

        private void 位图化_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("图片位图化", () =>
            {
                imageProcessor.ConvertToBitmap_Click(sender, e);
            });
        }

        private void 导出图片_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.ExportImage_Click(sender, e);
        }

        public void Cleanup()
        {
            imageProcessor.Cleanup();
        }
        #endregion

        #region 全文处理组
        private void TypesettingButton_Click(object sender, RibbonControlEventArgs e)
        {
            TypesettingTaskPane.TriggerShowOrHide();
        }

        private void 样式设置_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 非模态窗口：只打开一个实例，重复点击时激活已有窗口
                if (styleSettingsForm == null || styleSettingsForm.IsDisposed)
                {
                    styleSettingsForm = new StyleSettings();
                    styleSettingsForm.FormClosed += (s, args) => styleSettingsForm = null;
                    styleSettingsForm.Show();

                    // 设置 Word 主窗口为所有者，使窗口始终置于 Word 上方
                    try
                    {
                        var wordHwnd = new IntPtr(Globals.ThisAddIn.Application.ActiveWindow?.Hwnd ?? 0);
                        if (wordHwnd != IntPtr.Zero && styleSettingsForm.IsHandleCreated)
                        {
                            if (IntPtr.Size == 8)
                            {
                                NativeMethods.SetWindowLongPtr64(styleSettingsForm.Handle, NativeMethods.GWL_HWNDPARENT, wordHwnd);
                            }
                            else
                            {
                                NativeMethods.SetWindowLong32(styleSettingsForm.Handle, NativeMethods.GWL_HWNDPARENT, wordHwnd.ToInt32());
                            }
                        }
                    }
                    catch
                    {
                        // 设置所有者失败不影响窗口使用
                    }
                }
                else
                {
                    // 窗口已打开：激活并重新读取当前文档样式（可能已切换文档）
                    if (styleSettingsForm.WindowState == FormWindowState.Minimized)
                    {
                        styleSettingsForm.WindowState = FormWindowState.Normal;
                    }
                    styleSettingsForm.Activate();
                    styleSettingsForm.BringToFront();
                    try
                    {
                        styleSettingsForm.RefreshFromDocument();
                    }
                    catch
                    {
                        // 重新读取失败不影响窗口使用
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开样式设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 多级列表_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 如果表单已存在但已关闭，重新创建
                if (multiLevelListForm == null || multiLevelListForm.IsDisposed)
                {
                    multiLevelListForm = new MultiLevelListForm();
                }
                multiLevelListForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开多级列表设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 域名高亮_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("域名高亮", () =>
            {
                documentProcessor.HighlightFields(true);
            });
        }

        private void 取消高亮_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("取消高亮", () =>
            {
                documentProcessor.HighlightFields(false);
            });
        }

        private void 上标_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("设置上标", () =>
            {
                documentProcessor.SetFieldSuperscript(true);
            });
        }

        private void 正常_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("取消上标", () =>
            {
                documentProcessor.SetFieldSuperscript(false);
            });
        }

        private void 另存PDF_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.ExportToPDF();
        }

        private void 版本_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.ShowVersion();
        }

        private void 文档合并_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (documentMerger == null)
                {
                    documentMerger = new SplitAndMerge.DocumentMerger((Word.Application)Globals.ThisAddIn.Application);
                }
                documentMerger.ShowMergeDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文档合并失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 文档拆分_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (documentSplitter == null)
                {
                    documentSplitter = new DocumentSplitter(Globals.ThisAddIn.Application);
                }
                documentSplitter.ShowSplitDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文档拆分失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 公开_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("添加密级：公开", () =>
            {
                documentProcessor.AddSecurityLevel("公开");
            });
        }

        private void 内部_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("添加密级：内部", () =>
            {
                documentProcessor.AddSecurityLevel("内部★");
            });
        }

        private void 移除密级_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteWithUndoRecord("移除密级", () =>
            {
                documentProcessor.RemoveSecurityLevelFromCurrentPage();
            });
        }
        #endregion

    }

    /// <summary>
    /// Win32 原生方法辅助类（用于将样式设置窗口绑定到 Word 主窗口，使其始终置于 Word 上方）
    /// </summary>
    internal static class NativeMethods
    {
        /// <summary>
        /// GWL_HWNDPARENT - 设置窗口的所有者窗口
        /// </summary>
        public const int GWL_HWNDPARENT = -8;

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = CharSet.Auto)]
        public static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", CharSet = CharSet.Auto)]
        public static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
    }
}
