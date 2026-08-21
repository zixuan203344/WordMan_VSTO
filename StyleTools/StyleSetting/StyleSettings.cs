using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Word;
using WordMan;
using WordMan.MultiLevel;
using Color = System.Drawing.Color;
using Point = System.Drawing.Point;
using Rectangle = System.Drawing.Rectangle;
using Font = System.Drawing.Font;
using CheckBox = System.Windows.Forms.CheckBox;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    /// <summary>
    /// 样式设置窗体 - 按照WordFormatHelper的StyleSetGuider设计
    /// </summary>
    public partial class StyleSettings : Form
    {
        #region 私有字段

        // 使用 MultiLevelDataManager 的字号相关方法，避免重复定义

        private BindingList<string> StyleNames;
        private readonly List<CustomStyle> Styles = new List<CustomStyle>(17);
        private readonly Dictionary<string, CustomStyle> _styleNameMap = new Dictionary<string, CustomStyle>(); // 样式名称到样式的映射，提升查找性能
        private bool _isLoadingStyle = false;
        private int _currentStyleIndex = -1; // 当前选中的样式索引
        private bool _multiSelectMode = false; // 多选模式：选中多个样式
        private readonly HashSet<string> _modifiedControls = new HashSet<string>(); // 多选模式下被手动修改的设置项（控件名）
        private readonly HashSet<string> _multiplePlaceholders = new HashSet<string>(); // 已插入"多种"占位项的 DropDownList 下拉框

        // 样式颜色采样缓存：样式名 → 渲染后的实际颜色（窗口会话级，避免反复采样）
        private readonly Dictionary<string, Color> _styleColorCache = new Dictionary<string, Color>();

        // 从 WordOpenXML 解析的 样式名(NameLocal) → 颜色 映射（处理主题色编码，interop Font.Color 对主题色返回负数无法直接用 FromOle）
        private Dictionary<string, Color> _styleXmlColorMap;

        // 内置样式英文名(styles.xml 的 w:name) → WdBuiltinStyle 枚举值，用于反查本地化名 NameLocal
        private static readonly Dictionary<string, int> BuiltinStyleEnglishToEnum = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            {"Normal", -1}, {"heading 1", -2}, {"heading 2", -3}, {"heading 3", -4},
            {"heading 4", -5}, {"heading 5", -6}, {"heading 6", -7}, {"heading 7", -8},
            {"heading 8", -9}, {"heading 9", -10}, {"Title", -63}, {"Subtitle", -75},
            {"caption", -35}, {"Body Text", -67}, {"Body Text 2", -81}, {"Body Text 3", -82},
            {"Body Text Indent", -68}, {"Body Text First Indent", -78},
            {"Hyperlink", -72}, {"No Spacing", -156}, {"Quote", -187}, {"Intense Quote", -188},
            {"Subtle Emphasis", -189}, {"Emphasis", -26}, {"Strong", -25},
            {"List Paragraph", -180}, {"Book Title", -264}, {"Bibliography", -266},
            {"Intense Reference", -264}, {"Reference", -65}, {"Footnote Text", -29},
            {"Footnote Reference", -39}, {"Endnote Text", -43}, {"Endnote Reference", -55},
            {"TOC 1", -20}, {"TOC 2", -21}, {"TOC 3", -22}, {"TOC Heading", -289},
            {"Table Grid", -163}
        };

        #endregion

        #region 私有方法


        /// <summary>
        /// 通用下拉框初始化方法
        /// </summary>
        private void InitializeComboBox(StandardComboBox comboBox, IEnumerable<string> items)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.ToArray());
        }

        /// <summary>
        /// 通用选择索引设置方法
        /// </summary>
        private void SetComboBoxSelection(StandardComboBox comboBox, string value, IEnumerable<string> items)
        {
            if (comboBox == null || items == null)
                return;

            var itemList = items.ToList();
            int selectedIndex = itemList.IndexOf(value);

            if (selectedIndex >= 0)
            {
                comboBox.SelectedIndex = selectedIndex;
            }
            else
            {
                // 如果找不到匹配项，对于可编辑的下拉框，直接设置Text属性
                if (comboBox.DropDownStyle == ComboBoxStyle.DropDown)
                {
                    // 临时禁用TextChanged事件，避免在加载样式时触发保存
                    EventHandler textChangedHandler = GetTextChangedHandler(comboBox.Name);
                    if (textChangedHandler != null)
                    {
                        comboBox.TextChanged -= textChangedHandler;
                        comboBox.Text = value;
                        comboBox.TextChanged += textChangedHandler;
                    }
                    else
                    {
                        comboBox.Text = value;
                    }
                }
                else
                {
                    comboBox.SelectedIndex = -1;
                }
            }
        }

        /// <summary>
        /// 获取下拉框的TextChanged事件处理器
        /// </summary>
        private EventHandler GetTextChangedHandler(string comboBoxName)
        {
            switch (comboBoxName)
            {
                case "Cmb_LineSpacing":
                    return Cmb_LineSpacing_TextChanged;
                case "Cmb_BefreSpacing":
                case "Cmb_AfterSpacing":
                    return Cmb_Spacing_TextChanged;
                default:
                    return null;
            }
        }

        /// <summary>
        /// 通用文本验证和格式化
        /// </summary>
        private void ValidateAndFormatText(Control control, string unit)
        {
            if (control is StandardComboBox comboBox && comboBox.SelectedIndex != -1) return;

            if (control is StandardTextBox textBox)
            {
                string text = textBox.Text.Trim();
                if (!string.IsNullOrEmpty(text) && !text.EndsWith(unit))
                {
                    textBox.Text = text + unit;
                }
            }
        }

        /// <summary>
        /// 获取控件值
        /// </summary>
        private object GetControlValue(Control control)
        {
            if (control == null)
                return null;

            if (control is StandardComboBox comboBox)
                return comboBox.Text;
            
            if (control is StandardTextBox textBox)
                return textBox.Text;
            
            if (control is StandardNumericUpDown numericUpDown)
            {
                return $"{numericUpDown.GetValueInCentimeters():0.0} 厘米";
            }
            
            if (control is ToggleButton toggleButton)
                return toggleButton.Pressed;
            
            return null;
        }

        /// <summary>
        /// 获取对齐方式文本（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private string GetAlignmentText(int alignment)
        {
            string[] alignments = WordStyleInfo.HAlignments;
            return alignment >= 0 && alignment < alignments.Length ? alignments[alignment] : "左对齐";
        }

        // 使用 MultiLevelDataManager 的行距相关方法，避免重复实现

        /// <summary>
        /// 设置缩进值到StandardNumericUpDown控件（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private void SetIndentValue(StandardNumericUpDown numericUpDown, string indentText)
        {
            if (string.IsNullOrEmpty(indentText))
            {
                numericUpDown.Value = 0;
                return;
            }

            // 解析缩进文本，提取数值和单位
            var cleanText = indentText.Replace("厘米", "").Replace("磅", "").Replace("字符", "").Replace("行", "").Trim();
            if (decimal.TryParse(cleanText, out decimal value))
            {
                // 根据单位设置值
                if (indentText.Contains("字符"))
                {
                    numericUpDown.SetValueInUnit(value, "字符");
                }
                else if (indentText.Contains("厘米"))
                {
                    numericUpDown.SetValueInCentimeters(value);
                }
                else if (indentText.Contains("磅"))
                {
                    numericUpDown.SetValueInUnit(value, "磅");
                }
                else if (indentText.Contains("行"))
                {
                    numericUpDown.SetValueInUnit(value, "行");
                }
                else
                {
                    // 默认按厘米处理
                    numericUpDown.SetValueInCentimeters(value);
                }
            }
            else
            {
                numericUpDown.Value = 0;
            }
        }

        // 使用 MultiLevelDataManager.MultiLevelDataManager.ConvertFontSizeToString 方法，避免重复实现

        /// <summary>
        /// 首行缩进方式选择事件处理
        /// </summary>
        private void 首行缩进方式下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFirstLineIndentVisibility();
        }

        /// <summary>
        /// 预设样式选择事件处理
        /// </summary>
        private void 风格下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (风格下拉框.SelectedIndex >= 0)
            {
                LoadPresetStyle(风格下拉框.SelectedItem.ToString());
            }
        }

        /// <summary>
        /// 加载预设样式
        /// </summary>
        private void LoadPresetStyle(string presetName)
        {
            try
            {
                // 获取预设样式集合
                var presetStyles = StylePresetManager.GetPresetStyles(presetName);

                // 清空当前样式列表
                Styles.Clear();
                StyleNames.Clear();

                // 清空样式名称映射
                _styleNameMap.Clear();

                // 添加预设样式到样式列表
                foreach (var presetStyle in presetStyles)
                {
                    if (presetStyle != null && !string.IsNullOrEmpty(presetStyle.Name))
                    {
                        Styles.Add(presetStyle);
                        StyleNames.Add(presetStyle.Name);
                        _styleNameMap[presetStyle.Name] = presetStyle;
                    }
                }

                // 刷新样式列表显示
                Lst_Styles.DataSource = null;
                Lst_Styles.DataSource = StyleNames;

                // 如果当前有选中的样式，重新加载到控件
                if (Lst_Styles.SelectedIndex >= 0)
                {
                    string selectedStyle = Lst_Styles.SelectedItem.ToString();
                    var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                    if (style != null)
                    {
                        LoadStyleToControls(style);
                        UpdateStyleInfo(style);
                    }
                }
                else
                {
                    // 如果没有选中样式，选择第一个样式
                    Lst_Styles.SelectedIndex = 0;
                }

                // 根据预设样式更新显示标题数
                UpdateTitleCountForPresetStyle(presetName);

                // 刷新样式基准和后续段落样式下拉框选项
                RefreshBaseStyleComboBoxes();
                UpdateBaseStyleControlsVisibility();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用预设样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 根据预设样式更新显示标题数
        /// </summary>
        private void UpdateTitleCountForPresetStyle(string presetName)
        {
            // 根据预设样式设置合适的显示标题数
            switch (presetName)
            {
                case "公文风格":
                case "条文风格":
                    显示标题数下拉框.SelectedIndex = 6; // 显示6级标题
                    break;
                case "论文风格":
                    显示标题数下拉框.SelectedIndex = 3; // 显示3级标题
                    break;
                case "报告风格":
                    显示标题数下拉框.SelectedIndex = 4; // 显示4级标题
                    break;
                default:
                    显示标题数下拉框.SelectedIndex = 4; // 默认显示4级标题
                    break;
            }

            // 更新样式列表显示
            FilterStylesByTitleCount();
        }

        /// <summary>
        /// 更新首行缩进输入框的可见性
        /// </summary>
        private void UpdateFirstLineIndentVisibility()
        {
            if (首行缩进方式下拉框.SelectedIndex == 0) // 无
            {
                label7.Visible = false;
                Nud_FirstLineIndent.Visible = false;
            }
            else if (首行缩进方式下拉框.SelectedIndex == 1) // 悬挂缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                label7.Text = "悬挂缩进";
                Nud_FirstLineIndent.Unit = "厘米"; // 默认使用厘米单位
            }
            else if (首行缩进方式下拉框.SelectedIndex == 2) // 首行缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                label7.Text = "首行缩进";
                Nud_FirstLineIndent.Unit = "厘米"; // 默认使用厘米单位
            }
        }

        #endregion

        #region 构造函数

        public StyleSettings()
        {
            InitializeComponent();
            InitializeData();
            BindEvents();

            // 初始化默认样式，不自动读取文档样式
            InitializeDefaultStyles();
            FilterStylesByTitleCount();

            // 优化列表配色
            SetupListBoxStyling();

            // 底部按钮均匀排布，宽度自适应文字
            LayoutBottomButtons();
        }

        /// <summary>
        /// 设置列表框的样式和配色
        /// </summary>
        private void SetupListBoxStyling()
        {
            if (Lst_Styles != null)
            {
                Lst_Styles.DrawMode = DrawMode.OwnerDrawFixed;
                Lst_Styles.DrawItem += Lst_Styles_DrawItem;
            }
        }

        /// <summary>
        /// 列表项绘制事件处理
        /// </summary>
        private void Lst_Styles_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || Lst_Styles == null || e.Index >= Lst_Styles.Items.Count)
                return;

            e.DrawBackground();

            // 现代化配色
            Color selectedBg = Color.FromArgb(70, 130, 230);  // 柔和的蓝色
            Color alternatingRow = Color.FromArgb(249, 250, 252);  // 浅灰色
            Color textDark = Color.FromArgb(33, 37, 41);  // 深灰色文本

            // 绘制背景（交替行颜色和选中状态）
            Color backColor = (e.State & DrawItemState.Selected) == DrawItemState.Selected
                ? selectedBg
                : (e.Index % 2 == 0 ? Color.White : alternatingRow);
            
            e.Graphics.FillRectangle(new SolidBrush(backColor), e.Bounds);

            // 绘制文本
            string text = Lst_Styles.Items[e.Index].ToString();
            Color textColor = (e.State & DrawItemState.Selected) == DrawItemState.Selected 
                ? Color.White 
                : textDark;
            
            var textRect = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 10, e.Bounds.Height);
            
            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(text, Lst_Styles.Font, brush, textRect, StringFormat.GenericDefault);
            }

            e.DrawFocusRectangle();
        }

        /// <summary>
        /// 窗体显示时自动读取文档样式
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // 窗体显示时自动读取当前文档样式
            try
            {
                RefreshFromDocument();
                // 窗体尺寸最终确定后重排底部按钮
                LayoutBottomButtons();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"窗体显示时读取文档样式出错：{ex.Message}");
                // 如果读取失败，保持默认样式
            }
        }

        /// <summary>
        /// 从当前文档刷新样式：读取文档样式、填充样式列表并自动选中样式、
        /// 刷新样式基准/后续段落样式下拉框选项，在状态栏显示结果。
        /// 供功能区按钮重复点击时（窗口已打开）重新读取当前文档。
        /// </summary>
        public void RefreshFromDocument()
        {
            // 切换文档后清空样式颜色采样缓存，重建颜色映射（避免使用旧文档主题的颜色）
            _styleColorCache.Clear();
            _styleXmlColorMap = null;

            ReadDocumentStyles();
            // 刷新样式基准和后续段落样式下拉框选项
            RefreshBaseStyleComboBoxes();
            // 清空预设风格选择
            风格下拉框.SelectedIndex = -1;
            // 状态栏提示
            ShowStatusMessage($"已读取当前文档样式，共 {StyleNames.Count} 个");
        }

        #endregion

        #region 初始化方法

        private void InitializeData()
        {
            // 初始化控件状态

            // 初始化样式列表
            StyleNames = new BindingList<string>();
            Lst_Styles.DataSource = StyleNames;

            // 使用 MultiLevelDataManager 的统一方法初始化字体列表
            // 中文字体优先显示常用字体，西文字体优先显示"+西文标题"等特殊项（与 Word 原生字体设置窗口一致）
            InitializeComboBox(Cmb_ChnFontName, MultiLevelDataManager.GetChnFontItems());
            InitializeComboBox(Cmb_EngFontName, MultiLevelDataManager.GetEngFontItems());

            // 初始化其他下拉框
            InitializeComboBox(Cmb_FontSize, MultiLevelDataManager.GetFontSizes());
            InitializeComboBox(Cmb_ParaAligment, WordStyleInfo.HAlignments);
            InitializeComboBox(风格下拉框, StylePresetManager.GetPresetStyleNames());


            // 初始化行距下拉框 - 使用 WordStyleInfo.LineSpacings
            InitializeComboBox(Cmb_LineSpacing, WordStyleInfo.LineSpacings);

            // 初始化段落间距下拉框 - 使用 WordStyleInfo 的预设值
            InitializeComboBox(Cmb_BefreSpacing, WordStyleInfo.SpaceBeforeValues);
            InitializeComboBox(Cmb_AfterSpacing, WordStyleInfo.SpaceAfterValues);

            // 添加段落间距下拉框的文本变化事件处理
            Cmb_BefreSpacing.TextChanged += Cmb_Spacing_TextChanged;
            Cmb_AfterSpacing.TextChanged += Cmb_Spacing_TextChanged;

            // 设置显示标题数下拉框
            InitializeComboBox(显示标题数下拉框, new[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" });
            显示标题数下拉框.SelectedIndex = 4; // 默认显示4级标题
            显示标题数下拉框.SelectedIndexChanged += 显示标题数下拉框_SelectedIndexChanged;

            Lst_Styles.SelectedIndex = -1;
            Cmb_ParaAligment.SelectedIndex = -1;

            // 初始化首行缩进方式
            首行缩进方式下拉框.Items.Clear(); // 先清空避免重复
            首行缩进方式下拉框.Items.AddRange(new[] { "无", "悬挂缩进", "首行缩进" });
            首行缩进方式下拉框.SelectedIndex = 0; // 默认选择"无"

            // 设置首行缩进方式选择事件
            首行缩进方式下拉框.SelectedIndexChanged += 首行缩进方式下拉框_SelectedIndexChanged;

            // 设置预设样式选择事件
            风格下拉框.SelectedIndexChanged += 风格下拉框_SelectedIndexChanged;

            // 初始化样式基准和后续段落样式下拉框（前两项与 Word 原生修改样式窗口一致）
            InitializeComboBox(样式基准下拉框, new[] { "(无样式)", "正文" });
            InitializeComboBox(后续段落样式下拉框, new[] { "正文", "无间隔" });
            样式基准下拉框.Visible = false;
            后续段落样式下拉框.Visible = false;
            label20.Visible = false;
            label21.Visible = false;

            // 初始化样式名称输入框
            InitializeStyleNameTextBox();

            // 初始状态：隐藏首行缩进输入框
            UpdateFirstLineIndentVisibility();

            // 样式过滤在构造函数中已处理

            // 启用字体设置和段落设置面板
            Pal_Font.Enabled = true;
            Pal_ParaIndent.Enabled = true;

            // 启用按钮
            加载.Enabled = true;
            添加.Enabled = true;
            删除.Enabled = true;
        }

        /// <summary>
        /// 初始化样式名称输入框
        /// </summary>
        private void InitializeStyleNameTextBox()
        {
            Txt_AddStyleName.Text = "请输入需要增加的样式名称";
            Txt_AddStyleName.ForeColor = Color.Gray;

            // 添加焦点事件处理
            Txt_AddStyleName.Enter += Txt_AddStyleName_Enter;
            Txt_AddStyleName.Leave += Txt_AddStyleName_Leave;
        }

        /// <summary>
        /// 样式名称输入框获得焦点事件
        /// </summary>
        private void Txt_AddStyleName_Enter(object sender, EventArgs e)
        {
            if (Txt_AddStyleName.Text == "请输入需要增加的样式名称")
            {
                Txt_AddStyleName.Text = "";
                Txt_AddStyleName.ForeColor = Color.Black;
            }
        }

        /// <summary>
        /// 样式名称输入框失去焦点事件
        /// </summary>
        private void Txt_AddStyleName_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Txt_AddStyleName.Text))
            {
                Txt_AddStyleName.Text = "请输入需要增加的样式名称";
                Txt_AddStyleName.ForeColor = Color.Gray;
            }
        }

        private void InitializeDefaultStyles()
        {
            Styles.Clear();
            _styleNameMap.Clear();
            var defaultStyles = new CustomStyle[]
            {
                new CustomStyle(name: "正文", fontName: "宋体", engFontName: "宋体", fontSize: 10.5f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 2f, firstLineIndentByChar: 2, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 1", fontName: "宋体", engFontName: "宋体", fontSize: 16f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 2", fontName: "宋体", engFontName: "宋体", fontSize: 14f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 3", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 4", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 5", fontName: "宋体", engFontName: "宋体", fontSize: 10.5f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "标题 6", fontName: "宋体", engFontName: "宋体", fontSize: 10.5f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "题注", fontName: "宋体", engFontName: "宋体", fontSize: 9f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文"),
                new CustomStyle(name: "表内文字", fontName: "宋体", engFontName: "宋体", fontSize: 9f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false, baseStyle: "正文", nextParagraphStyle: "正文")
            };

            Styles.AddRange(defaultStyles);

            // 添加调试信息并更新字典
            foreach (var style in Styles)
            {
                if (style != null && !string.IsNullOrEmpty(style.Name))
                {
                    System.Diagnostics.Debug.WriteLine($"创建默认样式：{style.Name}，字号：{style.FontSize}，对象ID：{style.GetHashCode()}");
                    StyleNames.Add(style.Name);
                    _styleNameMap[style.Name] = style;
                }
            }
        }

        private void BindEvents()
        {
            Lst_Styles.SelectedIndexChanged += Lst_Styles_SelectedIndexChanged;

            // 添加样式修改事件，当用户修改样式时立即保存到样式对象
            Cmb_ChnFontName.SelectedIndexChanged += OnStyleChanged;
            Cmb_ChnFontName.TextChanged += OnStyleChanged; // 可输入模式下输入内容也实时保存
            Cmb_EngFontName.SelectedIndexChanged += OnStyleChanged;
            Cmb_EngFontName.TextChanged += OnStyleChanged;
            Cmb_FontSize.SelectedIndexChanged += OnStyleChanged;
            Cmb_ParaAligment.SelectedIndexChanged += OnStyleChanged;
            Cmb_LineSpacing.TextChanged += Cmb_LineSpacing_TextChanged;
            Cmb_LineSpacing.Validated += Cmb_LineSpacing_Validated;
            Nud_LeftIndent.ValueChanged += OnStyleChanged;
            Nud_RightIndent.ValueChanged += OnStyleChanged;
            Nud_FirstLineIndent.ValueChanged += OnStyleChanged;
            // 段落间距控件使用TextChanged事件，不需要SelectedIndexChanged
            Btn_Bold.PressedChanged += OnStyleChanged;
            Btn_Italic.PressedChanged += OnStyleChanged;
            Btn_UnderLine.PressedChanged += OnStyleChanged;
            Btn_FontColor.Click += Btn_FontColor_Click;
            首行缩进方式下拉框.SelectedIndexChanged += OnStyleChanged;
            样式基准下拉框.SelectedIndexChanged += OnStyleChanged;
            后续段落样式下拉框.SelectedIndexChanged += OnStyleChanged;

            添加.Click += 添加_Click;
            删除.Click += 删除_Click;
            Btn_ApplySet.Click += Btn_ApplySet_Click;
            应用当前.Click += 应用当前_Click;
            读取文档样式.Click += Btn_ReadDocumentStyle_Click;
            this.关闭.Click += 关闭_Click;
            this.加载.Click += 加载_Click;

            // 添加导入导出按钮事件
            导入.Click += Btn_Import_Click;
            导出.Click += Btn_Export_Click;
            存预设.Click += 存预设_Click;
        }

        #endregion

        #region 事件处理方法

        private void Lst_Styles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Lst_Styles == null)
                return;

            _isLoadingStyle = true; // 设置加载标志，防止触发保存事件

            try
            {
                // 多选：进入批量编辑模式，属性显示共同值或"多种"，修改后点"应用全部"批量设置
                if (Lst_Styles.SelectedIndices.Count > 1)
                {
                    _multiSelectMode = true;
                    _currentStyleIndex = -1;
                    _modifiedControls.Clear();

                    UpdateControlsForMultiSelect();

                    // 多选模式下允许编辑面板
                    Pal_Font.Enabled = true;
                    Pal_ParaIndent.Enabled = true;
                    删除.Enabled = false;      // 多选不删除
                    添加.Enabled = true;
                    应用当前.Enabled = false;  // 多选使用"应用全部"
                    UpdateBaseStyleControlsVisibility();
                    return;
                }

                // 退出多选模式
                if (_multiSelectMode)
                {
                    _multiSelectMode = false;
                    _modifiedControls.Clear();
                    RemoveAllMultiplePlaceholders();
                    应用当前.Enabled = true;
                }

                if (Lst_Styles.SelectedIndex >= 0 && Lst_Styles.SelectedItem != null)
                {
                    string selectedStyle = Lst_Styles.SelectedItem.ToString();
                    if (string.IsNullOrEmpty(selectedStyle))
                        return;

                    // 使用字典查找，提升性能
                    CustomStyle style;
                    if (!_styleNameMap.TryGetValue(selectedStyle, out style))
                    {
                        style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                        if (style != null)
                        {
                            _styleNameMap[selectedStyle] = style;
                        }
                    }

                    if (style != null)
                    {
                        // 找到样式在列表中的索引
                        _currentStyleIndex = Styles.IndexOf(style);

                        // 加载选中的样式到控件
                        LoadStyleToControls(style);
                        UpdateStyleInfo(style);

                        // 启用字体设置和段落设置面板
                        Pal_Font.Enabled = true;
                        Pal_ParaIndent.Enabled = true;

                        // 更新按钮状态
                        删除.Enabled = style.UserDefined; // 只有用户定义的样式才能删除
                        添加.Enabled = true;

                        // 更新样式基准/后续段落样式可见性（正文样式不显示）
                        UpdateBaseStyleControlsVisibility();

                        System.Diagnostics.Debug.WriteLine($"已选择样式：{selectedStyle}，用户定义：{style.UserDefined}，索引：{_currentStyleIndex}");
                    }
                }
                else
                {
                    // 如果没有选中样式，禁用编辑面板
                    _currentStyleIndex = -1;
                    Pal_Font.Enabled = false;
                    Pal_ParaIndent.Enabled = false;
                    删除.Enabled = false;
                    添加.Enabled = true;
                    UpdateBaseStyleControlsVisibility();
                }
            }
            finally
            {
                _isLoadingStyle = false; // 重置加载标志
            }
        }

        /// <summary>
        /// 样式修改事件处理 - 实时保存样式修改（参考StyleSetGuider实现）
        /// </summary>
        private void OnStyleChanged(object sender, EventArgs e)
        {
            // 如果正在加载样式到控件，跳过保存
            if (_isLoadingStyle)
            {
                System.Diagnostics.Debug.WriteLine($"跳过保存：正在加载样式到控件");
                return;
            }

            // 多选模式：只记录被手动修改的设置项（不修改单个样式对象），供"应用全部"批量应用
            if (_multiSelectMode)
            {
                if (sender is StandardComboBox multiCombo)
                {
                    // 用户选择了真实值，移除"多种"占位项
                    RemoveMultiplePlaceholder(multiCombo);
                    _modifiedControls.Add(multiCombo.Name);
                }
                else if (sender is StandardNumericUpDown multiNud)
                {
                    _modifiedControls.Add(multiNud.Name);
                }
                else if (sender is ToggleButton multiToggle)
                {
                    _modifiedControls.Add(multiToggle.Name);
                }
                else if (sender is StandardButton multiBtn && multiBtn.Name == "Btn_FontColor")
                {
                    _modifiedControls.Add("Btn_FontColor");
                }
                System.Diagnostics.Debug.WriteLine($"多选模式：记录修改的设置项 {sender?.GetType().Name}");
                return;
            }

            // 如果没有选中样式，跳过保存
            if (_currentStyleIndex < 0 || _currentStyleIndex >= Styles.Count)
            {
                System.Diagnostics.Debug.WriteLine($"跳过保存：没有选中样式或索引无效");
                return;
            }

            // 直接修改当前样式的属性（参考StyleSetGuider的实现方式）
            var currentStyle = Styles[_currentStyleIndex];
            System.Diagnostics.Debug.WriteLine($"用户修改样式，触发事件：{sender?.GetType().Name}，当前样式：{currentStyle.Name}");

            // 根据控件类型直接更新样式属性
            if (sender is StandardComboBox comboBox)
            {
                switch (comboBox.Name)
                {
                    case "Cmb_ChnFontName":
                        currentStyle.FontName = comboBox.SelectedItem?.ToString() ?? comboBox.Text;
                        break;
                    case "Cmb_EngFontName":
                        currentStyle.EngFontName = comboBox.SelectedItem?.ToString() ?? comboBox.Text;
                        break;
                    case "Cmb_FontSize":
                        if (comboBox.SelectedIndex >= 0)
                        {
                            currentStyle.FontSize = MultiLevelDataManager.ConvertFontSize(comboBox.SelectedItem?.ToString());
                        }
                        break;
                    case "Cmb_ParaAligment":
                        currentStyle.ParaAlignment = comboBox.SelectedIndex;
                        break;
                    case "Cmb_LineSpacing":
                        if (!string.IsNullOrEmpty(comboBox.Text))
                        {
                            currentStyle.LineSpacing = ConvertLineSpacingToFloat(comboBox.Text);
                        }
                        break;
                    case "Cmb_BefreSpacing":
                        currentStyle.BeforeSpacing = ConvertSpacingToPoints(comboBox.Text);
                        break;
                    case "Cmb_AfterSpacing":
                        currentStyle.AfterSpacing = ConvertSpacingToPoints(comboBox.Text);
                        break;
                    case "样式基准下拉框":
                        currentStyle.BaseStyle = comboBox.SelectedItem?.ToString() ?? comboBox.Text ?? "";
                        break;
                    case "后续段落样式下拉框":
                        currentStyle.NextParagraphStyle = comboBox.SelectedItem?.ToString() ?? comboBox.Text ?? "";
                        break;
                }
            }
            else if (sender is StandardNumericUpDown numericUpDown)
            {
                switch (numericUpDown.Name)
                {
                    case "Nud_LeftIndent":
                        currentStyle.LeftIndent = (float)numericUpDown.GetValueInCentimeters();
                        break;
                    case "Nud_RightIndent":
                        currentStyle.RightIndent = (float)numericUpDown.GetValueInCentimeters();
                        break;
                    case "Nud_FirstLineIndent":
                        // 根据首行缩进方式设置首行缩进值
                        if (首行缩进方式下拉框.SelectedIndex == 1) // 悬挂缩进
                        {
                            currentStyle.FirstLineIndent = -(float)numericUpDown.GetValueInCentimeters();
                        }
                        else if (首行缩进方式下拉框.SelectedIndex == 2) // 首行缩进
                        {
                            currentStyle.FirstLineIndent = (float)numericUpDown.GetValueInCentimeters();
                        }
                        break;
                }
            }
            else if (sender is ToggleButton toggleButton)
            {
                switch (toggleButton.Name)
                {
                    case "Btn_Bold":
                        currentStyle.Bold = toggleButton.Pressed;
                        break;
                    case "Btn_Italic":
                        currentStyle.Italic = toggleButton.Pressed;
                        break;
                    case "Btn_UnderLine":
                        currentStyle.Underline = toggleButton.Pressed;
                        break;
                }
            }
            else if (sender is StandardButton button && button.Name == "Btn_FontColor")
            {
                // 处理字体颜色变化
                currentStyle.FontColor = button.BackColor;
            }
            else if (sender is StandardComboBox && ((StandardComboBox)sender).Name == "首行缩进方式下拉框")
            {
                // 首行缩进方式改变时，重新设置首行缩进值
                if (首行缩进方式下拉框?.SelectedIndex == 0) // 无
                {
                    currentStyle.FirstLineIndent = 0f;
                }
                else if (首行缩进方式下拉框?.SelectedIndex == 1 && Nud_FirstLineIndent != null) // 悬挂缩进
                {
                    currentStyle.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters();
                }
                else if (首行缩进方式下拉框?.SelectedIndex == 2 && Nud_FirstLineIndent != null) // 首行缩进
                {
                    currentStyle.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
                }
            }

            // 更新样式信息显示
            UpdateStyleInfo(currentStyle);
            System.Diagnostics.Debug.WriteLine($"样式 {currentStyle.Name} 已自动保存修改");
        }

        private void StyleFontChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }

        private void IndentSpacingChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }

        private void FontStyleChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }



        /// <summary>
        /// 显示标题数下拉框选择变化事件
        /// </summary>
        private void 显示标题数下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 根据选择的标题数过滤样式列表
            FilterStylesByTitleCount();
        }

        /// <summary>
        /// 根据选择的标题数过滤样式列表
        /// </summary>
        private void FilterStylesByTitleCount()
        {
            int titleCount = 显示标题数下拉框.SelectedIndex; // 0=0级, 1=1级, 2=2级, ...

            // 清空当前样式列表
            StyleNames.Clear();

            foreach (var style in Styles)
            {
                bool shouldShow = false;

                if (titleCount == 0) // 不显示标题样式，只显示其他样式
                {
                    if (!style.Name.StartsWith("标题"))
                    {
                        shouldShow = true;
                    }
                }
                else if (style.Name.StartsWith("标题"))
                {
                    // 提取标题级别，只显示指定数量及以下的标题
                    // 标题级别从1开始，所以需要+1
                    int titleLevel = GetTitleLevel(style.Name);
                    if (titleLevel > 0 && titleLevel <= titleCount)
                    {
                        shouldShow = true;
                    }
                }
                else
                {
                    // 始终显示正文、题注、表内文字和其他非标题样式
                    shouldShow = true;
                }

                if (shouldShow)
                {
                    StyleNames.Add(style.Name);
                }
            }

            // 刷新列表显示
            Lst_Styles.DataSource = null;
            Lst_Styles.DataSource = StyleNames;

            // 如果当前选中的样式不在过滤后的列表中，重置当前样式索引
            if (_currentStyleIndex >= 0 && _currentStyleIndex < Styles.Count)
            {
                var currentStyle = Styles[_currentStyleIndex];
                if (currentStyle != null && !StyleNames.Contains(currentStyle.Name))
                {
                    _currentStyleIndex = -1;
                }
            }
        }

        /// <summary>
        /// 从样式名称中提取标题级别
        /// </summary>
        private int GetTitleLevel(string styleName)
        {
            if (string.IsNullOrEmpty(styleName) || !styleName.StartsWith("标题"))
                return 0;

            // 处理 "标题 1", "标题 2" 等格式
            if (styleName.StartsWith("标题 "))
            {
                string levelText = styleName.Substring(3).Trim(); // 去掉 "标题 " 前缀
                if (int.TryParse(levelText, out int level))
                {
                    return level;
                }
            }

            // 处理 "标题1", "标题2" 等格式（无空格）
            if (styleName.Length > 2 && char.IsDigit(styleName[2]))
            {
                string levelText = styleName.Substring(2); // 去掉 "标题" 前缀
                if (int.TryParse(levelText, out int level))
                {
                    return level;
                }
            }

            return 0;
        }


        private void 添加_Click(object sender, EventArgs e)
        {
            string styleName = Txt_AddStyleName.Text.Trim();

            // 检查是否为空或提示文本
            if (string.IsNullOrEmpty(styleName) || styleName == "请输入需要增加的样式名称")
            {
                MessageBox.Show("请输入样式名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (StyleNames.Contains(styleName))
            {
                MessageBox.Show("样式名称已存在", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 创建新样式
            var newStyle = CreateStyleFromControls(styleName);
            if (newStyle != null)
            {
                Styles.Add(newStyle);
                StyleNames.Add(styleName);
                _styleNameMap[styleName] = newStyle;
                
                // 更新当前样式索引
                _currentStyleIndex = Styles.Count - 1;
                
                Lst_Styles.SelectedItem = styleName;
            }

            // 清空输入框并恢复提示文本
            Txt_AddStyleName.Text = "请输入需要增加的样式名称";
            Txt_AddStyleName.ForeColor = Color.Gray;
        }

        private void 删除_Click(object sender, EventArgs e)
        {
            if (Lst_Styles == null || Lst_Styles.SelectedIndex < 0 || Lst_Styles.SelectedItem == null)
            {
                MessageBox.Show("请先选择要删除的样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedStyle = Lst_Styles.SelectedItem.ToString();
            if (string.IsNullOrEmpty(selectedStyle))
            {
                MessageBox.Show("请先选择要删除的样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 使用字典查找，提升性能
            CustomStyle style;
            if (!_styleNameMap.TryGetValue(selectedStyle, out style))
            {
                style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
            }

            if (style == null)
            {
                MessageBox.Show("找不到指定的样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 检查是否为内置样式
            if (!style.UserDefined)
            {
                MessageBox.Show("不能删除内置样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 确认删除
            var result = MessageBox.Show($"确定要删除样式 '{selectedStyle}' 吗？", "确认删除",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // 如果删除的是当前选中的样式，重置当前样式索引
                if (_currentStyleIndex == Styles.IndexOf(style))
                {
                    _currentStyleIndex = -1;
                }
                else if (_currentStyleIndex > Styles.IndexOf(style))
                {
                    // 如果删除的样式在当前样式之前，需要调整索引
                    _currentStyleIndex--;
                }

                Styles.Remove(style);
                StyleNames.Remove(selectedStyle);
                _styleNameMap.Remove(selectedStyle);
                ResetControlsToDefault();

                // 刷新样式列表显示
                Lst_Styles.DataSource = null;
                Lst_Styles.DataSource = StyleNames;

                MessageBox.Show("样式删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Btn_ApplySet_Click(object sender, EventArgs e)
        {
            try
            {
                // 多选批量应用：只修改被手动修改的设置项，其余设置项不变动
                if (_multiSelectMode)
                {
                    var selectedStyles = GetSelectedStyles();
                    if (selectedStyles.Count == 0)
                    {
                        ShowStatusMessage("请先选择要批量设置的样式");
                        return;
                    }

                    var modifiedNames = _modifiedControls.ToList();
                    if (modifiedNames.Count == 0)
                    {
                        ShowStatusMessage("请先修改要批量设置的属性（修改后值会应用到所有选中样式）");
                        return;
                    }

                    Globals.ThisAddIn.ExecuteWithUndoRecord("批量应用样式属性", () =>
                    {
                        foreach (var controlName in modifiedNames)
                        {
                            foreach (var style in selectedStyles)
                            {
                                ReadControlValueToStyle(controlName, style);
                            }
                        }
                        // 只应用选中的样式到文档
                        ApplyStylesToDocument(selectedStyles.Select(s => s.Name).ToList());
                    });

                    _modifiedControls.Clear();
                    // 应用后刷新显示：被修改的属性现在所有选中样式相同，显示共同值
                    UpdateControlsForMultiSelect();
                    ShowStatusMessage($"已批量设置 {selectedStyles.Count} 个样式，共修改 {modifiedNames.Count} 项属性");
                    return;
                }

                // 单选：样式已经实时保存，直接应用到文档（封装为一个可撤销步骤，支持 Ctrl+Z）
                Globals.ThisAddIn.ExecuteWithUndoRecord("应用样式设置", () =>
                {
                    ApplyStylesToDocument();
                });

                // 在状态栏显示成功信息（不再弹出提示框）
                int appliedCount = StyleNames.Count;
                ShowStatusMessage($"样式设置已成功应用到文档，共应用 {appliedCount} 个样式");
            }
            catch (Exception ex)
            {
                ShowStatusMessage($"应用样式时出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 应用当前样式 - 只应用当前选中的样式，其他样式不修改，应用后立即刷新
        /// </summary>
        private void 应用当前_Click(object sender, EventArgs e)
        {
            try
            {
                if (_currentStyleIndex < 0 || _currentStyleIndex >= Styles.Count)
                {
                    MessageBox.Show("请先选择要应用的样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var style = Styles[_currentStyleIndex];
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有打开的文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 应用当前样式（封装为一个可撤销步骤，支持 Ctrl+Z）
                Globals.ThisAddIn.ExecuteWithUndoRecord($"应用样式：{style.Name}", () =>
                {
                    ApplyStyleToDocument(doc, style);
                });

                // Word 立即刷新新样式
                app.ScreenRefresh();
                doc.Repaginate();

                // 在状态栏显示成功信息（不再弹出提示框）
                ShowStatusMessage($"样式 '{style.Name}' 已应用到文档");
            }
            catch (Exception ex)
            {
                ShowStatusMessage($"应用当前样式时出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 保存当前样式为预设 - 输入名称后持久化保存，并自动刷新预设风格下拉列表
        /// </summary>
        private void 存预设_Click(object sender, EventArgs e)
        {
            try
            {
                string name = "";
                using (var dialog = new SavePresetDialog())
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;
                    name = dialog.PresetName;
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    ShowStatusMessage("保存预设失败：预设名称不能为空");
                    return;
                }

                // 检查是否已存在同名预设
                var existingNames = StylePresetManager.GetPresetStyleNames();
                if (existingNames.Contains(name))
                {
                    var result = MessageBox.Show($"预设 '{name}' 已存在，是否覆盖？", "保存预设",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result != DialogResult.Yes)
                        return;
                }

                // 保存当前显示的样式集合（与导出逻辑一致）
                var stylesToSave = new List<CustomStyle>();
                foreach (var styleName in StyleNames)
                {
                    var style = Styles.FirstOrDefault(s => s.Name == styleName);
                    if (style != null)
                    {
                        stylesToSave.Add(style);
                    }
                }

                StylePresetManager.SavePreset(new StylePresetData
                {
                    Name = name,
                    Styles = stylesToSave,
                    TitleCount = 显示标题数下拉框.SelectedIndex,
                    CreateTime = DateTime.Now
                });

                // 刷新预设风格下拉列表，新预设立即可见
                InitializeComboBox(风格下拉框, StylePresetManager.GetPresetStyleNames());
                风格下拉框.SelectedItem = name;

                ShowStatusMessage($"预设 '{name}' 已保存，共 {stylesToSave.Count} 个样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存预设时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_ReadDocumentStyle_Click(object sender, EventArgs e)
        {
            try
            {
                ReadDocumentStyles();
                // 清空预设风格选择
                风格下拉框.SelectedIndex = -1;
                MessageBox.Show("文档样式读取完成", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取文档样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 关闭_Click(object sender, EventArgs e)
        {
            // 重置控件状态
            ResetControlsToDefault();
            this.Close();
        }

        private void 加载_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取文档中的所有样式
                var documentStyles = GetDocumentStyles();

                if (documentStyles.Count == 0)
                {
                    MessageBox.Show("文档中没有找到样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 创建样式选择窗体
                using (var styleSelectionForm = new StyleSelectionForm())
                {
                    // 初始化可用样式
                    styleSelectionForm.InitializeStyles(documentStyles);

                    // 设置当前已选择的样式
                    var currentSelectedStyles = StyleNames.ToList();
                    styleSelectionForm.SetSelectedStyles(currentSelectedStyles);

                    // 显示窗体
                    if (styleSelectionForm.ShowDialog() == DialogResult.OK)
                    {
                        // 更新样式列表
                        StyleNames.Clear();
                        _styleNameMap.Clear();
                        foreach (var selectedStyle in styleSelectionForm.SelectedStyles)
                        {
                            if (!string.IsNullOrEmpty(selectedStyle))
                            {
                                StyleNames.Add(selectedStyle);
                                // 从现有样式中查找并添加到字典
                                var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                                if (style != null)
                                {
                                    _styleNameMap[selectedStyle] = style;
                                }
                            }
                        }

                        // 刷新样式列表显示
                        Lst_Styles.DataSource = null;
                        Lst_Styles.DataSource = StyleNames;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出样式设置
        /// </summary>
        private void Btn_Export_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = StyleFileManager.ShowSaveFileDialog("样式设置", "XML文件|*.xml|所有文件|*.*");
                if (!string.IsNullOrEmpty(filePath))
                {
                    // 样式已经实时保存，无需额外保存

                    // 根据当前显示的样式列表进行导出，而不是导出所有样式
                    var stylesToExport = new List<CustomStyle>();
                    foreach (var styleName in StyleNames)
                    {
                        var style = Styles.FirstOrDefault(s => s.Name == styleName);
                        if (style != null)
                        {
                            stylesToExport.Add(style);
                        }
                    }

                    // 创建包含显示标题数信息的导出数据
                    var exportData = new StyleExportData
                    {
                        Styles = stylesToExport,
                        TitleCount = 显示标题数下拉框.SelectedIndex,
                        ExportTime = DateTime.Now
                    };

                    StyleFileManager.SerializeToXml(exportData, filePath);
                    MessageBox.Show($"样式设置导出成功，共导出 {stylesToExport.Count} 个样式，显示标题数：{显示标题数下拉框.SelectedIndex}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导入样式设置
        /// </summary>
        private void Btn_Import_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = StyleFileManager.ShowOpenFileDialog("XML文件|*.xml|所有文件|*.*");
                if (!string.IsNullOrEmpty(filePath))
                {
                    // 样式已经实时保存，无需额外保存

                    // 尝试导入新的导出格式
                    try
                    {
                        var exportData = StyleFileManager.DeserializeFromXml<StyleExportData>(filePath);
                        if (exportData != null && exportData.Styles != null && exportData.Styles.Count > 0)
                        {
                            // 清空现有样式
                            Styles.Clear();
                            StyleNames.Clear();
                            _styleNameMap.Clear();

                            // 加载导入的样式
                            if (exportData.Styles != null)
                            {
                                Styles.AddRange(exportData.Styles);
                                foreach (var style in exportData.Styles)
                                {
                                    if (style != null && !string.IsNullOrEmpty(style.Name))
                                    {
                                        StyleNames.Add(style.Name);
                                        _styleNameMap[style.Name] = style;
                                    }
                                }
                            }

                            // 恢复显示标题数
                            if (exportData.TitleCount >= 0 && exportData.TitleCount < 显示标题数下拉框.Items.Count)
                            {
                                显示标题数下拉框.SelectedIndex = exportData.TitleCount;
                            }
                            else
                            {
                                // 根据导入的样式更新显示标题数
                                UpdateTitleCountForImportedStyles(exportData.Styles);
                            }

                            // 刷新显示
                            Lst_Styles.DataSource = null;
                            Lst_Styles.DataSource = StyleNames;

                            // 清空预设风格选择
                            风格下拉框.SelectedIndex = -1;

                            MessageBox.Show($"成功导入 {exportData.Styles.Count} 个样式，显示标题数：{显示标题数下拉框.SelectedIndex}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch
                    {
                        // 如果新格式导入失败，尝试旧格式
                        var importedStyles = StyleFileManager.DeserializeListFromXml<CustomStyle>(filePath);
                        if (importedStyles != null && importedStyles.Count > 0)
                        {
                            // 清空现有样式
                            Styles.Clear();
                            StyleNames.Clear();
                            _styleNameMap.Clear();

                            // 加载导入的样式
                            if (importedStyles != null)
                            {
                                Styles.AddRange(importedStyles);
                                foreach (var style in importedStyles)
                                {
                                    if (style != null && !string.IsNullOrEmpty(style.Name))
                                    {
                                        StyleNames.Add(style.Name);
                                        _styleNameMap[style.Name] = style;
                                    }
                                }
                            }

                            // 根据导入的样式更新显示标题数
                            UpdateTitleCountForImportedStyles(importedStyles);

                            // 刷新显示
                            Lst_Styles.DataSource = null;
                            Lst_Styles.DataSource = StyleNames;

                            // 清空预设风格选择
                            风格下拉框.SelectedIndex = -1;

                            MessageBox.Show($"成功导入 {importedStyles.Count} 个样式（旧格式）", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 根据导入的样式更新显示标题数
        /// </summary>
        private void UpdateTitleCountForImportedStyles(List<CustomStyle> importedStyles)
        {
            // 统计导入的样式中有多少级标题
            int maxTitleLevel = 0;
            foreach (var style in importedStyles)
            {
                if (style.Name.StartsWith("标题"))
                {
                    int level = GetTitleLevel(style.Name);
                    if (level > maxTitleLevel)
                    {
                        maxTitleLevel = level;
                    }
                }
            }

            // 设置显示标题数
            if (maxTitleLevel > 0)
            {
                显示标题数下拉框.SelectedIndex = maxTitleLevel;
            }
            else
            {
                显示标题数下拉框.SelectedIndex = 4; // 默认显示4级标题
            }

            // 更新样式列表显示
            FilterStylesByTitleCount();
        }

        private void Cmb_LineSpacing_TextChanged(object sender, EventArgs e)
        {
            if (_isLoadingStyle) return; // 加载样式时忽略
            
            // 实时更新样式数据
            OnStyleChanged(sender, e);
        }

        private void Cmb_LineSpacing_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_LineSpacing, Cmb_LineSpacing.Text.EndsWith("行") ? "行" : "磅");
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 底部按钮均匀排布：宽度自适应文字，加载按钮靠左、关闭按钮靠右，
        /// 两端与窗口边距留白（sideMargin），中间按钮等间距分布
        /// </summary>
        private void LayoutBottomButtons()
        {
            var buttons = new[] { 加载, 添加, 删除, 读取文档样式, 导入, 导出, 存预设, 应用当前, Btn_ApplySet, 关闭 };
            const int buttonY = 356;        // 按钮行Y坐标（窗体底部预留状态栏空间）
            const int horizontalPadding = 14; // 按钮文字左右内边距
            const int sideMargin = 20;      // 加载/关闭与窗口左右边距的空白

            // 第一步：根据文字测量宽度，自适应每个按钮宽度
            int totalWidth = 0;
            foreach (var btn in buttons)
            {
                if (btn == null) continue;
                btn.AutoSize = false;
                int textWidth = TextRenderer.MeasureText(btn.Text, btn.Font).Width;
                btn.Width = textWidth + horizontalPadding;
                btn.Height = 31;
                // 窗口改变高度时，按钮保持贴底部（相对 groupBox3 底边固定距离）
                btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                totalWidth += btn.Width;
            }

            // 第二步：加载在最左（留边距），关闭在最右（留边距），中间按钮等间距
            int availableWidth = groupBox3.Width - sideMargin * 2;
            int gap = (availableWidth - totalWidth) / (buttons.Length - 1);
            if (gap < 6) gap = 6; // 最小间距，防止按钮重叠

            int startX = sideMargin;
            foreach (var btn in buttons)
            {
                if (btn == null) continue;
                btn.Location = new Point(startX, buttonY);
                startX += btn.Width + gap;
            }
        }

        /// <summary>
        /// 在窗口底部状态栏显示消息
        /// </summary>
        private void ShowStatusMessage(string message)
        {
            if (statusLabel != null)
            {
                statusLabel.Text = message;
            }
        }

        /// <summary>
        /// 更新样式基准/后续段落样式控件的可见性：
        /// 仅非"正文"样式显示（与 Word 原生修改样式窗口一致，正文样式无样式基准）
        /// </summary>
        private void UpdateBaseStyleControlsVisibility()
        {
            bool hasSelection = _currentStyleIndex >= 0 && _currentStyleIndex < Styles.Count;

            // 多选模式：所有选中样式均非"正文"时显示（含正文则隐藏）
            if (_multiSelectMode)
            {
                var selected = GetSelectedStyles();
                bool containsBody = selected.Any(s => s.Name == "正文");
                bool show = selected.Count > 0 && !containsBody;

                if (label20 != null) label20.Visible = show;
                if (样式基准下拉框 != null) 样式基准下拉框.Visible = show;
                if (label21 != null) label21.Visible = show;
                if (后续段落样式下拉框 != null) 后续段落样式下拉框.Visible = show;
                return;
            }

            bool isBodyStyle = hasSelection && Styles[_currentStyleIndex].Name == "正文";
            bool show2 = hasSelection && !isBodyStyle;

            if (label20 != null) label20.Visible = show2;
            if (样式基准下拉框 != null) 样式基准下拉框.Visible = show2;
            if (label21 != null) label21.Visible = show2;
            if (后续段落样式下拉框 != null) 后续段落样式下拉框.Visible = show2;
        }

        /// <summary>
        /// 获取当前列表中所有选中的样式对象
        /// </summary>
        private List<CustomStyle> GetSelectedStyles()
        {
            var result = new List<CustomStyle>();
            foreach (var item in Lst_Styles.SelectedItems)
            {
                string name = item?.ToString();
                if (string.IsNullOrEmpty(name)) continue;

                CustomStyle style;
                if (!_styleNameMap.TryGetValue(name, out style))
                {
                    style = Styles.FirstOrDefault(s => s.Name == name);
                    if (style != null)
                    {
                        _styleNameMap[name] = style;
                    }
                }
                if (style != null)
                {
                    result.Add(style);
                }
            }
            return result;
        }

        /// <summary>
        /// 多选模式：将各设置项显示为选中样式的共同值，或显示"多种"
        /// </summary>
        private void UpdateControlsForMultiSelect()
        {
            var selectedStyles = GetSelectedStyles();
            if (selectedStyles.Count == 0) return;

            // 中文字体
            ShowMultiValueOrCommon(Cmb_ChnFontName, selectedStyles.Select(s => s.FontName ?? "").ToList(), MultiLevelDataManager.GetChnFontItems());
            // 西文字体
            ShowMultiValueOrCommon(Cmb_EngFontName, selectedStyles.Select(s => !string.IsNullOrEmpty(s.EngFontName) ? s.EngFontName : (s.FontName ?? "")).ToList(), MultiLevelDataManager.GetEngFontItems());
            // 字号
            ShowMultiValueOrCommon(Cmb_FontSize, selectedStyles.Select(s => MultiLevelDataManager.ConvertFontSizeToString(s.FontSize)).ToList(), MultiLevelDataManager.GetFontSizes());

            // 粗体/斜体/下划线（ToggleButton 无三态，全部相同才按下，不同则未按下）
            SetToggleMulti(Btn_Bold, selectedStyles.Select(s => s.Bold).ToList());
            SetToggleMulti(Btn_Italic, selectedStyles.Select(s => s.Italic).ToList());
            SetToggleMulti(Btn_UnderLine, selectedStyles.Select(s => s.Underline).ToList());

            // 字体颜色：相同显示，不同显示灰色（"多种"）
            bool colorSame = selectedStyles.Select(s => s.FontColor.ToArgb()).Distinct().Count() == 1;
            Btn_FontColor.BackColor = colorSame ? selectedStyles[0].FontColor : Color.FromArgb(165, 165, 165);

            // 段落对齐
            ShowMultiValueOrCommon(Cmb_ParaAligment, selectedStyles.Select(s => GetAlignmentText(s.ParaAlignment)).ToList(), WordStyleInfo.HAlignments);

            // 左/右缩进
            ShowMultiValueOrCommon(Nud_LeftIndent, selectedStyles.Select(s => $"{s.LeftIndent:F1} 厘米").ToList());
            ShowMultiValueOrCommon(Nud_RightIndent, selectedStyles.Select(s => $"{s.RightIndent:F1} 厘米").ToList());

            // 首行缩进：方式与值
            var firstLineValues = selectedStyles.Select(s => s.FirstLineIndent).ToList();
            bool firstLineSame = firstLineValues.Distinct().Count() == 1;
            if (firstLineSame)
            {
                RemoveMultiplePlaceholder(首行缩进方式下拉框);
                float v = firstLineValues[0];
                if (v == 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 0; // 无
                    Nud_FirstLineIndent.Value = 0;
                }
                else if (v < 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 1; // 悬挂缩进
                    Nud_FirstLineIndent.Value = (decimal)Math.Abs(v);
                }
                else
                {
                    首行缩进方式下拉框.SelectedIndex = 2; // 首行缩进
                    Nud_FirstLineIndent.Value = (decimal)v;
                }
                UpdateFirstLineIndentVisibility();
            }
            else
            {
                // 方式不同：显示"多种"，缩进值归零
                ShowMultiValueOrCommon(首行缩进方式下拉框, new List<string>(), null);
                Nud_FirstLineIndent.Value = 0;
                UpdateFirstLineIndentVisibility();
            }

            // 行距
            ShowMultiValueOrCommon(Cmb_LineSpacing, selectedStyles.Select(s => ConvertLineSpacingToString(s.LineSpacing)).ToList(), WordStyleInfo.LineSpacings);
            // 段前/段后
            ShowMultiValueOrCommon(Cmb_BefreSpacing, selectedStyles.Select(s => s.BeforeSpacing.ToString("0.0 磅")).ToList(), WordStyleInfo.SpaceBeforeValues);
            ShowMultiValueOrCommon(Cmb_AfterSpacing, selectedStyles.Select(s => s.AfterSpacing.ToString("0.0 磅")).ToList(), WordStyleInfo.SpaceAfterValues);

            // 样式基准 / 后续段落样式
            ShowMultiValueOrCommon(样式基准下拉框, selectedStyles.Select(s => !string.IsNullOrEmpty(s.BaseStyle) ? s.BaseStyle : "(无样式)").ToList(), GetBaseStyleItems());
            ShowMultiValueOrCommon(后续段落样式下拉框, selectedStyles.Select(s => !string.IsNullOrEmpty(s.NextParagraphStyle) ? s.NextParagraphStyle : "正文").ToList(), GetNextParagraphStyleItems());
        }

        /// <summary>
        /// 下拉框/数值框显示共同值或"多种"：值全部相同则显示该值，否则显示"多种"
        /// </summary>
        private void ShowMultiValueOrCommon(StandardComboBox combo, List<string> values, IEnumerable<string> items)
        {
            bool allSame = values.Distinct().Count() <= 1;
            if (allSame)
            {
                RemoveMultiplePlaceholder(combo);
                if (items != null && values.Count > 0)
                {
                    SetComboBoxSelection(combo, values[0], items);
                }
            }
            else
            {
                if (combo.DropDownStyle == ComboBoxStyle.DropDownList)
                {
                    // DropDownList 无法直接显示文本，插入"多种"占位项并选中
                    if (combo.Items.Count == 0 || combo.Items[0] as string != "多种")
                    {
                        combo.Items.Insert(0, "多种");
                    }
                    _multiplePlaceholders.Add(combo.Name);
                    combo.SelectedIndex = 0;
                }
                else
                {
                    combo.Text = "多种";
                }
            }
        }

        /// <summary>
        /// 数值框显示共同值或归零（数值控件无法显示"多种"，不同时归零，用户输入即覆盖全部）
        /// </summary>
        private void ShowMultiValueOrCommon(StandardNumericUpDown nud, List<string> values)
        {
            bool allSame = values.Distinct().Count() <= 1;
            if (allSame && values.Count > 0)
            {
                SetIndentValue(nud, values[0]);
            }
            else
            {
                nud.Value = 0;
            }
        }

        /// <summary>
        /// 多选模式切换按钮：全部相同才按下（ToggleButton 无三态，无法显示"多种"）
        /// </summary>
        private void SetToggleMulti(ToggleButton button, List<bool> values)
        {
            bool allSame = values.Distinct().Count() <= 1;
            button.Pressed = allSame && values.Count > 0 && values[0];
        }

        /// <summary>
        /// 移除下拉框中的"多种"占位项（用户选择真实值或退出多选时调用）
        /// </summary>
        private void RemoveMultiplePlaceholder(StandardComboBox combo)
        {
            if (_multiplePlaceholders.Contains(combo.Name) &&
                combo.Items.Count > 0 && combo.Items[0] as string == "多种")
            {
                combo.Items.RemoveAt(0);
                _multiplePlaceholders.Remove(combo.Name);
            }
        }

        /// <summary>
        /// 移除所有"多种"占位项（退出多选模式时调用）
        /// </summary>
        private void RemoveAllMultiplePlaceholders()
        {
            var combos = new[] { Cmb_ChnFontName, Cmb_EngFontName, Cmb_FontSize, Cmb_ParaAligment,
                首行缩进方式下拉框, Cmb_LineSpacing, Cmb_BefreSpacing, Cmb_AfterSpacing,
                样式基准下拉框, 后续段落样式下拉框 };
            foreach (var combo in combos)
            {
                if (combo == null) continue;
                RemoveMultiplePlaceholder(combo);
            }
            _multiplePlaceholders.Clear();
        }

        /// <summary>
        /// 多选批量应用：将指定控件的当前值写入样式对象
        /// </summary>
        private void ReadControlValueToStyle(string controlName, CustomStyle style)
        {
            switch (controlName)
            {
                case "Cmb_ChnFontName":
                    style.FontName = Cmb_ChnFontName.Text;
                    break;
                case "Cmb_EngFontName":
                    style.EngFontName = Cmb_EngFontName.Text;
                    break;
                case "Cmb_FontSize":
                    if (!string.IsNullOrEmpty(Cmb_FontSize.Text))
                    {
                        style.FontSize = MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.Text);
                    }
                    break;
                case "Btn_Bold":
                    style.Bold = Btn_Bold.Pressed;
                    break;
                case "Btn_Italic":
                    style.Italic = Btn_Italic.Pressed;
                    break;
                case "Btn_UnderLine":
                    style.Underline = Btn_UnderLine.Pressed;
                    break;
                case "Btn_FontColor":
                    style.FontColor = Btn_FontColor.BackColor;
                    break;
                case "Cmb_ParaAligment":
                    if (Cmb_ParaAligment.SelectedIndex >= 0)
                    {
                        style.ParaAlignment = Cmb_ParaAligment.SelectedIndex;
                    }
                    break;
                case "Nud_LeftIndent":
                    style.LeftIndent = (float)Nud_LeftIndent.GetValueInCentimeters();
                    break;
                case "Nud_RightIndent":
                    style.RightIndent = (float)Nud_RightIndent.GetValueInCentimeters();
                    break;
                case "首行缩进方式下拉框":
                    if (首行缩进方式下拉框.SelectedIndex == 0)
                    {
                        style.FirstLineIndent = 0f;
                    }
                    else if (首行缩进方式下拉框.SelectedIndex == 1)
                    {
                        style.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    else if (首行缩进方式下拉框.SelectedIndex == 2)
                    {
                        style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    break;
                case "Nud_FirstLineIndent":
                    if (首行缩进方式下拉框.SelectedIndex == 1)
                    {
                        style.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    else if (首行缩进方式下拉框.SelectedIndex == 2)
                    {
                        style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    break;
                case "Cmb_LineSpacing":
                    if (!string.IsNullOrEmpty(Cmb_LineSpacing.Text))
                    {
                        style.LineSpacing = ConvertLineSpacingToFloat(Cmb_LineSpacing.Text);
                    }
                    break;
                case "Cmb_BefreSpacing":
                    style.BeforeSpacing = ConvertSpacingToPoints(Cmb_BefreSpacing.Text);
                    break;
                case "Cmb_AfterSpacing":
                    style.AfterSpacing = ConvertSpacingToPoints(Cmb_AfterSpacing.Text);
                    break;
                case "样式基准下拉框":
                    style.BaseStyle = 样式基准下拉框.SelectedItem?.ToString() ?? "";
                    break;
                case "后续段落样式下拉框":
                    style.NextParagraphStyle = 后续段落样式下拉框.SelectedItem?.ToString() ?? "";
                    break;
            }
        }

        /// <summary>
        /// 获取样式基准下拉框选项：(无样式)、正文 + 当前所有样式名
        /// </summary>
        private List<string> GetBaseStyleItems()
        {
            var items = new List<string> { "(无样式)", "正文" };
            foreach (var style in Styles)
            {
                if (style != null && !string.IsNullOrEmpty(style.Name) && !items.Contains(style.Name))
                {
                    items.Add(style.Name);
                }
            }
            return items;
        }

        /// <summary>
        /// 获取后续段落样式下拉框选项：正文、无间隔 + 当前所有样式名
        /// </summary>
        private List<string> GetNextParagraphStyleItems()
        {
            var items = new List<string> { "正文", "无间隔" };
            foreach (var style in Styles)
            {
                if (style != null && !string.IsNullOrEmpty(style.Name) && !items.Contains(style.Name))
                {
                    items.Add(style.Name);
                }
            }
            return items;
        }

        /// <summary>
        /// 刷新样式基准和后续段落样式下拉框选项（保留当前选中值）
        /// </summary>
        private void RefreshBaseStyleComboBoxes()
        {
            if (样式基准下拉框 == null || 后续段落样式下拉框 == null)
                return;

            string baseSelection = 样式基准下拉框.SelectedItem?.ToString() ?? 样式基准下拉框.Text;
            string nextSelection = 后续段落样式下拉框.SelectedItem?.ToString() ?? 后续段落样式下拉框.Text;

            var baseItems = GetBaseStyleItems();
            var nextItems = GetNextParagraphStyleItems();

            InitializeComboBox(样式基准下拉框, baseItems);
            if (!string.IsNullOrEmpty(baseSelection))
            {
                SetComboBoxSelection(样式基准下拉框, baseSelection, baseItems);
            }

            InitializeComboBox(后续段落样式下拉框, nextItems);
            if (!string.IsNullOrEmpty(nextSelection))
            {
                SetComboBoxSelection(后续段落样式下拉框, nextSelection, nextItems);
            }
        }

        /// <summary>
        /// 重置控件到默认状态
        /// </summary>
        private void ResetControlsToDefault()
        {
            // 设置加载标志，防止触发保存事件
            _isLoadingStyle = true;

            try
            {
                // 重置字体设置（使用与下拉框一致的优先字体列表，避免索引错位）
                SetComboBoxSelection(Cmb_ChnFontName, "宋体", MultiLevelDataManager.GetChnFontItems());
                SetComboBoxSelection(Cmb_EngFontName, "宋体", MultiLevelDataManager.GetEngFontItems());
                SetComboBoxSelection(Cmb_FontSize, "五号", MultiLevelDataManager.GetFontSizes());

                // 重置字体样式
                Btn_Bold.Pressed = false;
                Btn_Italic.Pressed = false;
                Btn_UnderLine.Pressed = false;
                Btn_FontColor.BackColor = Color.Black;

                // 重置段落设置
                SetComboBoxSelection(Cmb_ParaAligment, "左对齐", WordStyleInfo.HAlignments);
                SetComboBoxSelection(Cmb_LineSpacing, "单倍行距", WordStyleInfo.LineSpacings);

                // 重置缩进设置
                Nud_LeftIndent.Value = 0;
                Nud_RightIndent.Value = 0;
                Nud_FirstLineIndent.Value = 0;
                首行缩进方式下拉框.SelectedIndex = 0; // 无
                UpdateFirstLineIndentVisibility();

                // 重置段间距
                SetComboBoxSelection(Cmb_BefreSpacing, "0.0 磅", WordStyleInfo.SpaceBeforeValues);
                SetComboBoxSelection(Cmb_AfterSpacing, "0.0 磅", WordStyleInfo.SpaceAfterValues);

                // 清空样式列表选择
                Lst_Styles.SelectedIndex = -1;
                _currentStyleIndex = -1;

                // 禁用面板
                Pal_Font.Enabled = false;
                Pal_ParaIndent.Enabled = false;

                // 更新按钮状态
                删除.Enabled = false;
                添加.Enabled = true;

                // 隐藏样式基准/后续段落样式（无选中样式时不显示）
                UpdateBaseStyleControlsVisibility();
            }
            finally
            {
                // 重置加载标志
                _isLoadingStyle = false;
            }
        }

        private void LoadStyleToControls(CustomStyle style)
        {
            if (style == null)
                return;

            // 设置加载标志，防止触发保存事件
            _isLoadingStyle = true;

            try
            {
                // 加载样式的所有属性到控件；样式未定义的属性回退到合理默认值，
                // 避免控件残留上一个样式的值（串值）

                // 中文字体（空值回退宋体，与 ApplyStyleToDocument 逻辑一致）
                string chnFont = !string.IsNullOrEmpty(style.FontName) ? style.FontName : "宋体";
                SetComboBoxSelection(Cmb_ChnFontName, chnFont, MultiLevelDataManager.GetChnFontItems());
                // 西文字体（空值回退中文字体或宋体）
                string engFont = !string.IsNullOrEmpty(style.EngFontName) ? style.EngFontName : chnFont;
                SetComboBoxSelection(Cmb_EngFontName, engFont, MultiLevelDataManager.GetEngFontItems());

                // 字号（未定义/0 回退到正文默认字号）
                float fontSize = style.FontSize > 0 ? style.FontSize : 10.5f;
                SetComboBoxSelection(Cmb_FontSize, MultiLevelDataManager.ConvertFontSizeToString(fontSize), MultiLevelDataManager.GetFontSizes());

                // 字体样式
                Btn_Bold.Pressed = style.Bold;
                Btn_Italic.Pressed = style.Italic;
                Btn_UnderLine.Pressed = style.Underline;
                // 字体颜色
                Btn_FontColor.BackColor = style.FontColor;

                // 段落对齐（无效值回退左对齐）
                int alignment = style.ParaAlignment >= 0 ? style.ParaAlignment : 0;
                SetComboBoxSelection(Cmb_ParaAligment, GetAlignmentText(alignment), WordStyleInfo.HAlignments);

                // 行距（始终设置，默认单倍）
                float lineSpacing = style.LineSpacing > 0 ? style.LineSpacing : 1.0f;
                SetComboBoxSelection(Cmb_LineSpacing, ConvertLineSpacingToString(lineSpacing), WordStyleInfo.LineSpacings);

                // 设置缩进控件（使用厘米单位显示）
                SetIndentValue(Nud_LeftIndent, $"{style.LeftIndent:F1} 厘米");
                SetIndentValue(Nud_RightIndent, $"{style.RightIndent:F1} 厘米");

                // 根据首行缩进值设置首行缩进方式和控件
                if (style.FirstLineIndent == 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 0; // 无
                    UpdateFirstLineIndentVisibility();
                }
                else if (style.FirstLineIndent < 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 1; // 悬挂缩进
                    UpdateFirstLineIndentVisibility();
                    SetIndentValue(Nud_FirstLineIndent, $"{Math.Abs(style.FirstLineIndent):F1} 厘米");
                }
                else
                {
                    首行缩进方式下拉框.SelectedIndex = 2; // 首行缩进
                    UpdateFirstLineIndentVisibility();
                    SetIndentValue(Nud_FirstLineIndent, $"{style.FirstLineIndent:F1} 厘米");
                }

                // 段落间距（始终设置，0 磅也要显示，避免残留旧值）
                SetComboBoxSelection(Cmb_BefreSpacing, style.BeforeSpacing.ToString("0.0 磅"), WordStyleInfo.SpaceBeforeValues);
                SetComboBoxSelection(Cmb_AfterSpacing, style.AfterSpacing.ToString("0.0 磅"), WordStyleInfo.SpaceAfterValues);

                // 样式基准（空值回退"(无样式)"）
                string baseStyleValue = !string.IsNullOrEmpty(style.BaseStyle) ? style.BaseStyle : "(无样式)";
                SetComboBoxSelection(样式基准下拉框, baseStyleValue, GetBaseStyleItems());
                // 后续段落样式（空值回退"正文"）
                string nextStyleValue = !string.IsNullOrEmpty(style.NextParagraphStyle) ? style.NextParagraphStyle : "正文";
                SetComboBoxSelection(后续段落样式下拉框, nextStyleValue, GetNextParagraphStyleItems());

                // 更新样式基准/后续段落样式可见性（正文样式不显示）
                UpdateBaseStyleControlsVisibility();
            }
            finally
            {
                // 重置加载标志
                _isLoadingStyle = false;
            }
        }

        private void UpdateCurrentStyle()
        {
            if (Lst_Styles == null || Lst_Styles.SelectedIndex < 0 || Lst_Styles.SelectedItem == null)
                return;

            string selectedStyle = Lst_Styles.SelectedItem.ToString();
            if (string.IsNullOrEmpty(selectedStyle))
                return;

            // 使用字典查找，提升性能
            CustomStyle style;
            if (!_styleNameMap.TryGetValue(selectedStyle, out style))
            {
                style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                if (style != null)
                {
                    _styleNameMap[selectedStyle] = style;
                }
            }

            if (style != null)
            {
                UpdateStyleFromControls(style);
                UpdateStyleInfo(style);
            }
        }


        private void UpdateStyleFromControls(CustomStyle style)
        {
            if (style == null)
                return;

            // 分别获取中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
            style.FontName = Cmb_ChnFontName?.SelectedItem?.ToString() ?? Cmb_ChnFontName?.Text ?? "宋体";
            style.EngFontName = Cmb_EngFontName?.SelectedItem?.ToString() ?? Cmb_EngFontName?.Text ?? style.FontName;
            if (Cmb_FontSize.SelectedIndex >= 0)
            {
                style.FontSize = MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.SelectedItem?.ToString());
            }
            style.Bold = Btn_Bold?.Pressed ?? false;
            style.Italic = Btn_Italic?.Pressed ?? false;
            style.Underline = Btn_UnderLine?.Pressed ?? false;
            style.FontColor = Btn_FontColor?.BackColor ?? Color.Black;
            style.ParaAlignment = Cmb_ParaAligment?.SelectedIndex ?? 0;
            // 从控件读取缩进值，直接使用厘米单位
            style.LeftIndent = Nud_LeftIndent != null ? (float)Nud_LeftIndent.GetValueInCentimeters() : 0f;
            style.RightIndent = Nud_RightIndent != null ? (float)Nud_RightIndent.GetValueInCentimeters() : 0f;

            // 根据首行缩进方式设置首行缩进值
            int firstLineIndentType = 首行缩进方式下拉框?.SelectedIndex ?? 0;
            if (firstLineIndentType == 0) // 无
            {
                style.FirstLineIndent = 0f;
            }
            else if (firstLineIndentType == 1 && Nud_FirstLineIndent != null) // 悬挂缩进
            {
                style.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters(); // 悬挂缩进为负值
            }
            else if (firstLineIndentType == 2 && Nud_FirstLineIndent != null) // 首行缩进
            {
                style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
            }

            // 处理行距设置
            string lineSpacingText = Cmb_LineSpacing?.Text ?? string.Empty;
            if (!string.IsNullOrEmpty(lineSpacingText))
            {
                style.LineSpacing = ConvertLineSpacingToFloat(lineSpacingText);
            }
            else
            {
                style.LineSpacing = 1.0f; // 默认单倍行距
            }

            // 从控件读取段落间距值，直接使用磅单位
            style.BeforeSpacing = ConvertSpacingToPoints(Cmb_BefreSpacing?.Text ?? string.Empty);
            style.AfterSpacing = ConvertSpacingToPoints(Cmb_AfterSpacing?.Text ?? string.Empty);

            // 读取样式基准和后续段落样式
            style.BaseStyle = 样式基准下拉框?.SelectedItem?.ToString() ?? 样式基准下拉框?.Text ?? "";
            style.NextParagraphStyle = 后续段落样式下拉框?.SelectedItem?.ToString() ?? 后续段落样式下拉框?.Text ?? "";
        }

        private CustomStyle CreateStyleFromControls(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            // 分别获取中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
            string fontName = !string.IsNullOrEmpty(Cmb_ChnFontName?.Text) ? Cmb_ChnFontName.Text : (Cmb_EngFontName?.Text ?? "宋体");

            // 处理行距设置
            float lineSpacing = 1.0f; // 默认单倍行距
            string lineSpacingText = Cmb_LineSpacing?.Text ?? string.Empty;
            if (!string.IsNullOrEmpty(lineSpacingText))
            {
                lineSpacing = ConvertLineSpacingToFloat(lineSpacingText);
            }

            int firstLineIndentType = 首行缩进方式下拉框?.SelectedIndex ?? 0;
            float firstLineIndent = 0f;
            if (firstLineIndentType != 0 && Nud_FirstLineIndent != null)
            {
                firstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
            }

            return new CustomStyle(
                name: name,
                fontName: fontName,
                engFontName: Cmb_EngFontName?.SelectedItem?.ToString() ?? Cmb_EngFontName?.Text ?? fontName,
                fontSize: Cmb_FontSize?.SelectedIndex >= 0 ? MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.SelectedItem?.ToString()) : 0f,
                bold: Btn_Bold?.Pressed ?? false,
                italic: Btn_Italic?.Pressed ?? false,
                underline: Btn_UnderLine?.Pressed ?? false,
                fontColor: Btn_FontColor?.BackColor ?? Color.Black,
                paraAlignment: Cmb_ParaAligment?.SelectedIndex ?? 0,
                leftIndent: Nud_LeftIndent != null ? (float)Nud_LeftIndent.GetValueInCentimeters() : 0f,
                rightIndent: Nud_RightIndent != null ? (float)Nud_RightIndent.GetValueInCentimeters() : 0f,
                firstLineIndent: firstLineIndent,
                firstLineIndentByChar: 0,
                lineSpacing: lineSpacing,
                beforeBreak: false,
                beforeSpacing: ConvertSpacingToPoints(Cmb_BefreSpacing?.Text ?? string.Empty),
                afterSpacing: ConvertSpacingToPoints(Cmb_AfterSpacing?.Text ?? string.Empty),
                numberStyle: 0,
                numberFormat: null,
                userDefined: true,
                baseStyle: 样式基准下拉框?.SelectedItem?.ToString() ?? 样式基准下拉框?.Text ?? "正文",
                nextParagraphStyle: 后续段落样式下拉框?.SelectedItem?.ToString() ?? 后续段落样式下拉框?.Text ?? "正文"
            );
        }

        private void UpdateStyleInfo(CustomStyle style)
        {
            // 调试输出样式信息
            System.Diagnostics.Debug.WriteLine($"=== 样式信息更新 ===");
            System.Diagnostics.Debug.WriteLine($"样式名称：{style.Name}");
            System.Diagnostics.Debug.WriteLine($"字体：{style.FontName ?? "默认"}，大小：{style.FontSize}磅");
            System.Diagnostics.Debug.WriteLine($"格式：{(style.Bold ? "粗体 " : "")}{(style.Italic ? "斜体 " : "")}{(style.Underline ? "下划线" : "")}");
            System.Diagnostics.Debug.WriteLine($"对齐：{GetAlignmentText(style.ParaAlignment)}");
            System.Diagnostics.Debug.WriteLine($"缩进：左{style.LeftIndent:F1}厘米，首行{style.FirstLineIndent:F1}厘米");
            System.Diagnostics.Debug.WriteLine($"行距：{ConvertLineSpacingToString(style.LineSpacing)}");
            System.Diagnostics.Debug.WriteLine($"段前：{style.BeforeSpacing:F1}磅，段后：{style.AfterSpacing:F1}磅");
            System.Diagnostics.Debug.WriteLine($"==================");
        }

        // 行距文本到数值的映射字典（提升查找性能）
        private static readonly Dictionary<string, float> LineSpacingMap = new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase)
        {
            { "单倍行距", 1.0f },
            { "1.2倍行距", 1.2f },
            { "1.25倍行距", 1.25f },
            { "1.5倍行距", 1.5f },
            { "双倍行距", 2.0f }
        };

        /// <summary>
        /// 将行距文本转换为数值（复用多级段落设置的方法）
        /// </summary>
        private float ConvertLineSpacingToFloat(string lineSpacingText)
        {
            if (string.IsNullOrEmpty(lineSpacingText))
                return 1.0f;

            try
            {
                // 首先查找预定义的行距值
                if (LineSpacingMap.ContainsKey(lineSpacingText))
                {
                    return LineSpacingMap[lineSpacingText];
                }

                // 处理多倍行距格式（如 "2.5倍行距"）
                if (lineSpacingText.EndsWith("倍行距"))
                {
                    string valueText = lineSpacingText.Replace("倍行距", "").Trim();
                    if (float.TryParse(valueText, out float multipleValue))
                    {
                        return multipleValue;
                    }
                }
                // 处理固定行距格式（如 "12磅"）
                else if (lineSpacingText.EndsWith("磅"))
                {
                    string valueText = lineSpacingText.Replace("磅", "").Trim();
                    if (float.TryParse(valueText, out float exactValue))
                    {
                        return exactValue;
                    }
                }
                // 尝试直接解析为数值（可能是用户输入的自定义值）
                else if (float.TryParse(lineSpacingText, out float directValue))
                {
                    return directValue;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换行距时出错：{ex.Message}");
            }

            return 1.0f; // 默认单倍行距
        }

        /// <summary>
        /// 将行距数值转换为文本（复用多级段落设置的逻辑）
        /// </summary>
        private string ConvertLineSpacingToString(float lineSpacing)
        {
            try
            {
                // 复用多级段落设置的转换逻辑
                if (lineSpacing == 1.0f)
                {
                    return "单倍行距";
                }
                else if (Math.Abs(lineSpacing - 1.2f) < 0.01f)
                {
                    return "1.2倍行距";
                }
                else if (Math.Abs(lineSpacing - 1.25f) < 0.01f)
                {
                    return "1.25倍行距";
                }
                else if (lineSpacing == 1.5f)
                {
                    return "1.5倍行距";
                }
                else if (lineSpacing == 2.0f)
                {
                    return "双倍行距";
                }
                else
                {
                    return $"{lineSpacing:0.0} 倍行距";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换行距文本时出错：{ex.Message}");
                return "单倍行距";
            }
        }

        /// <summary>
        /// 将段落间距文本转换为磅单位（复用多级段落设置的逻辑）
        /// </summary>
        private float ConvertSpacingToPoints(string spacingText)
        {
            if (string.IsNullOrEmpty(spacingText))
                return 0f;

            try
            {
                string trimmedText = spacingText.Trim();
                
                // 处理"磅"单位
                if (trimmedText.EndsWith("磅"))
                {
                    string valueText = trimmedText.Replace("磅", "").Trim();
                    if (float.TryParse(valueText, out float result))
                    {
                        return result;
                    }
                }
                // 处理"行"单位
                else if (trimmedText.EndsWith("行"))
                {
                    string valueText = trimmedText.Replace("行", "").Trim();
                    if (float.TryParse(valueText, out float result))
                    {
                        return MultiLevelDataManager.LinesToPoints(result);
                    }
                }
                // 处理"厘米"单位
                else if (trimmedText.EndsWith("厘米"))
                {
                    string valueText = trimmedText.Replace("厘米", "").Trim();
                    if (float.TryParse(valueText, out float result))
                    {
                        return MultiLevelDataManager.CentimetersToPoints(result);
                    }
                }
                // 尝试直接解析为数值（可能是用户输入的自定义值）
                else if (float.TryParse(trimmedText, out float directValue))
                {
                    return directValue;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换段落间距时出错：{ex.Message}");
            }

            return 0f;
        }


        private void ApplyStylesToDocument(IEnumerable<string> styleNames = null)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;

            // 默认应用当前显示的样式（StyleNames中的样式）；指定时只应用指定的样式
            var namesToApply = styleNames ?? StyleNames.ToList();

            foreach (var styleName in namesToApply)
            {
                try
                {
                    // 从字典或列表中查找样式
                    CustomStyle style;
                    if (!_styleNameMap.TryGetValue(styleName, out style))
                    {
                        style = Styles.FirstOrDefault(s => s.Name == styleName);
                    }

                    if (style != null)
                    {
                        // 应用样式到文档
                        ApplyStyleToDocument(doc, style);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"应用样式 {styleName} 时出错：{ex.Message}");
                }
            }
        }

        private void ApplyStyleToDocument(Document doc, CustomStyle style)
        {
            try
            {
                // 检查样式是否存在，如果不存在则创建
                Style wordStyle = null;
                try
                {
                    wordStyle = doc.Styles[style.Name];
                }
                catch
                {
                    // 样式不存在，创建新样式
                    wordStyle = doc.Styles.Add(style.Name, WdStyleType.wdStyleTypeParagraph);
                }

                if (wordStyle != null)
                {
                    // 设置字体（分别设置中文字体和英文字体）
                    if (!string.IsNullOrEmpty(style.FontName))
                    {
                        wordStyle.Font.NameFarEast = style.FontName;
                    }
                    if (!string.IsNullOrEmpty(style.EngFontName))
                    {
                        wordStyle.Font.NameAscii = style.EngFontName;
                    }
                    else if (!string.IsNullOrEmpty(style.FontName))
                    {
                        // 如果英文字体为空，使用中文字体
                        wordStyle.Font.NameAscii = style.FontName;
                    }

                    if (style.FontSize > 0)
                    {
                        wordStyle.Font.Size = style.FontSize;
                    }

                    wordStyle.Font.Bold = style.Bold ? -1 : 0;
                    wordStyle.Font.Italic = style.Italic ? -1 : 0;
                    wordStyle.Font.Underline = style.Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;

                    // 设置字体颜色（使用ColorTranslator.ToOle方法，更可靠）
                    try
                    {
                        if (style.FontColor != null)
                        {
                            int oleColor = ColorTranslator.ToOle(style.FontColor);
                            wordStyle.Font.Color = (WdColor)oleColor;
                        }
                        else
                        {
                            wordStyle.Font.Color = WdColor.wdColorAutomatic;
                        }
                    }
                    catch
                    {
                        wordStyle.Font.Color = WdColor.wdColorAutomatic;
                    }

                    // 设置段落格式
                    var paragraphFormat = wordStyle.ParagraphFormat;

                    // 设置对齐方式
                    switch (style.ParaAlignment)
                    {
                        case 0: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case 1: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case 2: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; break;
                        case 3: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; break;
                        case 4: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute; break;
                        default: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                    }

                    // 设置缩进（转换为磅）
                    var app = Globals.ThisAddIn.Application;
                    paragraphFormat.LeftIndent = app.CentimetersToPoints(style.LeftIndent);
                    paragraphFormat.RightIndent = app.CentimetersToPoints(style.RightIndent);
                    paragraphFormat.FirstLineIndent = app.CentimetersToPoints(style.FirstLineIndent);

                    // 设置行距（复用多级段落设置的逻辑）
                    if (style.LineSpacing > 0)
                    {
                        if (style.LineSpacing == 1.0f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                        }
                        else if (style.LineSpacing == 1.2f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            paragraphFormat.LineSpacing = MultiLevelDataManager.LinesToPoints(1.2f);
                        }
                        else if (style.LineSpacing == 1.25f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            paragraphFormat.LineSpacing = MultiLevelDataManager.LinesToPoints(1.25f);
                        }
                        else if (style.LineSpacing == 1.5f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                        }
                        else if (style.LineSpacing == 2.0f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                        }
                        else if (style.LineSpacing < 1.0f)
                        {
                            // 固定行距（磅值）
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                            paragraphFormat.LineSpacing = style.LineSpacing;
                        }
                        else if (style.LineSpacing > 2.0f)
                        {
                            // 大于2倍的行距，使用多倍行距
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            paragraphFormat.LineSpacing = MultiLevelDataManager.LinesToPoints(style.LineSpacing);
                        }
                        else
                        {
                            // 多倍行距
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            paragraphFormat.LineSpacing = MultiLevelDataManager.LinesToPoints(style.LineSpacing);
                        }
                    }

                    // 设置段前段后间距（已经是磅单位，直接使用）
                    paragraphFormat.SpaceBefore = style.BeforeSpacing;
                    paragraphFormat.SpaceAfter = style.AfterSpacing;

                    // 设置分页
                    paragraphFormat.PageBreakBefore = style.BeforeBreak ? -1 : 0;

                    // 设置样式基准（与 Word 原生修改样式窗口一致；"(无样式)"表示不设置基准）
                    if (!string.IsNullOrEmpty(style.BaseStyle) && style.BaseStyle != "(无样式)")
                    {
                        try
                        {
                            object baseStyleRef = doc.Styles[style.BaseStyle];
                            wordStyle.set_BaseStyle(ref baseStyleRef);
                        }
                        catch
                        {
                            System.Diagnostics.Debug.WriteLine($"设置样式基准 {style.BaseStyle} 失败，可能不存在该样式");
                        }
                    }

                    // 设置后续段落样式（与 Word 原生修改样式窗口一致）
                    if (!string.IsNullOrEmpty(style.NextParagraphStyle))
                    {
                        try
                        {
                            object nextStyleRef = doc.Styles[style.NextParagraphStyle];
                            wordStyle.set_NextParagraphStyle(ref nextStyleRef);
                        }
                        catch
                        {
                            System.Diagnostics.Debug.WriteLine($"设置后续段落样式 {style.NextParagraphStyle} 失败，可能不存在该样式");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"应用样式 {style.Name} 失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档中的所有样式名称
        /// </summary>
        private List<string> GetDocumentStyles()
        {
            var styles = new List<string>();

            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveDocument?.Styles != null)
                {
                    // 保存当前选择状态
                    var originalSelection = app.Selection;
                    var originalRange = originalSelection.Range;

                    try
                    {
                        // 遍历样式列表，但不影响文档选择
                        var stylesCollection = app.ActiveDocument.Styles;
                        for (int i = 1; i <= stylesCollection.Count; i++)
                        {
                            try
                            {
                                var style = stylesCollection[i];
                                // 只添加用户定义的样式和内置样式
                                if (style.BuiltIn || style.InUse)
                                {
                                    styles.Add(style.NameLocal);
                                }
                            }
                            catch
                            {
                                // 忽略无法访问的样式
                                continue;
                            }
                        }
                    }
                    finally
                    {
                        // 恢复原始选择状态
                        try
                        {
                            originalRange.Select();
                        }
                        catch
                        {
                            // 如果无法恢复选择，忽略错误
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取文档样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return styles;
        }

        private void ReadDocumentStyles()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                // 先从 WordOpenXML 解析样式颜色映射（解决 interop Font.Color 对主题色返回负数的问题）
                BuildStyleColorMap(doc);

                // 清空现有样式列表
                Styles.Clear();
                StyleNames.Clear();

                // 定义要优先加载的样式名称（正文、标题、题注等）
                var priorityStyleNames = new[] { "正文", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9", "题注", "表内文字" };

                // 首先读取优先样式
                foreach (var styleName in priorityStyleNames)
                {
                    try
                    {
                        var wordStyle = doc.Styles[styleName];
                        if (wordStyle != null && wordStyle.Type == WdStyleType.wdStyleTypeParagraph)
                        {
                            var customStyle = CreateCustomStyleFromWordStyle(wordStyle, app, doc);
                            if (customStyle != null && !string.IsNullOrEmpty(customStyle.Name))
                            {
                                Styles.Add(customStyle);
                                StyleNames.Add(customStyle.Name);
                                _styleNameMap[customStyle.Name] = customStyle;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"读取优先样式 {styleName} 时出错：{ex.Message}");
                    }
                }

                // 然后读取文档中的其他段落样式
                try
                {
                    var stylesCollection = doc.Styles;
                    for (int i = 1; i <= stylesCollection.Count; i++)
                    {
                        try
                        {
                            var wordStyle = stylesCollection[i];

                            // 只处理段落样式，且不在优先列表中
                            if (wordStyle.Type == WdStyleType.wdStyleTypeParagraph &&
                                !priorityStyleNames.Contains(wordStyle.NameLocal))
                            {
                                // 只添加用户定义的样式或正在使用的样式
                                if (!wordStyle.BuiltIn || wordStyle.InUse)
                                {
                                    var customStyle = CreateCustomStyleFromWordStyle(wordStyle, app, doc);
                                    if (customStyle != null)
                                    {
                                        Styles.Add(customStyle);
                                        StyleNames.Add(customStyle.Name);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"读取样式 {i} 时出错：{ex.Message}");
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"遍历文档样式时出错：{ex.Message}");
                }

                // 根据当前选择的标题数过滤样式
                FilterStylesByTitleCount();

                // 刷新样式基准和后续段落样式下拉框选项
                RefreshBaseStyleComboBoxes();

                // 自动选中样式（优先"正文"），将读取到的样式属性填充到右侧面板
                if (StyleNames.Count > 0)
                {
                    int selectIndex = StyleNames.IndexOf("正文");
                    if (selectIndex < 0) selectIndex = 0;
                    Lst_Styles.SelectedIndex = selectIndex;
                }

                System.Diagnostics.Debug.WriteLine($"成功加载了 {Styles.Count} 个文档样式");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"读取文档样式时出错：{ex.Message}");
                // 如果读取失败，确保至少有一些默认样式
                if (Styles.Count == 0)
                {
                    InitializeDefaultStyles();
                    FilterStylesByTitleCount();
                }
            }
        }

        /// <summary>
        /// 从Word样式创建自定义样式对象
        /// </summary>
        private CustomStyle CreateCustomStyleFromWordStyle(Style wordStyle, Microsoft.Office.Interop.Word.Application app, Microsoft.Office.Interop.Word.Document doc)
        {
            try
            {
                // 字体名称回退链：NameFarEast(NameAscii) 可能为空字符串（如主题字体样式），
                // 逐级回退并剔除空白，避免下拉框显示为空
                string chnFontName = SafeGetFontName(wordStyle.Font.NameFarEast, wordStyle.Font.NameAscii, wordStyle.Font.Name);
                string engFontName = SafeGetFontName(wordStyle.Font.NameAscii, wordStyle.Font.Name, wordStyle.Font.NameFarEast);

                return new CustomStyle(
                    name: wordStyle.NameLocal,
                    fontName: chnFontName,
                    engFontName: engFontName,
                    fontSize: wordStyle.Font.Size,
                    bold: wordStyle.Font.Bold == (int)WdConstants.wdToggle,
                    italic: wordStyle.Font.Italic == (int)WdConstants.wdToggle,
                    underline: wordStyle.Font.Underline != WdUnderline.wdUnderlineNone,
                    fontColor: GetWordFontColor(wordStyle.Font, doc, wordStyle),
                    paraAlignment: (int)wordStyle.ParagraphFormat.Alignment,
                    leftIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.LeftIndent),
                    rightIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.RightIndent),
                    firstLineIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.FirstLineIndent),
                    firstLineIndentByChar: 0, // Word API不直接提供字符单位
                    lineSpacing: ConvertWordLineSpacingToFloat(wordStyle.ParagraphFormat),
                    beforeSpacing: (float)wordStyle.ParagraphFormat.SpaceBefore,
                    beforeBreak: wordStyle.ParagraphFormat.PageBreakBefore != 0,
                    afterSpacing: (float)wordStyle.ParagraphFormat.SpaceAfter,
                    numberStyle: 0,
                    numberFormat: null,
                    userDefined: !wordStyle.BuiltIn,
                    baseStyle: GetStyleBaseName(wordStyle),
                    nextParagraphStyle: GetStyleNextParagraphName(wordStyle)
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"创建自定义样式时出错：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 按顺序取第一个非空白字符串（Trim 后），全部空白返回空串。
        /// 用于字体名称回退链：NameFarEast / NameAscii / Name 可能返回 null 或空白。
        /// </summary>
        private string SafeGetFontName(params object[] candidates)
        {
            foreach (var candidate in candidates)
            {
                if (candidate == null) continue;
                string text = candidate.ToString();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text.Trim();
                }
            }
            return "";
        }

        /// <summary>
        /// 将Word行距格式转换为统一的行距值（复用多级段落设置的逻辑）
        /// </summary>
        private float ConvertWordLineSpacingToFloat(ParagraphFormat paragraphFormat)
        {
            try
            {
                var lineSpacingRule = paragraphFormat.LineSpacingRule;
                var lineSpacing = paragraphFormat.LineSpacing;

                // 复用多级段落设置的转换逻辑
                switch (lineSpacingRule)
                {
                    case WdLineSpacing.wdLineSpaceSingle:
                        return 1.0f;
                    case WdLineSpacing.wdLineSpace1pt5:
                        return 1.5f;
                    case WdLineSpacing.wdLineSpaceDouble:
                        return 2.0f;
                    case WdLineSpacing.wdLineSpaceMultiple:
                        // 多倍行距，保存具体的倍数值
                        float multipleValue = MultiLevelDataManager.PointsToLines(lineSpacing);
                        if (Math.Abs(multipleValue - 1.2f) < 0.01f)
                            return 1.2f;
                        else if (Math.Abs(multipleValue - 1.25f) < 0.01f)
                            return 1.25f;
                        else
                            return multipleValue;
                    case WdLineSpacing.wdLineSpaceExactly:
                        // 固定行距，保存磅值
                        return lineSpacing;
                    default:
                        return 1.0f; // 默认单倍行距
                }
            }
            catch
            {
                return 1.0f; // 出错时返回默认单倍行距
            }
        }

        /// <summary>
        /// 获取Word字体颜色。
        /// 样式级颜色若是主题色（Word 内置样式如标题 1 常见），Font.Color 返回"自动"（wdColorAutomatic），
        /// 无法直接读取；此时将采样字符应用该样式，Word 在文本层会把主题色展开为实际 RGB（与界面一致）。
        /// </summary>
        /// <summary>
        /// 获取Word字体颜色。
        /// Word 对主题色样式（如标题 1，styles.xml 中 w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"）
        /// 返回的 Font.Color 是负数编码（如 -738148353），ColorTranslator.FromOle 对负数会算出错误颜色。
        /// 正解：从 WordOpenXML 解析的 _styleXmlColorMap 按 NameLocal 查 w:val 直接色（与 Word 界面一致）；
        /// 映射缺失或非主题色时回退 interop Font.Color（仅正数有效）。
        /// </summary>
        private Color GetWordFontColor(Microsoft.Office.Interop.Word.Font font, Microsoft.Office.Interop.Word.Document doc, Style wordStyle)
        {
            // 优先从 WordOpenXML 解析的映射读（处理主题色编码）
            if (wordStyle != null && _styleXmlColorMap != null)
            {
                try
                {
                    string localName = wordStyle.NameLocal;
                    if (!string.IsNullOrEmpty(localName) && _styleXmlColorMap.TryGetValue(localName, out Color xmlColor))
                    {
                        return xmlColor;
                    }
                }
                catch { }
            }

            // 回退：interop Font.Color（仅正数有效；负数是主题色编码，FromOle 会算错，跳过返回黑色）
            try
            {
                int colorVal = (int)font.Color;
                if (font.Color != WdColor.wdColorAutomatic && colorVal > 0 && colorVal != 9999999)
                {
                    return ColorTranslator.FromOle(colorVal);
                }
            }
            catch { }
            return Color.Black;
        }

        /// <summary>
        /// 从文档的 WordOpenXML 解析所有样式的颜色，建立 样式名(NameLocal) → 颜色 映射。
        /// 优先用 w:color w:val 直接色（与 Word 界面一致），解决 interop Font.Color 对主题色返回负数编码的问题。
        /// </summary>
        private void BuildStyleColorMap(Microsoft.Office.Interop.Word.Document doc)
        {
            _styleXmlColorMap = new Dictionary<string, Color>();
            if (doc == null) return;

            try
            {
                string xml = doc.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return;

                var docXml = new XmlDocument();
                docXml.LoadXml(xml);
                var ns = new XmlNamespaceManager(docXml.NameTable);
                ns.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/mar/package");
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                // 定位 /word/styles.xml part
                var stylesNode = docXml.SelectSingleNode("//pkg:part[@pkg:name='/word/styles.xml']/pkg:xmlData/w:styles", ns);
                if (stylesNode == null) return;

                foreach (XmlNode styleNode in stylesNode.SelectNodes("w:style", ns))
                {
                    var nameNode = styleNode.SelectSingleNode("w:name", ns);
                    var colorNode = styleNode.SelectSingleNode("w:rPr/w:color", ns);
                    if (nameNode == null || colorNode == null) continue;

                    string engName = nameNode.Attributes["w:val"]?.Value;
                    Color color = ResolveColorFromXmlNode(colorNode);
                    if (string.IsNullOrEmpty(engName)) continue;

                    // 内置样式：styles.xml 用英文名(heading 1)，需用枚举值反查本地化名 NameLocal
                    string localName = null;
                    if (BuiltinStyleEnglishToEnum.TryGetValue(engName, out int enumVal))
                    {
                        try
                        {
                            object idx = enumVal;
                            var st = doc.Styles.get_Item(ref idx);
                            localName = st?.NameLocal;
                        }
                        catch { }
                    }
                    // 自定义样式：w:name 通常就是显示名（与 NameLocal 一致）
                    if (string.IsNullOrEmpty(localName))
                    {
                        localName = engName;
                    }

                    if (!string.IsNullOrEmpty(localName) && !_styleXmlColorMap.ContainsKey(localName))
                    {
                        _styleXmlColorMap[localName] = color;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"解析样式颜色XML失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 从 w:color 节点解析颜色：优先 w:val（RRGGBB 直接色，与 Word 界面一致）；无 val 时返回黑色。
        /// </summary>
        private Color ResolveColorFromXmlNode(XmlNode colorNode)
        {
            try
            {
                var valAttr = colorNode.Attributes["w:val"];
                if (valAttr != null && !string.IsNullOrEmpty(valAttr.Value) &&
                    !string.Equals(valAttr.Value, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    // w:val 是 6 位十六进制 RRGGBB（big-endian）
                    int rgb = Convert.ToInt32(valAttr.Value, 16);
                    int r = (rgb >> 16) & 0xFF;
                    int g = (rgb >> 8) & 0xFF;
                    int b = rgb & 0xFF;
                    return Color.FromArgb(r, g, b);
                }
            }
            catch { }
            return Color.Black;
        }

        /// <summary>
        /// 获取样式基准名称（读取失败返回空字符串）
        /// </summary>
        private string GetStyleBaseName(Style wordStyle)
        {
            try
            {
                var baseStyle = wordStyle.get_BaseStyle() as Style;
                return baseStyle != null ? baseStyle.NameLocal : "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 获取后续段落样式名称（读取失败返回空字符串）
        /// </summary>
        private string GetStyleNextParagraphName(Style wordStyle)
        {
            try
            {
                var nextStyle = wordStyle.get_NextParagraphStyle() as Style;
                return nextStyle != null ? nextStyle.NameLocal : "";
            }
            catch
            {
                return "";
            }
        }

        #endregion

        private void label16_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 字体颜色选择事件（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private void Btn_FontColor_Click(object sender, EventArgs e)
        {
            if (Btn_FontColor == null)
                return;

            ColorDialog colorDialog = new ColorDialog
            {
                Color = Btn_FontColor.BackColor,
                AnyColor = true,
                SolidColorOnly = true,
                AllowFullOpen = true,
                FullOpen = true
            };

            if (colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                Btn_FontColor.BackColor = colorDialog.Color;

                // 实时更新样式数据（与其他控件保持一致）
                OnStyleChanged(sender, e);
            }
        }

        /// <summary>
        /// 段落间距下拉框文本变化事件处理
        /// </summary>
        private void Cmb_Spacing_TextChanged(object sender, EventArgs e)
        {
            if (_isLoadingStyle) return; // 加载样式时忽略

            if (sender is StandardComboBox comboBox)
            {
                string text = comboBox.Text.Trim();
                if (!string.IsNullOrEmpty(text) && !text.EndsWith("磅") && !text.EndsWith("行") && !text.EndsWith("厘米"))
                {
                    // 如果输入的是纯数字，自动添加"磅"单位
                    if (float.TryParse(text, out _))
                    {
                        comboBox.Text = text + " 磅";
                    }
                }
            }
            
            // 实时更新样式数据
            OnStyleChanged(sender, e);
        }
    }

    #region CustomStyle 类

    /// <summary>
    /// 自定义样式类
    /// </summary>
    public class CustomStyle
    {
        public string Name { get; set; }
        public string FontName { get; set; }
        public string EngFontName { get; set; }
        public float FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        [XmlIgnore]
        public Color FontColor { get; set; }

        /// <summary>
        /// 用于XML序列化的字体颜色属性
        /// </summary>
        [XmlElement("FontColor")]
        public string FontColorString
        {
            get { return FontColor.ToArgb().ToString(); }
            set 
            { 
                try
                {
                    if (!string.IsNullOrEmpty(value) && int.TryParse(value, out int argb))
                    {
                        FontColor = Color.FromArgb(argb);
                    }
                    else
                    {
                        FontColor = Color.Black; // 默认黑色
                    }
                }
                catch
                {
                    FontColor = Color.Black; // 解析失败时使用默认黑色
                }
            }
        }
        public int ParaAlignment { get; set; }
        public float LeftIndent { get; set; }
        public float RightIndent { get; set; }
        public float FirstLineIndent { get; set; }
        public int FirstLineIndentByChar { get; set; }
        public float LineSpacing { get; set; }
        public float BeforeSpacing { get; set; }
        public bool BeforeBreak { get; set; }
        public float AfterSpacing { get; set; }
        public int NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public bool UserDefined { get; set; }

        /// <summary>
        /// 样式基准（"(无样式)"表示无基准，其他为样式名称）
        /// </summary>
        public string BaseStyle { get; set; }

        /// <summary>
        /// 后续段落样式（样式名称，如"正文"、"无间隔"）
        /// </summary>
        public string NextParagraphStyle { get; set; }

        /// <summary>
        /// 无参数构造函数，用于XML序列化
        /// </summary>
        public CustomStyle()
        {
            // 设置默认值
            Name = "";
            FontName = "宋体";
            EngFontName = "宋体";
            FontSize = 10.5f;
            Bold = false;
            Italic = false;
            Underline = false;
            FontColor = Color.Black;
            ParaAlignment = 0; // 左对齐
            LeftIndent = 0f;
            RightIndent = 0f;
            FirstLineIndent = 0f;
            FirstLineIndentByChar = 0;
            LineSpacing = 1f;
            BeforeSpacing = 0f;
            BeforeBreak = false;
            AfterSpacing = 0f;
            NumberStyle = 0;
            NumberFormat = "";
            UserDefined = false;
            BaseStyle = "";
            NextParagraphStyle = "";
        }

        /// <summary>
        /// 带参数的构造函数
        /// </summary>
        public CustomStyle(string name, string fontName, string engFontName, float fontSize, bool bold, bool italic, bool underline, Color fontColor,
            int paraAlignment, float leftIndent, float rightIndent, float firstLineIndent, int firstLineIndentByChar, float lineSpacing,
            float beforeSpacing, bool beforeBreak, float afterSpacing, int numberStyle, string numberFormat, bool userDefined,
            string baseStyle = "", string nextParagraphStyle = "")
        {
            Name = name;
            FontName = fontName;
            EngFontName = engFontName;
            FontSize = fontSize;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            FontColor = fontColor;
            ParaAlignment = paraAlignment;
            LeftIndent = leftIndent;
            RightIndent = rightIndent;
            FirstLineIndent = firstLineIndent;
            FirstLineIndentByChar = firstLineIndentByChar;
            LineSpacing = lineSpacing;
            BeforeSpacing = beforeSpacing;
            BeforeBreak = beforeBreak;
            AfterSpacing = afterSpacing;
            NumberStyle = numberStyle;
            NumberFormat = numberFormat;
            UserDefined = userDefined;
            BaseStyle = baseStyle;
            NextParagraphStyle = nextParagraphStyle;
        }
    }

    /// <summary>
    /// 样式导出数据类
    /// </summary>
    public class StyleExportData
    {
        public List<CustomStyle> Styles { get; set; }
        public int TitleCount { get; set; }
        public DateTime ExportTime { get; set; }

        public StyleExportData()
        {
            Styles = new List<CustomStyle>();
            TitleCount = 4; // 默认显示4级标题
            ExportTime = DateTime.Now;
        }
    }
}
    #endregion
