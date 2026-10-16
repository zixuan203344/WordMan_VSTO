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
        private Word.Document _sourceDocument;
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
            if (comboBox == null) return;
            // Use the actual items, including custom values and mixed-value placeholders.
            int index = comboBox.Items.IndexOf(value ?? "");
            if (index < 0 && comboBox.DropDownStyle == ComboBoxStyle.DropDownList && !string.IsNullOrEmpty(value))
                index = comboBox.Items.Add(value);
            comboBox.SelectedIndex = index;
            if (index < 0) comboBox.Text = value ?? "";
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
                ReplaceStyles(StylePresetManager.GetPresetStyles(presetName), StylePresetManager.GetPresetTitleCount(presetName));
                ShowStatusMessage($"已加载预设：{presetName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"加载预设样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReplaceStyles(List<CustomStyle> styles, int titleCount)
        {
            ValidateStyles(styles);
            bool wasLoading = _isLoadingStyle;
            _isLoadingStyle = true;
            try
            {
                _currentStyleIndex = -1;
                _multiSelectMode = false;
                _modifiedControls.Clear();
                RemoveAllMultiplePlaceholders();
                Styles.Clear();
                Styles.AddRange(styles);
                foreach (var style in Styles)
                    if (style.FirstLineIndentByChar != 0)
                        style.FirstLineIndent = style.FirstLineIndentByChar * style.FontSize * 2.54f / 72f;
                _styleNameMap.Clear();
                foreach (var style in Styles) _styleNameMap.Add(style.Name, style);
                RefreshBaseStyleComboBoxes();
                显示标题数下拉框.SelectedIndex = Math.Max(0, Math.Min(9, titleCount));
                FilterStylesByTitleCount();
            }
            finally { _isLoadingStyle = wasLoading; }
            if (!wasLoading) Lst_Styles_SelectedIndexChanged(Lst_Styles, EventArgs.Empty);
        }

        private void ValidateStyles(List<CustomStyle> styles)
        {
            if (styles == null || styles.Count == 0) throw new ArgumentException("样式文件中没有有效样式");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var style in styles)
            {
                if (style == null || string.IsNullOrWhiteSpace(style.Name) || style.Name.Length > 253 ||
                    style.Name.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0 || !names.Add(style.Name))
                    throw new ArgumentException($"样式名称为空、重复或包含无效字符：{style?.Name}");
                float[] numbers = { style.FontSize, style.LineSpacing, style.LeftIndent, style.RightIndent,
                    style.FirstLineIndent, style.BeforeSpacing, style.AfterSpacing };
                if (numbers.Any(v => float.IsNaN(v) || float.IsInfinity(v)) ||
                    style.FontSize < 1 || style.FontSize > 1638 || style.LineSpacing <= 0 || style.LineSpacing > 1584 ||
                    Math.Abs(style.LeftIndent) > 55.87f || Math.Abs(style.RightIndent) > 55.87f ||
                    Math.Abs(style.FirstLineIndent) > 55.87f || Math.Abs((long)style.FirstLineIndentByChar) > 100 || style.BeforeSpacing < 0 || style.BeforeSpacing > 1584 ||
                    style.AfterSpacing < 0 || style.AfterSpacing > 1584 || style.ParaAlignment < 0 || style.ParaAlignment > 4 ||
                    style.LineSpacingRule < -1 || style.LineSpacingRule > 5)
                    throw new ArgumentException($"样式 '{style.Name}' 含有无效的格式数值");
            }
            foreach (var style in styles)
            {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { style.Name };
                var current = style;
                while (!string.IsNullOrEmpty(current.BaseStyle) && current.BaseStyle != "(无样式)")
                {
                    if (!seen.Add(current.BaseStyle)) throw new ArgumentException($"样式 '{style.Name}' 存在循环基准");
                    current = styles.FirstOrDefault(x => string.Equals(x.Name, current.BaseStyle, StringComparison.OrdinalIgnoreCase));
                    if (current == null) break;
                }
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
            string mode = 首行缩进方式下拉框.Text;
            bool show = mode == "悬挂缩进" || mode == "首行缩进";
            label7.Visible = show;
            Nud_FirstLineIndent.Visible = show;
            if (show) label7.Text = mode;
        }

        #endregion

        #region 构造函数

        public StyleSettings()
        {
            InitializeComponent();
            InitializeData();
            Nud_FirstLineIndent.Minimum = 0;
            BindEvents();
            groupBox3.SizeChanged += (sender, args) => LayoutBottomButtons();

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
        public void RefreshFromDocument(bool onlyIfDocumentChanged = false)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) throw new InvalidOperationException("当前没有打开的文档");
            if (onlyIfDocumentChanged && _sourceDocument == app.ActiveDocument) return;
            _styleColorCache.Clear();
            _styleXmlColorMap = null;
            ReadDocumentStyles();
            风格下拉框.SelectedIndex = -1;
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
            Cmb_FontSize.TextChanged += OnStyleChanged;
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
            if (Lst_Styles == null || _isLoadingStyle)
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
                        应用当前.Enabled = true;
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
                    应用当前.Enabled = false;
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
            if (_isLoadingStyle) return;
            var control = sender as Control;
            if (control == null) return;
            var combo = control as StandardComboBox;
            if (combo != null)
            {
                if (combo.Text == "多种") return;
                _isLoadingStyle = true;
                try { RemoveMultiplePlaceholder(combo); }
                finally { _isLoadingStyle = false; }
            }
            float value;
            int rule;
            if (control == Cmb_FontSize && !TryParseFontSize(Cmb_FontSize.Text, out value)) return;
            if (control == Cmb_LineSpacing && !TryParseLineSpacing(Cmb_LineSpacing.Text, out value, out rule)) return;
            if ((control == Cmb_BefreSpacing || control == Cmb_AfterSpacing) &&
                !TryParseSpacing(control.Text, out value)) return;
            var targets = _multiSelectMode ? GetSelectedStyles() :
                (_currentStyleIndex >= 0 && _currentStyleIndex < Styles.Count
                    ? new List<CustomStyle> { Styles[_currentStyleIndex] } : new List<CustomStyle>());
            // Keep batch changes in the model when exporting or changing selection.
            if (_multiSelectMode) _modifiedControls.Add(control.Name);
            foreach (var style in targets)
            {
                ReadControlValueToStyle(control.Name, style);
                UpdateStyleInfo(style);
            }
            if (control == Cmb_FontSize && !_multiSelectMode && targets.Count == 1 && targets[0].FirstLineIndentByChar != 0)
            {
                _isLoadingStyle = true;
                try { Nud_FirstLineIndent.SetValueInCentimeters((decimal)Math.Abs(targets[0].FirstLineIndent)); }
                finally { _isLoadingStyle = false; }
            }
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
            var selected = Lst_Styles.SelectedItems.Cast<object>().Select(x => x.ToString()).ToList();
            var pendingProperties = _modifiedControls.ToList();
            bool wasMultiSelect = _multiSelectMode;
            int titleCount = Math.Max(0, 显示标题数下拉框.SelectedIndex);
            bool wasLoading = _isLoadingStyle;
            _isLoadingStyle = true;
            Lst_Styles.BeginUpdate();
            try
            {
                StyleNames.Clear();
                foreach (var style in Styles)
                {
                    int level = GetTitleLevel(style.Name);
                    if (level == 0 || level <= titleCount) StyleNames.Add(style.Name);
                }
                Lst_Styles.ClearSelected();
                foreach (var name in selected)
                {
                    int index = StyleNames.IndexOf(name);
                    if (index >= 0) Lst_Styles.SetSelected(index, true);
                }
                if (Lst_Styles.SelectedIndices.Count == 0 && StyleNames.Count > 0) Lst_Styles.SelectedIndex = 0;
            }
            finally
            {
                Lst_Styles.EndUpdate();
                _isLoadingStyle = wasLoading;
            }
            if (!wasLoading)
            {
                Lst_Styles_SelectedIndexChanged(Lst_Styles, EventArgs.Empty);
                var currentSelection = Lst_Styles.SelectedItems.Cast<object>().Select(x => x.ToString());
                if (wasMultiSelect && _multiSelectMode && new HashSet<string>(selected).SetEquals(currentSelection))
                    _modifiedControls.UnionWith(pendingProperties);
            }
        }

        /// <summary>
        /// 从样式名称中提取标题级别
        /// </summary>
        private int GetTitleLevel(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return 0;
            string suffix = styleName.StartsWith("标题", StringComparison.Ordinal) ? styleName.Substring(2) :
                styleName.StartsWith("Heading ", StringComparison.OrdinalIgnoreCase) ? styleName.Substring(8) : "";
            int level;
            return int.TryParse(suffix.Trim(), out level) && level >= 1 && level <= 9 ? level : 0;
        }


        private void 添加_Click(object sender, EventArgs e)
        {
            if (!ValidateEditorInputs()) return;
            string styleName = Txt_AddStyleName.Text.Trim();

            // 检查是否为空或提示文本
            if (string.IsNullOrEmpty(styleName) || styleName == "请输入需要增加的样式名称")
            {
                MessageBox.Show("请输入样式名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Styles.Any(s => string.Equals(s.Name, styleName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("样式名称已存在", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 创建新样式
            var newStyle = CreateStyleFromControls(styleName);
            if (newStyle != null)
            {
                try { ValidateStyles(Styles.Concat(new[] { newStyle }).ToList()); }
                catch (ArgumentException ex) { ShowStatusMessage(ex.Message); return; }
                Styles.Add(newStyle);
                StyleNames.Add(styleName);
                _styleNameMap[styleName] = newStyle;
                
                // 更新当前样式索引
                _currentStyleIndex = Styles.Count - 1;
                
                Lst_Styles.ClearSelected();
                RefreshBaseStyleComboBoxes();
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

                var remaining = Styles.Where(x => x != style).ToList();
                foreach (var item in remaining)
                {
                    if (item.BaseStyle == selectedStyle) item.BaseStyle = "(无样式)";
                    if (item.NextParagraphStyle == selectedStyle) item.NextParagraphStyle = item.Name;
                }
                if (remaining.Count > 0) ReplaceStyles(remaining, 显示标题数下拉框.SelectedIndex);
                else
                {
                    Styles.Clear();
                    _styleNameMap.Clear();
                    StyleNames.Clear();
                    ResetControlsToDefault();
                    RefreshBaseStyleComboBoxes();
                }

                MessageBox.Show("样式删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Btn_ApplySet_Click(object sender, EventArgs e)
        {
            if (!ValidateEditorInputs()) return;
            try
            {
                var targets = _multiSelectMode ? GetSelectedStyles() :
                    Styles.Where(x => StyleNames.Contains(x.Name)).ToList();
                if (targets.Count == 0) throw new InvalidOperationException("请先加载或选择样式");
                ValidateStyles(Styles);
                GetTargetDocument();
                // Batch editing changes only the properties explicitly edited by the user.
                var properties = _multiSelectMode ? _modifiedControls.ToList() : null;
                if (properties != null && properties.Count == 0)
                {
                    ShowStatusMessage("请先修改要批量设置的属性");
                    return;
                }
                Globals.ThisAddIn.ExecuteWithUndoRecord("应用样式设置", () =>
                    ApplyStylesToDocument(targets.Select(x => x.Name), properties));
                _modifiedControls.Clear();
                if (_multiSelectMode) UpdateControlsForMultiSelect();
                ShowStatusMessage($"已应用 {targets.Count} 个样式");
            }
            catch (Exception ex) { ShowStatusMessage($"应用样式时出错：{ex.Message}"); }
        }

        /// <summary>
        /// 应用当前样式 - 只应用当前选中的样式，其他样式不修改，应用后立即刷新
        /// </summary>
        private void 应用当前_Click(object sender, EventArgs e)
        {
            if (!ValidateEditorInputs()) return;
            try
            {
                if (_multiSelectMode || _currentStyleIndex < 0 || _currentStyleIndex >= Styles.Count)
                    throw new InvalidOperationException("请先选择一个要应用的样式");
                ValidateStyles(Styles);
                var style = Styles[_currentStyleIndex];
                GetTargetDocument();
                Globals.ThisAddIn.ExecuteWithUndoRecord($"应用样式：{style.Name}", () =>
                    ApplyStylesToDocument(new[] { style.Name }));
                ShowStatusMessage($"样式 '{style.Name}' 已应用到文档");
            }
            catch (Exception ex) { ShowStatusMessage($"应用当前样式时出错：{ex.Message}"); }
        }

        /// <summary>
        /// 保存当前样式为预设 - 输入名称后持久化保存，并自动刷新预设风格下拉列表
        /// </summary>
        private void 存预设_Click(object sender, EventArgs e)
        {
            if (!ValidateEditorInputs()) return;
            try
            {
                string name = "";
                using (var dialog = new SavePresetDialog(StylePresetManager.GetPresetStyleNames(),
                    风格下拉框.SelectedItem as string))
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

                if (stylesToSave.Count == 0) throw new InvalidOperationException("没有可保存的样式");
                ValidateStyles(stylesToSave);
                StylePresetManager.SavePreset(new StylePresetData
                {
                    Name = name,
                    Styles = stylesToSave,
                    TitleCount = 显示标题数下拉框.SelectedIndex,
                    CreateTime = DateTime.Now
                });

                // 刷新预设风格下拉列表，新预设立即可见
                风格下拉框.SelectedIndexChanged -= 风格下拉框_SelectedIndexChanged;
                try
                {
                    InitializeComboBox(风格下拉框, StylePresetManager.GetPresetStyleNames());
                    风格下拉框.SelectedItem = name;
                }
                finally { 风格下拉框.SelectedIndexChanged += 风格下拉框_SelectedIndexChanged; }

                ShowStatusMessage($"预设 '{name}' 已保存，共 {stylesToSave.Count} 个样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"保存预设时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_ReadDocumentStyle_Click(object sender, EventArgs e)
        {
            try { RefreshFromDocument(); }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"读取文档样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 关闭_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void 加载_Click(object sender, EventArgs e)
        {
            try
            {
                var doc = GetTargetDocument();
                var names = GetDocumentStyles();
                if (names.Count == 0) throw new InvalidOperationException("文档中没有可加载的段落样式");
                using (var dialog = new StyleSelectionForm())
                {
                    dialog.InitializeStyles(names);
                    dialog.SetSelectedStyles(StyleNames.ToList());
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    if (dialog.SelectedStyles.Count == 0)
                    {
                        ShowStatusMessage("请至少选择一个样式");
                        return;
                    }
                    var loaded = new List<CustomStyle>();
                    foreach (string name in dialog.SelectedStyles)
                    {
                        var style = Styles.FirstOrDefault(x => x.Name == name);
                        if (style == null)
                            style = CreateCustomStyleFromWordStyle(doc.Styles[name], Globals.ThisAddIn.Application, doc);
                        if (style == null) throw new InvalidOperationException($"无法读取样式：{name}");
                        loaded.Add(style);
                    }
                    int titleCount = Math.Max(显示标题数下拉框.SelectedIndex, loaded.Select(x => GetTitleLevel(x.Name)).Max());
                    ReplaceStyles(loaded, titleCount);
                    风格下拉框.SelectedIndex = -1;
                    ShowStatusMessage($"已加载 {loaded.Count} 个样式");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"加载样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出样式设置
        /// </summary>
        private void Btn_Export_Click(object sender, EventArgs e)
        {
            if (!ValidateEditorInputs()) return;
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

                    ValidateStyles(stylesToExport);
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
                string path = StyleFileManager.ShowOpenFileDialog();
                if (string.IsNullOrEmpty(path)) return;
                List<CustomStyle> imported;
                int titleCount;
                // Determine the format first. Validation errors must not fall back to another
                // format after partially replacing the current model.
                var xml = new XmlDocument { XmlResolver = null };
                xml.Load(path);
                if (xml.DocumentElement != null && xml.DocumentElement.LocalName == "StyleExportData")
                {
                    var data = StyleFileManager.DeserializeFromXml<StyleExportData>(path);
                    imported = data.Styles;
                    titleCount = data.TitleCount;
                }
                else
                {
                    imported = StyleFileManager.DeserializeListFromXml<CustomStyle>(path);
                    titleCount = imported.Where(x => x != null).Select(x => GetTitleLevel(x.Name)).DefaultIfEmpty(0).Max();
                }
                if (titleCount < 0 || titleCount > 9) throw new ArgumentException("显示标题数必须为 0 到 9");
                ReplaceStyles(imported, titleCount);
                风格下拉框.SelectedIndex = -1;
                ShowStatusMessage($"成功导入 {imported.Count} 个样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"导入样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            float value;
            int rule;
            if (!_isLoadingStyle && TryParseLineSpacing(Cmb_LineSpacing.Text, out value, out rule))
                Cmb_LineSpacing.Text = ConvertLineSpacingToString(value, rule);
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
            if (groupBox3 == null || buttons.Any(b => b == null)) return;
            int totalWidth = 0;
            foreach (var button in buttons)
            {
                button.AutoSize = false;
                button.Width = TextRenderer.MeasureText(button.Text, button.Font).Width + 14;
                button.Height = 31;
                button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                totalWidth += button.Width;
            }
            int requiredWidth = totalWidth + 40 + (buttons.Length - 1) * 6;
            int borderWidth = Width - ClientSize.Width;
            MinimumSize = new Size(Math.Max(720, requiredWidth + borderWidth + 12), Math.Max(MinimumSize.Height, 460));
            int gap = Math.Max(6, (groupBox3.ClientSize.Width - 40 - totalWidth) / (buttons.Length - 1));
            int x = 20;
            int y = groupBox3.ClientSize.Height - 43;
            foreach (var button in buttons)
            {
                button.Location = new Point(x, y);
                x += button.Width + gap;
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
            bool wasLoading = _isLoadingStyle;
            _isLoadingStyle = true;
            try
            {
            var selectedStyles = GetSelectedStyles();
            if (selectedStyles.Count == 0) return;

            // 中文字体
            ShowMultiValueOrCommon(Cmb_ChnFontName, selectedStyles.Select(s => s.FontName ?? "").ToList(), MultiLevelDataManager.GetChnFontItems());
            // 西文字体
            ShowMultiValueOrCommon(Cmb_EngFontName, selectedStyles.Select(s => !string.IsNullOrEmpty(s.EngFontName) ? s.EngFontName : (s.FontName ?? "")).ToList(), MultiLevelDataManager.GetEngFontItems());
            // 字号
            ShowMultiValueOrCommon(Cmb_FontSize, selectedStyles.Select(s => FormatFontSize(s.FontSize)).ToList(), MultiLevelDataManager.GetFontSizes());

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
            ShowMultiValueOrCommon(Nud_LeftIndent, selectedStyles.Select(s => $"{s.LeftIndent:R} 厘米").ToList());
            ShowMultiValueOrCommon(Nud_RightIndent, selectedStyles.Select(s => $"{s.RightIndent:R} 厘米").ToList());

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
                    Nud_FirstLineIndent.SetValueInCentimeters((decimal)Math.Abs(v));
                }
                else
                {
                    首行缩进方式下拉框.SelectedIndex = 2; // 首行缩进
                    Nud_FirstLineIndent.SetValueInCentimeters((decimal)v);
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
            ShowMultiValueOrCommon(Cmb_LineSpacing, selectedStyles.Select(s => ConvertLineSpacingToString(s.LineSpacing, s.LineSpacingRule)).ToList(), WordStyleInfo.LineSpacings);
            // 段前/段后
            ShowMultiValueOrCommon(Cmb_BefreSpacing, selectedStyles.Select(s => s.BeforeSpacing.ToString("0.0 磅")).ToList(), WordStyleInfo.SpaceBeforeValues);
            ShowMultiValueOrCommon(Cmb_AfterSpacing, selectedStyles.Select(s => s.AfterSpacing.ToString("0.0 磅")).ToList(), WordStyleInfo.SpaceAfterValues);

            // 样式基准 / 后续段落样式
            ShowMultiValueOrCommon(样式基准下拉框, selectedStyles.Select(s => !string.IsNullOrEmpty(s.BaseStyle) ? s.BaseStyle : "(无样式)").ToList(), GetBaseStyleItems());
            ShowMultiValueOrCommon(后续段落样式下拉框, selectedStyles.Select(s => !string.IsNullOrEmpty(s.NextParagraphStyle) ? s.NextParagraphStyle : "正文").ToList(), GetNextParagraphStyleItems());
            }
            finally { _isLoadingStyle = wasLoading; }
        }

        /// <summary>
        /// 下拉框/数值框显示共同值或"多种"：值全部相同则显示该值，否则显示"多种"
        /// </summary>
        private void ShowMultiValueOrCommon(StandardComboBox combo, List<string> values, IEnumerable<string> items)
        {
            bool allSame = values.Count > 0 && values.Distinct().Count() == 1;
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
                    combo.SelectedIndex = -1;
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
                        float fontSize;
                        if (TryParseFontSize(Cmb_FontSize.Text, out fontSize))
                        {
                            style.FontSize = fontSize;
                            if (style.FirstLineIndentByChar != 0)
                                style.FirstLineIndent = style.FirstLineIndentByChar * fontSize * 2.54f / 72f;
                        }
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
                    int alignment = Array.IndexOf(WordStyleInfo.HAlignments, Cmb_ParaAligment.Text);
                    if (alignment >= 0) style.ParaAlignment = alignment;
                    break;
                case "Nud_LeftIndent":
                    style.LeftIndent = (float)Nud_LeftIndent.GetValueInCentimeters();
                    break;
                case "Nud_RightIndent":
                    style.RightIndent = (float)Nud_RightIndent.GetValueInCentimeters();
                    break;
                case "首行缩进方式下拉框":
                    if (首行缩进方式下拉框.Text == "无")
                    {
                        style.FirstLineIndentByChar = 0;
                        style.FirstLineIndent = 0f;
                    }
                    else if (首行缩进方式下拉框.Text == "悬挂缩进")
                    {
                        style.FirstLineIndentByChar = 0;
                        style.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    else if (首行缩进方式下拉框.Text == "首行缩进")
                    {
                        style.FirstLineIndentByChar = 0;
                        style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    break;
                case "Nud_FirstLineIndent":
                    if (首行缩进方式下拉框.Text == "悬挂缩进")
                    {
                        style.FirstLineIndentByChar = 0;
                        style.FirstLineIndent = -(float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    else if (首行缩进方式下拉框.Text == "首行缩进")
                    {
                        style.FirstLineIndentByChar = 0;
                        style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
                    }
                    break;
                case "Cmb_LineSpacing":
                    if (!string.IsNullOrEmpty(Cmb_LineSpacing.Text))
                    {
                        float spacing;
                        int rule;
                        if (TryParseLineSpacing(Cmb_LineSpacing.Text, out spacing, out rule))
                        {
                            style.LineSpacing = spacing;
                            style.LineSpacingRule = rule;
                        }
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
            if (样式基准下拉框 == null || 后续段落样式下拉框 == null) return;
            bool wasLoading = _isLoadingStyle;
            _isLoadingStyle = true;
            try
            {
                string baseSelection = 样式基准下拉框.Text;
                string nextSelection = 后续段落样式下拉框.Text;
                InitializeComboBox(样式基准下拉框, GetBaseStyleItems());
                SetComboBoxSelection(样式基准下拉框, baseSelection, null);
                InitializeComboBox(后续段落样式下拉框, GetNextParagraphStyleItems());
                SetComboBoxSelection(后续段落样式下拉框, nextSelection, null);
            }
            finally { _isLoadingStyle = wasLoading; }
        }

        /// <summary>
        /// 重置控件到默认状态
        /// </summary>
        private void ResetControlsToDefault()
        {
            // 设置加载标志，防止触发保存事件
            bool wasLoading = _isLoadingStyle;
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
                _isLoadingStyle = wasLoading;
            }
        }

        private void LoadStyleToControls(CustomStyle style)
        {
            if (style == null)
                return;

            // 设置加载标志，防止触发保存事件
            bool wasLoading = _isLoadingStyle;
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
                SetComboBoxSelection(Cmb_FontSize, FormatFontSize(fontSize), MultiLevelDataManager.GetFontSizes());

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
                SetComboBoxSelection(Cmb_LineSpacing, ConvertLineSpacingToString(lineSpacing, style.LineSpacingRule), WordStyleInfo.LineSpacings);

                // 设置缩进控件（使用厘米单位显示）
                SetIndentValue(Nud_LeftIndent, $"{style.LeftIndent:R} 厘米");
                SetIndentValue(Nud_RightIndent, $"{style.RightIndent:R} 厘米");

                // 根据首行缩进值设置首行缩进方式和控件
                if (style.FirstLineIndent == 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 0; // 无
                    Nud_FirstLineIndent.Value = 0;
                    UpdateFirstLineIndentVisibility();
                }
                else if (style.FirstLineIndent < 0f)
                {
                    首行缩进方式下拉框.SelectedIndex = 1; // 悬挂缩进
                    UpdateFirstLineIndentVisibility();
                    SetIndentValue(Nud_FirstLineIndent, $"{Math.Abs(style.FirstLineIndent):R} 厘米");
                }
                else
                {
                    首行缩进方式下拉框.SelectedIndex = 2; // 首行缩进
                    UpdateFirstLineIndentVisibility();
                    SetIndentValue(Nud_FirstLineIndent, $"{style.FirstLineIndent:R} 厘米");
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
                _isLoadingStyle = wasLoading;
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
            foreach (string name in new[] { "Cmb_ChnFontName", "Cmb_EngFontName", "Cmb_FontSize", "Btn_Bold", "Btn_Italic",
                "Btn_UnderLine", "Btn_FontColor", "Cmb_ParaAligment", "Nud_LeftIndent", "Nud_RightIndent",
                "首行缩进方式下拉框", "Cmb_LineSpacing", "Cmb_BefreSpacing", "Cmb_AfterSpacing", "样式基准下拉框", "后续段落样式下拉框" })
                ReadControlValueToStyle(name, style);
        }

        private CustomStyle CreateStyleFromControls(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            var style = new CustomStyle { Name = name, UserDefined = true };
            if (!_multiSelectMode && Lst_Styles.SelectedIndices.Count == 1)
                UpdateStyleFromControls(style);
            return style;
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
        private float ConvertLineSpacingToFloat(string text)
        {
            float value;
            int rule;
            if (!TryParseLineSpacing(text, out value, out rule))
                throw new ArgumentException("行距请输入正数，例如 1.25倍行距、18磅或最小值18磅");
            return value;
        }

        /// <summary>
        /// 将行距数值转换为文本（复用多级段落设置的逻辑）
        /// </summary>
        private string ConvertLineSpacingToString(float value, int rule = -1)
        {
            if (rule == (int)WdLineSpacing.wdLineSpaceExactly || (rule < 0 && value >= 10))
                return value.ToString("0.###") + " 磅";
            if (rule == (int)WdLineSpacing.wdLineSpaceAtLeast)
                return "最小值 " + value.ToString("0.###") + " 磅";
            if (value == 1f) return "单倍行距";
            if (value == 1.5f) return "1.5倍行距";
            if (value == 2f) return "双倍行距";
            return value.ToString("0.###") + "倍行距";
        }

        private bool TryParseLineSpacing(string text, out float value, out int rule)
        {
            value = 0;
            rule = (int)WdLineSpacing.wdLineSpaceMultiple;
            text = (text ?? "").Trim();
            if (LineSpacingMap.TryGetValue(text, out value)) return true;
            bool atLeast = text.StartsWith("最小值", StringComparison.Ordinal);
            bool exact = text.StartsWith("固定值", StringComparison.Ordinal) ||
                text.EndsWith("磅", StringComparison.Ordinal) || text.EndsWith("pt", StringComparison.OrdinalIgnoreCase);
            bool multiple = text.EndsWith("倍行距", StringComparison.Ordinal) || text.EndsWith("倍", StringComparison.Ordinal);
            string number = text.Replace("最小值", "").Replace("固定值", "").Replace("倍行距", "")
                .Replace("倍", "").Replace("磅", "").Replace("pt", "").Trim();
            if (!TryParseFiniteNumber(number, out value) || value <= 0 || value > 1584) return false;
            if (atLeast) rule = (int)WdLineSpacing.wdLineSpaceAtLeast;
            else if (exact || (!multiple && value >= 10)) rule = (int)WdLineSpacing.wdLineSpaceExactly;
            return true;
        }

        private static bool TryParseFiniteNumber(string text, out float value)
        {
            return float.TryParse(text, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CurrentCulture, out value) && !float.IsNaN(value) && !float.IsInfinity(value);
        }

        private bool TryParseFontSize(string text, out float value)
        {
            var sizes = new Dictionary<string, float>
            {
                {"初号",42}, {"小初",36}, {"一号",26}, {"小一",24}, {"二号",22}, {"小二",18},
                {"三号",16}, {"小三",15}, {"四号",14}, {"小四",12}, {"五号",10.5f}, {"小五",9},
                {"六号",7.5f}, {"小六",6.5f}, {"七号",5.5f}, {"八号",5}
            };
            text = (text ?? "").Trim();
            if (sizes.TryGetValue(text, out value)) return true;
            return TryParseFiniteNumber(text.Replace("磅", "").Replace("pt", "").Trim(), out value) && value >= 1 && value <= 1638;
        }

        private string FormatFontSize(float value)
        {
            string label = MultiLevelDataManager.ConvertFontSizeToString(value);
            float parsed;
            return TryParseFontSize(label, out parsed) && Math.Abs(parsed - value) < 0.001f
                ? label : value.ToString("0.###") + " 磅";
        }

        private bool TryParseSpacing(string text, out float value)
        {
            text = (text ?? "").Trim();
            string number = text.Replace("厘米", "").Replace("磅", "").Replace("行", "").Trim();
            if (!TryParseFiniteNumber(number, out value) || value < 0) return false;
            if (text.EndsWith("行")) value = (float)UnitConverter.UnitConvert(value, "行", "磅");
            else if (text.EndsWith("厘米")) value = (float)UnitConverter.UnitConvert(value, "厘米", "磅");
            return value <= 1584;
        }

        private bool ValidateEditorInputs()
        {
            if (Lst_Styles.SelectedIndices.Count == 0) return true;
            float value;
            int rule;
            var controls = new[] { Cmb_FontSize, Cmb_LineSpacing, Cmb_BefreSpacing, Cmb_AfterSpacing };
            foreach (var control in controls)
            {
                if (_multiSelectMode && control.Text == "多种") continue;
                bool valid = control == Cmb_FontSize ? TryParseFontSize(control.Text, out value) :
                    control == Cmb_LineSpacing ? TryParseLineSpacing(control.Text, out value, out rule) :
                    TryParseSpacing(control.Text, out value);
                if (!valid)
                {
                    ShowStatusMessage("请输入有效的字号、行距或非负段落间距");
                    control.Focus();
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 将段落间距文本转换为磅单位（复用多级段落设置的逻辑）
        /// </summary>
        private float ConvertSpacingToPoints(string text)
        {
            float value;
            if (!TryParseSpacing(text, out value)) throw new ArgumentException("段落间距请输入非负数值，可使用磅、行或厘米");
            return value;
        }


        private void ApplyStylesToDocument(IEnumerable<string> styleNames = null, ICollection<string> properties = null)
        {
            var doc = GetTargetDocument();
            var names = (styleNames ?? StyleNames.ToList()).ToList();
            var targets = names.Select(name => Styles.First(x => x.Name == name)).ToList();
            // Create all target names first so forward base/next-style references resolve.
            foreach (var style in targets)
            {
                try { var existing = doc.Styles[style.Name]; }
                catch (System.Runtime.InteropServices.COMException)
                { doc.Styles.Add(style.Name, WdStyleType.wdStyleTypeParagraph); }
            }
            var errors = new List<string>();
            foreach (var style in targets)
            {
                try { ApplyStyleToDocument(doc, style, properties); }
                catch (Exception ex) { errors.Add(ex.Message); }
            }
            if (errors.Count > 0)
                throw new InvalidOperationException($"有 {errors.Count} 个样式未完整应用：{string.Join("；", errors)}");
            Globals.ThisAddIn.Application.ScreenRefresh();
        }

        private void ApplyStyleToDocument(Document doc, CustomStyle style, ICollection<string> properties = null)
        {
            try
            {
                Style wordStyle = doc.Styles[style.Name];
                Func<string, bool> changed = name => properties == null || properties.Contains(name);
                // Set inheritance before explicit formatting to avoid inherited values
                // replacing the font/paragraph values the user just chose.
                if (changed("样式基准下拉框") && !string.IsNullOrEmpty(style.BaseStyle))
                {
                    object baseRef = style.BaseStyle == "(无样式)" ? (object)"" : doc.Styles[style.BaseStyle];
                    wordStyle.set_BaseStyle(ref baseRef);
                }
                if (changed("后续段落样式下拉框") && !string.IsNullOrEmpty(style.NextParagraphStyle))
                {
                    object nextRef = doc.Styles[style.NextParagraphStyle];
                    wordStyle.set_NextParagraphStyle(ref nextRef);
                }
                if (changed("Cmb_ChnFontName") && !string.IsNullOrWhiteSpace(style.FontName))
                    wordStyle.Font.NameFarEast = style.FontName;
                if (changed("Cmb_EngFontName") && !string.IsNullOrWhiteSpace(style.EngFontName))
                    wordStyle.Font.NameAscii = style.EngFontName;
                if (changed("Cmb_FontSize")) wordStyle.Font.Size = style.FontSize;
                if (changed("Btn_Bold")) wordStyle.Font.Bold = style.Bold ? -1 : 0;
                if (changed("Btn_Italic")) wordStyle.Font.Italic = style.Italic ? -1 : 0;
                if (changed("Btn_UnderLine")) wordStyle.Font.Underline = style.Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                if (changed("Btn_FontColor")) wordStyle.Font.Color = (WdColor)ColorTranslator.ToOle(style.FontColor);
                var paragraph = wordStyle.ParagraphFormat;
                var app = Globals.ThisAddIn.Application;
                if (changed("Cmb_ParaAligment")) paragraph.Alignment = (WdParagraphAlignment)style.ParaAlignment;
                if (changed("Nud_LeftIndent")) paragraph.LeftIndent = app.CentimetersToPoints(style.LeftIndent);
                if (changed("Nud_RightIndent")) paragraph.RightIndent = app.CentimetersToPoints(style.RightIndent);
                if (changed("Nud_FirstLineIndent") || changed("首行缩进方式下拉框"))
                {
                    if (style.FirstLineIndentByChar != 0)
                        paragraph.CharacterUnitFirstLineIndent = style.FirstLineIndentByChar;
                    else
                        paragraph.FirstLineIndent = app.CentimetersToPoints(style.FirstLineIndent);
                }
                if (changed("Cmb_LineSpacing"))
                {
                    int rule = style.LineSpacingRule;
                    if (rule < 0) rule = style.LineSpacing >= 10 ? (int)WdLineSpacing.wdLineSpaceExactly : (int)WdLineSpacing.wdLineSpaceMultiple;
                    paragraph.LineSpacingRule = (WdLineSpacing)rule;
                    if (rule == (int)WdLineSpacing.wdLineSpaceExactly || rule == (int)WdLineSpacing.wdLineSpaceAtLeast)
                        paragraph.LineSpacing = style.LineSpacing;
                    else if (rule == (int)WdLineSpacing.wdLineSpaceMultiple)
                        paragraph.LineSpacing = app.LinesToPoints(style.LineSpacing);
                }
                if (changed("Cmb_BefreSpacing")) { paragraph.SpaceBeforeAuto = 0; paragraph.SpaceBefore = style.BeforeSpacing; }
                if (changed("Cmb_AfterSpacing")) { paragraph.SpaceAfterAuto = 0; paragraph.SpaceAfter = style.AfterSpacing; }
                if (properties == null) paragraph.PageBreakBefore = style.BeforeBreak ? -1 : 0;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"应用样式 '{style.Name}' 失败：{ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取文档中的所有样式名称
        /// </summary>
        private List<string> GetDocumentStyles()
        {
            var result = new List<string>();
            var doc = GetTargetDocument();
            foreach (Style style in doc.Styles)
            {
                if (style.Type == WdStyleType.wdStyleTypeParagraph || style.Type == WdStyleType.wdStyleTypeLinked)
                    result.Add(style.NameLocal);
            }
            return result.Distinct().ToList();
        }

        private void ReadDocumentStyles()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) throw new InvalidOperationException("当前没有打开的文档");
            var doc = app.ActiveDocument;
            BuildStyleColorMap(doc);
            var loaded = new List<CustomStyle>();
            var priority = new[] { WdBuiltinStyle.wdStyleNormal, WdBuiltinStyle.wdStyleHeading1,
                WdBuiltinStyle.wdStyleHeading2, WdBuiltinStyle.wdStyleHeading3, WdBuiltinStyle.wdStyleHeading4,
                WdBuiltinStyle.wdStyleHeading5, WdBuiltinStyle.wdStyleHeading6, WdBuiltinStyle.wdStyleHeading7,
                WdBuiltinStyle.wdStyleHeading8, WdBuiltinStyle.wdStyleHeading9, WdBuiltinStyle.wdStyleCaption };
            foreach (var id in priority)
            {
                var style = CreateCustomStyleFromWordStyle(doc.Styles[id], app, doc);
                if (style != null) loaded.Add(style);
            }
            foreach (Style wordStyle in doc.Styles)
            {
                if ((wordStyle.Type == WdStyleType.wdStyleTypeParagraph || wordStyle.Type == WdStyleType.wdStyleTypeLinked) &&
                    (!wordStyle.BuiltIn || wordStyle.InUse) && !loaded.Any(x => x.Name == wordStyle.NameLocal))
                {
                    var style = CreateCustomStyleFromWordStyle(wordStyle, app, doc);
                    if (style != null) loaded.Add(style);
                }
            }
            ReplaceStyles(loaded, Math.Max(0, 显示标题数下拉框.SelectedIndex));
            _sourceDocument = doc;
        }

        private Document GetTargetDocument()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) throw new InvalidOperationException("当前没有打开的文档");
            var doc = app.ActiveDocument;
            if (_sourceDocument != null && _sourceDocument != doc)
                throw new InvalidOperationException("当前文档已切换，请先点击“读取文档样式”再操作");
            return doc;
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
                    bold: wordStyle.Font.Bold == -1,
                    italic: wordStyle.Font.Italic == -1,
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
                ) { LineSpacingRule = (int)wordStyle.ParagraphFormat.LineSpacingRule };
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
                    case WdLineSpacing.wdLineSpaceAtLeast:
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
                return baseStyle != null ? baseStyle.NameLocal : Convert.ToString(wordStyle.get_BaseStyle());
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
                return nextStyle != null ? nextStyle.NameLocal : Convert.ToString(wordStyle.get_NextParagraphStyle());
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

            using (var colorDialog = new ColorDialog
            {
                Color = Btn_FontColor.BackColor,
                AnyColor = true,
                SolidColorOnly = true,
                AllowFullOpen = true,
                FullOpen = true
            })
            {
                if (colorDialog.ShowDialog(this) == DialogResult.OK)
                {
                    Btn_FontColor.BackColor = colorDialog.Color;
                    OnStyleChanged(Btn_FontColor, e);
                }
            }
        }

        /// <summary>
        /// 段落间距下拉框文本变化事件处理
        /// </summary>
        private void Cmb_Spacing_TextChanged(object sender, EventArgs e)
        {
            // Formatting while typing moves the caret and prevents entering decimals.
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
        // -1 keeps older XML files compatible; otherwise a Word WdLineSpacing value.
        public int LineSpacingRule { get; set; }
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
            LineSpacingRule = -1;
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
            FirstLineIndent = firstLineIndentByChar != 0 ? firstLineIndentByChar * fontSize * 2.54f / 72f : firstLineIndent;
            FirstLineIndentByChar = firstLineIndentByChar;
            LineSpacing = lineSpacing;
            LineSpacingRule = -1;
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
