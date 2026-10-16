using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;

namespace WordMan.MultiLevel
{
    public partial class LevelStyleSettingsForm : Form
    {
        private readonly List<WordStyleInfo> LevelStyles = new List<WordStyleInfo>();
        private readonly List<string> FontNames = new List<string>();
        private bool userChange;
        private int maxLevel = 9;

        /// <summary>
        /// 获取当前样式设置
        /// </summary>
        public List<WordStyleInfo> GetLevelStyles()
        {
            return LevelStyles.Select(x => x.Clone()).ToList();
        }

        /// <summary>
        /// 加载现有的样式设置
        /// </summary>
        public void LoadExistingStyles(List<WordStyleInfo> existingStyles)
        {
            if (existingStyles == null) return;
            userChange = false;
            try
            {
                // 对副本编辑，取消或右上角关闭时不改动父窗口的待应用设置。
                for (int i = 0; i < LevelStyles.Count && i < existingStyles.Count; i++)
                    if (existingStyles[i] != null) LevelStyles[i] = existingStyles[i].Clone();
                Dta_StyleList.DataSource = null;
                Dta_StyleList.DataSource = LevelStyles;
                if (LevelStyles.Count > 0) SetValueByStyle(LevelStyles[0]);
            }
            finally { userChange = true; }
        }

        public LevelStyleSettingsForm(int maxLevel = 9)
        {
            if (maxLevel < 1 || maxLevel > 9) throw new ArgumentOutOfRangeException("maxLevel");
            this.maxLevel = maxLevel;
            InitializeComponent();
            ShowInTaskbar = false;
            AcceptButton = Btn_SetStyles;
            CancelButton = Btn_Cancel;
            Btn_Cancel.DialogResult = DialogResult.Cancel;
            Btn_Cancel.CausesValidation = false;
            Txt_LeftIndent.DecimalPlaces = 3;
            Txt_RightIndent.DecimalPlaces = 3;
            Txt_LeftIndent.Minimum = Txt_RightIndent.Minimum = -1000000m;
            Txt_LeftIndent.Maximum = Txt_RightIndent.Maximum = 1000000m;
            Cmb_Alignment.AllowInput = false;
            InitializeForm();
        }

        private void InitializeForm()
        {
            userChange = false;

            // 获取系统字体
            LoadSystemFonts();

            // 初始化级别样式（只有在没有现有样式时才初始化）
            if (LevelStyles.Count == 0)
            {
                InitializeLevelStyles();
            }

            // 初始化控件
            InitializeControls();

            // 初始化事件处理
            InitializeEventHandlers();

            if (LevelStyles.Count > 0) SetValueByStyle(LevelStyles[0]);
            userChange = true;
        }

        #region 通用方法

        /// <summary>
        /// 加载系统字体
        /// </summary>
        private void LoadSystemFonts()
        {
            using (var fonts = new InstalledFontCollection())
                FontNames.AddRange(fonts.Families.Select(family => family.Name));
        }

        /// <summary>
        /// 通用下拉框初始化方法
        /// </summary>
        private void InitializeComboBox(StandardComboBox comboBox, IEnumerable<string> items)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.ToArray());
        }

        /// <summary>
        /// 通用DataGridView更新方法
        /// </summary>
        private void UpdateDataGridView(string columnName, int rowIndex)
        {
            if (Dta_StyleList.Columns.Contains(columnName))
            {
                int columnIndex = Dta_StyleList.Columns[columnName].Index;
                Dta_StyleList.UpdateCellValue(columnIndex, rowIndex);

                // 对于颜色列，需要强制刷新显示
                if (columnName == "Col_FontColor")
                {
                    Dta_StyleList.InvalidateCell(columnIndex, rowIndex);
                }
            }
        }

        /// <summary>
        /// 通用选择索引设置方法
        /// </summary>
        private void SetComboBoxSelection(StandardComboBox comboBox, string value, IEnumerable<string> items)
        {
            int index = comboBox.Items.IndexOf(value ?? "");
            if (index < 0 && !string.IsNullOrEmpty(value) && comboBox.DropDownStyle == ComboBoxStyle.DropDownList)
                index = comboBox.Items.Add(value);
            comboBox.SelectedIndex = index;
            if (index < 0) comboBox.Text = value ?? "";
        }

        /// <summary>
        /// 通用文本验证和格式化
        /// </summary>
        private void ValidateAndFormatText(Control control, string unit)
        {
            // 只格式化已验证的值，不能将非法输入静默替换为默认值。
            string text = control.Text.Trim();
            if (control == Cmb_FontSize)
            {
                float size;
                if (!MultiLevelDataManager.TryParseFontSize(text, out size)) return;
                if (!text.Contains("号")) control.Text = size.ToString("0.##") + " 磅";
            }
            else if (control == Cmb_LineSpacing)
            {
                if (!WordStyleInfo.IsValidLineSpacing(text)) return;
                control.Text = text;
            }
            else
            {
                float value;
                string suffix = text.EndsWith("行", StringComparison.Ordinal) ? "行" : "磅";
                if (float.TryParse(text.Replace("行", "").Replace("磅", "").Trim(), out value) &&
                    !float.IsNaN(value) && !float.IsInfinity(value) && value >= 0 && value <= 1584)
                    control.Text = value.ToString("0.###") + " " + suffix;
            }
        }

        /// <summary>
        /// 检查是否为有效的行距格式
        /// </summary>
        private bool IsValidLineSpacing(string text)
        {
            return WordStyleInfo.IsValidLineSpacing(text);
        }

        /// <summary>
        /// 获取控件值
        /// </summary>
        private object GetControlValue(Control control)
        {
            if (control is StandardComboBox comboBox)
                return comboBox.Text;
            if (control is StandardTextBox textBox)
                return textBox.Text;
            if (control is StandardNumericUpDown numericUpDown)
            {
                // 使用磅保存物理缩进，避免字符换算与 Word 写入不一致。
                return numericUpDown.GetValueInUnit("磅").ToString("0.###") + " 磅";
            }
            if (control is ToggleButton toggleButton)
                return toggleButton.Pressed;
            if (control is Button button && button.Name == "Btn_FontColor")
                return button.BackColor;
            return null;
        }

        /// <summary>
        /// 设置缩进值到StandardNumericUpDown控件
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
                    // 默认按字符处理（左缩进右缩进）
                    numericUpDown.SetValueInUnit(value, "字符");
                }
            }
            else
            {
                numericUpDown.Value = 0;
            }
        }

        #endregion

        /// <summary>
        /// 设置所有控件字体为微软雅黑
        /// </summary>
        private void SetAllControlsFont()
        {
            // 设置窗体字体
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);

            // 设置GroupBox字体
            Grp_SetSelectedStyle.Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold);
            Grp_SetSelectedStyle.ForeColor = Color.FromArgb(51, 51, 51);

            // 设置DataGridView字体
            Dta_StyleList.Font = new Font("Microsoft YaHei", 9F);
            Dta_StyleList.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold);
            Dta_StyleList.DefaultCellStyle.Font = new Font("Microsoft YaHei", 9F);

            // 设置所有标签字体
            SetControlFont(label1, 9F, FontStyle.Bold);
            SetControlFont(label2, 9F, FontStyle.Bold);
            SetControlFont(label3, 9F, FontStyle.Bold);
            SetControlFont(label4, 9F, FontStyle.Bold);
            SetControlFont(label5, 9F, FontStyle.Bold);
            SetControlFont(label6, 9F, FontStyle.Bold);
            SetControlFont(label7, 9F, FontStyle.Bold);
            SetControlFont(label8, 9F, FontStyle.Bold);
            SetControlFont(label9, 9F, FontStyle.Bold);
            SetControlFont(label10, 9F, FontStyle.Bold);
            SetControlFont(label11, 9F, FontStyle.Bold);
            SetControlFont(label12, 9F, FontStyle.Bold);
            SetControlFont(label13, 9F, FontStyle.Bold);
            SetControlFont(label14, 9F, FontStyle.Bold);

            // 设置所有下拉框字体
            SetControlFont(Cmb_ChnFontName, 9F);
            SetControlFont(Cmb_EngFontName, 9F);
            SetControlFont(Cmb_FontSize, 9F);
            SetControlFont(Cmb_Alignment, 9F);
            SetControlFont(Cmb_LineSpacing, 9F);
            SetControlFont(Cmb_SpaceBefore, 9F);
            SetControlFont(Cmb_SpaceAfter, 9F);

            // 设置所有数值输入框字体
            SetControlFont(Txt_LeftIndent, 9F);
            SetControlFont(Txt_RightIndent, 9F);

            // 设置所有按钮字体
            SetControlFont(Btn_FontColor, 9F, FontStyle.Bold);
            SetControlFont(Btn_Bold, 8F, FontStyle.Bold);
            SetControlFont(Btn_Italic, 8F, FontStyle.Bold);
            SetControlFont(Btn_Underline, 8F, FontStyle.Bold);
            SetControlFont(Btn_BreakBefore, 8F, FontStyle.Bold);
            SetControlFont(Btn_SetStyles, 9F, FontStyle.Bold);
            SetControlFont(Btn_Cancel, 9F, FontStyle.Bold);
        }

        /// <summary>
        /// 设置控件字体
        /// </summary>
        private void SetControlFont(Control control, float size, FontStyle style = FontStyle.Regular)
        {
            if (control != null)
            {
                control.Font = new Font("Microsoft YaHei", size, style);
            }
        }

        private void InitializeLevelStyles()
        {
            try
            {
                var headingStyles = new[]
                {
                    WdBuiltinStyle.wdStyleHeading1, WdBuiltinStyle.wdStyleHeading2, WdBuiltinStyle.wdStyleHeading3,
                    WdBuiltinStyle.wdStyleHeading4, WdBuiltinStyle.wdStyleHeading5, WdBuiltinStyle.wdStyleHeading6,
                    WdBuiltinStyle.wdStyleHeading7, WdBuiltinStyle.wdStyleHeading8, WdBuiltinStyle.wdStyleHeading9
                };

                var styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;

                for (int i = 0; i < maxLevel && i < headingStyles.Length; i++)
                {
                    object index = headingStyles[i];
                    Style style = styles[ref index];

                    // 创建样式对象，从Word文档样式读取所有属性包括颜色
                    var wordStyleInfo = new WordStyleInfo(style, headingStyles[i]);

                    // 注释掉强制设置黑色，让颜色从文档样式自动读取
                    // wordStyleInfo.FontColor = Color.Black;

                    LevelStyles.Add(wordStyleInfo);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("初始化级别样式失败：" + ex.Message, ex);
            }
        }

        private void InitializeControls()
        {
            // 初始化DataGridView
            InitializeDataGridView();

            // 设置所有控件字体为微软雅黑
            SetAllControlsFont();

            // 使用通用方法初始化下拉框
            InitializeComboBox(Cmb_ChnFontName, FontNames);
            InitializeComboBox(Cmb_EngFontName, FontNames);
            InitializeComboBox(Cmb_FontSize, WordStyleInfo.FontSizes);
            InitializeComboBox(Cmb_Alignment, WordStyleInfo.HAlignments);
            InitializeComboBox(Cmb_LineSpacing, WordStyleInfo.LineSpacings);
            InitializeComboBox(Cmb_SpaceBefore, WordStyleInfo.SpaceBeforeValues);
            InitializeComboBox(Cmb_SpaceAfter, WordStyleInfo.SpaceAfterValues);
        }

        /// <summary>
        /// 初始化事件处理
        /// </summary>
        private void InitializeEventHandlers()
        {
            // 字体相关控件
            Cmb_ChnFontName.TextChanged += (s, e) => OnControlValueChanged(s, "ChnFontName", "Col_ChnFontName");
            Cmb_EngFontName.TextChanged += (s, e) => OnControlValueChanged(s, "EngFontName", "Col_EngFontName");
            Cmb_FontSize.TextChanged += (s, e) => OnControlValueChanged(s, "FontSize", "Col_FontSize");

            // 格式相关控件
            Btn_Bold.PressedChanged += (s, e) => OnControlValueChanged(s, "Bold", "Col_FontBold");
            Btn_Italic.PressedChanged += (s, e) => OnControlValueChanged(s, "Italic", "Col_FontItalic");
            Btn_Underline.PressedChanged += (s, e) => OnControlValueChanged(s, "Underline", "Col_FontUnderline");

            // 缩进相关控件
            Txt_LeftIndent.ValueChanged += (s, e) => OnControlValueChanged(s, "LeftIndent", "Col_LeftIndent");
            Txt_RightIndent.ValueChanged += (s, e) => OnControlValueChanged(s, "RightIndent", "Col_RightIndent");

            // 间距相关控件
            Cmb_LineSpacing.TextChanged += (s, e) => OnControlValueChanged(s, "LineSpace", "Col_LineSpace");
            Cmb_SpaceBefore.TextChanged += (s, e) => OnControlValueChanged(s, "SpaceBefore", "Col_SpaceBefore");
            Cmb_SpaceAfter.TextChanged += (s, e) => OnControlValueChanged(s, "SpaceAfter", "Col_SpaceAfter");

            // 对齐相关控件
            Cmb_Alignment.SelectedIndexChanged += (s, e) => OnControlValueChanged(s, "HAlignment", "Col_HAlignment");

            // 分页相关控件
            Btn_BreakBefore.PressedChanged += (s, e) => OnControlValueChanged(s, "BreakBefore", "Col_BreakBefore");

            // 颜色控件
            Btn_FontColor.BackColorChanged += (s, e) => OnControlValueChanged(s, "FontColor", "Col_FontColor");
        }

        /// <summary>
        /// 通用控件值变化事件处理
        /// </summary>
        private void OnControlValueChanged(object sender, string propertyName, string columnName)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0) return;
            var sourceControl = sender as Control;
            object value = GetControlValue(sourceControl);
            if (value == null) return;
            string text = value as string;
            float parsed;
            if (propertyName == "FontSize" && !MultiLevelDataManager.TryParseFontSize(text, out parsed)) return;
            if (propertyName == "LineSpace" && !WordStyleInfo.IsValidLineSpacing(text)) return;
            if (propertyName == "SpaceBefore" || propertyName == "SpaceAfter")
            {
                string suffix = (text ?? "").Trim().EndsWith("行", StringComparison.Ordinal) ? "行" : "磅";
                if (!float.TryParse((text ?? "").Replace("行", "").Replace("磅", "").Trim(), out parsed) ||
                    float.IsNaN(parsed) || float.IsInfinity(parsed) || parsed < 0 || parsed > 1584) return;
                value = parsed.ToString("0.###") + " " + suffix;
            }

            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows.Cast<DataGridViewRow>().ToArray())
            {
                var style = selectedRow.DataBoundItem as WordStyleInfo;
                if (style == null) continue;

                // 使用反射设置属性值
                var property = typeof(WordStyleInfo).GetProperty(propertyName);
                if (property != null)
                {
                    property.SetValue(style, value, null);
                    if (propertyName == "Underline")
                        style.UnderlineStyle = style.Underline ? (int)WdUnderline.wdUnderlineSingle : 0;

                    // 对于颜色属性，确保数据绑定正确更新
                    if (propertyName == "FontColor")
                    {
                        // 强制刷新整个行
                        Dta_StyleList.Refresh();
                    }
                    else
                    {
                        UpdateDataGridView(columnName, selectedRow.Index);
                    }
                }
            }
        }

        private void InitializeDataGridView()
        {
            Dta_StyleList.Columns.Clear();
            Dta_StyleList.ReadOnly = true;
            Dta_StyleList.AllowUserToAddRows = false;
            Dta_StyleList.AllowUserToDeleteRows = false;
            Dta_StyleList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Dta_StyleList.AutoGenerateColumns = false;

            // 设置默认单元格样式为居中
            var centerAlignment = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };

            // 添加列
            Dta_StyleList.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_StyleName",
                    DataPropertyName = "StyleName",
                    Frozen = true,
                    HeaderText = "样式名",
                    ReadOnly = true,
                    Width = 80,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_ChnFontName",
                    DataPropertyName = "ChnFontName",
                    HeaderText = "中文字体",
                    Width = 90,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_EngFontName",
                    DataPropertyName = "EngFontName",
                    HeaderText = "西文字体",
                    Width = 90,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_FontSize",
                    DataPropertyName = "FontSize",
                    HeaderText = "字体大小",
                    Width = 90,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_FontColor",
                    DataPropertyName = "FontColor",
                    HeaderText = "颜色",
                    Width = 50,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontBold",
                    DataPropertyName = "Bold",
                    HeaderText = "粗体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontItalic",
                    DataPropertyName = "Italic",
                    HeaderText = "斜体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontUnderline",
                    DataPropertyName = "Underline",
                    HeaderText = "下划线",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 60,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_LeftIndent",
                    DataPropertyName = "LeftIndent",
                    HeaderText = "左缩进",
                    Width = 80,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_RightIndent",
                    DataPropertyName = "RightIndent",
                    HeaderText = "右缩进",
                    Width = 80,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_LineSpace",
                    DataPropertyName = "LineSpace",
                    HeaderText = "行距",
                    Width = 80,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_SpaceBefore",
                    DataPropertyName = "SpaceBefore",
                    HeaderText = "段前间距",
                    Width = 100,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_SpaceAfter",
                    DataPropertyName = "SpaceAfter",
                    HeaderText = "段后间距",
                    Width = 100,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_HAlignment",
                    DataPropertyName = "HAlignment",
                    HeaderText = "水平对齐",
                    Width = 100,
                    DefaultCellStyle = centerAlignment
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_BreakBefore",
                    DataPropertyName = "BreakBefore",
                    HeaderText = "段前分页",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50,
                    DefaultCellStyle = centerAlignment
                }
            });

            Dta_StyleList.DataSource = LevelStyles;
            Dta_StyleList.CellFormatting += Dta_StyleList_CellFormatting;
            Dta_StyleList.SelectionChanged += Dta_StyleList_SelectionChanged;
        }

        private void Dta_StyleList_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex != Dta_StyleList.Columns["Col_FontColor"].Index || e.RowIndex < 0) return;
            var style = Dta_StyleList.Rows[e.RowIndex].DataBoundItem as WordStyleInfo;
            if (style == null) return;
            e.Value = "";
            e.CellStyle.BackColor = style.FontColor;
            e.CellStyle.SelectionBackColor = style.FontColor;
            e.FormattingApplied = true;
        }

        private void Dta_StyleList_SelectionChanged(object sender, EventArgs e)
        {
            if (userChange && Dta_StyleList.SelectedRows.Count > 0 &&
                Dta_StyleList.SelectedRows[0].DataBoundItem is WordStyleInfo selectedStyle)
                SetValueByStyle(selectedStyle);
        }

        private void SetValueByStyle(WordStyleInfo style)
        {
            bool previousUserChange = userChange;
            userChange = false;
            try
            {

                // 使用通用方法设置控件值
                SetComboBoxSelection(Cmb_ChnFontName, style.ChnFontName, FontNames);
                SetComboBoxSelection(Cmb_EngFontName, style.EngFontName, FontNames);
                SetComboBoxSelection(Cmb_FontSize, style.FontSize, WordStyleInfo.FontSizes);
                SetComboBoxSelection(Cmb_LineSpacing, style.LineSpace, WordStyleInfo.LineSpacings);
                SetComboBoxSelection(Cmb_SpaceBefore, style.SpaceBefore, WordStyleInfo.SpaceBeforeValues);
                SetComboBoxSelection(Cmb_SpaceAfter, style.SpaceAfter, WordStyleInfo.SpaceAfterValues);
                SetComboBoxSelection(Cmb_Alignment, style.HAlignment, WordStyleInfo.HAlignments);

                // 设置其他控件
                Btn_FontColor.BackColor = style.FontColor;
                Btn_Bold.Pressed = style.Bold;
                Btn_Italic.Pressed = style.Italic;
                Btn_Underline.Pressed = style.Underline;
                Btn_BreakBefore.Pressed = style.BreakBefore;

                // 设置缩进控件（从字符串解析数值）
                SetIndentValue(Txt_LeftIndent, style.LeftIndent);
                SetIndentValue(Txt_RightIndent, style.RightIndent);

            }
            finally { userChange = previousUserChange; }
        }

        private void ToggleButton_PressedChanged(object sender, EventArgs e)
        {
            if (sender is ToggleButton toggleButton)
            {
                toggleButton.Text = toggleButton.Pressed ? "是" : "否";
            }
        }

        private void Btn_FontColor_Click(object sender, EventArgs e)
        {
            using (var dialog = new ColorDialog
            {
                Color = Btn_FontColor.BackColor, AnyColor = true, SolidColorOnly = true, AllowFullOpen = true, FullOpen = true
            })
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    Btn_FontColor.BackColor = dialog.Color;
                    // 所选颜色与首行相同也要应用到其他选中行。
                    OnControlValueChanged(Btn_FontColor, "FontColor", "Col_FontColor");
                }
        }


        private void Cmb_FontSize_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_FontSize, "磅");
        }


        private void Cmb_LineSpace_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_LineSpacing, Cmb_LineSpacing.Text.EndsWith("行") ? "行" : "磅");
        }

        private void Cmb_SpaceValue_Validated(object sender, EventArgs e)
        {
            if (sender is StandardComboBox comboBox)
            {
                ValidateAndFormatText(comboBox, comboBox.Text.EndsWith("行") ? "行" : "磅");
            }
        }

        private void Btn_SetStyles_Click(object sender, EventArgs e)
        {
            try
            {
                float size;
                if (!MultiLevelDataManager.TryParseFontSize(Cmb_FontSize.Text, out size))
                    throw new ArgumentException("请输入 1 至 1638 磅之间的有效字号。");
                if (!WordStyleInfo.IsValidLineSpacing(Cmb_LineSpacing.Text)) throw new ArgumentException("请输入有效的行距，例如 1.25倍行距、18磅或最小值18磅。");
                ValidateChildren();
                foreach (var field in new[] { Cmb_SpaceBefore, Cmb_SpaceAfter })
                {
                    float number;
                    string text = field.Text.Replace("磅", "").Replace("行", "").Trim();
                    if (!float.TryParse(text, out number) || float.IsNaN(number) || float.IsInfinity(number) || number < 0 || number > 1584)
                        throw new ArgumentException("请输入有效的段前、段后间距。");
                }
                foreach (var style in LevelStyles) style.Validate();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "样式设置无效：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        #region 设计器事件处理方法（保持兼容性）



        #endregion

        private void label9_Click(object sender, EventArgs e)
        {
            // 段落行距标签点击事件 - 可以添加帮助信息或重置功能
        }

    }
}
