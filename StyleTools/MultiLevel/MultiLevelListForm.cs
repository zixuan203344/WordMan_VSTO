using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using WordMan;
using WordMan.MultiLevel;

namespace WordMan
{

    public partial class MultiLevelListForm : Form
    {
        #region 常量和配置
        private const int MAX_LEVELS = 9;
        
        // 魔法数字常量定义
        private const float INVALID_TAB_POSITION = 9999999f;
        private const int MAX_INDENT_CM = 50;
        
        // 使用MultiLevelDataManager中的统一常量定义
        private static readonly WdListNumberStyle[] LevelNumStyle = new WdListNumberStyle[]
        {
            WdListNumberStyle.wdListNumberStyleArabic,           // 0: 1,2,3...
            WdListNumberStyle.wdListNumberStyleLegalLZ,          // 1: 01,02,03...
            WdListNumberStyle.wdListNumberStyleUppercaseLetter,  // 2: A,B,C...
            WdListNumberStyle.wdListNumberStyleLowercaseLetter,  // 3: a,b,c...
            WdListNumberStyle.wdListNumberStyleUppercaseRoman,   // 4: I,II,III...
            WdListNumberStyle.wdListNumberStyleLowercaseRoman,   // 5: i,ii,iii...
            WdListNumberStyle.wdListNumberStyleSimpChinNum1,     // 6: 一,二,三...
            WdListNumberStyle.wdListNumberStyleSimpChinNum2,     // 7: 壹,贰,叁...
            WdListNumberStyle.wdListNumberStyleZodiac1,          // 8: 甲,乙,丙...
            WdListNumberStyle.wdListNumberStyleArabic            // 9: 正规编号
        };
        
        // 使用MultiLevelDataManager中的统一常量定义
        private static readonly string[] NumberStyleOptions = MultiLevelDataManager.ValidationConstants.ValidNumberStyles;
        private static readonly string[] AfterNumberOptions = MultiLevelDataManager.ValidationConstants.ValidAfterNumberTypes;
        private static readonly string[] LinkedStyleOptions = MultiLevelDataManager.ValidationConstants.ValidLinkedStyles;
        #endregion
        
        #region 私有字段
        private int currentLevels = 0;
        private List<LevelData> levelDataList = new List<LevelData>();
        private Microsoft.Office.Interop.Word.Application app;
        private List<WordStyleInfo> levelStyleSettings = new List<WordStyleInfo>();
        private Document sourceDocument;
        private bool updatingControls;
        #endregion
        
        /// <summary>
        /// 初始化多级列表表单
        /// </summary>
        public MultiLevelListForm()
        {
            InitializeComponent();
            
            // 添加空引用检查
            if (Globals.ThisAddIn?.Application == null)
            {
                throw new InvalidOperationException("Word应用程序不可用，无法初始化多级列表表单。");
            }
            
            app = Globals.ThisAddIn.Application;
            if (app.Documents.Count > 0) sourceDocument = app.ActiveDocument;
            ShowInTaskbar = false;
            // 左侧原有 1050 像素宽度与右侧 350 像素重叠，导致按钮及链接列被遮挡。
            leftPanel.Dock = DockStyle.Fill;
            leftPanel.BringToFront();
            ClientSize = new Size(1250, 480);
            levelsScrollPanel.Width = leftPanel.ClientSize.Width - 30;
            cmbLevelCount.Left = 125;
            cmbLevelCount.Width = 60;
            levelsContainer.MinimumSize = new Size(840, 0);
            foreach (var input in new[] { nudNumberIndent, nudTextIndent, nudTabPosition, nudFirstLevelIndent, nudIncrementIndent })
            {
                // 范围由换算后的厘米值校验；不能用显示单位的 Maximum=50 限制磅值。
                input.Maximum = 1000000m;
                input.Minimum = -1000000m;
                input.DecimalPlaces = 3;
            }
            
            // 初始化表单数据、事件处理和控件
            InitializeData();
            SetupEventHandlers();
            CreateLevelControls();
            InitializePresetControls();
        }


        /// <summary>
        /// 初始化表单数据 - 创建默认的级别数据配置
        /// </summary>
        private void InitializeData()
        {
            levelDataList.AddRange(MultiLevelPresetManager.CreateDefaultLevels());

            // 设置默认不显示任何级别
            cmbLevelCount.SetSelectedItem("0");
            currentLevels = 0;
        }

        /// <summary>
        /// 创建级别控件 - 动态生成多级列表配置界面
        /// </summary>
        private void CreateLevelControls()
        {
            updatingControls = true;
            levelsContainer.SuspendLayout();
            try
            {
                foreach (Control oldControl in levelsContainer.Controls.Cast<Control>().ToArray())
                    oldControl.Dispose();

                // 动态创建级别控件 - 按正确顺序（标题在上，1级在下）
                for (int i = currentLevels; i >= 1; i--)
                {
                    CreateLevelRow(i);
                }

                // 添加列标题 - 使用MultiLevelListControlFactory创建
                var headerPanel = new Panel();
                headerPanel.Height = 30;
                headerPanel.Dock = DockStyle.Top;
                headerPanel.BackColor = Color.Transparent;
            
                var lblLevel = MultiLevelListControlFactory.CreateColumnHeader("级别", new Point(10, 8), new Size(50, 20));
                var lblNumberStyle = MultiLevelListControlFactory.CreateColumnHeader("编号样式", new Point(70, 8), new Size(100, 20));
                var lblNumberFormat = MultiLevelListControlFactory.CreateColumnHeader("编号格式", new Point(180, 8), new Size(100, 20));
                var lblNumberIndent = MultiLevelListControlFactory.CreateColumnHeader("编号缩进", new Point(290, 8), new Size(100, 20));
                var lblTextIndent = MultiLevelListControlFactory.CreateColumnHeader("文本缩进", new Point(400, 8), new Size(100, 20));
                var lblAfterNumber = MultiLevelListControlFactory.CreateColumnHeader("编号之后", new Point(510, 8), new Size(100, 20));
                var lblTabPosition = MultiLevelListControlFactory.CreateColumnHeader("制表位位置", new Point(620, 8), new Size(100, 20));
                var lblLinkedStyle = MultiLevelListControlFactory.CreateColumnHeader("链接样式", new Point(730, 8), new Size(100, 20));
            
                headerPanel.Controls.AddRange(new Control[] { lblLevel, lblNumberStyle, lblNumberFormat, lblNumberIndent, lblTextIndent, lblAfterNumber, lblTabPosition, lblLinkedStyle });
                levelsContainer.Controls.Add(headerPanel);
            
                // 设置所有级别的制表位位置启用状态
                for (int i = 1; i <= currentLevels; i++)
                {
                    var controls = GetLevelControls(i);
                    if (controls.AfterNumber != null && controls.TabPosition != null)
                        controls.TabPosition.Enabled = (controls.AfterNumber.Text == "制表位");
                }
            }
            finally
            {
                levelsContainer.ResumeLayout(true);
                updatingControls = false;
            }
            btnSetMultiLevelList.Enabled = currentLevels > 0;
            btnSetLevelStyle.Enabled = currentLevels > 0;
            btnApplySettings.Enabled = currentLevels > 0;
        }


        /// <summary>
        /// 创建单个级别的控件行
        /// </summary>
        /// <param name="level">级别编号（1-9）</param>
        private void CreateLevelRow(int level)
        {
            var rowPanel = new Panel();
            rowPanel.Height = 35;
            rowPanel.Dock = DockStyle.Top;
            rowPanel.BackColor = Color.Transparent;
            rowPanel.BorderStyle = BorderStyle.None;

            // 使用MultiLevelListControlFactory创建控件
            var lblLevel = MultiLevelListControlFactory.CreateLevelLabel(level, new Point(10, 8));
            var cmbNumberStyle = MultiLevelListControlFactory.CreateNumberStyleCombo("CmbNumStyle" + level, new Point(70, 5), 0, null, false);
            var txtNumberFormat = MultiLevelListControlFactory.CreateNumberFormatTextBox("TextBoxNumFormat" + level, new Point(180, 5), "");
            var nudNumberIndent = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxNumIndent" + level, new Point(290, 5), new Size(100, 25));
            var nudTextIndent = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxTextIndent" + level, new Point(400, 5), new Size(100, 25));
            var cmbAfterNumber = MultiLevelListControlFactory.CreateAfterNumberCombo("CmbAfterNumber" + level, new Point(510, 5), 1, null, false);
            var nudTabPosition = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxTabPosition" + level, new Point(620, 5), new Size(100, 25));
            var cmbLinkedStyle = MultiLevelListControlFactory.CreateLinkedStyleCombo("CmbLinkedStyle" + level, new Point(730, 5), 0, null, false);

            // 从levelDataList中获取数据并设置
            var levelData = levelDataList[level - 1];
            
            // 设置控件值
            int styleIndex = Array.IndexOf(NumberStyleOptions, levelData.NumberStyle);
            cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
            txtNumberFormat.Text = levelData.NumberFormat;
            int afterIndex = Array.IndexOf(AfterNumberOptions, levelData.AfterNumberType);
            cmbAfterNumber.SelectedIndex = afterIndex >= 0 ? afterIndex : 1;
            int linkedIndex = Array.IndexOf(LinkedStyleOptions, levelData.LinkedStyle);
            if (linkedIndex < 0) linkedIndex = cmbLinkedStyle.Items.Add(levelData.LinkedStyle);
            cmbLinkedStyle.SelectedIndex = linkedIndex;
            foreach (var input in new[] { nudNumberIndent, nudTextIndent, nudTabPosition })
            {
                input.Minimum = -1000000m;
                input.Maximum = 1000000m;
                input.DecimalPlaces = 3;
            }

            // 设置数值（使用Word API转换）
            nudNumberIndent.SetValueInCentimeters(levelData.NumberIndent);
            nudTextIndent.SetValueInCentimeters(levelData.TextIndent);
            nudTabPosition.SetValueInCentimeters(levelData.TabPosition);
            nudTabPosition.Enabled = (levelData.AfterNumberType == "制表位");

            // 添加控件到行面板
            rowPanel.Controls.AddRange(new Control[] { 
                lblLevel, cmbNumberStyle, txtNumberFormat, 
                nudNumberIndent, nudTextIndent, cmbAfterNumber, nudTabPosition, cmbLinkedStyle 
            });

            // 添加事件处理
            cmbNumberStyle.SelectedIndexChanged += (s, e) => UpdateLevelData(level);
            txtNumberFormat.TextChanged += (s, e) => UpdateLevelData(level);
            nudNumberIndent.ValueChanged += (s, e) => UpdateLevelData(level);
            nudTextIndent.ValueChanged += (s, e) => UpdateLevelData(level);
            cmbAfterNumber.SelectedIndexChanged += (s, e) => {
                UpdateLevelData(level);
                var controls = GetLevelControls(level);
                if (controls.AfterNumber != null && controls.TabPosition != null)
                    controls.TabPosition.Enabled = (controls.AfterNumber.Text == "制表位");
            };
            nudTabPosition.ValueChanged += (s, e) => UpdateLevelData(level);
            cmbLinkedStyle.SelectedIndexChanged += (s, e) => UpdateLevelData(level);

            levelsContainer.Controls.Add(rowPanel);
        }

        #region 控件操作辅助方法
        private T FindControl<T>(string name, int level) where T : Control
        {
            return levelsContainer.Controls.Find(name + level, true).FirstOrDefault() as T;
        }
        
        private LevelControls GetLevelControls(int level)
        {
            return new LevelControls
            {
                NumberStyle = FindControl<StandardComboBox>("CmbNumStyle", level),
                NumberFormat = FindControl<StandardTextBox>("TextBoxNumFormat", level),
                AfterNumber = FindControl<StandardComboBox>("CmbAfterNumber", level),
                LinkedStyle = FindControl<StandardComboBox>("CmbLinkedStyle", level),
                NumberIndent = FindControl<StandardNumericUpDown>("TxtBoxNumIndent", level),
                TextIndent = FindControl<StandardNumericUpDown>("TxtBoxTextIndent", level),
                TabPosition = FindControl<StandardNumericUpDown>("TxtBoxTabPosition", level)
            };
        }
        
        
        private class LevelControls
        {
            public StandardComboBox NumberStyle { get; set; }
            public StandardTextBox NumberFormat { get; set; }
            public StandardComboBox AfterNumber { get; set; }
            public StandardComboBox LinkedStyle { get; set; }
            public StandardNumericUpDown NumberIndent { get; set; }
            public StandardNumericUpDown TextIndent { get; set; }
            public StandardNumericUpDown TabPosition { get; set; }
        }
        #endregion
        
        private void UpdateLevelData(int level)
        {
            if (updatingControls || level < 1 || level > levelDataList.Count) return;

            var levelData = levelDataList[level - 1];
            var controls = GetLevelControls(level);
            
            // 使用MultiLevelListControlFactory获取数值（自动使用Word API转换）
            var inputValues = MultiLevelListControlFactory.GetInputValues(levelsContainer, level);
            
            // 更新数据
            levelData.NumberStyle = controls.NumberStyle?.Text ?? levelData.NumberStyle;
            levelData.NumberFormat = controls.NumberFormat?.Text ?? levelData.NumberFormat;
            levelData.AfterNumberType = controls.AfterNumber?.Text ?? levelData.AfterNumberType;
            levelData.LinkedStyle = controls.LinkedStyle?.Text ?? levelData.LinkedStyle;
            levelData.NumberIndent = inputValues.NumberIndent;
            levelData.TextIndent = inputValues.TextIndent;
            levelData.TabPosition = inputValues.TabPosition;
        }




        private void SetupEventHandlers()
        {
            // 底部控制按钮事件
            cmbLevelCount.SelectedIndexChanged += CmbLevelCount_SelectedIndexChanged;
            // btnSetLevelStyle.Click 事件已在设计器中绑定
            btnLoadCurrentList.Click += BtnLoadCurrentList_Click;
            btnSetMultiLevelList.Click += BtnApply_Click;
            // 关闭事件已由设计器绑定，避免重复执行。
            btnImport.Click += btnImport_Click;
            btnExport.Click += btnExport_Click;
            btnApplySettings.Click += BtnApplySettings_Click;

            // 右侧快捷设置事件
            chkNumberIndent.CheckedChanged += CheckBox_CheckedChanged; // 编号缩进
            chkTextIndent.CheckedChanged += CheckBox_CheckedChanged; // 文本缩进
            chkTabPosition.CheckedChanged += CheckBox_CheckedChanged; // 制表位位置
            chkProgressiveIndent.CheckedChanged += ProgressiveIndent_CheckedChanged; // 递进缩进设置
            chkLinkTitles.CheckedChanged += LinkTitles_CheckedChanged; // 链接到标题样式
            chkUnlinkTitles.CheckedChanged += UnlinkTitles_CheckedChanged; // 不链接标题样式
        }

        private void CmbLevelCount_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (updatingControls) return;
            int selectedLevels;
            if (!int.TryParse(cmbLevelCount.GetSelectedText(), out selectedLevels) || selectedLevels < 0 || selectedLevels > MAX_LEVELS) return;
            currentLevels = selectedLevels;
            CreateLevelControls();
        }



        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = ConfigurationManager.ShowImportDialog(this);
                if (string.IsNullOrEmpty(filePath)) return;
                LoadConfigurationFromFile(filePath);
                MessageBox.Show(this, "配置导入成功！", "导入成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "导入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = ConfigurationManager.ShowExportDialog(this);
                if (string.IsNullOrEmpty(filePath)) return;
                SaveConfigurationToFile(filePath);
                MessageBox.Show(this, "配置导出成功！", "导出成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLoadCurrentList_Click(object sender, EventArgs e)
        {
            try
            {
                if (app.Documents.Count == 0) throw new InvalidOperationException("请先打开 Word 文档。");
                var document = app.ActiveDocument;
                var template = app.Selection.Range.ListFormat.ListTemplate;
                if (template == null) throw new InvalidOperationException("当前位置没有可载入的多级列表。");
                // 先完整读取到临时集合；读取失败时保留原编辑内容。
                var loaded = new List<LevelData>();
                for (int i = 1; i <= MAX_LEVELS; i++)
                {
                    var level = template.ListLevels[i];
                    if (string.IsNullOrEmpty(level.NumberFormat) || level.NumberStyle == WdListNumberStyle.wdListNumberStyleNone) break;
                    int styleIndex = GetNumberStyleIndex(level.NumberStyle);
                    if (styleIndex < 0) throw new InvalidOperationException("第 " + i + " 级的编号样式暂不支持，未替换当前设置。");
                    string linkedStyle = string.IsNullOrEmpty(level.LinkedStyle) ? "无" : level.LinkedStyle;
                    for (int heading = 1; heading <= MAX_LEVELS; heading++)
                        if (linkedStyle == document.Styles[GetBuiltinStyleEnum(heading)].NameLocal)
                            linkedStyle = "标题 " + heading;
                    loaded.Add(new LevelData
                    {
                        Level = i,
                        NumberStyle = NumberStyleOptions[styleIndex],
                        NumberFormat = level.NumberFormat,
                        NumberIndent = (decimal)app.PointsToCentimeters(level.NumberPosition),
                        TextIndent = (decimal)app.PointsToCentimeters(level.TextPosition),
                        AfterNumberType = level.TrailingCharacter == WdTrailingCharacter.wdTrailingTab ? "制表位" :
                            level.TrailingCharacter == WdTrailingCharacter.wdTrailingSpace ? "空格" : "无",
                        TabPosition = level.TabPosition >= INVALID_TAB_POSITION ? 0m : (decimal)app.PointsToCentimeters(level.TabPosition),
                        LinkedStyle = linkedStyle,
                        StartAt = level.StartAt,
                        ResetOnHigher = level.ResetOnHigher
                    });
                }
                List<string> errors;
                if (!MultiLevelDataManager.ValidateLevelDataList(loaded, out errors))
                    throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
                sourceDocument = document;
                levelStyleSettings.Clear();
                ReplaceLevelData(loaded, loaded.Count);
                ClearPresetSelection();
                ClearQuickSettings();
                MessageBox.Show(this, "已载入当前多级列表设置", "成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "载入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int GetNumberStyleIndex(WdListNumberStyle numberStyle)
        {
            return Array.IndexOf(LevelNumStyle, numberStyle);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            try
            {
                GetTargetDocument();
                var activeLevels = GetValidatedLevels();
                // 在改动 Word 之前解析所有样式引用，避免错误名称导致部分写入。
                foreach (var level in activeLevels)
                    if (level.LinkedStyle != "无") GetWordStyleName(level.LinkedStyle);
                foreach (var style in levelStyleSettings.Take(currentLevels)) style.Validate();
                Globals.ThisAddIn.ExecuteWithUndoRecord("应用多级列表", () =>
                    CreateListTemplate(currentLevels,
                        activeLevels.Select(x => Array.IndexOf(NumberStyleOptions, x.NumberStyle)).ToArray(),
                        activeLevels.Select(x => x.NumberFormat).ToArray(),
                        activeLevels.Select(x => (float)x.NumberIndent).ToArray(),
                        activeLevels.Select(x => (float)x.TextIndent).ToArray(),
                        activeLevels.Select(x => x.AfterNumberType).ToArray(),
                        activeLevels.Select(x => (float)x.TabPosition).ToArray(),
                        activeLevels.Select(x => x.LinkedStyle).ToArray()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "应用多级列表失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Document GetTargetDocument()
        {
            if (app == null || app.Documents.Count == 0)
                throw new InvalidOperationException("请先打开 Word 文档。");
            var currentDocument = app.ActiveDocument;
            if (sourceDocument != null && !ReferenceEquals(currentDocument, sourceDocument))
                throw new InvalidOperationException("当前文档已切换。请返回原文档、重新打开本窗口，或点击“载入当前列表”后再应用。");
            if (sourceDocument == null) sourceDocument = currentDocument;
            return currentDocument;
        }

        private List<LevelData> GetValidatedLevels()
        {
            ValidateChildren();
            for (int level = 1; level <= currentLevels; level++) UpdateLevelData(level);
            var activeLevels = levelDataList.Take(currentLevels).ToList();
            List<string> errors;
            if (!MultiLevelDataManager.ValidateLevelDataList(activeLevels, out errors))
                throw new ArgumentException(string.Join(Environment.NewLine, errors));
            return activeLevels;
        }

        private void ReplaceLevelData(List<LevelData> loaded, int count)
        {
            // 始终保留九级数据槽，导入低级数后增加级数也不会越界。
            updatingControls = true;
            try
            {
                levelDataList.Clear();
                InitializeData();
                foreach (var level in loaded) levelDataList[level.Level - 1] = level;
                currentLevels = count;
                cmbLevelCount.SetSelectedItem(count.ToString());
            }
            finally { updatingControls = false; }
            CreateLevelControls();
        }

        private void CreateListTemplate(int levelCount, int[] numberStyles, string[] numberFormats,
            float[] numberIndents, float[] textIndents, string[] afterNumberTypes, float[] tabPositions, string[] linkedStyles)
        {
            var document = GetTargetDocument();
            // 设置链接样式可能立即影响段落，必须先记录原选区及每段级别。
            var targets = new List<KeyValuePair<Microsoft.Office.Interop.Word.Range, int>>();
            var paragraphs = app.Selection.Range.Duplicate.Paragraphs;
            for (int index = 1; index <= paragraphs.Count; index++)
            {
                var range = paragraphs[index].Range;
                int level = range.ListFormat.ListLevelNumber;
                if (level < 1 || level > levelCount) level = 1;
                var style = range.get_Style() as Style;
                if (style != null)
                    for (int candidate = 1; candidate <= levelCount; candidate++)
                        if (linkedStyles[candidate - 1] != "无" && style.NameLocal == GetWordStyleName(linkedStyles[candidate - 1]))
                        { level = candidate; break; }
                targets.Add(new KeyValuePair<Microsoft.Office.Interop.Word.Range, int>(range, level));
            }
            // 新建文档级模板，避免修改共享图库模板或选区之外的旧列表。
            var template = document.ListTemplates.Add(true);
            ConfigureListLevels(template, levelCount, numberStyles, numberFormats,
                numberIndents, textIndents, afterNumberTypes, tabPositions, linkedStyles);
            ApplyLevelStyleSettings();
            ApplyListTemplate(template, targets);
        }

        /// <summary>
        /// 配置列表级别属性
        /// </summary>
        private void ConfigureListLevels(ListTemplate listTemplate, int levelCount, int[] numberStyles, 
            string[] numberFormats, float[] numberIndents, float[] textIndents, 
            string[] afterNumberTypes, float[] tabPositions, string[] linkedStyles)
        {
            // 设置各级属性
            for (int i = 1; i <= levelCount; i++)
            {
                ListLevel listLevel = listTemplate.ListLevels[i];
                
                if (numberStyles[i - 1] != -1)
                {
                    listLevel.NumberStyle = LevelNumStyle[numberStyles[i - 1]];
                }
                
                listLevel.NumberFormat = numberFormats[i - 1];
                
                // 设置链接样式
                SetLinkedStyle(listLevel, linkedStyles[i - 1]);
                
                listLevel.NumberPosition = app.CentimetersToPoints(numberIndents[i - 1]);
                listLevel.TextPosition = app.CentimetersToPoints(textIndents[i - 1]);
                
                // 设置编号之后的字符类型和制表位位置
                SetTrailingCharacter(listLevel, afterNumberTypes[i - 1], tabPositions[i - 1]);
                
                listLevel.StartAt = levelDataList[i - 1].StartAt;
                listLevel.ResetOnHigher = levelDataList[i - 1].ResetOnHigher < 0 ? i - 1 : levelDataList[i - 1].ResetOnHigher;
            }
            
            // 清空未使用的级别
            for (int j = levelCount + 1; j <= 9; j++)
            {
                ListLevel listLevel = listTemplate.ListLevels[j];
                listLevel.NumberFormat = "";
                listLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleNone;
                listLevel.LinkedStyle = "";
            }
        }

        /// <summary>
        /// 设置链接样式
        /// </summary>
        private void SetLinkedStyle(ListLevel listLevel, string linkedStyleName)
        {
            listLevel.LinkedStyle = GetWordStyleName(linkedStyleName);
        }

        /// <summary>
        /// 获取内置样式枚举
        /// </summary>
        private WdBuiltinStyle GetBuiltinStyleEnum(int level)
        {
            switch (level)
            {
                case 1: return WdBuiltinStyle.wdStyleHeading1;
                case 2: return WdBuiltinStyle.wdStyleHeading2;
                case 3: return WdBuiltinStyle.wdStyleHeading3;
                case 4: return WdBuiltinStyle.wdStyleHeading4;
                case 5: return WdBuiltinStyle.wdStyleHeading5;
                case 6: return WdBuiltinStyle.wdStyleHeading6;
                case 7: return WdBuiltinStyle.wdStyleHeading7;
                case 8: return WdBuiltinStyle.wdStyleHeading8;
                case 9: return WdBuiltinStyle.wdStyleHeading9;
                default: return WdBuiltinStyle.wdStyleHeading1;
            }
        }

        /// <summary>
        /// 设置尾随字符
        /// </summary>
        private void SetTrailingCharacter(ListLevel listLevel, string afterNumberType, float tabPosition)
        {
            if (afterNumberType == "制表位")
            {
                listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                listLevel.TabPosition = app.CentimetersToPoints(tabPosition);
            }
            else if (afterNumberType == "空格")
            {
                listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
            }
            else
            {
                listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
            }
        }

        /// <summary>
        /// 应用列表模板到选区
        /// </summary>
        private void ApplyListTemplate(ListTemplate listTemplate,
            List<KeyValuePair<Microsoft.Office.Interop.Word.Range, int>> targets)
        {
            for (int index = 0; index < targets.Count; index++)
            {
                object continuePrevious = index > 0;
                object applyTo = WdListApplyTo.wdListApplyToSelection;
                object behavior = WdDefaultListBehavior.wdWord9ListBehavior;
                object applyLevel = targets[index].Value;
                targets[index].Key.ListFormat.ApplyListTemplateWithLevel(listTemplate, ref continuePrevious, ref applyTo, ref behavior, ref applyLevel);
            }
        }

        /// <summary>
        /// 应用级别样式设置
        /// </summary>
        private void ApplyLevelStyleSettings()
        {
            var document = GetTargetDocument();
            foreach (var style in levelStyleSettings.Take(currentLevels))
                if (!style.SetStyle(document))
                    throw new InvalidOperationException("样式应用失败：" + style.StyleName);
        }

        private void BtnApplySettings_Click(object sender, EventArgs e)
        {
            try { ApplyQuickSettings(); }
            catch (Exception ex) { MessageBox.Show(this, "快捷设置失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ApplyQuickSettings()
        {
            if (currentLevels == 0) throw new InvalidOperationException("请先选择列表级数。");
            ValidateChildren();
            // 先检查所有级别的结果，防止递进溢出或单位换算后越界导致部分更新。
            for (int level = 1; level <= currentLevels; level++)
            {
                if (chkNumberIndent.Checked) ValidateQuickIndent(nudNumberIndent.GetValueInCentimeters());
                if (chkTextIndent.Checked) ValidateQuickIndent(nudTextIndent.GetValueInCentimeters());
                if (chkTabPosition.Checked) ValidateQuickIndent(nudTabPosition.GetValueInCentimeters());
                if (chkProgressiveIndent.Checked)
                    ValidateQuickIndent(nudFirstLevelIndent.GetValueInCentimeters() + nudIncrementIndent.GetValueInCentimeters() * (level - 1));
            }
            for (int level = 1; level <= currentLevels; level++)
            {
                var controls = GetLevelControls(level);
                
                // 1. 统一缩进设置
                if (chkNumberIndent.Checked && controls.NumberIndent != null)
                    controls.NumberIndent.SetValueInCentimeters(nudNumberIndent.GetValueInCentimeters());
                
                if (chkTextIndent.Checked && controls.TextIndent != null)
                    controls.TextIndent.SetValueInCentimeters(nudTextIndent.GetValueInCentimeters());
                
                if (chkTabPosition.Checked && controls.TabPosition != null)
                {
                    controls.AfterNumber.SelectedItem = "制表位";
                    controls.TabPosition.SetValueInCentimeters(nudTabPosition.GetValueInCentimeters());
                }

                // 2. 递进缩进设置
                if (chkProgressiveIndent.Checked && controls.NumberIndent != null)
                {
                    if (level == 1)
                    {
                        controls.NumberIndent.SetValueInCentimeters(nudFirstLevelIndent.GetValueInCentimeters());
                    }
                    else
                    {
                        // 使用Word API进行递进计算
                        decimal baseIndent = nudFirstLevelIndent.GetValueInCentimeters();
                        decimal increment = nudIncrementIndent.GetValueInCentimeters();
                        controls.NumberIndent.SetValueInCentimeters(baseIndent + increment * (level - 1));
                    }
                }

                // 3. 链接标题样式
                if (chkLinkTitles.Checked && controls.LinkedStyle != null)
                    controls.LinkedStyle.SelectedIndex = level;
                else if (chkUnlinkTitles.Checked && controls.LinkedStyle != null)
                    controls.LinkedStyle.SelectedIndex = 0;
            }

            // 清空快捷设置
            ClearQuickSettings();
        }

        private void ClearQuickSettings()
        {
            chkNumberIndent.Checked = false;
            chkTextIndent.Checked = false;
            chkTabPosition.Checked = false;
            chkProgressiveIndent.Checked = false;
            chkLinkTitles.Checked = false;
            chkUnlinkTitles.Checked = false;
            nudNumberIndent.Enabled = false;
            nudFirstLevelIndent.Enabled = false;
            nudIncrementIndent.Enabled = false;
            nudTextIndent.Enabled = false;
            nudTabPosition.Enabled = false;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (chkNumberIndent.Checked && sender == chkNumberIndent) chkProgressiveIndent.Checked = false;
            nudNumberIndent.Enabled = chkNumberIndent.Checked;
            nudTextIndent.Enabled = chkTextIndent.Checked;
            nudTabPosition.Enabled = chkTabPosition.Checked;
        }

        private void ValidateQuickIndent(decimal value)
        {
            if (value < 0m || value > MAX_INDENT_CM)
                throw new ArgumentOutOfRangeException("value", "缩进和制表位换算后必须在 0 至 50 厘米之间。");
        }

        private void ProgressiveIndent_CheckedChanged(object sender, EventArgs e)
        {
            if (chkProgressiveIndent.Checked) chkNumberIndent.Checked = false;
            nudFirstLevelIndent.Enabled = chkProgressiveIndent.Checked;
            nudIncrementIndent.Enabled = chkProgressiveIndent.Checked;
        }

        private void LinkTitles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLinkTitles.Checked)
            {
                chkUnlinkTitles.Checked = false;
            }
        }

        private void UnlinkTitles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUnlinkTitles.Checked)
            {
                chkLinkTitles.Checked = false;
            }
        }




        private string GetWordStyleName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName) || displayName == "无") return "";
            int level;
            var document = GetTargetDocument();
            if (int.TryParse(displayName.Replace("标题 ", "").Replace("标题", ""), out level) && level >= 1 && level <= MAX_LEVELS)
                return document.Styles[GetBuiltinStyleEnum(level)].NameLocal;
            var style = document.Styles[displayName];
            if (style.Type != WdStyleType.wdStyleTypeParagraph && style.Type != WdStyleType.wdStyleTypeLinked)
                throw new ArgumentException("链接样式不是段落样式：" + displayName);
            return style.NameLocal;
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        private void SaveConfigurationToFile(string filePath)
        {
            if (currentLevels > 0) GetValidatedLevels();
            ConfigurationManager.SaveConfigurationToFile(filePath, levelDataList, currentLevels, levelStyleSettings);
        }


        /// <summary>
        /// 从文件加载配置
        /// </summary>
        private void LoadConfigurationFromFile(string filePath)
        {
            List<LevelData> loaded;
            List<WordStyleInfo> loadedStyles;
            int count;
            ConfigurationManager.LoadConfigurationFromFile(filePath, out loaded, out count, out loadedStyles);
            ReplaceLevelData(loaded, count);
            levelStyleSettings = loadedStyles;
            ClearQuickSettings();
            ClearPresetSelection();
        }

        /// <summary>
        /// 设置每级样式按钮点击事件
        /// </summary>
        private void BtnSetLevelStyle_Click(object sender, EventArgs e)
        {
            OpenLevelStyleSettingsForm();
        }

        /// <summary>
        /// 打开级别样式设置窗体
        /// </summary>
        private void OpenLevelStyleSettingsForm()
        {
            try
            {
                GetTargetDocument();
                if (currentLevels == 0) return;
                using (var dialog = new LevelStyleSettingsForm(currentLevels))
                {
                    dialog.LoadExistingStyles(levelStyleSettings);
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    var hidden = levelStyleSettings.Skip(currentLevels).ToList();
                    levelStyleSettings = dialog.GetLevelStyles();
                    levelStyleSettings.AddRange(hidden);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "设置每级样式失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
