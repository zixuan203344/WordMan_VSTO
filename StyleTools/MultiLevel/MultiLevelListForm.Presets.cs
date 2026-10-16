using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordMan
{
    public partial class MultiLevelListForm
    {
        private readonly MultiLevelPresetManager presetManager = new MultiLevelPresetManager();
        private ComboBox cmbPresetStyle;
        private StandardButton btnLoadPreset;
        private StandardButton btnSavePreset;
        private StandardButton btnDeletePreset;
        private Label lblPresetStatus;
        private bool refreshingPresets;
        private string selectedPresetName;

        private void InitializePresetControls()
        {
            // 独立顶栏，不挤占原有底部按钮，也不增加窗口高度。
            var presetBar = new Panel
            {
                Name = "presetBar", Location = new Point(15, 12),
                Size = new Size(leftPanel.ClientSize.Width - 30, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, TabIndex = 0
            };
            var label = new Label
            { Text = "预设风格：", AutoSize = true, Location = new Point(0, 9) };
            cmbPresetStyle = new ComboBox
            {
                Name = "cmbPresetStyle", DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(85, 5), Size = new Size(200, 25), TabIndex = 0
            };
            btnLoadPreset = new StandardButton
            {
                Name = "btnLoadPreset", Text = "载入", Location = new Point(295, 0),
                Size = new Size(65, 35), Enabled = false, TabIndex = 1
            };
            btnSavePreset = new StandardButton
            {
                Name = "btnSavePreset", Text = "存预设", Location = new Point(370, 0),
                Size = new Size(85, 35), TabIndex = 2
            };
            btnDeletePreset = new StandardButton
            {
                Name = "btnDeletePreset", Text = "删预设", Location = new Point(465, 0),
                Size = new Size(100, 35), Enabled = false, TabIndex = 3
            };
            lblPresetStatus = new Label
            {
                Name = "lblPresetStatus", Text = "选择预设载入，点击“应用”写入文档",
                Location = new Point(580, 4), Size = new Size(275, 32), AutoEllipsis = true
            };
            presetBar.Controls.AddRange(new Control[] { label, cmbPresetStyle, btnLoadPreset, btnSavePreset, btnDeletePreset, lblPresetStatus });
            leftPanel.Controls.Add(presetBar);
            levelsScrollPanel.Top = 58;
            levelsScrollPanel.Height = 317;
            cmbPresetStyle.SelectedIndexChanged += CmbPresetStyle_SelectedIndexChanged;
            btnLoadPreset.Click += CmbPresetStyle_SelectedIndexChanged;
            btnSavePreset.Click += BtnSavePreset_Click;
            btnDeletePreset.Click += BtnDeletePreset_Click;
            try { RefreshPresetChoices(null); }
            catch (Exception ex)
            {
                // 文件损坏时不把旧内容当作空集合覆盖，普通列表编辑继续可用。
                cmbPresetStyle.Enabled = false;
                btnSavePreset.Enabled = false;
                lblPresetStatus.Text = "预设读取失败，重新打开窗口重试";
                string message = "读取多级列表预设失败：" + ex.Message;
                // 等到 Word 所属窗口建立后再显示子窗口，避免构造期间抢占前台。
                Shown += (sender, args) => MessageBox.Show(this, message, "预设", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshPresetChoices(string selection)
        {
            var names = presetManager.GetPresetNames();
            bool canDelete = selection != null && presetManager.HasSavedPreset(selection);
            refreshingPresets = true;
            try
            {
                cmbPresetStyle.Items.Clear();
                cmbPresetStyle.Items.AddRange(names.ToArray());
                cmbPresetStyle.SelectedItem = selection;
                selectedPresetName = cmbPresetStyle.SelectedItem as string;
                btnLoadPreset.Enabled = selectedPresetName != null;
                btnDeletePreset.Text = MultiLevelPresetManager.IsBuiltIn(selectedPresetName) ? "恢复默认" : "删预设";
                btnDeletePreset.Enabled = canDelete;
            }
            finally { refreshingPresets = false; }
        }

        private void CmbPresetStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (refreshingPresets) return;
            string name = cmbPresetStyle.SelectedItem as string;
            if (name == null) return;
            try
            {
                // 先完整读取和校验，再替换当前编辑内容；此处不操作 Word 文档。
                var preset = presetManager.LoadPreset(name);
                bool canDelete = presetManager.HasSavedPreset(name);
                ReplaceLevelData(preset.Levels, preset.LevelCount);
                levelStyleSettings = preset.Styles;
                ClearQuickSettings();
                selectedPresetName = name;
                btnLoadPreset.Enabled = true;
                btnDeletePreset.Text = MultiLevelPresetManager.IsBuiltIn(name) ? "恢复默认" : "删预设";
                btnDeletePreset.Enabled = canDelete;
                lblPresetStatus.Text = "已载入：" + name;
            }
            catch (Exception ex)
            {
                refreshingPresets = true;
                try { cmbPresetStyle.SelectedItem = selectedPresetName; }
                finally { refreshingPresets = false; }
                MessageBox.Show(this, "载入预设失败：" + ex.Message, "预设", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearPresetSelection()
        {
            refreshingPresets = true;
            try
            {
                cmbPresetStyle.SelectedIndex = -1;
                selectedPresetName = null;
                btnLoadPreset.Enabled = false;
                btnDeletePreset.Enabled = false;
                btnDeletePreset.Text = "删预设";
                if (cmbPresetStyle.Enabled) lblPresetStatus.Text = "可将当前配置存为预设";
            }
            finally { refreshingPresets = false; }
        }

        private void BtnSavePreset_Click(object sender, EventArgs e)
        {
            try
            {
                // 零级或非法输入不创建空预设；所有九级样式槽随当前有效列表一起保存。
                var levels = GetValidatedLevels();
                using (var dialog = new SavePresetDialog(presetManager.GetPresetNames(), selectedPresetName))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    presetManager.SavePreset(dialog.PresetName, levels, currentLevels, levelStyleSettings);
                    RefreshPresetChoices(dialog.PresetName);
                    lblPresetStatus.Text = "已保存：" + dialog.PresetName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存预设失败：" + ex.Message, "预设", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeletePreset_Click(object sender, EventArgs e)
        {
            string name = selectedPresetName;
            if (name == null) return;
            try
            {
                bool builtIn = MultiLevelPresetManager.IsBuiltIn(name);
                string prompt = builtIn ? "删除对“" + name + "”的覆盖，恢复内置默认？" : "删除预设“" + name + "”？";
                if (MessageBox.Show(this, prompt + "\n当前编辑内容将保留。", "预设",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                bool deleted = presetManager.DeletePreset(name);
                // 不自动重新载入，删除/恢复操作不会丢失尚未保存的编辑。
                RefreshPresetChoices(null);
                lblPresetStatus.Text = deleted ? (builtIn ? "已恢复默认，请选择预设载入" : "已删除：" + name) : "该预设没有可删除的保存内容";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "删除预设失败：" + ex.Message, "预设", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
