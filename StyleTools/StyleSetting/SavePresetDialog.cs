using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordMan
{
    /// <summary>
    /// 保存预设对话框：上方名称输入框，下方一行两个按钮（确定/取消）
    /// </summary>
    public class SavePresetDialog : Form
    {
        private TextBox txtName;
        private Button btnCancel;
        private Button btnOk;

        /// <summary>
        /// 用户输入的预设名称
        /// </summary>
        public string PresetName
        {
            get { return txtName != null ? txtName.Text.Trim() : ""; }
        }

        public SavePresetDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.txtName = new TextBox();
            this.btnCancel = new Button();
            this.btnOk = new Button();
            this.SuspendLayout();

            // 窗体
            this.ClientSize = new Size(340, 118);
            this.Font = new Font("微软雅黑", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "保存预设";
            this.BackColor = Color.FromArgb(250, 250, 252);
            this.AcceptButton = this.btnOk;
            this.CancelButton = this.btnCancel;

            // 名称标签
            var label = new Label
            {
                Text = "预设名称：",
                Location = new Point(20, 18),
                Size = new Size(80, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(label);

            // 名称输入框
            this.txtName.Location = new Point(100, 18);
            this.txtName.Size = new Size(220, 25);
            this.txtName.BorderStyle = BorderStyle.FixedSingle;
            this.txtName.Font = new Font("微软雅黑", 9F);
            this.txtName.TabIndex = 0;
            this.Controls.Add(this.txtName);

            // 取消按钮
            this.btnCancel.Text = "取消";
            this.btnCancel.Location = new Point(182, 68);
            this.btnCancel.Size = new Size(70, 30);
            this.btnCancel.FlatStyle = FlatStyle.Flat;
            this.btnCancel.FlatAppearance.BorderSize = 1;
            this.btnCancel.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 205);
            this.btnCancel.BackColor = Color.FromArgb(245, 245, 247);
            this.btnCancel.ForeColor = Color.FromArgb(33, 37, 41);
            this.btnCancel.Font = new Font("微软雅黑", 9F);
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.TabIndex = 1;
            this.Controls.Add(this.btnCancel);

            // 确定按钮（主按钮，蓝色）
            this.btnOk.Text = "确定";
            this.btnOk.Location = new Point(258, 68);
            this.btnOk.Size = new Size(70, 30);
            this.btnOk.FlatStyle = FlatStyle.Flat;
            this.btnOk.FlatAppearance.BorderSize = 0;
            this.btnOk.BackColor = Color.FromArgb(70, 130, 230);
            this.btnOk.ForeColor = Color.White;
            this.btnOk.Font = new Font("微软雅黑", 9F, FontStyle.Bold);
            this.btnOk.DialogResult = DialogResult.OK;
            this.btnOk.TabIndex = 2;
            this.Controls.Add(this.btnOk);

            this.ResumeLayout(false);
        }
    }
}
