using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Windows.Forms;
using WordMan;
using Word = Microsoft.Office.Interop.Word;

// Run in an STA .NET Framework process. Add --word for the COM integration checks.
// All Word operations use a new, hidden instance and disposable, unsaved documents.
internal static class StyleSettingsRegression
{
    private static int checks;
    private static readonly BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr handle, uint command);
    private static object Call(object target, string name, params object[] args)
    {
        try { return target.GetType().GetMethods(Private).Single(m => m.Name == name && m.GetParameters().Length == args.Length).Invoke(target, args); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    private static T Field<T>(object target, string name)
    { return (T)target.GetType().GetField(name, Private).GetValue(target); }
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception("FAIL: " + message);
        checks++;
        Console.WriteLine("PASS: " + message);
    }
    private static List<CustomStyle> Models(StyleSettings form)
    { return Field<List<CustomStyle>>(form, "Styles"); }
    private static void Select(StyleSettings form, params string[] names)
    {
        var list = Field<ListBox>(form, "Lst_Styles");
        list.ClearSelected();
        foreach (var name in names) list.SetSelected(list.Items.IndexOf(name), true);
    }
    private static void Click(StyleSettings form, string name)
    { Call(form, name, null, EventArgs.Empty); }
    private static string StyleName(object value)
    { return value is Word.Style ? ((Word.Style)value).NameLocal : Convert.ToString(value); }

    private static void CheckSavePresetDialog()
    {
        using (var dialog = new SavePresetDialog(new[] { "公文风格", "已有预设", "已有预设", " 已有预设 ", " ", null }, "已有预设"))
        {
            var names = Field<ComboBox>(dialog, "cmbName");
            var ok = Field<Button>(dialog, "btnOk");
            Check(names.DropDownStyle == ComboBoxStyle.DropDown && names.Items.Count == 2, "preset name is editable and existing choices are unique");
            Check(dialog.PresetName == "已有预设" && ok.Enabled, "current preset name is prefilled");
            names.SelectedItem = "公文风格";
            Check(dialog.PresetName == "公文风格" && ok.Enabled, "built-in preset name can be selected for overwrite");
            names.Text = "  新预设  ";
            Check(dialog.PresetName == "新预设" && ok.Enabled, "new preset name can be typed and trimmed");
            names.Text = "已有预设";
            Check(dialog.PresetName == "已有预设", "existing preset name can also be typed");
            names.Text = "   ";
            Check(!ok.Enabled, "blank preset name cannot be confirmed");
            Check(!dialog.ShowInTaskbar && dialog.AcceptButton == ok && dialog.CancelButton == Field<Button>(dialog, "btnCancel"), "save dialog retains Enter/Escape buttons without a separate taskbar window");
        }
        using (var dialog = new SavePresetDialog(null, null))
            Check(dialog.PresetName == "" && !Field<Button>(dialog, "btnOk").Enabled, "empty preset list still permits entering a new name");
    }

    private static void CheckNestedWindowOwnership()
    {
        using (var host = new Form())
        using (var settings = new StyleSettings())
        {
            host.ShowInTaskbar = false;
            host.Show();
            var ownerType = typeof(StyleSettings).Assembly.GetType("WordMan.WordWindowOwner");
            var owner = (IWin32Window)Activator.CreateInstance(ownerType, new object[] { host.Handle });
            settings.ShowInTaskbar = false;
            settings.Show(owner);
            Check(GetWindow(settings.Handle, 4) == host.Handle, "modeless settings window is owned by the supplied native host");
            // Cover Cancel, title-bar close, and successful save-dialog close.
            for (int action = 0; action < 3; action++)
            {
                int closeAction = action;
                using (var dialog = new SavePresetDialog(new[] { "已有预设" }, "已有预设"))
                {
                    Exception callbackError = null;
                    dialog.Shown += (sender, e) => dialog.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            Check(GetWindow(dialog.Handle, 4) == settings.Handle, "save dialog is owned by settings, close path " + closeAction);
                            if (closeAction == 0) Field<Button>(dialog, "btnCancel").PerformClick();
                            else if (closeAction == 1) dialog.Close();
                            else Field<Button>(dialog, "btnOk").PerformClick();
                        }
                        catch (Exception ex)
                        {
                            callbackError = ex;
                            dialog.DialogResult = DialogResult.Cancel;
                        }
                    }));
                    var result = dialog.ShowDialog(settings);
                    if (callbackError != null) throw callbackError;
                    Check(result == (closeAction == 2 ? DialogResult.OK : DialogResult.Cancel), "save dialog returns expected result, close path " + closeAction);
                }
                Check(settings.Enabled && GetWindow(settings.Handle, 4) == host.Handle, "nested dialog preserves settings owner and enabled state, close path " + closeAction);
            }
            settings.Close();
            Check(host.Visible && host.Enabled && host.WindowState == FormWindowState.Normal, "closing settings leaves host visible and enabled");
            host.Close();
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Word.Application app = null;
        StyleSettings form = null;
        string folder = Path.Combine(Path.GetTempPath(), "WordManStyleRegression-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            Application.EnableVisualStyles();
            CheckSavePresetDialog();
            CheckNestedWindowOwnership();
            form = new StyleSettings();
            var normal = Models(form).First(x => x.Name == "正文");
            Select(form, "正文");
            Field<StandardComboBox>(form, "首行缩进方式下拉框").SelectedItem = "悬挂缩进";
            Field<StandardNumericUpDown>(form, "Nud_FirstLineIndent").Value = 0.75m;
            Check(normal.FirstLineIndent == -0.75f, "hanging indent updates model with negative value");
            Field<StandardTextBox>(form, "Txt_AddStyleName").Text = "AuditAdded";
            Click(form, "添加_Click");
            Check(Models(form).Single(x => x.Name == "AuditAdded").FirstLineIndent == -0.75f, "new style retains hanging indent sign");
            Select(form, "正文");
            Field<StandardComboBox>(form, "首行缩进方式下拉框").SelectedItem = "无";
            Check(normal.FirstLineIndent == 0, "selecting no indentation clears the value");
            Field<StandardComboBox>(form, "Cmb_FontSize").Text = "13.5 磅";
            Check(normal.FontSize == 13.5f, "typed custom font size is retained");
            Field<StandardComboBox>(form, "Cmb_BefreSpacing").Text = "12.5";
            Check(normal.BeforeSpacing == 12.5f && Field<StandardComboBox>(form, "Cmb_BefreSpacing").Text == "12.5", "decimal spacing input is not rewritten while typing");
            Field<StandardComboBox>(form, "Cmb_LineSpacing").Text = "18磅";
            Check(normal.LineSpacing == 18 && normal.LineSpacingRule == (int)Word.WdLineSpacing.wdLineSpaceExactly, "fixed line spacing has explicit rule");
            Select(form, "标题 1");
            Select(form, "正文");
            Check(Field<StandardComboBox>(form, "Cmb_LineSpacing").Text == "18 磅", "fixed line spacing survives selection changes");
            Check(Field<StandardComboBox>(form, "Cmb_FontSize").Text == "13.5 磅", "fractional font size survives selection changes");
            Select(form, "正文", "标题 1");
            Check(Field<StandardComboBox>(form, "首行缩进方式下拉框").Text == "多种" || Models(form).First(x => x.Name == "标题 1").FirstLineIndent == 0, "mixed indentation display is consistent");
            Field<StandardComboBox>(form, "Cmb_ParaAligment").SelectedItem = "右对齐";
            Check(normal.ParaAlignment == 2 && Models(form).First(x => x.Name == "标题 1").ParaAlignment == 2, "mixed placeholder does not offset alignment selection");
            Check(Field<HashSet<string>>(form, "_modifiedControls").SetEquals(new[] { "Cmb_ParaAligment" }), "display refresh does not mark unrelated batch properties");
            Field<StandardComboBox>(form, "显示标题数下拉框").SelectedIndex = 3;
            Check(Field<HashSet<string>>(form, "_modifiedControls").SetEquals(new[] { "Cmb_ParaAligment" }), "heading filter retains pending batch edits when selected styles stay the same");
            Select(form, "正文");
            Check(normal.ParaAlignment == 2, "batch edits survive selection changes");
            var title = Models(form).First(x => x.Name == "标题 1");
            title.FirstLineIndent = 1;
            Select(form, "正文", "标题 1");
            Check(Field<StandardComboBox>(form, "首行缩进方式下拉框").Text == "多种", "different indent modes show mixed placeholder: " + Field<StandardComboBox>(form, "首行缩进方式下拉框").Text + "; values=" + normal.FirstLineIndent + "," + title.FirstLineIndent);
            Field<StandardComboBox>(form, "首行缩进方式下拉框").SelectedItem = "悬挂缩进";
            Field<StandardNumericUpDown>(form, "Nud_FirstLineIndent").Value = 0.6m;
            Check(normal.FirstLineIndent == -0.6f && title.FirstLineIndent == -0.6f, "mixed first-line dropdown applies correct hanging mode");
            var indentControl = Field<StandardNumericUpDown>(form, "Nud_FirstLineIndent");
            Call(indentControl, "UnitLabel_Click", null, EventArgs.Empty);
            Check(indentControl.Unit == "磅" && Math.Abs(normal.FirstLineIndent + 0.6f) < 0.001, "switching indent units preserves the physical value");
            Field<StandardComboBox>(form, "首行缩进方式下拉框").SelectedItem = "首行缩进";
            Check(indentControl.Unit == "磅" && Math.Abs(normal.FirstLineIndent - 0.6f) < 0.001, "switching indent mode preserves the chosen unit");
            Select(form, "标题 1");
            string baseName = title.BaseStyle;
            Call(form, "RefreshBaseStyleComboBoxes");
            Check(title.BaseStyle == baseName, "refreshing relation choices does not overwrite model");
            Field<StandardComboBox>(form, "显示标题数下拉框").SelectedIndex = 3;
            Check(Field<ListBox>(form, "Lst_Styles").SelectedItem.ToString() == "标题 1", "heading filter retains selected style");
            var before = Models(form).ToList();
            try { Call(form, "ReplaceStyles", new List<CustomStyle> { new CustomStyle { Name = "dup" }, new CustomStyle { Name = "dup" } }, 4); }
            catch (ArgumentException) { }
            Check(Models(form).SequenceEqual(before), "invalid imported collection leaves editor unchanged");
            var cycle = new List<CustomStyle> { new CustomStyle { Name = "A", BaseStyle = "B" }, new CustomStyle { Name = "B", BaseStyle = "A" } };
            bool cycleBlocked = false;
            try { Call(form, "ReplaceStyles", cycle, 4); } catch (ArgumentException) { cycleBlocked = true; }
            Check(cycleBlocked && Models(form).SequenceEqual(before), "circular style inheritance is rejected without replacing editor state");
            Select(form, "正文");
            Field<StandardComboBox>(form, "Cmb_FontSize").Text = "NaN";
            Check(!(bool)Call(form, "ValidateEditorInputs") && normal.FontSize == 13.5f, "invalid numeric input cannot corrupt the model or be applied");
            Field<StandardComboBox>(form, "Cmb_FontSize").Text = "13.5磅";
            string xml = Path.Combine(folder, "styles.xml");
            StyleFileManager.SerializeToXml(new StyleExportData { Styles = before, TitleCount = 3 }, xml);
            var restored = StyleFileManager.DeserializeFromXml<StyleExportData>(xml);
            Check(restored.Styles[0].LineSpacingRule == normal.LineSpacingRule && restored.TitleCount == 3, "XML round trip retains line spacing rule and title count");
            normal.FontColor = Color.FromArgb(12, 34, 56);
            StyleFileManager.SerializeListToXml(new List<CustomStyle> { normal }, xml);
            var legacy = StyleFileManager.DeserializeListFromXml<CustomStyle>(xml);
            Check(legacy.Single().FontColor.ToArgb() == normal.FontColor.ToArgb(), "legacy list XML and font colors round trip");
            var expectedPresetBody = StylePresetManager.GetPresetStyles("公文风格").First(x => x.Name == "正文");
            int expectedTitleCount = Math.Max(0, Math.Min(9, StylePresetManager.GetPresetTitleCount("公文风格")));
            float expectedIndent = expectedPresetBody.FirstLineIndentByChar != 0
                ? expectedPresetBody.FirstLineIndentByChar * expectedPresetBody.FontSize * 2.54f / 72f
                : expectedPresetBody.FirstLineIndent;
            Call(form, "LoadPresetStyle", "公文风格");
            var officialBody = Models(form).First(x => x.Name == "正文");
            Check(Field<StandardComboBox>(form, "显示标题数下拉框").SelectedIndex == expectedTitleCount && Math.Abs(officialBody.FirstLineIndent - expectedIndent) < 0.001, "preset loads saved metadata and character indent, including overridden built-in presets");
            Call(form, "ReplaceStyles", new List<CustomStyle> { new CustomStyle { Name = "标题 7" } }, 7);
            Check(Field<StandardComboBox>(form, "显示标题数下拉框").SelectedIndex == 7 && Field<ListBox>(form, "Lst_Styles").Items.Contains("标题 7"), "restored custom preset title count retains higher heading levels");
            Check(StylePresetManager.IsBuiltInPreset("公文风格") && !StylePresetManager.IsBuiltInPreset("AuditCustomPreset"), "built-in names remain identifiable after allowing saved overrides");
            form.Size = new Size(form.Width + 180, form.Height + 140);
            foreach (Control button in new[] { "加载", "添加", "删除", "读取文档样式", "导入", "导出", "存预设", "应用当前", "Btn_ApplySet", "关闭" }.Select(name => Field<StandardButton>(form, name)))
                Check(button.Left >= 0 && button.Right <= button.Parent.ClientSize.Width && button.Bottom <= button.Parent.ClientSize.Height, "bottom button fits: " + button.Name);

            if (!args.Contains("--word"))
            {
                Console.WriteLine("Completed " + checks + " checks. Word integration: not requested.");
                return 0;
            }
            // Integration tests run the actual COM implementation, outside the VSTO host.
            Console.WriteLine("Starting disposable Word instance...");
            app = new Word.Application();
            Console.WriteLine("Word instance created.");
            app.Visible = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            var addinType = typeof(StyleSettings).Assembly.GetType("WordMan.ThisAddIn");
            object addin = FormatterServices.GetUninitializedObject(addinType);
            addinType.GetField("Application", Private).SetValue(addin, app);
            typeof(StyleSettings).Assembly.GetType("WordMan.Globals").GetField("_ThisAddIn", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, addin);
            var doc = app.Documents.Add();
            Console.WriteLine("Disposable document created.");
            var wordNormal = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            wordNormal.Font.Bold = -1;
            wordNormal.Font.Italic = -1;
            form.RefreshFromDocument();
            string normalName = wordNormal.NameLocal;
            Select(form, normalName);
            Check(Models(form).First(x => x.Name == normalName).Bold && Models(form).First(x => x.Name == normalName).Italic, "document read recognizes Word true (-1) for bold and italic");
            Field<StandardComboBox>(form, "Cmb_LineSpacing").Text = "18磅";
            Field<StandardComboBox>(form, "Cmb_FontSize").Text = "13.5磅";
            Click(form, "应用当前_Click");
            Check(wordNormal.ParagraphFormat.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly && Math.Abs(wordNormal.ParagraphFormat.LineSpacing - 18) < 0.1, "apply current uses fixed spacing in Word");
            Check(Math.Abs(wordNormal.Font.Size - 13.5) < 0.01, "apply current writes custom font size to Word");
            Field<StandardComboBox>(form, "Cmb_LineSpacing").Text = "最小值20磅";
            Click(form, "应用当前_Click");
            Check(wordNormal.ParagraphFormat.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceAtLeast, "minimum line spacing applies to Word");
            Field<StandardComboBox>(form, "Cmb_LineSpacing").Text = "1.25倍行距";
            Click(form, "应用当前_Click");
            Check(wordNormal.ParagraphFormat.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceMultiple && Math.Abs(wordNormal.ParagraphFormat.LineSpacing - app.LinesToPoints(1.25f)) < 0.1, "multiple line spacing applies to Word");
            wordNormal.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            wordNormal.ParagraphFormat.LineSpacing = 15;
            wordNormal.Font.Underline = Word.WdUnderline.wdUnderlineDouble;
            var heading = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
            form.RefreshFromDocument();
            Select(form, normalName, heading.NameLocal);
            Field<ToggleButton>(form, "Btn_Italic").Pressed = !Field<ToggleButton>(form, "Btn_Italic").Pressed;
            Click(form, "Btn_ApplySet_Click");
            Check(wordNormal.Font.Underline == Word.WdUnderline.wdUnderlineDouble && Math.Abs(wordNormal.ParagraphFormat.LineSpacing - 15) < 0.1, "batch apply preserves unedited Word formatting");
            Check(Field<HashSet<string>>(form, "_modifiedControls").Count == 0, "batch refresh clears pending properties");
            var created = doc.Styles.Add("AuditUnusedParagraph", Word.WdStyleType.wdStyleTypeParagraph);
            created.Font.Bold = -1;
            var available = (List<string>)Call(form, "GetDocumentStyles");
            Check(available.Contains(created.NameLocal) && !available.Contains(doc.Styles[Word.WdBuiltinStyle.wdStyleHyperlink].NameLocal), "load choices include unused custom paragraphs and exclude character styles");
            form.RefreshFromDocument();
            Select(form, created.NameLocal);
            Field<StandardComboBox>(form, "样式基准下拉框").SelectedItem = "(无样式)";
            Field<StandardComboBox>(form, "后续段落样式下拉框").SelectedItem = normalName;
            Click(form, "应用当前_Click");
            Check(StyleName(created.get_BaseStyle()) == "" && StyleName(created.get_NextParagraphStyle()) == normalName,
                "no-base and next-paragraph choices are written to Word: base=" + StyleName(created.get_BaseStyle()) + "; next=" + StyleName(created.get_NextParagraphStyle()) + "; " + Field<ToolStripStatusLabel>(form, "statusLabel").Text);
            Field<StandardComboBox>(form, "Cmb_EngFontName").Text = "Arial";
            var colorButton = Field<StandardButton>(form, "Btn_FontColor");
            colorButton.BackColor = Color.FromArgb(12, 34, 56);
            Call(form, "OnStyleChanged", colorButton, EventArgs.Empty);
            Field<StandardComboBox>(form, "Cmb_ParaAligment").SelectedItem = "右对齐";
            Field<StandardNumericUpDown>(form, "Nud_LeftIndent").SetValueInCentimeters(0.5m);
            Field<StandardNumericUpDown>(form, "Nud_RightIndent").SetValueInCentimeters(0.3m);
            Field<StandardComboBox>(form, "Cmb_BefreSpacing").Text = "6磅";
            Field<StandardComboBox>(form, "Cmb_AfterSpacing").Text = "8磅";
            Click(form, "应用当前_Click");
            Check(created.Font.NameAscii == "Arial" && (int)created.Font.Color == ColorTranslator.ToOle(colorButton.BackColor), "font family and color apply to Word");
            Check(created.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphRight && Math.Abs(created.ParagraphFormat.LeftIndent - app.CentimetersToPoints(0.5f)) < 0.1 && Math.Abs(created.ParagraphFormat.RightIndent - app.CentimetersToPoints(0.3f)) < 0.1, "alignment and left/right indentation apply to Word");
            Check(Math.Abs(created.ParagraphFormat.SpaceBefore - 6) < 0.01 && Math.Abs(created.ParagraphFormat.SpaceAfter - 8) < 0.01, "paragraph before/after spacing applies to Word");
            Field<StandardTextBox>(form, "Txt_AddStyleName").Text = "AuditNewParagraph";
            Click(form, "添加_Click");
            Field<ToggleButton>(form, "Btn_UnderLine").Pressed = true;
            Click(form, "应用当前_Click");
            Check(doc.Styles["AuditNewParagraph"].Font.Underline == Word.WdUnderline.wdUnderlineSingle, "newly added styles are created and applied to Word");
            doc.Undo();
            bool removedByUndo = false;
            try { var ignored = doc.Styles["AuditNewParagraph"]; } catch (System.Runtime.InteropServices.COMException) { removedByUndo = true; }
            Check(removedByUndo, "apply current is a single undo step");
            // A broken relation must be reported as a failure, never as success.
            Select(form, created.NameLocal);
            Models(form).First(x => x.Name == created.NameLocal).NextParagraphStyle = "MissingAuditStyle";
            Click(form, "应用当前_Click");
            Check(Field<ToolStripStatusLabel>(form, "statusLabel").Text.Contains("出错"), "COM application failures are surfaced in the status bar");
            Models(form).First(x => x.Name == created.NameLocal).NextParagraphStyle = normalName;
            Select(form, normalName);
            Click(form, "Btn_ApplySet_Click");
            Check(!Field<ToolStripStatusLabel>(form, "statusLabel").Text.Contains("出错"), "single-selection apply all updates the visible style collection: " + Field<ToolStripStatusLabel>(form, "statusLabel").Text);
            Field<StandardComboBox>(form, "Cmb_FontSize").Text = "14磅";
            form.RefreshFromDocument(true);
            Check(Models(form).First(x => x.Name == normalName).FontSize == 14, "reopening same document retains edits");
            var second = app.Documents.Add();
            bool blocked = false;
            try { Call(form, "GetTargetDocument"); } catch (InvalidOperationException) { blocked = true; }
            Check(blocked, "document switch blocks applying stale editor values");
            form.RefreshFromDocument(true);
            Check(Call(form, "GetTargetDocument") != null, "refresh permits editing the new document");
            second.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Console.WriteLine("Completed " + checks + " checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (form != null) form.Dispose();
            if (app != null) app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }
}
