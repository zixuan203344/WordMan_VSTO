using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Windows.Forms;
using WordMan;
using WordMan.MultiLevel;
using Word = Microsoft.Office.Interop.Word;

// Source-only regression harness. This change does not compile or execute it.
// The optional --word checks use an independent Word instance and disposable documents.
internal static class MultiLevelListRegression
{
    private static int checks;
    private static readonly BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
    private static void Check(bool value, string name)
    {
        if (!value) throw new Exception("FAIL: " + name);
        checks++;
        Console.WriteLine("PASS: " + name);
    }
    private static T Field<T>(object target, string name)
    { return (T)target.GetType().GetField(name, Private).GetValue(target); }
    private static object Call(object target, string name, params object[] args)
    {
        try { return target.GetType().GetMethod(name, Private).Invoke(target, args); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    private static void Reject(Action action, string name)
    {
        bool rejected = false;
        try { action(); } catch (ArgumentException) { rejected = true; } catch (FormatException) { rejected = true; }
        Check(rejected, name);
    }
    private static LevelData Level(int value)
    {
        return new LevelData
        {
            Level = value, NumberStyle = "1,2,3...", NumberFormat = MultiLevelDataManager.GenerateNumberFormat(value),
            NumberIndent = 0.25m, TextIndent = 0.5m, TabPosition = 0.75m,
            AfterNumberType = "制表位", LinkedStyle = "标题 " + value
        };
    }

    private static void CheckPresets(string folder)
    {
        string path = Path.Combine(folder, "named-presets.xml");
        var manager = new MultiLevelPresetManager(path);
        var builtIns = MultiLevelPresetManager.GetBuiltInNames();
        Check(manager.GetPresetNames().SequenceEqual(builtIns), "missing preset file exposes built-ins without writing a file");
        Check(!File.Exists(path), "listing default presets is read-only");
        foreach (var name in builtIns)
        {
            var preset = manager.LoadPreset(name);
            List<string> errors;
            Check(MultiLevelDataManager.ValidateLevelDataList(preset.Levels, out errors), "built-in validates: " + name);
            preset.Levels[0].NumberFormat = "changed %1";
            Check(manager.LoadPreset(name).Levels[0].NumberFormat != "changed %1", "built-in snapshots are independent: " + name);
        }
        var levels = new List<LevelData> { Level(1), Level(2) };
        levels[0].StartAt = 7;
        levels[1].ResetOnHigher = 0;
        levels[1].LinkedStyle = "Custom;\"Style\"";
        var style = new WordStyleInfo("标题 1") { FontColor = Color.FromArgb(12, 34, 56), Bold = true };
        manager.SavePreset("  自定义  ", levels, 2, new[] { style });
        var reopened = new MultiLevelPresetManager(path);
        var saved = reopened.LoadPreset("自定义");
        Check(saved.LevelCount == 2 && saved.Levels[0].StartAt == 7 && saved.Levels[1].ResetOnHigher == 0 &&
            saved.Levels[1].LinkedStyle == levels[1].LinkedStyle, "named preset round-trips levels, restart and custom style links");
        Check(saved.Styles.Single().FontColor.ToArgb() == style.FontColor.ToArgb() && saved.Styles.Single().Bold,
            "named preset round-trips per-level formatting and color");
        levels[0].NumberFormat = "第%1节";
        style.Bold = false;
        Check(saved.Levels[0].NumberFormat != levels[0].NumberFormat && saved.Styles[0].Bold,
            "saving and editing input objects cannot mutate stored snapshots");
        saved.Levels[0].NumberFormat = "changed %1";
        saved.Styles[0].Bold = false;
        Check(reopened.LoadPreset("自定义").Levels[0].NumberFormat != "changed %1" && reopened.LoadPreset("自定义").Styles[0].Bold,
            "loading returns independent level and style snapshots");
        manager.SavePreset("自定义", new[] { Level(1) }, 1, null);
        Check(manager.GetPresetNames().Count(x => x == "自定义") == 1 && manager.LoadPreset("自定义").LevelCount == 1 &&
            manager.LoadPreset("自定义").Styles.Count == 0, "same-name save replaces whole preset without stale levels or styles");
        string originalFormat = manager.LoadPreset(builtIns[0]).Levels[0].NumberFormat;
        manager.SavePreset(builtIns[0], new[] { Level(1) }, 1, null);
        Check(manager.LoadPreset(builtIns[0]).LevelCount == 1 && manager.HasSavedPreset(builtIns[0]) &&
            manager.GetPresetNames().Count(x => x == builtIns[0]) == 1, "built-in names support a saved override without duplicate entries");
        Check(manager.DeletePreset(builtIns[0]) && !manager.HasSavedPreset(builtIns[0]) &&
            manager.LoadPreset(builtIns[0]).Levels[0].NumberFormat == originalFormat,
            "deleting built-in override restores underlying defaults");
        Check(!manager.DeletePreset(builtIns[0]) && manager.GetPresetNames().Contains(builtIns[0]), "original built-ins remain available");
        string before = File.ReadAllText(path);
        Reject(() => manager.SavePreset(" ", new[] { Level(1) }, 1, null), "blank preset name rejected");
        Reject(() => manager.SavePreset("bad", new LevelData[0], 0, null), "zero-level preset rejected");
        Reject(() => manager.SavePreset("bad", new[] { Level(1) }, 2, null), "mismatched level count rejected");
        var invalid = Level(1); invalid.NumberFormat = "%2";
        Reject(() => manager.SavePreset("bad", new[] { invalid }, 1, null), "invalid numbering format rejected before saving");
        Reject(() => manager.SavePreset("bad", new[] { Level(1) }, 1, new WordStyleInfo[] { null }), "null style rejected before saving");
        Reject(() => manager.SavePreset("bad", new[] { Level(1) }, 1, new[] { style, style.Clone() }), "duplicate style names rejected");
        var invalidStyle = style.Clone(); invalidStyle.FontSize = "bad";
        Reject(() => manager.SavePreset("bad", new[] { Level(1) }, 1, new[] { invalidStyle }), "invalid per-level formatting rejected");
        Check(File.ReadAllText(path) == before, "rejected saves preserve existing preset file");
        var hiddenStyle = new WordStyleInfo("标题 9") { Italic = true };
        manager.SavePreset("含隐藏样式", new[] { Level(1) }, 1, new[] { style, hiddenStyle });
        Check(new MultiLevelPresetManager(path).LoadPreset("含隐藏样式").Styles.Last().Italic,
            "hidden higher-level style settings survive a lower-level preset save");
        Check(manager.DeletePreset("自定义") && !manager.GetPresetNames().Contains("自定义"), "custom preset deletion persists");
        Reject(() => manager.LoadPreset("自定义"), "deleted preset is no longer loadable");
        File.WriteAllText(path, "<broken>");
        bool rejected = false;
        try { manager.SavePreset("new", new[] { Level(1) }, 1, null); }
        catch (System.Xml.XmlException) { rejected = true; }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected && File.ReadAllText(path) == "<broken>", "corrupt preset file cannot be silently overwritten");
        Check(!Directory.GetFiles(folder, ".MultiLevelPresets-*.tmp").Any(), "preset writes leave no temporary files");
    }

    [STAThread]
    private static int Main(string[] args)
    {
        string folder = Path.Combine(Path.GetTempPath(), "WordManMultiLevel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        Word.Application app = null;
        object previousAddin = null;
        FieldInfo globals = null;
        try
        {
            CheckPresets(folder);
            float size;
            Check(MultiLevelDataManager.TryParseFontSize("5", out size) && size == 5, "numeric size is points, not a Chinese size index");
            Check(MultiLevelDataManager.TryParseFontSize("五号", out size) && size == 10.5f, "Chinese font size remains supported");
            Check(MultiLevelDataManager.ConvertFontSizeToString(13.5f) == "13.5 磅", "fractional size survives display");
            foreach (var text in new[] { "NaN", "Infinity", "-1", "0", "1639", "bad" })
                Check(!MultiLevelDataManager.TryParseFontSize(text, out size), "invalid font size rejected: " + text);
            foreach (var text in new[] { "%0", "%10", "%2", "%1%", "%1\\n" })
                Check(!MultiLevelDataManager.ValidateNumberFormat(text.Replace("\\n", "\n"), 1), "invalid numbering placeholder rejected: " + text);
            Check(MultiLevelDataManager.ValidateNumberFormat("第%1章.%2", 2), "ancestor and current level placeholders accepted");
            List<string> errors;
            Check(!MultiLevelDataManager.ValidateLevelDataList(new List<LevelData> { Level(2) }, out errors), "missing first level is rejected");
            var duplicate = Level(2); duplicate.LinkedStyle = "标题 1";
            Check(!MultiLevelDataManager.ValidateLevelDataList(new List<LevelData> { Level(1), duplicate }, out errors), "duplicate linked styles are rejected");
            Check(WordStyleInfo.IsValidLineSpacing("1.25倍行距") && WordStyleInfo.IsValidLineSpacing("18磅") &&
                WordStyleInfo.IsValidLineSpacing("最小值18磅") && !WordStyleInfo.IsValidLineSpacing("NaN磅"), "line spacing validates finite values and retains rule");
            var style = new WordStyleInfo("标题 1"); style.FontColor = Color.FromArgb(12, 34, 56);
            Check(style != null && (WordStyleInfo)null == null && new WordStyleInfo("标题 1") == style, "style null and equality checks cannot recurse");
            var clone = style.Clone(); clone.Bold = !style.Bold;
            Check(style.Bold != clone.Bold, "editing a cloned style cannot alter parent data");
            var levels = new List<LevelData> { Level(1), Level(2), Level(3) };
            levels[0].NumberFormat = "第%1;章";
            levels[0].StartAt = 7;
            levels[1].ResetOnHigher = 0;
            levels[2].LinkedStyle = "Custom;\"Style\"";
            string file = Path.Combine(folder, "roundtrip.txt");
            ConfigurationManager.SaveConfigurationToFile(file, levels, 3, new List<WordStyleInfo> { style });
            List<LevelData> restored; List<WordStyleInfo> restoredStyles; int count;
            ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count, out restoredStyles);
            Check(count == 3 && restored[0].StartAt == 7 && restored[1].ResetOnHigher == 0, "start and restart rules survive export/import");
            Check(restored[0].NumberFormat == levels[0].NumberFormat && restored[2].LinkedStyle == levels[2].LinkedStyle, "semicolons and quotes survive export/import");
            Check(restoredStyles.Single().FontColor.ToArgb() == style.FontColor.ToArgb(), "per-level style color survives export/import");
            ConfigurationManager.SaveConfigurationToFile(file, new List<LevelData>(), 0);
            ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count);
            Check(count == 0 && restored.Count == 0, "zero-level configuration round trips");
            string header = "列表级数;2\n级别;编号样式;编号格式;编号缩进;文本缩进;编号之后类型;制表位位置;链接样式\n";
            string first = "1;1,2,3...;%1;0.25;0.5;空格;0;标题 1\n";
            string second = "2;1,2,3...;%1.%2;0.5;0.75;制表位;1;标题 2\n";
            File.WriteAllText(file, header + second + first);
            ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count);
            Check(restored[0].Level == 1 && restored[1].Level == 2 && restored[1].ResetOnHigher == -1, "legacy eight-column rows are sorted and default restart is preserved");
            File.WriteAllText(file, header + first + first);
            Reject(() => ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count), "duplicate imported levels rejected");
            File.WriteAllText(file, header + first);
            Reject(() => ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count), "missing imported level rejected");
            foreach (int invalid in new[] { -1, 10 })
            {
                File.WriteAllText(file, "列表级数;" + invalid + "\nheader\n");
                Reject(() => ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count), "invalid count rejected: " + invalid);
            }
            var previousCulture = CultureInfo.CurrentCulture;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
                File.WriteAllText(file, header + first.Replace("0.25", "0,25") + second);
                ConfigurationManager.LoadConfigurationFromFile(file, out restored, out count);
                Check(restored[0].NumberIndent == 0.25m, "legacy decimal comma is not interpreted as a thousands separator");
            }
            finally { System.Threading.Thread.CurrentThread.CurrentCulture = previousCulture; }
            if (!args.Contains("--word")) { Console.WriteLine("Completed " + checks + " checks; Word checks not requested."); return 0; }
            Application.EnableVisualStyles();
            app = new Word.Application(); app.Visible = false; app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            var addinType = typeof(MultiLevelListForm).Assembly.GetType("WordMan.ThisAddIn");
            var addin = FormatterServices.GetUninitializedObject(addinType);
            addinType.GetField("Application", Private).SetValue(addin, app);
            globals = typeof(MultiLevelListForm).Assembly.GetType("WordMan.Globals").GetField("_ThisAddIn", BindingFlags.Static | BindingFlags.NonPublic);
            previousAddin = globals.GetValue(null); globals.SetValue(null, addin);
            var doc = app.Documents.Add();
            using (var form = new MultiLevelListForm())
            {
                var presetCombo = Field<ComboBox>(form, "cmbPresetStyle");
                Check(presetCombo.Items.Contains("默认风格") && !Field<StandardButton>(form, "btnLoadPreset").Enabled,
                    "preset toolbar starts with choices without replacing current editor");
                using (var saveDialog = new SavePresetDialog(new[] { "默认风格", "自定义" }, "自定义"))
                {
                    Check(saveDialog.PresetName == "自定义" && Field<ComboBox>(saveDialog, "cmbName").Items.Count == 2,
                        "shared save dialog uses multi-level names and current selection");
                }
                var levelCount = Field<StandardComboBox>(form, "cmbLevelCount");
                var defaultPreset = Field<MultiLevelPresetManager>(form, "presetManager").LoadPreset("默认风格");
                string documentText = doc.Content.Text;
                presetCombo.SelectedItem = "默认风格";
                Check(levelCount.GetSelectedText() == defaultPreset.LevelCount.ToString() && doc.Content.Text == documentText,
                    "selecting preset populates editor without writing to document");
                var firstFormat = (StandardTextBox)Field<Panel>(form, "levelsContainer").Controls.Find("TextBoxNumFormat1", true).Single();
                firstFormat.Text = "changed %1";
                Call(form, "CmbPresetStyle_SelectedIndexChanged", Field<StandardButton>(form, "btnLoadPreset"), EventArgs.Empty);
                Check(Field<List<LevelData>>(form, "levelDataList")[0].NumberFormat == defaultPreset.Levels[0].NumberFormat,
                    "load button reloads same preset after edits");
                Call(form, "ClearPresetSelection");
                Check(presetCombo.SelectedIndex == -1 && !Field<StandardButton>(form, "btnLoadPreset").Enabled &&
                    !Field<StandardButton>(form, "btnDeletePreset").Enabled, "clearing preset selection resets dependent buttons");
                levelCount.SelectedItem = "3";
                var container = Field<Panel>(form, "levelsContainer");
                var number = (StandardNumericUpDown)container.Controls.Find("TxtBoxNumIndent1", true).Single();
                number.SetValueInCentimeters(1m);
                levelCount.SelectedItem = "1"; levelCount.SelectedItem = "3";
                Check(Field<List<LevelData>>(form, "levelDataList")[0].NumberIndent == 1m, "changing level count retains edits");
                Field<CheckBox>(form, "chkNumberIndent").Checked = true;
                Field<CheckBox>(form, "chkTextIndent").Checked = true;
                Field<CheckBox>(form, "chkNumberIndent").Checked = false;
                Check(!Field<StandardNumericUpDown>(form, "nudNumberIndent").Enabled && Field<StandardNumericUpDown>(form, "nudTextIndent").Enabled, "each quick-setting checkbox controls its own input");
                Field<CheckBox>(form, "chkTabPosition").Checked = true;
                Field<StandardNumericUpDown>(form, "nudTabPosition").Value = 0;
                Call(form, "ApplyQuickSettings");
                Check(Field<List<LevelData>>(form, "levelDataList").Take(3).All(x => x.AfterNumberType == "制表位"), "quick tab setting also selects a tab suffix, including zero position");
                ConfigurationManager.SaveConfigurationToFile(file, new List<LevelData> { Level(1) }, 1);
                Call(form, "LoadConfigurationFromFile", file);
                levelCount.SelectedItem = "9";
                Check(Field<List<LevelData>>(form, "levelDataList").Count == 9, "importing one level then increasing to nine never indexes past data");
                foreach (var value in new[] { Word.WdListNumberStyle.wdListNumberStyleLegalLZ, Word.WdListNumberStyle.wdListNumberStyleSimpChinNum1,
                    Word.WdListNumberStyle.wdListNumberStyleSimpChinNum2, Word.WdListNumberStyle.wdListNumberStyleZodiac1 })
                    Check((int)Call(form, "GetNumberStyleIndex", value) >= 0, "loaded numbering type has the same mapping: " + value);
                using (var editor = new LevelStyleSettingsForm(3))
                {
                    var originals = editor.GetLevelStyles();
                    editor.LoadExistingStyles(originals);
                    var grid = Field<DataGridView>(editor, "Dta_StyleList");
                    grid.ClearSelection(); grid.Rows[0].Selected = true;
                    bool oldBold = originals[0].Bold;
                    Field<ToggleButton>(editor, "Btn_Bold").Pressed = !oldBold;
                    Check(originals[0].Bold == oldBold && editor.GetLevelStyles()[0].Bold != oldBold, "cancelable child editing does not mutate existing styles");
                }
                var other = app.Documents.Add();
                bool blocked = false;
                try { Call(form, "GetTargetDocument"); } catch (InvalidOperationException) { blocked = true; }
                Check(blocked, "modeless editor cannot apply stale settings to a different document");
                other.Close(Word.WdSaveOptions.wdDoNotSaveChanges); doc.Activate();
                doc.Content.Text = "One\rTwo\rThree\r";
                var template = doc.ListTemplates.Add(true);
                doc.Paragraphs[1].Range.ListFormat.ApplyListTemplateWithLevel(template, false, Word.WdListApplyTo.wdListApplyToSelection, Word.WdDefaultListBehavior.wdWord9ListBehavior, 1);
                doc.Paragraphs[2].Range.ListFormat.ApplyListTemplateWithLevel(template, true, Word.WdListApplyTo.wdListApplyToSelection, Word.WdDefaultListBehavior.wdWord9ListBehavior, 2);
                levelCount.SelectedItem = "3";
                for (int i = 1; i <= 3; i++) ((StandardComboBox)container.Controls.Find("CmbLinkedStyle" + i, true).Single()).SelectedItem = "无";
                doc.Range(doc.Paragraphs[1].Range.Start, doc.Paragraphs[2].Range.End).Select();
                Call(form, "BtnApply_Click", null, EventArgs.Empty);
                Check(doc.Paragraphs[1].Range.ListFormat.ListLevelNumber == 1 && doc.Paragraphs[2].Range.ListFormat.ListLevelNumber == 2, "applying keeps selected paragraphs at their own levels");
                doc.Undo();
                Check(doc.Paragraphs[2].Range.ListFormat.ListLevelNumber == 2, "list application is one undo operation");
            }
            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Console.WriteLine("Completed " + checks + " checks."); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (globals != null) globals.SetValue(null, previousAddin);
            if (app != null) app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }
}
