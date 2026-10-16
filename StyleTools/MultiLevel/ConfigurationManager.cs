using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using WordMan.MultiLevel;

namespace WordMan
{
    /// <summary>兼容原有分号分隔配置，并完整保存编号起始值和每级样式。</summary>
    public static class ConfigurationManager
    {
        public static void SaveConfigurationToFile(string filePath, List<LevelData> levels, int count)
        {
            SaveConfigurationToFile(filePath, levels, count, null);
        }

        public static void SaveConfigurationToFile(string filePath, List<LevelData> levels, int count,
            List<WordStyleInfo> styles)
        {
            if (count < 0 || count > 9 || levels == null) throw new ArgumentException("列表级数必须在 0 至 9 之间");
            var active = levels.Where(x => x != null && x.Level <= count).OrderBy(x => x.Level).ToList();
            List<string> errors;
            if (count > 0 && (active.Count != count || !MultiLevelDataManager.ValidateLevelDataList(active, out errors)))
                throw new ArgumentException("级别数据不完整或包含无效值，请检查后再导出。");
            var content = new StringBuilder();
            content.AppendLine("列表级数;" + count);
            content.AppendLine("级别;编号样式;编号格式;编号缩进;文本缩进;编号之后类型;制表位位置;链接样式;起始编号;重新编号级别");
            foreach (var level in active)
                content.AppendLine(string.Join(";", new[]
                {
                    level.Level.ToString(CultureInfo.InvariantCulture), Escape(level.NumberStyle), Escape(level.NumberFormat),
                    level.NumberIndent.ToString(CultureInfo.InvariantCulture), level.TextIndent.ToString(CultureInfo.InvariantCulture),
                    Escape(level.AfterNumberType), level.TabPosition.ToString(CultureInfo.InvariantCulture), Escape(level.LinkedStyle),
                    level.StartAt.ToString(CultureInfo.InvariantCulture), level.ResetOnHigher.ToString(CultureInfo.InvariantCulture)
                }));
            if (styles != null && styles.Count > 0)
            {
                foreach (var style in styles)
                {
                    if (style == null) throw new ArgumentException("每级样式不能为空。");
                    style.Validate();
                }
                using (var buffer = new MemoryStream())
                {
                    new XmlSerializer(typeof(List<WordStyleInfo>)).Serialize(buffer, styles);
                    content.AppendLine("每级样式;" + Convert.ToBase64String(buffer.ToArray()));
                }
            }
            // 完整写入临时文件后再替换目标，防止写入失败破坏已有配置。
            string target = Path.GetFullPath(filePath);
            string temporary = Path.Combine(Path.GetDirectoryName(target), ".WordMan-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllText(temporary, content.ToString(), Encoding.UTF8);
                if (File.Exists(target)) File.Replace(temporary, target, null);
                else File.Move(temporary, target);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        public static void LoadConfigurationFromFile(string filePath, out List<LevelData> levels, out int count)
        {
            List<WordStyleInfo> styles;
            LoadConfigurationFromFile(filePath, out levels, out count, out styles);
        }

        public static void LoadConfigurationFromFile(string filePath, out List<LevelData> levels, out int count,
            out List<WordStyleInfo> styles)
        {
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            int parsedCount;
            if (lines.Length < 2 || !lines[0].StartsWith("列表级数;", StringComparison.Ordinal) ||
                !int.TryParse(lines[0].Substring("列表级数;".Length), out parsedCount) || parsedCount < 0 || parsedCount > 9)
                throw new FormatException("配置文件的列表级数必须在 0 至 9 之间。");
            var loaded = new List<LevelData>();
            var loadedStyles = new List<WordStyleInfo>();
            bool hasStyles = false;
            for (int row = 2; row < lines.Length; row++)
            {
                if (string.IsNullOrWhiteSpace(lines[row])) continue;
                if (lines[row].StartsWith("每级样式;", StringComparison.Ordinal))
                {
                    if (hasStyles) throw new FormatException("每级样式配置重复。");
                    hasStyles = true;
                    byte[] xml = Convert.FromBase64String(lines[row].Substring("每级样式;".Length));
                    using (var stream = new MemoryStream(xml))
                    using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings
                    { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null }))
                        loadedStyles = (List<WordStyleInfo>)new XmlSerializer(typeof(List<WordStyleInfo>)).Deserialize(reader);
                    continue;
                }
                var fields = SplitRecord(lines[row]);
                if (fields.Count != 8 && fields.Count != 10) throw new FormatException("配置第 " + (row + 1) + " 行列数无效。");
                loaded.Add(new LevelData
                {
                    Level = int.Parse(fields[0], CultureInfo.InvariantCulture), NumberStyle = fields[1], NumberFormat = fields[2],
                    NumberIndent = ParseDecimal(fields[3]), TextIndent = ParseDecimal(fields[4]), AfterNumberType = fields[5],
                    TabPosition = ParseDecimal(fields[6]), LinkedStyle = fields[7],
                    StartAt = fields.Count == 10 ? int.Parse(fields[8], CultureInfo.InvariantCulture) : 1,
                    ResetOnHigher = fields.Count == 10 ? int.Parse(fields[9], CultureInfo.InvariantCulture) : -1
                });
            }
            List<string> errors;
            if (loaded.Count != parsedCount || (parsedCount > 0 && !MultiLevelDataManager.ValidateLevelDataList(loaded, out errors)))
                throw new FormatException("级别必须从 1 开始连续，数量须与列表级数一致，且格式、缩进和链接样式必须有效。");
            if (loadedStyles == null || loadedStyles.Count > 9 || loadedStyles.Any(x => x == null || string.IsNullOrWhiteSpace(x.StyleName)) ||
                loadedStyles.Select(x => x.StyleName).Distinct(StringComparer.OrdinalIgnoreCase).Count() != loadedStyles.Count)
                throw new FormatException("每级样式数据无效。");
            foreach (var style in loadedStyles) style.Validate();
            levels = loaded.OrderBy(x => x.Level).ToList();
            count = parsedCount;
            styles = loadedStyles;
        }

        private static decimal ParseDecimal(string text)
        {
            decimal value;
            // 兼容旧文件的逗号小数；禁止把它误作千位分隔符。
            if (decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                decimal.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                decimal.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return value;
            throw new FormatException("无效数值：" + text);
        }

        private static string Escape(string value)
        {
            value = value ?? "";
            if (value.IndexOfAny(new[] { '\r', '\n' }) >= 0) throw new FormatException("配置字段不能包含换行。");
            return value.IndexOfAny(new[] { ';', '"' }) < 0 ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static List<string> SplitRecord(string line)
        {
            var fields = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < line.Length; i++)
            {
                char current = line[i];
                if (current == '"' && (quoted || field.Length == 0))
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"') { field.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (current == ';' && !quoted) { fields.Add(field.ToString()); field.Clear(); }
                else field.Append(current);
            }
            if (quoted) throw new FormatException("配置字段引号不匹配。");
            fields.Add(field.ToString());
            return fields;
        }

        public static string ShowImportDialog(IWin32Window owner = null)
        {
            using (var dialog = new OpenFileDialog { Filter = "TXT文件|*.txt|所有文件|*.*", Title = "导入多级列表配置", DefaultExt = "txt" })
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : string.Empty;
        }

        public static string ShowExportDialog(IWin32Window owner = null)
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "TXT文件|*.txt|所有文件|*.*", Title = "导出多级列表配置", DefaultExt = "txt",
                FileName = "多级列表配置_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
            })
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : string.Empty;
        }
    }
}
