using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using WordMan.MultiLevel;

namespace WordMan
{
    public class MultiLevelPreset
    {
        public string Name { get; set; }
        public int LevelCount { get; set; }
        public List<LevelData> Levels { get; set; } = new List<LevelData>();
        public List<WordStyleInfo> Styles { get; set; } = new List<WordStyleInfo>();
    }

    /// <summary>命名预设与文本导入导出独立；读取和保存均使用副本，避免编辑污染预设。</summary>
    public sealed class MultiLevelPresetManager
    {
        private readonly string filePath;
        private static readonly string[] BuiltInNames = { "默认风格", "多级数字", "中文层级" };

        public MultiLevelPresetManager() : this(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordMan", "MultiLevelPresets.xml")) { }

        public MultiLevelPresetManager(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("预设文件路径不能为空。", nameof(filePath));
            this.filePath = Path.GetFullPath(filePath);
        }

        public static bool IsBuiltIn(string name) { return BuiltInNames.Contains(name, StringComparer.Ordinal); }
        public static string[] GetBuiltInNames() { return (string[])BuiltInNames.Clone(); }

        public List<string> GetPresetNames()
        {
            return BuiltInNames.Concat(ReadStored().Select(x => x.Name)).Distinct(StringComparer.Ordinal).ToList();
        }

        public bool HasSavedPreset(string name) { return ReadStored().Any(x => x.Name == name); }

        public MultiLevelPreset LoadPreset(string name)
        {
            name = NormalizeName(name);
            var saved = ReadStored().FirstOrDefault(x => x.Name == name);
            if (saved != null) return Snapshot(saved);
            if (!IsBuiltIn(name)) throw new ArgumentException("找不到预设：" + name);
            return CreateBuiltIn(name);
        }

        public void SavePreset(string name, IEnumerable<LevelData> levels, int count, IEnumerable<WordStyleInfo> styles)
        {
            var preset = Snapshot(new MultiLevelPreset
            {
                Name = NormalizeName(name), LevelCount = count,
                Levels = levels == null ? null : levels.ToList(),
                Styles = styles == null ? new List<WordStyleInfo>() : styles.ToList()
            });
            var stored = ReadStored();
            stored.RemoveAll(x => x.Name == preset.Name);
            stored.Add(preset);
            WriteStored(stored);
        }

        // 内置预设只删除用户覆盖，保留原始默认配置；自定义预设直接删除。
        public bool DeletePreset(string name)
        {
            name = NormalizeName(name);
            var stored = ReadStored();
            if (stored.RemoveAll(x => x.Name == name) == 0) return false;
            WriteStored(stored);
            return true;
        }

        public static List<LevelData> CreateDefaultLevels()
        {
            string[] formats = { "第%1章", "%1.%2", "%1.%2.%3", "%4.", "(%5)", "%6.", "(%7)", "%8.", "(%9)" };
            return Enumerable.Range(1, 9).Select(level => new LevelData
            {
                Level = level, NumberFormat = formats[level - 1],
                NumberStyle = level >= 8 ? "a,b,c..." : level >= 6 ? "A,B,C..." : "1,2,3...",
                NumberIndent = level >= 4 ? 0.8m : 0m, TextIndent = 0m,
                AfterNumberType = "空格", TabPosition = 0m, LinkedStyle = "标题 " + level,
                StartAt = 1, ResetOnHigher = -1
            }).ToList();
        }

        private static MultiLevelPreset CreateBuiltIn(string name)
        {
            var levels = CreateDefaultLevels();
            if (name == "多级数字")
                foreach (var level in levels)
                {
                    level.NumberStyle = "1,2,3...";
                    level.NumberFormat = string.Join(".", Enumerable.Range(1, level.Level).Select(i => "%" + i));
                    level.NumberIndent = 0m;
                }
            if (name == "中文层级")
            {
                levels = levels.Take(4).ToList();
                string[] formats = { "%1、", "（%2）", "%3.", "（%4）" };
                foreach (var level in levels)
                {
                    level.NumberFormat = formats[level.Level - 1];
                    level.NumberStyle = level.Level <= 2 ? "一,二,三..." : "1,2,3...";
                    level.NumberIndent = 0m;
                }
            }
            return Snapshot(new MultiLevelPreset { Name = name, LevelCount = levels.Count, Levels = levels });
        }

        private static string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("预设名称不能为空。");
            return name.Trim();
        }

        private static MultiLevelPreset Snapshot(MultiLevelPreset preset)
        {
            if (preset == null) throw new FormatException("预设数据不能为空。");
            string name = NormalizeName(preset.Name);
            List<string> errors;
            if (preset.LevelCount < 1 || preset.LevelCount > 9 || preset.Levels == null ||
                preset.Levels.Count != preset.LevelCount || !MultiLevelDataManager.ValidateLevelDataList(preset.Levels, out errors))
                throw new FormatException("预设“" + name + "”的级别数据无效，请设置 1 至 9 级有效配置。");
            if (preset.Styles == null || preset.Styles.Count > 9 ||
                preset.Styles.Any(x => x == null || string.IsNullOrWhiteSpace(x.StyleName)) ||
                preset.Styles.Select(x => x.StyleName).Distinct(StringComparer.OrdinalIgnoreCase).Count() != preset.Styles.Count)
                throw new FormatException("预设“" + name + "”的每级样式数据无效。");
            foreach (var style in preset.Styles) style.Validate();
            return new MultiLevelPreset
            {
                Name = name, LevelCount = preset.LevelCount,
                Levels = preset.Levels.OrderBy(x => x.Level).Select(x => new LevelData
                {
                    Level = x.Level, NumberStyle = x.NumberStyle, NumberFormat = x.NumberFormat,
                    NumberIndent = x.NumberIndent, TextIndent = x.TextIndent, AfterNumberType = x.AfterNumberType,
                    TabPosition = x.TabPosition, LinkedStyle = x.LinkedStyle, StartAt = x.StartAt, ResetOnHigher = x.ResetOnHigher
                }).ToList(),
                Styles = preset.Styles.Select(x => x.Clone()).ToList()
            };
        }

        private List<MultiLevelPreset> ReadStored()
        {
            if (!File.Exists(filePath)) return new List<MultiLevelPreset>();
            // 不吞掉损坏文件错误，防止随后保存静默丢失原预设。
            using (var reader = XmlReader.Create(filePath, new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            {
                var stored = (List<MultiLevelPreset>)new XmlSerializer(typeof(List<MultiLevelPreset>)).Deserialize(reader);
                if (stored == null) throw new FormatException("预设文件内容为空。");
                var result = stored.Select(Snapshot).ToList();
                if (result.Select(x => x.Name).Distinct(StringComparer.Ordinal).Count() != result.Count)
                    throw new FormatException("预设文件包含重复名称。");
                return result;
            }
        }

        private void WriteStored(List<MultiLevelPreset> presets)
        {
            string directory = Path.GetDirectoryName(filePath);
            Directory.CreateDirectory(directory);
            string temporary = Path.Combine(directory, ".MultiLevelPresets-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = File.Create(temporary))
                    new XmlSerializer(typeof(List<MultiLevelPreset>)).Serialize(stream, presets);
                if (File.Exists(filePath)) File.Replace(temporary, filePath, null);
                else File.Move(temporary, filePath);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
