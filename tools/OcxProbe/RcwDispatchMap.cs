using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace OcxProbe
{
    internal static class RcwDispatchMap
    {
        private static readonly string[] HotMethods = new[]
        {
            "Acknowledge","AckSelected","Shelve","Unshelve","Refresh","LoadMessages","ApplyFilter"
        };

        private static string ResolveOutDir()
        {
            var env = Environment.GetEnvironmentVariable("OCXPROBE_OUTDIR");
            if (!string.IsNullOrWhiteSpace(env))
            {
                try { Directory.CreateDirectory(env); return env; } catch { }
            }
            return AppDomain.CurrentDomain.BaseDirectory;
        }

        public static (string jsonPath, string csvPath) Export(object comObj, string fileStem = null)
        {
            if (comObj == null) throw new ArgumentNullException(nameof(comObj));

            var t = comObj.GetType();
            var map = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var itf in t.GetInterfaces())
            {
                foreach (var m in itf.GetMethods(BindingFlags.Public | BindingFlags.Instance))
                {
                    var disp = (DispIdAttribute)Attribute.GetCustomAttribute(m, typeof(DispIdAttribute));
                    if (disp != null) map[m.Name] = disp.Value;
                }
                foreach (var p in itf.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    var get = p.GetGetMethod();
                    var set = p.GetSetMethod();
                    var disp = (DispIdAttribute)(get != null
                        ? Attribute.GetCustomAttribute(get, typeof(DispIdAttribute))
                        : (set != null ? Attribute.GetCustomAttribute(set, typeof(DispIdAttribute)) : null));
                    if (disp != null) map["prop:" + p.Name] = disp.Value;
                }
            }

            var baseDir = ResolveOutDir();
            var stem = string.IsNullOrWhiteSpace(fileStem) ? t.Name + "_map" : fileStem;
            var jsonPath = Path.Combine(baseDir, stem + ".json");
            var csvPath = Path.Combine(baseDir, stem + ".csv");

            var sb = new StringBuilder();
            sb.AppendLine("{");
            int i = 0, n = map.Count;
            foreach (var kv in map)
            {
                sb.Append("  \"").Append(EscapeJson(kv.Key)).Append("\": ").Append(kv.Value);
                if (++i < n) sb.Append(",");
                sb.AppendLine();
            }
            var presentHots = HotMethods.Where(m => map.ContainsKey(m)).ToArray();
            if (presentHots.Length > 0)
            {
                sb.Append("  ,\"__hot__\": [");
                for (int k = 0; k < presentHots.Length; k++)
                {
                    if (k > 0) sb.Append(", ");
                    sb.Append("\"").Append(EscapeJson(presentHots[k])).Append("\"");
                }
                sb.AppendLine("]");
            }
            sb.AppendLine("}");
            File.WriteAllText(jsonPath, sb.ToString(), Encoding.UTF8);

            var lines = new List<string> { "name;dispid" };
            lines.AddRange(map.Select(kv => $"{kv.Key};{kv.Value}"));
            File.WriteAllLines(csvPath, lines, Encoding.UTF8);

            return (jsonPath, csvPath);
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        public static (string json, string csv) ExportSummary()
            => Export(new FTAlarmEventSummary.AlarmEventSummaryClass(), "ftae_summary_map");

        public static (string json, string csv) ExportBanner()
            => Export(new FTAlarmEventSummary.AlarmEventBannerClass(), "ftae_banner_map");
    }
}
