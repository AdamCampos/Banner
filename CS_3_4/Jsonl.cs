using System;
using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ComHookLib
{
    internal static class Jsonl
    {
        // Profundidade máxima para evitar ciclos acidentais
        private const int MaxDepth = 6;

        public static string Serialize(object obj)
        {
            try
            {
                var sb = new StringBuilder(256);
                WriteValue(sb, obj, 0);
                return sb.ToString();
            }
            catch
            {
                // Fallback mínimo para não perder logs
                return "{\"error\":\"json_serialize_failed\"}";
            }
        }

        private static void WriteValue(StringBuilder sb, object value, int depth)
        {
            if (depth > MaxDepth)
            {
                sb.Append("\"#depth\"");
                return;
            }

            if (value == null)
            {
                sb.Append("null");
                return;
            }

            // Tipos primários rápidos
            var t = value.GetType();

            if (value is string s)
            {
                WriteString(sb, s);
                return;
            }

            if (value is bool b)
            {
                sb.Append(b ? "true" : "false");
                return;
            }

            if (IsNumber(t))
            {
                sb.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
                return;
            }

            if (value is DateTime dt)
            {
                // ISO 8601
                sb.Append('"').Append(dt.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)).Append('"');
                return;
            }

            if (value is DateTimeOffset dto)
            {
                sb.Append('"').Append(dto.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)).Append('"');
                return;
            }

            if (value is Guid g)
            {
                sb.Append('"').Append(g.ToString()).Append('"');
                return;
            }

            if (t.IsEnum)
            {
                WriteString(sb, Convert.ToString(value, CultureInfo.InvariantCulture));
                return;
            }

            if (value is byte[] bytes)
            {
                WriteString(sb, Convert.ToBase64String(bytes));
                return;
            }

            // IDictionary
            var asDict = value as IDictionary;
            if (asDict != null)
            {
                WriteDictionary(sb, asDict, depth);
                return;
            }

            // IEnumerable (array/list)
            var asEnum = value as IEnumerable;
            if (asEnum != null && !(value is string))
            {
                WriteArray(sb, asEnum, depth);
                return;
            }

            // Objeto com propriedades públicas
            WriteObject(sb, value, depth);
        }

        private static void WriteDictionary(StringBuilder sb, IDictionary dict, int depth)
        {
            sb.Append('{');
            bool first = true;
            foreach (DictionaryEntry kv in dict)
            {
                if (!first) sb.Append(',');
                first = false;
                // Forçamos chave como string
                WriteString(sb, Convert.ToString(kv.Key, CultureInfo.InvariantCulture) ?? string.Empty);
                sb.Append(':');
                WriteValue(sb, kv.Value, depth + 1);
            }
            sb.Append('}');
        }

        private static void WriteArray(StringBuilder sb, IEnumerable seq, int depth)
        {
            sb.Append('[');
            bool first = true;
            foreach (var item in seq)
            {
                if (!first) sb.Append(',');
                first = false;
                WriteValue(sb, item, depth + 1);
            }
            sb.Append(']');
        }

        private static void WriteObject(StringBuilder sb, object obj, int depth)
        {
            sb.Append('{');
            bool first = true;

            // Propriedades públicas com getter, sem indexadores
            var props = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < props.Length; i++)
            {
                var p = props[i];
                if (!p.CanRead) continue;
                if (p.GetIndexParameters().Length != 0) continue;

                object val;
                try
                {
                    val = p.GetValue(obj, null);
                }
                catch
                {
                    continue; // ignora propriedades que lançam
                }

                if (!first) sb.Append(',');
                first = false;

                WriteString(sb, p.Name);
                sb.Append(':');
                WriteValue(sb, val, depth + 1);
            }

            sb.Append('}');
        }

        private static void WriteString(StringBuilder sb, string s)
        {
            if (s == null) { sb.Append("null"); return; }

            sb.Append('"');
            for (int i = 0; i < s.Length; i++)
            {
                var ch = s[i];
                switch (ch)
                {
                    case '\"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 32 || ch > 0x7E)
                        {
                            sb.Append("\\u").Append(((int)ch).ToString("x4"));
                        }
                        else
                        {
                            sb.Append(ch);
                        }
                        break;
                }
            }
            sb.Append('"');
        }

        private static bool IsNumber(Type t)
        {
            // cobre inteiros e ponto flutuante
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;
                default:
                    return false;
            }
        }
    }
}
