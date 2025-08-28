using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace ComHookLib
{
    // NÃO declarar ILogSink aqui — já existe (ex.: em ComLogIpc.cs):
    // public interface ILogSink { void WriteLine(string line); }

    public interface ILogger
    {
        void Info(string message, params object[] args);
        void Warn(string message, params object[] args);
        void Error(string message, params object[] args);
        void Error(Exception ex, string message, params object[] args);
        void Log(object dto);
    }

    internal sealed class ComLogger : ILogger
    {
        private readonly ILogSink _sink;
        private readonly string _sessionId;

        // Metadados fixos do log
        private const string Schema = "comlog.v2";
        private static readonly string BuildVersion = GetBuildVersion();

        public ComLogger(ILogSink sink, string sessionId)
        {
            if (sink == null) throw new ArgumentNullException(nameof(sink));
            _sink = sink;
            _sessionId = sessionId ?? "unknown";
        }

        public void Info(string message, params object[] args) { Write("info", null, message, args); }
        public void Warn(string message, params object[] args) { Write("warn", null, message, args); }
        public void Error(string message, params object[] args) { Write("error", null, message, args); }
        public void Error(Exception ex, string message, params object[] args) { Write("error", ex, message, args); }

        /// <summary>
        /// Ponto único de saída de eventos (DTOs). Aplica enriquecimento (HRESULT/IID/ProgID/kind/CLSCTX).
        /// </summary>
        public void Log(object dto)
        {
            if (dto == null) return;

            object enriched = dto;
            try
            {
                enriched = EnrichIfComEvent(dto);
            }
            catch (Exception ex)
            {
                _sink.WriteLine(Jsonl.Serialize(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    level = "warn",
                    schema = "comlog.v2",
                    build = BuildVersion,  // <-- corrigido
                    session_id = _sessionId,
                    pid = System.Diagnostics.Process.GetCurrentProcess().Id,
                    tid = Native.GetCurrentThreadId(),
                    msg = "Falha no EnrichIfComEvent.",
                    err = ex.ToString()
                }));
                return;
            }

            try
            {
                _sink.WriteLine(Jsonl.Serialize(enriched));
            }
            catch (Exception ex)
            {
                _sink.WriteLine(Jsonl.Serialize(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    level = "warn",
                    schema = "comlog.v2",
                    build = BuildVersion, // <-- corrigido
                    session_id = _sessionId,
                    pid = System.Diagnostics.Process.GetCurrentProcess().Id,
                    tid = Native.GetCurrentThreadId(),
                    msg = "Falha ao logar payload COM.",
                    err = ex.ToString()
                }));
            }
        }


        // =========================================================================
        // Enriquecimento COM: resolve hr_name, iid_name, progId (com fallback T4), kind e normaliza clsctx (T5)
        // =========================================================================
        // Dentro de ComLogger.cs
        // =========================================================================
        // Enriquecimento COM tolerante a tipos anônimos (sem setter)
        // - Se tiver setter, grava na própria DTO;
        // - Caso contrário, cria um Dictionary<string,object> com os campos extra.
        // =========================================================================
        private object EnrichIfComEvent(object dto)
        {
            if (dto == null) return null;

            var t = dto.GetType();

            // Só enriquece COM se tiver "api" e "clsid"
            var apiProp = t.GetProperty("api");
            var clsidProp = t.GetProperty("clsid");
            if (apiProp == null || clsidProp == null)
                return dto;

            var clsid = clsidProp.GetValue(dto) as string;

            // (1) ProgID (T3 + T4)
            var progIdProp = t.GetProperty("progId");
            var progIdIn = progIdProp != null ? progIdProp.GetValue(dto) as string : null;
            var progIdResolved = ComDictionary.TryResolveProgId(clsid, progIdIn);
            if (progIdProp != null && progIdProp.CanWrite &&
                !string.Equals(progIdResolved, progIdIn, StringComparison.Ordinal))
            {
                progIdProp.SetValue(dto, progIdResolved);
            }

            // (2) HRESULT legível
            var hrProp = t.GetProperty("hr");
            var hrNameProp = t.GetProperty("hr_name");
            if (hrProp != null && hrNameProp != null && hrNameProp.CanWrite)
            {
                var hrStr = hrProp.GetValue(dto) as string;
                var hrName = ComDictionary.TryResolveHResult(hrStr);
                if (!string.IsNullOrEmpty(hrName))
                    hrNameProp.SetValue(dto, hrName);
            }

            // (3) IID legível
            var iidProp = t.GetProperty("iid");
            var iidNameProp = t.GetProperty("iid_name");
            if (iidProp != null && iidNameProp != null && iidNameProp.CanWrite)
            {
                var iid = iidProp.GetValue(dto) as string;
                var iidName = ComDictionary.TryResolveIid(iid);
                if (!string.IsNullOrEmpty(iidName))
                    iidNameProp.SetValue(dto, iidName);
            }

            // (3b) Lista de iids (CoCreateInstanceEx): popular iid_names se existir
            var iidsProp = t.GetProperty("iids");
            var iidNamesListProp = t.GetProperty("iid_names");
            if (iidsProp != null && iidNamesListProp != null && iidNamesListProp.CanWrite)
            {
                var iidsVal = iidsProp.GetValue(dto) as System.Collections.IEnumerable;
                if (iidsVal != null)
                {
                    var names = new System.Collections.Generic.List<string>();
                    foreach (var o in iidsVal)
                        names.Add(ComDictionary.TryResolveIid(o as string));

                    // 👇 AQUI está o pulo do gato:
                    iidNamesListProp.SetValue(dto, names.ToArray()); // em vez de List<string>
                }
            }

            // (4) CLSCTX normalizado (T5)
            var clsctxProp = t.GetProperty("clsctx");
            if (clsctxProp != null && clsctxProp.CanWrite)
            {
                var raw = clsctxProp.GetValue(dto) as string;
                var norm = ComDecode.NormalizeClsctx(raw); // sua função existente
                if (!string.IsNullOrEmpty(norm))
                    clsctxProp.SetValue(dto, norm);
            }

            // (5) kind (clsid + progId resolvido)
            var kindProp = t.GetProperty("kind");
            if (kindProp != null && kindProp.CanWrite)
            {
                var kind = ComDictionary.TryResolveKind(clsid, progIdResolved);
                if (!string.IsNullOrEmpty(kind))
                    kindProp.SetValue(dto, kind);
            }

            return dto;
        }




        // =========================================================================
        // Infra de logging "texto" → JSONL
        // =========================================================================
        private void Write(string level, Exception ex, string message, params object[] args)
        {
            string text = SafeFormat(message, args);
            var payload = new
            {
                ts = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                level = level,
                schema = Schema,
                build = BuildVersion,
                session_id = _sessionId,
                pid = Process.GetCurrentProcess().Id,
                tid = Native.GetCurrentThreadId(),
                msg = text,
                err = ex != null ? ex.ToString() : null
            };
            _sink.WriteLine(Jsonl.Serialize(payload));
        }

        private static string SafeFormat(string message, object[] args)
        {
            try
            {
                return (args != null && args.Length > 0)
                    ? string.Format(CultureInfo.InvariantCulture, message, args)
                    : message;
            }
            catch
            {
                return message + " " + string.Join(" | ", args ?? Array.Empty<object>());
            }
        }

        // =========================================================================
        // Helpers de DTO → Dictionary<string, object> (mantidos por compatibilidade)
        // =========================================================================
        private static Dictionary<string, object> ToDictionary(object dto)
        {
            if (dto is IDictionary asDict)
            {
                var copy = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry kv in asDict)
                {
                    var key = kv.Key?.ToString() ?? string.Empty;
                    copy[key] = kv.Value;
                }
                return copy;
            }

            var map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var props = dto.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < props.Length; i++)
            {
                var p = props[i];
                if (!p.CanRead) continue;
                if (p.GetIndexParameters().Length != 0) continue;

                try { map[p.Name] = p.GetValue(dto, null); }
                catch { /* ignora propriedades que lançam */ }
            }
            return map;
        }

        private static string GetString(Dictionary<string, object> map, string key)
        {
            object v;
            if (map.TryGetValue(key, out v) && v != null)
                return v as string ?? Convert.ToString(v, CultureInfo.InvariantCulture);
            return null;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            if (values == null) return null;
            for (int i = 0; i < values.Length; i++)
                if (!string.IsNullOrWhiteSpace(values[i]))
                    return values[i];
            return null;
        }

        private static string GetBuildVersion()
        {
            try
            {
                var asm = typeof(ComLogger).Assembly;
                var fvi = FileVersionInfo.GetVersionInfo(asm.Location);
                return string.IsNullOrWhiteSpace(fvi.FileVersion)
                    ? (asm.GetName().Version != null ? asm.GetName().Version.ToString() : "0.0.0.0")
                    : fvi.FileVersion;
            }
            catch
            {
                return "0.0.0.0";
            }
        }
    }
}
