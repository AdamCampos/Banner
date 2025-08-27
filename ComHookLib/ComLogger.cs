using System;
using System.Diagnostics;
using System.Globalization;

namespace ComHookLib
{
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

        public ComLogger(ILogSink sink, string sessionId)
        {
            if (sink == null) throw new ArgumentNullException("sink");
            _sink = sink;
            _sessionId = sessionId ?? "unknown";
        }

        public void Info(string message, params object[] args)
        {
            Write("info", null, message, args);
        }

        public void Warn(string message, params object[] args)
        {
            Write("warn", null, message, args);
        }

        public void Error(string message, params object[] args)
        {
            Write("error", null, message, args);
        }

        public void Error(Exception ex, string message, params object[] args)
        {
            Write("error", ex, message, args);
        }

        public void Log(object dto)
        {
            if (dto == null) return;
            _sink.WriteLine(Jsonl.Serialize(dto));
        }

        private void Write(string level, Exception ex, string message, params object[] args)
        {
            string text = SafeFormat(message, args);
            var payload = new
            {
                ts = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                level = level,
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
                // Em caso de erro de formatação, retorna concat simples
                return message + " " + string.Join(" | ", args ?? new object[0]);
            }
        }
    }
}
