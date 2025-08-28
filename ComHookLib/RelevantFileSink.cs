// ComHookLib - RelevantFileSink.cs
// Grava apenas linhas "relevantes": kind == "ftaerel" OU hr_name != "S_OK"

using System;
using System.IO;

namespace ComHookLib
{
    internal sealed class RelevantFileSink : ILogSink, IDisposable
    {
        private readonly StreamWriter _sw;

        internal RelevantFileSink(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("filePath inválido.", nameof(filePath));

            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            _sw = new StreamWriter(filePath, true);
            _sw.AutoFlush = true;
        }

        public void WriteLine(string line)
        {
            if (line == null) return;

            // critério simples e performático (sem JSON parse):
            // 1) kind == "ftaerel"
            bool isRel = line.IndexOf("\"kind\":\"ftaerel\"", StringComparison.OrdinalIgnoreCase) >= 0;

            // 2) hr_name != "S_OK" (se houver hr_name)
            if (!isRel)
            {
                int idx = line.IndexOf("\"hr_name\":\"", StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    // pega o valor de hr_name rapidamente
                    int start = idx + "\"hr_name\":\"".Length;
                    int end = line.IndexOf('"', start);
                    if (end > start)
                    {
                        var hrName = line.Substring(start, end - start);
                        if (!string.Equals(hrName, "S_OK", StringComparison.OrdinalIgnoreCase))
                            isRel = true;
                    }
                }
            }

            if (isRel)
                _sw.WriteLine(line);
        }

        public void Dispose()
        {
            try { _sw.Dispose(); } catch { /* ignore */ }
        }
    }
}
