using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ComHookLib
{
    /// <summary>
    /// Logger simples compatível com C# 7.3.
    /// - Gera arquivo de log e CSV no diretório %TEMP% (ou pasta atual se falhar).
    /// - Oferece múltiplas sobrecargas de Csv para cobrir todas as chamadas do ComHooks.
    /// - Métodos auxiliares: Write (texto livre) e Err (erro com exceção).
    /// </summary>
    public static class ComLogger
    {
        private static readonly object _sync = new object();
        private static readonly string _baseDir;
        private static readonly string _txtPath;
        private static readonly string _csvPath;

        static ComLogger()
        {
            try
            {
                _baseDir = Path.GetTempPath();
                if (string.IsNullOrEmpty(_baseDir) || !Directory.Exists(_baseDir))
                    _baseDir = AppDomain.CurrentDomain.BaseDirectory;
            }
            catch
            {
                _baseDir = AppDomain.CurrentDomain.BaseDirectory;
            }

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            _txtPath = Path.Combine(_baseDir, $"com_activations_{stamp}.log");
            _csvPath = Path.Combine(_baseDir, $"com_activations_{stamp}.csv");

            // Cabeçalho do CSV
            SafeAppend(_csvPath,
                "ts;api;arg1;arg2;ctx;hr;threadId;serverResolved;extra;modulePath" + Environment.NewLine);
            // Cabeçalho do TXT
            SafeAppend(_txtPath,
                $"# COM Activation Log  {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}");
        }

        // ========== API de texto ==========
        public static void Write(string message)
        {
            try
            {
                var line = $"{DateTime.Now:HH:mm:ss.fff}  {message}{Environment.NewLine}";
                SafeAppend(_txtPath, line);
            }
            catch { /* nunca propaga */ }
        }

        public static void Err(string where, Exception ex)
        {
            try
            {
                var line = $"{DateTime.Now:HH:mm:ss.fff}  [ERR] {where}: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}";
                SafeAppend(_txtPath, line);
            }
            catch { /* nunca propaga */ }
        }

        // ========== API CSV ==========
        // Assinatura "clássica" (8 parâmetros) usada anteriormente
        public static void Csv(string api, string arg1, string arg2, int ctx, uint hr, int threadId, string serverResolved, string extra)
        {
            Csv(api, arg1, arg2, ctx, hr, threadId, serverResolved, extra, null);
        }

        // Assinatura estendida (9 parâmetros) com modulePath nomeado
        public static void Csv(string api, string arg1, string arg2, int ctx, uint hr, int threadId, string serverResolved, string extra, string modulePath)
        {
            try
            {
                var ts = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
                var sb = new StringBuilder();
                sb.Append(ts).Append(';')
                  .Append(S(api)).Append(';')
                  .Append(S(arg1)).Append(';')
                  .Append(S(arg2)).Append(';')
                  .Append(ctx.ToString(CultureInfo.InvariantCulture)).Append(';')
                  .Append(hr.ToString("X8", CultureInfo.InvariantCulture)).Append(';')
                  .Append(threadId.ToString(CultureInfo.InvariantCulture)).Append(';')
                  .Append(S(serverResolved)).Append(';')
                  .Append(S(extra)).Append(';')
                  .Append(S(modulePath))
                  .Append(Environment.NewLine);

                SafeAppend(_csvPath, sb.ToString());

                // Também escreve um resumo no TXT para acompanhamento humano
                var human = $"{ts}  {api} {arg1}  ctx={ctx}  hr=0x{hr:X8}  {(threadId != 0 ? $"[T{threadId}]" : "")}  {extra}".Trim();
                SafeAppend(_txtPath, human + Environment.NewLine);
            }
            catch { /* nunca propaga */ }
        }

        // Sobrecarga simplificada usada quando só queremos registrar o modulePath
        public static void Csv(string api, string arg1, string arg2, int ctx, uint hr, string modulePath)
        {
            // threadId=0; serverResolved vazio; extra vazio
            Csv(api, arg1, arg2, ctx, hr, 0, "", "", modulePath);
        }

        // ======== util ========
        private static string S(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            // Sanitiza ponto-e-vírgula e quebras de linha para manter CSV estável
            return s.Replace(";", ",").Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static void SafeAppend(string path, string text)
        {
            lock (_sync)
            {
                File.AppendAllText(path, text, Encoding.UTF8);
            }
        }
    }
}
