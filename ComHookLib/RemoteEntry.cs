using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using EasyHook;

namespace ComHookLib
{
    public sealed class RemoteEntry : IEntryPoint, IDisposable
    {
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly string _sessionId;
        private readonly string _logPath;
        private readonly ILogger _logger;
        private readonly ILogSink _sink; // encadeado com RelevantFileSink quando habilitado
        private UiHook _uiHook;

        private static void SafeAppendShared(string path, string text)
        {
            try
            {
                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
                    sw.Write(text);
            }
            catch { }
        }

        private static string BuildRelevantPath(string mainPath)
        {
            try
            {
                var dir = Path.GetDirectoryName(mainPath) ?? "";
                var name = Path.GetFileNameWithoutExtension(mainPath);
                if (string.IsNullOrEmpty(name)) name = "ftaelog";
                var ext = Path.GetExtension(mainPath);
                if (string.IsNullOrEmpty(ext)) ext = ".log";
                return Path.Combine(dir, name + ".relevant" + ext);
            }
            catch
            {
                return mainPath + ".relevant";
            }
        }

        public RemoteEntry(RemoteHooking.IContext context, string logPath)
        {
            _logPath = string.IsNullOrWhiteSpace(logPath)
                ? Path.Combine(Path.GetTempPath(), "ftaelog.log")
                : logPath;

            _sessionId = $"{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}-{Process.GetCurrentProcess().Id}-{Native.GetCurrentThreadId()}";

            // --- T7: bootstrap do sink secundário "relevante" via env var FTA_LOG_RELEVANT ---
            ILogSink baseSink = new ComLogIpc("FTAEPipe", _logPath);
            var enableRelevant = Environment.GetEnvironmentVariable("FTA_LOG_RELEVANT");
            if (string.Equals(enableRelevant, "1", StringComparison.Ordinal))
            {
                var relPath = BuildRelevantPath(_logPath);
                ILogSink relevantSink = new RelevantFileSink(relPath);
                _sink = new MultiSink(baseSink, relevantSink);
                SafeAppendShared(_logPath, "[REMOTE OK] RelevantFileSink habilitado: " + relPath + Environment.NewLine);
            }
            else
            {
                _sink = baseSink;
            }
            // -------------------------------------------------------------------------------

            _logger = new ComLogger(_sink, _sessionId);

            SafeAppendShared(_logPath, "[REMOTE OK] IEntryPoint carregado. PID=" + Process.GetCurrentProcess().Id + " session=" + _sessionId + Environment.NewLine);
            _logger.Info("[REMOTE OK] RemoteEntry iniciado. PID={0} session={1}", Process.GetCurrentProcess().Id, _sessionId);

            _uiHook = new UiHook(_logger);
            _uiHook.Start();

            // Etapa 2/4: instalar hooks COM reais
            ComHooks.Install(_logger);

            _logger.Info("[REMOTE OK] Hooks (UI + COM) instalados.");
        }

        public void Run(RemoteHooking.IContext context, string logPathFromInjector)
        {
            try
            {
                while (!_cts.IsCancellationRequested)
                    Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "RemoteEntry.Run falhou.");
            }
        }

        public void Dispose()
        {
            try
            {
                _cts.Cancel();
                _uiHook?.Dispose();

                var d = _sink as IDisposable;
                if (d != null) d.Dispose();
            }
            catch { }
        }
    }
}
