using System;
using System.IO;
using System.Threading;
using EasyHook;

namespace ComHook
{
    /// <summary>
    /// Ponto de entrada remoto do EasyHook.
    /// Mantém o processo-alvo vivo e instala/desinstala hooks com segurança.
    /// </summary>
    public class RemoteEntry : IEntryPoint
    {
        private readonly string _logPath;

        // O Injector envia: (context, logPath)
        public RemoteEntry(RemoteHooking.IContext context, string logPath)
        {
            _logPath = logPath;
            try
            {
                if (!string.IsNullOrEmpty(_logPath))
                {
                    var dir = Path.GetDirectoryName(_logPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);

                    File.AppendAllText(_logPath,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REMOTE OK] RemoteEntry ctor carregado. PID={RemoteHooking.GetCurrentProcessId()}{Environment.NewLine}");
                }
            }
            catch { /* nunca propaga no remoto */ }
        }

        public void Run(RemoteHooking.IContext context, string logPathFromInjector)
        {
            try
            {
                // Instala os hooks
                ComHook.ComHooks.Install();

                // Mais um ping de prova-de-vida
                try
                {
                    if (!string.IsNullOrEmpty(_logPath))
                    {
                        File.AppendAllText(_logPath,
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REMOTE OK] Hooks instalados.{Environment.NewLine}");
                    }
                }
                catch { }

                // Mantém o remoto vivo
                while (true)
                    Thread.Sleep(500);
            }
            catch (ThreadAbortException)
            {
                try { ComHook.ComHooks.Uninstall(); } catch { }
            }
            catch (Exception)
            {
                try { ComHook.ComHooks.Uninstall(); } catch { }
            }
        }

        public static void Uninstall()
        {
            try { ComHook.ComHooks.Uninstall(); } catch { }
        }
    }
}
