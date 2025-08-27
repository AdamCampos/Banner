using System;
using System.Threading;
using ComHookLib;
using EasyHook;

namespace ComHook
{
    /// <summary>
    /// Ponto de entrada remoto do EasyHook.
    /// Mantém o processo-alvo vivo e instala/desinstala hooks com segurança.
    /// </summary>
    public class RemoteEntry : IEntryPoint
    {
        // Canal/IPC não utilizado neste exemplo; mantido para compatibilidade
        public RemoteEntry(RemoteHooking.IContext context, string channelName)
        {
            // No ctor apenas registramos o início
            ComLogger.Write("[Remote] Starting hooks");
        }

        public void Run(RemoteHooking.IContext context, string channelName)
        {
            try
            {
                // Instala os hooks
                ComHooks.Install();

                // Mantém o remoto vivo até ser descarregado
                // (em implementações com IPC, você poderia aguardar sinal)
                while (true)
                {
                    Thread.Sleep(500);
                }
            }
            catch (ThreadAbortException)
            {
                // Descarregado pelo injetor — tenta desinstalar
                try { ComHooks.Uninstall(); } catch (Exception ex) { ComLogger.Err("RemoteEntry.Run/Abort", ex); }
            }
            catch (Exception ex)
            {
                ComLogger.Err("RemoteEntry.Run", ex);
                try { ComHooks.Uninstall(); } catch (Exception ex2) { ComLogger.Err("RemoteEntry.Run/Uninstall", ex2); }
            }
        }

        /// <summary>
        /// Método auxiliar para cenários em que o injetor chame explicitamente.
        /// </summary>
        public static void Uninstall()
        {
            try { ComHooks.Uninstall(); }
            catch (Exception ex) { ComLogger.Err("RemoteEntry.Uninstall", ex); }
        }
    }
}
