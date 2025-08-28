using System;

namespace ComHookLib.Dto
{
    internal class ComEventBase
    {
        public string ts { get; set; } = DateTime.UtcNow.ToString("o");
        public int pid { get; set; } = System.Diagnostics.Process.GetCurrentProcess().Id;
        public uint tid { get; set; } = (uint)Native.GetCurrentThreadId();
        public string schema { get; set; } = "comlog.v2";
        // Observação: build/session_id são adicionados pelo ComLogger.Write nos logs textuais;
        // aqui focamos no payload de evento COM.
    }
}
