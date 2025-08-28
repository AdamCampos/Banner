using System;

namespace ComHookLib
{
    /// <summary>
    /// DTOs usados no log: COM (CoGetClassObject/CoCreateInstance/Ex) e UI (WinEvent).
    /// Mantém nomes de propriedades iguais ao JSON atual para retrocompatibilidade.
    /// </summary>
    public static class ComEvents
    {
        // --------------------------- COM ---------------------------

        public sealed class CoGetClassObjectEvent
        {
            public string ts { get; set; }
            public string api { get; set; } = "CoGetClassObject";
            public int pid { get; set; }
            public int tid { get; set; }
            public string clsid { get; set; }
            public string progId { get; set; }
            public string iid { get; set; }
            public string iid_name { get; set; }
            public string clsctx { get; set; }
            public string hr { get; set; }
            public string hr_name { get; set; }
            public double elapsed_ms { get; set; }
            public string kind { get; set; }
        }

        public sealed class CoCreateInstanceEvent
        {
            public string ts { get; set; }
            public string api { get; set; } = "CoCreateInstance";
            public int pid { get; set; }
            public int tid { get; set; }
            public string clsid { get; set; }
            public string progId { get; set; }
            public string iid { get; set; }
            public string iid_name { get; set; }
            public string clsctx { get; set; }
            public string hr { get; set; }
            public string hr_name { get; set; }
            public double elapsed_ms { get; set; }
            public string kind { get; set; }
        }

        public sealed class CoCreateInstanceExEvent
        {
            public string ts { get; set; }
            public string api { get; set; } = "CoCreateInstanceEx";
            public int pid { get; set; }
            public int tid { get; set; }
            public string clsid { get; set; }
            public string progId { get; set; }
            public string clsctx { get; set; }
            public uint count { get; set; }
            public string[] iids { get; set; }
            public string[] iid_names { get; set; }
            public int[] multiqi_hr { get; set; }
            public string hr { get; set; }
            public string hr_name { get; set; }
            public double elapsed_ms { get; set; }
            public string kind { get; set; }
        }

        // ---------------------------- UI ---------------------------

        public sealed class UiWindowEvent
        {
            public string ts { get; set; }
            public string evt { get; set; }
            public string kind { get; set; }
            public int pid { get; set; }
            public int tid { get; set; }
            public string hwnd { get; set; }
            public string cls { get; set; }
            public string title { get; set; }
            public uint threadId { get; set; }
        }
    }
}
