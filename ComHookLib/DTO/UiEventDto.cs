using System;

namespace ComHookLib.Dto
{
    internal class UiEventDto
    {
        public string ts { get; set; } = DateTime.UtcNow.ToString("o");
        public string evt { get; set; }              // "window.create" | "window.show" | "window.foreground"
        public string kind { get; set; }             // "dialog" | "ftaebanner" | "window"
        public int pid { get; set; } = System.Diagnostics.Process.GetCurrentProcess().Id;
        public uint tid { get; set; } = (uint)Native.GetCurrentThreadId();
        public string hwnd { get; set; }             // "0xABCDEF..."
        public string cls { get; set; }              // class name
        public string title { get; set; }            // window text
        public uint threadId { get; set; }
    }
}
