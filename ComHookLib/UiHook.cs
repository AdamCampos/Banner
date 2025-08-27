using System;

namespace ComHookLib
{
    /// <summary>
    /// Captura eventos de janela para “banner/diálogo” via WinEventHook.
    /// Roda dentro do processo-alvo. Usa ILogger do ComLogger.cs.
    /// </summary>
    internal sealed class UiHook : IDisposable
    {
        private readonly ILogger _logger;

        private Native.WinEventDelegate _cbCreate;
        private Native.WinEventDelegate _cbShow;
        private Native.WinEventDelegate _cbForeground;

        private IntPtr _hCreate;
        private IntPtr _hShow;
        private IntPtr _hForeground;

        public UiHook(ILogger logger)
        {
            if (logger == null) throw new ArgumentNullException("logger");
            _logger = logger;
        }

        public void Start()
        {
            _cbCreate = OnEventCreate;
            _cbShow = OnEventShow;
            _cbForeground = OnEventForeground;

            _hCreate = Native.SetWinEventHook(Native.EVENT_OBJECT_CREATE, Native.EVENT_OBJECT_CREATE, IntPtr.Zero, _cbCreate, 0, 0, Native.WINEVENT_OUTOFCONTEXT);
            _hShow = Native.SetWinEventHook(Native.EVENT_OBJECT_SHOW, Native.EVENT_OBJECT_SHOW, IntPtr.Zero, _cbShow, 0, 0, Native.WINEVENT_OUTOFCONTEXT);
            _hForeground = Native.SetWinEventHook(Native.EVENT_SYSTEM_FOREGROUND, Native.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _cbForeground, 0, 0, Native.WINEVENT_OUTOFCONTEXT);

            _logger.Info("UiHook iniciado: create={0} show={1} foreground={2}", _hCreate, _hShow, _hForeground);
        }

        private void OnEventCreate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != 0 || hwnd == IntPtr.Zero) return;
            LogWindow("window.create", hwnd, dwEventThread);
        }

        private void OnEventShow(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != 0 || hwnd == IntPtr.Zero) return;
            LogWindow("window.show", hwnd, dwEventThread);
        }

        private void OnEventForeground(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != 0 || hwnd == IntPtr.Zero) return;
            LogWindow("window.foreground", hwnd, dwEventThread);
        }

        private void LogWindow(string evt, IntPtr hwnd, uint threadId)
        {
            string cls = Native.GetClassNameSafe(hwnd);
            string title = Native.GetWindowTextSafe(hwnd);

            string kind =
                (!string.IsNullOrEmpty(cls) && (cls.IndexOf("#32770", StringComparison.OrdinalIgnoreCase) >= 0 || cls.IndexOf("Dialog", StringComparison.OrdinalIgnoreCase) >= 0)) ? "dialog" :
                ((!string.IsNullOrEmpty(cls) && cls.IndexOf("FT", StringComparison.OrdinalIgnoreCase) >= 0) || (!string.IsNullOrEmpty(title) && title.IndexOf("Alarm", StringComparison.OrdinalIgnoreCase) >= 0)) ? "ftaebanner" :
                "window";

            var record = new
            {
                ts = DateTime.UtcNow.ToString("o"),
                evt = evt,
                kind = kind,
                pid = System.Diagnostics.Process.GetCurrentProcess().Id,
                tid = Native.GetCurrentThreadId(),
                hwnd = "0x" + ((long)hwnd).ToString("X"),
                cls = cls,
                title = title,
                threadId = threadId
            };

            _logger.Log(record);
            _logger.Info("UI {0} {1} hwnd={2} cls='{3}' title='{4}'", evt, kind, record.hwnd, cls, title);
        }

        public void Dispose()
        {
            try { if (_hCreate != IntPtr.Zero) Native.UnhookWinEvent(_hCreate); } catch { }
            try { if (_hShow != IntPtr.Zero) Native.UnhookWinEvent(_hShow); } catch { }
            try { if (_hForeground != IntPtr.Zero) Native.UnhookWinEvent(_hForeground); } catch { }
        }
    }
}
