using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ComHookLib
{
    internal static class Native
    {
        // ---------- Threads ----------
        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        // ---------- WinEvent hooks ----------
        internal const uint EVENT_OBJECT_CREATE = 0x8000;
        internal const uint EVENT_OBJECT_SHOW = 0x8002;
        internal const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        internal const uint WINEVENT_OUTOFCONTEXT = 0x0000;

        internal delegate void WinEventDelegate(
            IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        internal static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        // ---------- Windows ----------
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        internal static string GetClassNameSafe(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            try { GetClassNameW(hwnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        internal static string GetWindowTextSafe(IntPtr hwnd)
        {
            var sb = new StringBuilder(1024);
            try { GetWindowTextW(hwnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        // ---------- COM (sem [MarshalAs], usando IntPtr) ----------
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int ProgIDFromCLSID(ref Guid clsid, out IntPtr lplpszProgID);

        internal static string ProgIdFromClsidSafe(Guid clsid)
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                int hr = ProgIDFromCLSID(ref clsid, out ptr);
                if (hr != 0 || ptr == IntPtr.Zero)
                    return null;
                // Unicode por CharSet=Unicode
                string s = Marshal.PtrToStringUni(ptr);
                // Doc da Win32: liberar com CoTaskMemFree
                Marshal.FreeCoTaskMem(ptr);
                ptr = IntPtr.Zero;
                return s;
            }
            catch
            {
                try { if (ptr != IntPtr.Zero) Marshal.FreeCoTaskMem(ptr); } catch { }
                return null;
            }
        }
    }
}
