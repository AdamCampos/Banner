using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using EasyHook;
using ComHookLib.Hooking;

namespace ComHookLib
{
    internal static class ComHooks
    {
        private static bool _installed;
        private static ILogger _logger;

        private static LocalHook _hkCCI, _hkCCIEx, _hkCGCO;
        private static CoCreateInstance_Delegate _origCCI;
        private static CoCreateInstanceEx_Delegate _origCCIEx;
        private static CoGetClassObject_Delegate _origCGCO;

        public static void Install(ILogger logger)
        {
            if (_installed) return;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            try
            {
                _origCCI = EasyHookHelper.GetOriginalDelegate<CoCreateInstance_Delegate>("ole32.dll", "CoCreateInstance");
                _origCCIEx = EasyHookHelper.GetOriginalDelegate<CoCreateInstanceEx_Delegate>("ole32.dll", "CoCreateInstanceEx");
                _origCGCO = EasyHookHelper.GetOriginalDelegate<CoGetClassObject_Delegate>("ole32.dll", "CoGetClassObject");

                _hkCCI = EasyHookHelper.CreateHook("ole32.dll", "CoCreateInstance", new CoCreateInstance_Delegate(CCICallback), null);
                _hkCCIEx = EasyHookHelper.CreateHook("ole32.dll", "CoCreateInstanceEx", new CoCreateInstanceEx_Delegate(CCIExCallback), null);
                _hkCGCO = EasyHookHelper.CreateHook("ole32.dll", "CoGetClassObject", new CoGetClassObject_Delegate(CGCOCallback), null);

                _installed = true;
                _logger.Info("COM hooks instalados com sucesso (CoCreateInstance/Ex, CoGetClassObject).");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Falha ao instalar hooks COM.");
            }
        }

        private static int CCICallback(ref Guid rclsid, IntPtr pUnkOuter, CLSCTX ctx, ref Guid riid, out IntPtr ppv)
        {
            var pid = Process.GetCurrentProcess().Id;
            var tid = Native.GetCurrentThreadId();
            string progId = Native.ProgIdFromClsidSafe(rclsid);

            var sw = Stopwatch.StartNew();
            int hr = 0;
            ppv = IntPtr.Zero;

            try
            {
                hr = _origCCI(ref rclsid, pUnkOuter, ctx, ref riid, out ppv);
                return hr;
            }
            finally
            {
                sw.Stop();
                SafeLog(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoCreateInstance",
                    pid = pid,
                    tid = tid,
                    clsid = rclsid,
                    progId = progId,
                    iid = riid,
                    clsctx = ctx.ToString(),
                    hr = $"0x{hr:X8}",
                    elapsed_ms = sw.Elapsed.TotalMilliseconds
                });
            }
        }

        private static int CCIExCallback(ref Guid rclsid, IntPtr pUnkOuter, CLSCTX ctx, IntPtr pServerInfo, uint dwCount, IntPtr pResults)
        {
            var pid = Process.GetCurrentProcess().Id;
            var tid = Native.GetCurrentThreadId();
            string progId = Native.ProgIdFromClsidSafe(rclsid);

            Guid[] iidsIn = ReadIidsFromMultiQi(pResults, dwCount);
            int hr = 0;
            var sw = Stopwatch.StartNew();

            try
            {
                hr = _origCCIEx(ref rclsid, pUnkOuter, ctx, pServerInfo, dwCount, pResults);
                return hr;
            }
            finally
            {
                sw.Stop();
                int[] hrsPerQi = ReadHResultsFromMultiQi(pResults, dwCount);

                SafeLog(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoCreateInstanceEx",
                    pid = pid,
                    tid = tid,
                    clsid = rclsid,
                    progId = progId,
                    clsctx = ctx.ToString(),
                    count = dwCount,
                    iids = iidsIn,
                    multiqi_hr = hrsPerQi,
                    hr = $"0x{hr:X8}",
                    elapsed_ms = sw.Elapsed.TotalMilliseconds
                });
            }
        }

        private static int CGCOCallback(ref Guid rclsid, CLSCTX ctx, IntPtr pServerInfo, ref Guid riid, out IntPtr ppv)
        {
            var pid = Process.GetCurrentProcess().Id;
            var tid = Native.GetCurrentThreadId();
            string progId = Native.ProgIdFromClsidSafe(rclsid);

            var sw = Stopwatch.StartNew();
            int hr = 0;
            ppv = IntPtr.Zero;

            try
            {
                hr = _origCGCO(ref rclsid, ctx, pServerInfo, ref riid, out ppv);
                return hr;
            }
            finally
            {
                sw.Stop();
                SafeLog(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoGetClassObject",
                    pid = pid,
                    tid = tid,
                    clsid = rclsid,
                    progId = progId,
                    iid = riid,
                    clsctx = ctx.ToString(),
                    hr = $"0x{hr:X8}",
                    elapsed_ms = sw.Elapsed.TotalMilliseconds
                });
            }
        }

        private static Guid[] ReadIidsFromMultiQi(IntPtr pResults, uint count)
        {
            if (pResults == IntPtr.Zero || count == 0) return new Guid[0];
            var list = new Guid[count];
            int size = Marshal.SizeOf(typeof(MULTI_QI));
            for (int i = 0; i < count; i++)
            {
                IntPtr cur = IntPtr.Add(pResults, i * size);
                MULTI_QI qi;
                try { qi = Marshal.PtrToStructure<MULTI_QI>(cur); }
                catch { continue; }

                try
                {
                    if (qi.pIID != IntPtr.Zero)
                    {
                        var iid = Marshal.PtrToStructure<Guid>(qi.pIID);
                        list[i] = iid;
                    }
                }
                catch { }
            }
            return list;
        }

        private static int[] ReadHResultsFromMultiQi(IntPtr pResults, uint count)
        {
            if (pResults == IntPtr.Zero || count == 0) return new int[0];
            var list = new int[count];
            int size = Marshal.SizeOf(typeof(MULTI_QI));
            for (int i = 0; i < count; i++)
            {
                IntPtr cur = IntPtr.Add(pResults, i * size);
                try
                {
                    var qi = Marshal.PtrToStructure<MULTI_QI>(cur);
                    list[i] = qi.hr;
                }
                catch { }
            }
            return list;
        }

        private static void SafeLog(object payload)
        {
            try { _logger?.Log(payload); }
            catch { try { _logger?.Warn("Falha ao logar payload COM."); } catch { } }
        }
    }
}
