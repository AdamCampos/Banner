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
            var tid = (int)Native.GetCurrentThreadId();
            string progIdRaw = Native.ProgIdFromClsidSafe(rclsid);
            string clsidStr = rclsid.ToString("D");
            string iidStr = riid.ToString("D");

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
                var dto = new ComEvents.CoCreateInstanceEvent
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoCreateInstance",
                    pid = pid,
                    tid = tid,
                    clsid = clsidStr,
                    progId = ComDictionary.TryResolveProgId(clsidStr, progIdRaw),
                    iid = iidStr,
                    iid_name = ComDictionary.TryResolveIid(iidStr),
                    clsctx = ctx.ToString(),
                    hr = $"0x{hr:X8}",
                    hr_name = ComDictionary.TryResolveHResult($"0x{hr:X8}"),
                    elapsed_ms = sw.Elapsed.TotalMilliseconds,
                    kind = ComDictionary.TryResolveKind(clsidStr, progIdRaw)
                };
                SafeLog(dto);
            }
        }

        private static int CCIExCallback(ref Guid rclsid, IntPtr pUnkOuter, CLSCTX ctx, IntPtr pServerInfo, uint dwCount, IntPtr pResults)
        {
            var pid = Process.GetCurrentProcess().Id;
            var tid = (int)Native.GetCurrentThreadId();
            string progIdRaw = Native.ProgIdFromClsidSafe(rclsid);
            string clsidStr = rclsid.ToString("D");

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

                string[] iidStrs = new string[iidsIn.Length];
                string[] iidNames = new string[iidsIn.Length];
                for (int i = 0; i < iidsIn.Length; i++)
                {
                    var s = iidsIn[i].ToString("D");
                    iidStrs[i] = s;
                    iidNames[i] = ComDictionary.TryResolveIid(s);
                }

                var dto = new ComEvents.CoCreateInstanceExEvent
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoCreateInstanceEx",
                    pid = pid,
                    tid = tid,
                    clsid = clsidStr,
                    progId = ComDictionary.TryResolveProgId(clsidStr, progIdRaw),
                    clsctx = ctx.ToString(),
                    count = dwCount,
                    iids = iidStrs,
                    iid_names = iidNames,
                    multiqi_hr = hrsPerQi,
                    hr = $"0x{hr:X8}",
                    hr_name = ComDictionary.TryResolveHResult($"0x{hr:X8}"),
                    elapsed_ms = sw.Elapsed.TotalMilliseconds,
                    kind = ComDictionary.TryResolveKind(clsidStr, progIdRaw)
                };
                SafeLog(dto);
            }
        }

        private static int CGCOCallback(ref Guid rclsid, CLSCTX ctx, IntPtr pServerInfo, ref Guid riid, out IntPtr ppv)
        {
            var pid = Process.GetCurrentProcess().Id;
            var tid = (int)Native.GetCurrentThreadId();
            string progIdRaw = Native.ProgIdFromClsidSafe(rclsid);
            string clsidStr = rclsid.ToString("D");
            string iidStr = riid.ToString("D");

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
                var dto = new ComEvents.CoGetClassObjectEvent
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoGetClassObject",
                    pid = pid,
                    tid = tid,
                    clsid = clsidStr,
                    progId = ComDictionary.TryResolveProgId(clsidStr, progIdRaw),
                    iid = iidStr,
                    iid_name = ComDictionary.TryResolveIid(iidStr),
                    clsctx = ctx.ToString(),
                    hr = $"0x{hr:X8}",
                    hr_name = ComDictionary.TryResolveHResult($"0x{hr:X8}"),
                    elapsed_ms = sw.Elapsed.TotalMilliseconds,
                    kind = ComDictionary.TryResolveKind(clsidStr, progIdRaw)
                };
                SafeLog(dto);
            }
        }

        private static Guid[] ReadIidsFromMultiQi(IntPtr pResults, uint count)
        {
            if (pResults == IntPtr.Zero || count == 0) return Array.Empty<Guid>();
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
            if (pResults == IntPtr.Zero || count == 0) return Array.Empty<int>();
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
