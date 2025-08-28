using System;
using System.Collections.Concurrent;
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

        // ======== NOVO: GUIDs e delegates de interfaces complementares ========
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly Guid IID_IConnectionPointContainer = new Guid("B196B284-BAB4-101A-B69C-00AA00341D07");

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IConnectionPointContainer_FindConnectionPoint(
            IntPtr pThis, ref Guid riid, out IntPtr ppCP /* IConnectionPoint** */);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IConnectionPoint_Advise(
            IntPtr pThis, IntPtr pUnkSink, out int pdwCookie);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IConnectionPoint_Unadvise(
            IntPtr pThis, int dwCookie);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IDispatch_GetIDsOfNames(
            IntPtr pDisp, ref Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
            int cNames, uint lcid,
            [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IDispatch_Invoke(
            IntPtr pDisp, int dispIdMember, ref Guid riid, uint lcid, ushort wFlags,
            IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);

        // ======== NOVO: hooks/estado de interfaces ========
        private static LocalHook _hkFindCP, _hkAdvise, _hkUnadvise, _hkInvoke;
        private static IConnectionPointContainer_FindConnectionPoint _origFindCP;
        private static IConnectionPoint_Advise _origAdvise;
        private static IConnectionPoint_Unadvise _origUnadvise;
        private static IDispatch_GetIDsOfNames _getIDs;
        private static IDispatch_Invoke _origInvoke;

        // pDisp -> (dispId -> nome)
        private static readonly ConcurrentDictionary<IntPtr, ConcurrentDictionary<int, string>> _dispMaps =
            new ConcurrentDictionary<IntPtr, ConcurrentDictionary<int, string>>(1, 64);

        // Métodos alvo (pode expandir)
        private static readonly string[] _names = new[]
        {
            "Acknowledge","AckSelected","Shelve","Unshelve","Refresh","LoadMessages","ApplyFilter"
        };

        // Helper de vtable (sem unsafe)
        private static class Vtbl
        {
            internal static IntPtr GetFunc(IntPtr comPtr, int slot)
            {
                if (comPtr == IntPtr.Zero) return IntPtr.Zero;
                var vtbl = Marshal.ReadIntPtr(comPtr);
                return Marshal.ReadIntPtr(vtbl, slot * IntPtr.Size);
            }
        }

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

                // OBS: hooks de ConnectionPoint/IDispatch serão instalados sob demanda (lazy),
                // quando detectarmos o primeiro objeto relevante (via EnsureInterfaceHooks/TryHookIDispatch...).
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

                // ======== NOVO: instalar ganchos complementares quando criarmos algo relevante ========
                try
                {
                    if (hr == 0 && ppv != IntPtr.Zero)
                    {
                        // classifica a relevância por progId/clsid
                        var progIdResolved = ComDictionary.TryResolveProgId(clsidStr, progIdRaw);
                        var kind = ComDictionary.TryResolveKind(clsidStr, progIdRaw);

                        // Só tentamos quando parecer relevante p/ FTAE/Alarm (mas não é obrigatório filtrar aqui)
                        EnsureInterfaceHooks(ppv);
                        TryHookIDispatchForRelevantObject(progIdResolved, ppv);
                    }
                }
                catch { /* melhor esforço */ }

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
                    clsctx = ComDecode.NormalizeClsctx(ctx.ToString()),
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

                // ======== NOVO: pós-instalação p/ objetos relevantes quando alguma MULTI_QI foi S_OK ========
                try
                {
                    // varre pResults novamente para pegar ppv e hr individuais
                    if (pResults != IntPtr.Zero && dwCount > 0)
                    {
                        int structSize = Marshal.SizeOf(typeof(MULTI_QI));
                        for (int i = 0; i < dwCount; i++)
                        {
                            IntPtr cur = IntPtr.Add(pResults, i * structSize);
                            MULTI_QI qi;
                            try { qi = Marshal.PtrToStructure<MULTI_QI>(cur); } catch { continue; }

                            if (qi.hr == 0 && qi.pItf != IntPtr.Zero)
                            {
                                var progIdResolved = ComDictionary.TryResolveProgId(clsidStr, progIdRaw);
                                EnsureInterfaceHooks(qi.pItf);
                                TryHookIDispatchForRelevantObject(progIdResolved, qi.pItf);
                            }
                        }
                    }
                }
                catch { /* melhor esforço */ }

                var dto = new ComEvents.CoCreateInstanceExEvent
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "CoCreateInstanceEx",
                    pid = pid,
                    tid = tid,
                    clsid = clsidStr,
                    progId = ComDictionary.TryResolveProgId(clsidStr, progIdRaw),
                    clsctx = ComDecode.NormalizeClsctx(ctx.ToString()),
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
                    clsctx = ComDecode.NormalizeClsctx(ctx.ToString()),
                    hr = $"0x{hr:X8}",
                    hr_name = ComDictionary.TryResolveHResult($"0x{hr:X8}"),
                    elapsed_ms = sw.Elapsed.TotalMilliseconds,
                    kind = ComDictionary.TryResolveKind(clsidStr, progIdRaw)
                };
                SafeLog(dto);
            }
        }

        // ======== utilitários MULTI_QI (existentes) ========
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

        // ======== NOVO: instalação lazy de ConnectionPoint ========
        private static void EnsureInterfaceHooks(IntPtr anyComPtr)
        {
            if (_hkFindCP != null && _hkInvoke != null) return; // já instalados ou em uso

            try
            {
                // tenta obter IConnectionPointContainer do objeto atual
                IntPtr pCpc;
                var iidCpc = IID_IConnectionPointContainer; // C# 7.3: evitar ref em static readonly
                int hrqi = Marshal.QueryInterface(anyComPtr, ref iidCpc, out pCpc);

                if (hrqi != 0 || pCpc == IntPtr.Zero) return;

                var addrFind = Vtbl.GetFunc(pCpc, 3); // IConnectionPointContainer::FindConnectionPoint
                Marshal.Release(pCpc);

                if (_hkFindCP == null && addrFind != IntPtr.Zero)
                {
                    _origFindCP = (IConnectionPointContainer_FindConnectionPoint)
                        Marshal.GetDelegateForFunctionPointer(addrFind, typeof(IConnectionPointContainer_FindConnectionPoint));

                    _hkFindCP = LocalHook.Create(addrFind, new IConnectionPointContainer_FindConnectionPoint(FindCP_Hook), null);
                    _hkFindCP.ThreadACL.SetExclusiveACL(Array.Empty<int>());
                    _logger.Info("Hook: IConnectionPointContainer::FindConnectionPoint instalado.");
                }
            }
            catch (Exception ex)
            {
                _logger.Warn("EnsureInterfaceHooks (FindCP) falhou: " + ex.Message);
            }
        }

        private static int FindCP_Hook(IntPtr pThis, ref Guid riid, out IntPtr ppCP)
        {
            int rc = _origFindCP(pThis, ref riid, out ppCP);
            try
            {
                if (rc == 0 && ppCP != IntPtr.Zero)
                {
                    // Instalar (uma vez) Advise/Unadvise nos endereços da vtable do IConnectionPoint retornado
                    var addrAdvise = Vtbl.GetFunc(ppCP, 3); // IConnectionPoint::Advise
                    var addrUnadv = Vtbl.GetFunc(ppCP, 4); // IConnectionPoint::Unadvise

                    if (_hkAdvise == null && addrAdvise != IntPtr.Zero)
                    {
                        _origAdvise = (IConnectionPoint_Advise)Marshal.GetDelegateForFunctionPointer(addrAdvise, typeof(IConnectionPoint_Advise));
                        _hkAdvise = LocalHook.Create(addrAdvise, new IConnectionPoint_Advise(Advise_Hook), null);
                        _hkAdvise.ThreadACL.SetExclusiveACL(Array.Empty<int>());
                        _logger.Info("Hook: IConnectionPoint::Advise instalado.");
                    }

                    if (_hkUnadvise == null && addrUnadv != IntPtr.Zero)
                    {
                        _origUnadvise = (IConnectionPoint_Unadvise)Marshal.GetDelegateForFunctionPointer(addrUnadv, typeof(IConnectionPoint_Unadvise));
                        _hkUnadvise = LocalHook.Create(addrUnadv, new IConnectionPoint_Unadvise(Unadvise_Hook), null);
                        _hkUnadvise.ThreadACL.SetExclusiveACL(Array.Empty<int>());
                        _logger.Info("Hook: IConnectionPoint::Unadvise instalado.");
                    }

                    _logger.Log(new
                    {
                        ts = DateTime.UtcNow.ToString("o"),
                        api = "FindConnectionPoint",
                        iid = riid.ToString("D"),
                        kind = "ftaerel"
                    });
                }
            }
            catch { /* melhor esforço */ }

            return rc;
        }

        private static int Advise_Hook(IntPtr pThis, IntPtr pUnkSink, out int cookie)
        {
            int rc = _origAdvise(pThis, pUnkSink, out cookie);
            try
            {
                _logger.Log(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "IConnectionPoint.Advise",
                    cookie = cookie,
                    kind = "ftaerel"
                });
            }
            catch { }
            return rc;
        }

        private static int Unadvise_Hook(IntPtr pThis, int cookie)
        {
            int rc = _origUnadvise(pThis, cookie);
            try
            {
                _logger.Log(new
                {
                    ts = DateTime.UtcNow.ToString("o"),
                    api = "IConnectionPoint.Unadvise",
                    cookie = cookie,
                    kind = "ftaerel"
                });
            }
            catch { }
            return rc;
        }

        // ======== NOVO: hook de IDispatch::Invoke (alvo: Summary/Banner/AlarmMux) ========
        private static void TryHookIDispatchForRelevantObject(string progId, IntPtr ppvObj)
        {
            if (string.IsNullOrEmpty(progId) || ppvObj == IntPtr.Zero) return;

            // filtro simples por ProgID (pode combinar com 'kind' interno)
            bool isRelevant =
                progId.StartsWith("FTAlarmSummary.", StringComparison.OrdinalIgnoreCase) ||
                progId.StartsWith("RnaAlarmMux.", StringComparison.OrdinalIgnoreCase);

            if (!isRelevant) return;
            IntPtr pDisp;
            var iidDisp = IID_IDispatch; // C# 7.3: evitar ref em static readonly
            var hrqi = Marshal.QueryInterface(ppvObj, ref iidDisp, out pDisp);
            if (hrqi != 0 || pDisp == IntPtr.Zero) return;

            try
            {
                var addrGetIds = Vtbl.GetFunc(pDisp, 5); // IDispatch::GetIDsOfNames
                var addrInvoke = Vtbl.GetFunc(pDisp, 6); // IDispatch::Invoke

                if (addrGetIds != IntPtr.Zero)
                    _getIDs = (IDispatch_GetIDsOfNames)Marshal.GetDelegateForFunctionPointer(addrGetIds, typeof(IDispatch_GetIDsOfNames));
                if (addrInvoke != IntPtr.Zero)
                    _origInvoke = (IDispatch_Invoke)Marshal.GetDelegateForFunctionPointer(addrInvoke, typeof(IDispatch_Invoke));

                if (_hkInvoke == null && addrInvoke != IntPtr.Zero)
                {
                    _hkInvoke = LocalHook.Create(addrInvoke, new IDispatch_Invoke(Invoke_Hook), null);
                    _hkInvoke.ThreadACL.SetExclusiveACL(Array.Empty<int>());
                    _logger.Info("Hook: IDispatch::Invoke instalado em " + progId);
                }

                PrimeDispIds(pDisp);
            }
            catch (Exception ex)
            {
                _logger.Warn("TryHookIDispatchForRelevantObject falhou em " + progId + ": " + ex.Message);
            }
            finally
            {
                Marshal.Release(pDisp);
            }
        }

        private static void PrimeDispIds(IntPtr pDisp)
        {
            if (_getIDs == null) return;

            var map = _dispMaps.GetOrAdd(pDisp, _ => new ConcurrentDictionary<int, string>());
            var dispIds = new int[_names.Length];
            var riidNull = Guid.Empty;

            try
            {
                int hr = _getIDs(pDisp, ref riidNull, _names, _names.Length, 0, dispIds);

                for (int i = 0; i < dispIds.Length; i++)
                {
                    if (dispIds[i] != -1) map[dispIds[i]] = _names[i];
                }

                _logger.Info($"Prime IDs pDisp=0x{pDisp.ToInt64():X}: {string.Join(",", map.Values)}");
            }
            catch (Exception ex)
            {
                _logger.Warn("PrimeDispIds falhou: " + ex.Message);
            }
        }

        private static int Invoke_Hook(
            IntPtr pDisp, int dispIdMember, ref Guid riid, uint lcid, ushort wFlags,
            IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr)
        {
            try
            {
                ConcurrentDictionary<int, string> map;
                string name;
                if (_dispMaps.TryGetValue(pDisp, out map) && map.TryGetValue(dispIdMember, out name))
                {
                    _logger.Log(new
                    {
                        ts = DateTime.UtcNow.ToString("o"),
                        api = "IDispatch.Invoke",
                        method = name,
                        dispId = dispIdMember,
                        wFlags = (int)wFlags,
                        kind = "ftaerel"
                    });
                }
            }
            catch { /* melhor esforço */ }

            return _origInvoke(pDisp, dispIdMember, ref riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
        }
    }
}
