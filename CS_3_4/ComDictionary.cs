// ComHookLib - ComDictionary.cs
// T3/T4: resolvers e fallback de registro (x86) para CLSID -> ProgID
// Compatível com C# 7.3

using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Win32;

namespace ComHookLib
{
    /// <summary>
    /// Dicionários e utilidades de resolução de nomes de COM (HRESULT, IID, ProgID, "kind").
    /// </summary>
    public static partial class ComDictionary
    {
        // --------------------------------------------------------------------
        // T4) Fallback de registro para CLSID→ProgID (vista x86)
        // --------------------------------------------------------------------

        /// <summary>
        /// Lê HKCR\CLSID\{clsid}\ProgID (merge), depois HKCU/HKLM\Software\Classes na vista de 32-bit.
        /// Retorna o ProgID ou null se não encontrado.
        /// </summary>
        public static string TryRegistryProgId(string clsid)
        {
            var norm = NormalizeGuidD(clsid);
            if (norm == null) return null;

            var rel = @"CLSID\{" + norm + @"}\ProgID";

            // 1) HKCR (merge) – em processo x86 normalmente já reflete a vista correta
            try
            {
                using (var k = Registry.ClassesRoot.OpenSubKey(rel))
                {
                    var v = k == null ? null : k.GetValue(null) as string; // (Default)
                    if (!string.IsNullOrWhiteSpace(v))
                        return v;
                }
            }
            catch
            {
                // ignore
            }

            // 2) HKCU\Software\Classes (x86)
            var vUser = ReadClasses32(RegistryHive.CurrentUser, rel);
            if (!string.IsNullOrWhiteSpace(vUser))
                return vUser;

            // 3) HKLM\Software\Classes (x86)
            var vMachine = ReadClasses32(RegistryHive.LocalMachine, rel);
            if (!string.IsNullOrWhiteSpace(vMachine))
                return vMachine;

            return null;
        }

        private static string ReadClasses32(RegistryHive hive, string subKeyUnderClasses)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry32))
                using (var k = baseKey.OpenSubKey(@"Software\Classes\" + subKeyUnderClasses))
                {
                    return k == null ? null : (k.GetValue(null) as string);
                }
            }
            catch
            {
                return null;
            }
        }

        // --------------------------------------------------------------------
        // Resolvedores públicos usados pelo logger/enrichment
        // --------------------------------------------------------------------

        /// <summary>
        /// Traduz HRESULT ("0x........") para um nome amigável, ex.: "S_OK".
        /// </summary>
        public static string TryResolveHResult(string hr)
        {
            if (string.IsNullOrWhiteSpace(hr)) return null;

            var canon = CanonHResult(hr);
            string name;
            if (_hresultNames.TryGetValue(canon, out name))
                return name;

            return null;
        }

        /// <summary>
        /// Traduz IID (GUID) para nome conhecido. Fallback para "IID_XXXXXXXX" (8 hex iniciais).
        /// </summary>
        public static string TryResolveIid(string iid)
        {
            var norm = NormalizeGuidD(iid);
            if (norm == null) return null;

            string name;
            if (_iidNames.TryGetValue(norm, out name))
                return name;

            // Fallback legível que você já vinha usando
            return "IID_" + norm.Substring(0, 8);
        }

        /// <summary>
        /// Resolve ProgID a partir do valor já capturado OU via dicionário interno OU (T4) via registro x86.
        /// </summary>
        public static string TryResolveProgId(string clsid, string progId)
        {
            if (!string.IsNullOrWhiteSpace(progId))
                return progId;

            var norm = NormalizeGuidD(clsid);

            // 1) Mapa interno
            if (!string.IsNullOrEmpty(norm))
            {
                string mapped;
                if (_clsid2ProgId.TryGetValue(norm, out mapped) && !string.IsNullOrWhiteSpace(mapped))
                    return mapped;
            }

            // 2) Fallback (T4) – Registro (x86)
            return TryRegistryProgId(clsid);
        }

        /// <summary>
        /// Classificador leve do "kind" para filtros/relatórios.
        /// - "ftaerel" quando ProgID denota FTAlarm*/RnaAe*/AlarmMux*/FTAE
        /// - "maybe" caso contrário.
        /// </summary>
        public static string TryResolveKind(string clsid, string progId)
        {
            // Tenta com o progId fornecido
            if (IsFtaeRelevant(progId))
                return "ftaerel";

            // Sem progId: dá uma olhada no mapa interno para inferir
            var norm = NormalizeGuidD(clsid);
            if (!string.IsNullOrEmpty(norm))
            {
                string mapped;
                if (_clsid2ProgId.TryGetValue(norm, out mapped))
                {
                    if (IsFtaeRelevant(mapped))
                        return "ftaerel";
                }
            }

            return "maybe";
        }

        private static bool IsFtaeRelevant(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId)) return false;
            var p = progId.Trim().ToLowerInvariant();

            // Heurística simples e estável:
            return p.StartsWith("ftalarm", StringComparison.Ordinal) ||
                   p.StartsWith("rnaae", StringComparison.Ordinal) ||
                   p.Contains("ftae") ||
                   p.Contains("alarmmux");
        }

        // --------------------------------------------------------------------
        // Helpers de normalização
        // --------------------------------------------------------------------

        private static string NormalizeGuidD(string v)
        {
            if (string.IsNullOrWhiteSpace(v)) return null;
            v = v.Trim().Trim('{', '}');
            Guid g;
            if (!Guid.TryParse(v, out g))
                return null;

            return g.ToString("D").ToLowerInvariant();
        }

        private static string CanonHResult(string hr)
        {
            var s = hr.Trim();
            if (s.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                uint x;
                if (uint.TryParse(s.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out x))
                    return "0x" + x.ToString("x8");
                return s.ToLowerInvariant();
            }
            // Tenta decimal
            int dec;
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out dec))
            {
                unchecked
                {
                    return "0x" + ((uint)dec).ToString("x8");
                }
            }
            return s.ToLowerInvariant();
        }

        // --------------------------------------------------------------------
        // Dicionários (T3) - estenda conforme necessário
        // --------------------------------------------------------------------

        // HRESULT básicos usados recorrentemente
        private static readonly Dictionary<string, string> _hresultNames =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "0x00000000", "S_OK" },
                { "0x80004001", "E_NOTIMPL" },
                { "0x80004002", "E_NOINTERFACE" },
                { "0x80004003", "E_POINTER" },
                { "0x80004005", "E_FAIL" },
                { "0x80070005", "E_ACCESSDENIED" },
                { "0x80040110", "CLASS_E_NOAGGREGATION" },
                { "0x80040111", "CLASS_E_CLASSNOTAVAILABLE" },
                { "0x80040112", "CLASS_E_NOTLICENSED" },
                { "0x80040154", "REGDB_E_CLASSNOTREG" },
                { "0x800401F0", "CO_E_NOTINITIALIZED" },
                { "0x800401F3", "CO_E_CLASSSTRING" },
                { "0x800401F6", "CO_E_INVALIDAPARTMENT" },
                { "0x8007000E", "E_OUTOFMEMORY" }
            };

        // Alguns IIDs comuns + o do seu log recente (000001fd…)
        // IIDs básicos e comuns (corrigidos + ampliados)
        private static readonly Dictionary<string, string> _iidNames =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
    // Núcleo COM
    { "00000000-0000-0000-c000-000000000046", "IUnknown" },
    { "00000001-0000-0000-c000-000000000046", "IClassFactory" },
    { "00000003-0000-0000-c000-000000000046", "IMarshal" },

    // Persistência / Monikers / ROT
    { "0000000b-0000-0000-c000-000000000046", "IStorage" },
    { "0000000c-0000-0000-c000-000000000046", "IStream" },
    { "0000000e-0000-0000-c000-000000000046", "IBindCtx" },
    { "0000000f-0000-0000-c000-000000000046", "IMoniker" },
    { "00000010-0000-0000-c000-000000000046", "IRunningObjectTable" },

    { "00000109-0000-0000-c000-000000000046", "IPersistStream" },
    { "0000010b-0000-0000-c000-000000000046", "IPersistFile" },
    { "0000010c-0000-0000-c000-000000000046", "IPersist" },

    // OLE / Dados
    { "0000010e-0000-0000-c000-000000000046", "IDataObject" },
    { "00000112-0000-0000-c000-000000000046", "IOleObject" },

    // Automation / TypeLib
    { "00020400-0000-0000-c000-000000000046", "IDispatch" },
    { "00020401-0000-0000-c000-000000000046", "ITypeInfo" },
    { "00020402-0000-0000-c000-000000000046", "ITypeLib" },
    { "00020403-0000-0000-c000-000000000046", "ITypeComp" },
    { "00020404-0000-0000-c000-000000000046", "IEnumVARIANT" },

    // Serviços / Class Info / Conexões
    { "6d5140c1-7436-11ce-8034-00aa006009fa", "IServiceProvider" },
    { "b196b283-bab4-101a-b69c-00aa00341d07", "IProvideClassInfo" },
    { "b196b284-bab4-101a-b69c-00aa00341d07", "IConnectionPointContainer" },
    { "b196b285-bab4-101a-b69c-00aa00341d07", "IEnumConnectionPoints" },
    { "b196b286-bab4-101a-b69c-00aa00341d07", "IConnectionPoint" },
    { "b196b287-bab4-101a-b69c-00aa00341d07", "IEnumConnections" },

    // Erros (Automation)
    { "1cf2b120-547d-101b-8e65-08002b2bd119", "IErrorInfo" },
    { "df0b3d60-548f-101b-8e65-08002b2bd119", "ISupportErrorInfo" },

    // Extensão útil
    { "a6ef9860-c720-11d0-9337-00a0c90dcaa9", "IDispatchEx" },

    // O que você já tinha citado no log
    { "00000146-0000-0000-c000-000000000046", "IGlobalInterfaceTable" },
    { "000001fd-0000-0000-c000-000000000046", "IComCatalogSCM" } // se aplicável no seu ambiente
        };


        // Mapeamentos internos CLSID -> ProgID (adicione os seus FTAE aqui conforme descubra/precise)
        private static readonly Dictionary<string, string> _clsid2ProgId =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Exemplos ilustrativos:
                // { "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx", "FTAlarm.Manager" },
                // { "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy", "RnaAe.Provider" },
                // { "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz", "AlarmMux.Core" }
            };
    }
}
