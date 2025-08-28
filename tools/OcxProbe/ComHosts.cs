using System;
using System.Windows.Forms;
using FTAlarmEventSummary; // <- referência ao RCW que você gerou com aximp

namespace OcxProbe
{
    // Host do Banner (usa o CLSID do coclass AlarmEventBannerClass)
    public sealed class AlarmEventBannerHost : AxHost
    {
        public AlarmEventBannerHost()
            : base(typeof(AlarmEventBannerClass).GUID.ToString("B")) { }

        public AlarmEventBanner Ocx => (AlarmEventBanner)GetOcx();
    }

    // Host do Summary (usa o CLSID do coclass AlarmEventSummaryClass)
    public sealed class AlarmEventSummaryHost : AxHost
    {
        public AlarmEventSummaryHost()
            : base(typeof(AlarmEventSummaryClass).GUID.ToString("B")) { }

        public AlarmEventSummary Ocx => (AlarmEventSummary)GetOcx();
    }
}
