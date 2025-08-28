using System.Windows.Forms;
using FTAlarmEventSummary;

namespace OcxProbe
{
    // CLSID = e7a3bbdf-71da-4688-8f39-671355104ea5 (AlarmEventSummaryClass)
    public class AlarmEventSummaryAx : AxHost
    {
        public AlarmEventSummaryAx() : base("e7a3bbdf-71da-4688-8f39-671355104ea5") { }
        public AlarmEventSummary GetTyped() { return (AlarmEventSummary)base.GetOcx(); }
    }

    // CLSID = c6379734-8380-40d1-838d-f898ddcc8c1b (AlarmEventBannerClass)
    public class AlarmEventBannerAx : AxHost
    {
        public AlarmEventBannerAx() : base("c6379734-8380-40d1-838d-f898ddcc8c1b") { }
        public AlarmEventBanner GetTyped() { return (AlarmEventBanner)base.GetOcx(); }
    }
}
