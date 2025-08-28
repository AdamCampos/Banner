using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using FTAlarmEventSummary;

namespace OcxProbe
{
    public partial class Form1 : Form
    {
        private AlarmEventBannerAx _banAx;
        private AlarmEventSummaryAx _sumAx;

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Layout simples: Banner em cima (35%), Summary embaixo (65%)
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 65));
            this.Controls.Add(layout);

            _banAx = new AlarmEventBannerAx();
            _sumAx = new AlarmEventSummaryAx();

            layout.Controls.Add(_banAx, 0, 0);
            layout.Controls.Add(_sumAx, 0, 1);

            // Garante criação de janela/ClientSite/ConnectionPoints
            _banAx.CreateControl();
            _sumAx.CreateControl();

            // Instâncias COM tipadas (RCW)
            var banner = _banAx.GetTyped();
            var summary = _sumAx.GetTyped();

            // (Opcional) ajustes “seguros” antes de testar chamadas
            TrySet(summary, "DisplayErrors", true);
            TrySet(summary, "AutomaticUpdate", false);

            // Assinar TODOS os eventos dinamicamente (usa o seu AttachAllEvents)
            RcwEventProbe.AttachAllEvents(banner, Sink);
            RcwEventProbe.AttachAllEvents(summary, Sink);

            // Exercitar alguns getters “baratos” (não forçam conexão)
            TryLogGet(summary, "ActiveEventCount");
            TryLogGet(summary, "DisplayedEventCount");
            TryLogGet(banner, "ActiveEventCount");

            // Se quiser abrir as propriedades do controle:
            // TryInvoke(summary, "ShowPropertyPages");
        }

        private void Sink(string evName, object[] args)
        {
            var line = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + evName + "  " +
                       (args == null ? "" : string.Join(", ", args));
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rcw_events.txt");
            File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
            Debug.WriteLine(line);
        }

        private static void TrySet(object o, string prop, object val)
        {
            try { var p = o.GetType().GetProperty(prop); if (p != null && p.CanWrite) p.SetValue(o, val, null); }
            catch { /* silencioso para não travar o fluxo */ }
        }

        private static void TryLogGet(object o, string prop)
        {
            try
            {
                var p = o.GetType().GetProperty(prop);
                if (p != null && p.CanRead)
                {
                    var v = p.GetValue(o, null);
                    var line = $"+ GET {o.GetType().Name}.{prop} = {v}";
                    File.AppendAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rcw_run.txt"),
                                       line + Environment.NewLine, Encoding.UTF8);
                    Debug.WriteLine(line);
                }
            }
            catch { }
        }

        // Se quiser reaproveitar o TryInvoke do seu RcwEventProbe, ok;
        // deixei só os utilitários essenciais aqui.
    }
}
