using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OcxProbe
{
    static class Program
    {
        private static string OutDir()
        {
            var env = Environment.GetEnvironmentVariable("OCXPROBE_OUTDIR");
            if (!string.IsNullOrWhiteSpace(env))
            {
                try { Directory.CreateDirectory(env); return env; } catch { }
            }
            return AppDomain.CurrentDomain.BaseDirectory;
        }

        [STAThread]
        static void Main()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            Application.ThreadException += (s, e) =>
            {
                try
                {
                    File.AppendAllText(Path.Combine(OutDir(), "rcw_unhandled.txt"),
                        $"[UI] {DateTime.Now:O} {e.Exception}\r\n", Encoding.UTF8);
                }
                catch { }
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try
                {
                    File.AppendAllText(Path.Combine(OutDir(), "rcw_unhandled.txt"),
                        $"[AppDomain] {DateTime.Now:O} {e.ExceptionObject}\r\n", Encoding.UTF8);
                }
                catch { }
            };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
