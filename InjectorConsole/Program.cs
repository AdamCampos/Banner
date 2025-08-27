using System;
using System.IO;
using System.Reflection;
using EasyHook;

class Injector
{
    static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Uso: InjectorConsole <PID> <LogPath> [<DllPath>]");
            Console.WriteLine(@"Ex.: InjectorConsole 17656 C:\Temp\ftaelog.log C:\...\ComHookLib_abcd.dll");
            return 1;
        }

        int pid = int.Parse(args[0]);
        string logPath = args[1];

        string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        if (string.IsNullOrEmpty(baseDir))
            baseDir = AppDomain.CurrentDomain.BaseDirectory;

        string hookLib = (args.Length >= 3 && !string.IsNullOrWhiteSpace(args[2]))
            ? args[2]
            : Path.Combine(baseDir, "ComHookLib.dll");

        if (!File.Exists(hookLib))
        {
            Console.WriteLine("ERRO: ComHookLib.dll não encontrado: " + hookLib);
            return 2;
        }

        var dir = Path.GetDirectoryName(logPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        Console.WriteLine($"Injetando no PID {pid} com log em: {logPath}");
        try
        {
            RemoteHooking.Inject(
                pid,
                InjectionOptions.Default,
                hookLib, // x86
                hookLib, // mesmo caminho passado duas vezes (assinatura exige 32/64)
                logPath  // argumento entregue ao remoto
            );
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Falha na injeção: " + ex);
            return 3;
        }
    }
}
