$ts = Get-Date -Format "yyyyMMdd_HHmmss"
.\Test-Injection.ps1 -Mode Attach -TargetPid 3604 `
  -InjectorExe 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe' `
  -DllSource  'C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll' `
  -BinOutDir  'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release' `
  -LogPath    "C:\Temp\ftaelog_$ts.log" `
  -ComTimeoutSeconds 600  # aumente se o ciclo for > 2 min
