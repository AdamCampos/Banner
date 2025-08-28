.\Test-Injection.ps1 -Mode Smoke `
  -InjectorExe 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe' `
  -DllSource  'C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll' `
  -BinOutDir  'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release' `
  -LogPath    "C:\Temp\ftaelog_smoke_$(Get-Date -Format yyyyMMdd_HHmmss).log"
