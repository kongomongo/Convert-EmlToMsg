Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile "nuget.exe"

.\nuget.exe install MsgKit -OutputDirectory .\Libs