Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile "nuget.exe"

.\nuget.exe sources add -Name "nuget.org" -Source "https://api.nuget.org/v3/index.json"

.\nuget.exe install MsgKit -OutputDirectory .\Libs