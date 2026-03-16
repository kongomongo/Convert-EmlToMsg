Invoke-PS2EXE -InputFile "Convert-EmlToMsg.ps1" `
              -OutputFile "EmlToMsgConverter.exe" `
              -NoConsole:$false `
              -IconFile "C:\path\to\your\icon.ico" `   # optional
              -Title "EML to MSG Converter"