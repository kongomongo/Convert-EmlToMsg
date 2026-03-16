try {

    # === ROBUST SCRIPT DIRECTORY (works in .ps1 AND compiled .exe - even when only filename is returned) ===
    $ScriptDir = if (-not $PSScriptRoot) {
        # Compiled .exe mode
        $exePath = [Environment]::GetCommandLineArgs()[0]
        if (-not (Split-Path -Parent $exePath)) {
            $exePath = Convert-Path $exePath   # forces full absolute path
        }
        Split-Path -Parent $exePath
    } else {
        # Direct .ps1 mode
        $PSScriptRoot
    }
    $libs = Join-Path $scriptDir "Libs"

    # Load libraries (adjust path if you keep them elsewhere)
    Get-ChildItem $libs -Filter *.dll | ForEach-Object { Add-Type -Path $_.FullName }

    # ====================== PARSE ARGUMENTS ======================
    $files = @()
    $deleteOriginal = $false

    foreach ($arg in $args) {
        if ($arg -ieq "/DEL" -or $arg -ieq "-DEL") {
            $deleteOriginal = $true
            continue
        }
        if (Test-Path -LiteralPath $arg -PathType Leaf) {
            if ($arg -like "*.eml") {
                $files += $arg
            } else {
                Write-Host "⚠️  Skipping non-.eml file: $arg" -ForegroundColor Yellow
            }
        } else {
            Write-Host "❌ File not found: $arg" -ForegroundColor Red
        }
    }

    if ($files.Count -eq 0) {
        Write-Host "No .eml files provided." -ForegroundColor Yellow
        Write-Host "Usage: Drag .eml files onto the .exe or use SendTo" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 0
    }

    # ====================== CONVERSION LOOP ======================
    foreach ($emlPath in $files) {
        $msgPath = [System.IO.Path]::ChangeExtension($emlPath, ".msg")
        
        $emlStream = [System.IO.File]::OpenRead($emlPath)
        $msgStream = [System.IO.File]::Create($msgPath)
        
        [MsgKit.Converter]::ConvertEmlToMsg($emlStream, $msgStream)

        $msgStream.Close()
        $emlStream.Close()
        
        Write-Host "✅ Converted: $msgPath" -ForegroundColor Green
        
        if ($deleteOriginal) {
            Remove-Item -LiteralPath $emlPath -Force
            Write-Host "   (original .eml deleted)" -ForegroundColor DarkGray
        }
    }
}
catch {
    Write-Host "`n=== ERRORS OCCURRED ===" -ForegroundColor Red
    $errorInfo = [PSCustomObject]@{
        Timestamp       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ErrorMessage    = $_.Exception.Message
        FullyQualifiedErrorId = $_.FullyQualifiedErrorId
        Category        = $_.CategoryInfo.Category
        Reason          = $_.CategoryInfo.Reason
        TargetObject    = $_.TargetObject
        ScriptName      = $_.InvocationInfo.ScriptName
        LineNumber      = $_.InvocationInfo.ScriptLineNumber
        Line            = $_.InvocationInfo.Line.Trim()
        PositionMessage = $_.InvocationInfo.PositionMessage
        ExceptionType   = $_.Exception.GetType().FullName
        InnerException  = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $null }
        StackTrace      = $_.ScriptStackTrace
        FullError       = $_ | Out-String  # Complete error record as string
    }

    # Output to console (clear and readable)
    Write-Host "Time:     $($errorInfo.Timestamp)"
    Write-Host "Message:  $($errorInfo.ErrorMessage)"
    Write-Host "Line:     $($errorInfo.Line)"
    Write-Host "Script:   $($errorInfo.ScriptName):$($errorInfo.LineNumber)"
    Write-Host "Position: $($errorInfo.PositionMessage)"
    Write-Host "`nPress Enter to close the window..." -ForegroundColor Yellow
    Read-Host
}
# If everything succeeded → console closes immediately (perfect for SendTo)
