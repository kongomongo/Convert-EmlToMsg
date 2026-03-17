try {

    # === ROBUST SCRIPT DIRECTORY (works in .ps1 AND compiled .exe - even when only filename is returned) ===
    $ScriptDir = if (-not $PSScriptRoot) {
        # Compiled .exe mode
        $exePath = [Environment]::GetCommandLineArgs()[0]
        if (-not (Split-Path -Parent $exePath)) {
            $exePath = Convert-Path $exePath   # forces full absolute path
        }
        Split-Path -Parent $exePath
    }
    else {
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
        # Skip empty arguments
        if ([string]::IsNullOrWhiteSpace($arg)) { continue }

        # Try to expand as wildcard / folder / file
        $expanded = $null

        try {
            # -LiteralPath when no wildcard → exact match (good for filenames with [])
            # -Path when it looks like it contains wildcards
            if ($arg -match '[\?\*\[\]]') {
                $expanded = Get-ChildItem -Path $arg -File -ErrorAction SilentlyContinue
            }
            else {
                $expanded = Get-ChildItem -LiteralPath $arg -File -ErrorAction SilentlyContinue
            }
        }
        catch {
            # probably invalid path syntax — just skip
        }

        if ($expanded) {
            # Filter only .eml files (case insensitive)
            $emlFiles = $expanded | Where-Object { $_.Extension -ieq '.eml' }
            
            if ($emlFiles.Count -gt 0) {
                $files += $emlFiles.FullName
            }
            else {
                Write-Host "⚠️  No .eml files found in: $arg" -ForegroundColor Yellow
            }
        }
        else {
            # Last chance: maybe it's a single file without wildcard
            if (Test-Path -LiteralPath $arg -PathType Leaf) {
                if ($arg -like "*.eml" -or $arg -like "*.EML") {

                    $files += $arg
                }
                else {
                    Write-Host "⚠️  Skipping non-.eml file: $arg" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "❌ Not found or inaccessible: $arg" -ForegroundColor Red
            }
        }
    }

    # Remove duplicates (in case same file came from multiple patterns)
    $files = $files | Sort-Object -Unique

    if ($files.Count -eq 0) {
        Write-Host "Convert-EmlToMsg v1.1" -ForegroundColor Yellow
        Write-Host "No .eml files provided." -ForegroundColor Yellow
        Write-Host "Usage: Drag .eml files onto the .exe or use SendTo" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 0
    }

    Write-Host "Found $($files.Count) .eml file(s) to convert" -ForegroundColor Cyan

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
            Remove-Item -LiteralPath $emlPath -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path -LiteralPath $emlPath)) {
                Write-Host "   (original .eml deleted)" -ForegroundColor DarkGray
            }
            else {
                Write-Host "   (failed to delete original)" -ForegroundColor DarkYellow
            }
        }
    }

    Write-Host "`nDone." -ForegroundColor Green
}
catch {
    Write-Host "`n=== ERRORS OCCURRED ===" -ForegroundColor Red
    $errorInfo = [PSCustomObject]@{
        Timestamp             = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ErrorMessage          = $_.Exception.Message
        FullyQualifiedErrorId = $_.FullyQualifiedErrorId
        Category              = $_.CategoryInfo.Category
        Reason                = $_.CategoryInfo.Reason
        TargetObject          = $_.TargetObject
        ScriptName            = $_.InvocationInfo.ScriptName
        LineNumber            = $_.InvocationInfo.ScriptLineNumber
        Line                  = $_.InvocationInfo.Line.Trim()
        PositionMessage       = $_.InvocationInfo.PositionMessage
        ExceptionType         = $_.Exception.GetType().FullName
        InnerException        = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $null }
        StackTrace            = $_.ScriptStackTrace
        FullError             = $_ | Out-String  # Complete error record as string
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
