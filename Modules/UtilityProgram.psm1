#region Features
function Copy-Files ([Array]$path_array, [string]$dest) {
    if (!(Test-Path -Path $dest)) {
        New-Item -ItemType directory -Path $dest
    }
    
    foreach ( $path in $path_array) {
        if ((Test-Path -Path $path) -or (Test-Path -Path $path -PathType Leaf)) {
            write-host "# Backing up $path to $dest"
            Copy-Item -Path $path -Destination $dest -Recurse -Force
        }
        else {
            Write-Warning "$path not found"
        }
    }
} 

function Compress-Folder ([string]$dir) {
    # Compress
    write-host "# Compressing $dir"
    $zip_file = "$dir.zip"
    Compress-Archive -Path $dir -DestinationPath $zip_file
    
    if ((Test-Path -Path $zip_file -PathType Leaf)) {
        # Delete folder
        Remove-Item $dir -Recurse -Force
        Write-Host "Compress complete !!`n$zip_file" -ForegroundColor Green
    }
    else {
        Write-Host "Compress Fail !!" -ForegroundColor Red
    }
    
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Show-Message ([string]$msg, [string]$title) {
    # Messagebox
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.Popup($msg, 0, $title, 0 + 64)
}

function Copy-WithProgress {
    <#
        .NOTES
        https://stackoverflow.com/a/21209726
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source,
        [Parameter(Mandatory = $true)]
        [string] $Destination,
        [int] $Gap = 200,
        [int] $ReportGap = 2000
    )
    # Define regular expression that will gather number of bytes copied
    $RegexBytes = '(?<=\s+)\d+(?=\s+)';

    #region Robocopy params
    # MIR = Mirror mode
    # NP  = Don't show progress percentage in log
    # NC  = Don't log file classes (existing, new file, etc.)
    # BYTES = Show file sizes in bytes
    # NJH = Do not display robocopy job header (JH)
    # NJS = Do not display robocopy job summary (JS)
    # TEE = Display log in stdout AND in target log file
    $CommonRobocopyParams = '/MIR /NP /NDL /NC /BYTES /NJH /NJS';
    #endregion Robocopy params

    #region Robocopy Staging
    Write-Verbose -Message 'Analyzing robocopy job ...';
    $StagingLogPath = '{0}\temp\{1} robocopy staging.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');

    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
    Write-Verbose -Message ('Staging arguments: {0}' -f $StagingArgumentList);
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow;
    # Get the total number of files that will be copied
    $StagingContent = Get-Content -Path $StagingLogPath;
    $TotalFileCount = $StagingContent.Count - 1;

    # Get the total number of bytes to be copied
    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
    Write-Verbose -Message ('Total bytes to be copied: {0}' -f $BytesTotal);
    #endregion Robocopy Staging

    #region Start Robocopy
    # Begin the robocopy process
    $RobocopyLogPath = '{0}\temp\{1} robocopy.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
    $ArgumentList = '"{0}" "{1}" /LOG:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams;
    Write-Verbose -Message ('Beginning the robocopy process with arguments: {0}' -f $ArgumentList);
    $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -NoNewWindow;
    Start-Sleep -Milliseconds 100;
    #endregion Start Robocopy

    #region Progress bar loop
    while (!$Robocopy.HasExited) {
        Start-Sleep -Milliseconds $ReportGap;
        $BytesCopied = 0;
        $LogContent = Get-Content -Path $RobocopyLogPath;
        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
        $CopiedFileCount = $LogContent.Count - 1;
        Write-Verbose -Message ('Bytes copied: {0}' -f $BytesCopied);
        Write-Verbose -Message ('Files copied: {0}' -f $LogContent.Count);
        $Percentage = 0;
        if ($BytesCopied -gt 0) {
            $Percentage = (($BytesCopied / $BytesTotal) * 100)
        }
        Write-Progress -Activity Robocopy -Status ("Copied {0} of {1} files; Copied {2} of {3} bytes [{4}]" -f $CopiedFileCount, $TotalFileCount, $BytesCopied, $BytesTotal, $Source) -PercentComplete $Percentage
    }
    #endregion Progress loop

    #region Function output
    [PSCustomObject]@{
        Source = $Source;
        Destination = $Destination;
        BytesCopied = $BytesCopied;
        FilesCopied = $CopiedFileCount;
    };
    #endregion Function output
}

#endregion