<# Approved Verbs for PowerShell Commands
https://learn.microsoft.com/en-us/powershell/scripting/developer/cmdlet/approved-verbs-for-windows-powershell-commands?view=powershell-7.4

Import-Module "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Modules\UtilityProgram.psm1"
#>

#region Public Porperty
$global:w_dir = "W:" # ("\\tpint60002\ECAD", "\\10.35.35.6\ECAD")
$global:cadance_dir = "C:\Cadence"
# When $Env:HOME has not set yet
if ([string]::IsNullOrEmpty($Env:HOME)) {
    Write-Host "Environment variable 'HOME' not found!" -ForegroundColor Red
    Write-Host "Setting 'HOME' to $global:cadance_dir"
    [Environment]::SetEnvironmentVariable("HOME", $global:cadance_dir, [System.EnvironmentVariableTarget]::User)
    $Env:HOME = [Environment]::GetEnvironmentVariable("HOME", [EnvironmentVariableTarget]::User)
}

$global:ver_array = 'SPB_16.6', 'SPB_17.2', 'SPB_17.4', 'SPB_22.1'
$global:pcbenv_path = "$Env:HOME\pcbenv"
$global:cdssetup_path = "$Env:HOME\cdssetup"
$global:menus_path = "share\local\pcb\menus"
$global:capture_path = "tools\capture"
$global:nclegend_path = "share\pcb\text\nclegend"
$global:pcb_path = "share\local\pcb"
$global:skill_path = "$global:pcb_path\skill"
#endregion

class SelectMenu {
    #region Property
    [string] $Title
    [string[]] $Options
    [string] $Question = "Select an option and press Enter"
    [string] $DefaultKey = "x"
    #endregion

    #region Main
    SelectMenu() { $this.Init(@{}) }
    SelectMenu([hashtable]$Properties) { $this.Init($Properties) }
    #endregion

    #region Function
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }

    [string]Show() {
        Clear-Host
        $banner_line = $(('=') * 16)
        Write-Host "$banner_line " $this.Title " $banner_line"
        foreach ( $option in $this.Options) {
            Write-Host $option
        }
        $input_key = Read-Host $this.Question
        if (-not $input_key) {
            $input_key = $this.DefaultKey
        }

        return $input_key
    }
    #endregion
}

function New-SelectMenu ([string]$Title, [string[]]$Options, [string]$DefaultKey) {
    $menu = [SelectMenu]::new()
    if (-not [string]::IsNullOrEmpty($Title)) {
        $menu.Title = $Title
    }
    if ($Options.Count -gt 0) {
        $menu.Options = $Options
    }
    if (-not [string]::IsNullOrEmpty($DefaultKey)) {
        $menu.DefaultKey = $DefaultKey
    }
    $menu
}

#region Features

#region Check connection
function Test-PathExist ([string]$Path) {
    if (-not (Test-Path $Path)) {
        Write-Host "Path not found! $Path" -ForegroundColor Red
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}
#endregion

function Test-Allegro {
    # check Allegro not running
    while (Get-Process allegro -ErrorAction SilentlyContinue) {
        Write-Host "Please close Allegro before initialize Allegro settings!" -ForegroundColor Red
        Write-Host "Close Allegro and press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

#region get process and kill it
function Close-Process([string]$ProcessName) {
    $process = Get-Process $ProcessName -ErrorAction SilentlyContinue
    if ($process) {
        # try gracefully first
        #$process.CloseMainThread()
        $process.CloseMainWindow()
        # kill after five seconds
        Start-Sleep 5
        if (-not $process.HasExited) {
            $process | Stop-Process -Force
        }
    }
}
#endregion

function Get-NowFolder ([string]$Directory) {
    $now_dir = Join-Path -Path $Directory -ChildPath (Get-Date -Format "yyyy-MM-dd_HHmmss")
    if (-not (Test-Path $now_dir -PathType Container)) {
        New-Item -Path $now_dir -ItemType Directory -Force | Out-Null
    }

    return $now_dir
}

function Compress-Folder ([string]$Directory) {
    # Compress
    write-host "# Compressing $Directory"
    #$zip_path = "$Directory.zip"
    #& "C:\Program Files\7-Zip\7z.exe" a -tzip -mx=9 $zip_path $Directory

    # Compress-Archive will have error
    ## ZipArchiveHelper : The specified path, file name, or both are too long. The fully qualified file name must be less than 260 characters, and the directory name must be less than 248 characters.
    $parent_dir = (Get-Item $Directory).Parent.FullName
    # Set dir to temp Z:\
    subst.exe Z: $parent_dir
    $temp_dir = "Z:\$((Get-Item $Directory).name)"
    Compress-Archive -Path $temp_dir -DestinationPath "$temp_dir.zip"
    # Delete temp Z:\
    subst.exe Z: /d

    $zip_path = "$Directory.zip"
    if (Test-Path $zip_path -PathType Leaf) {
        # Delete folder
        Remove-Item $Directory -Recurse -Force
        Write-Host "Compress complete!!`n$zip_path" -ForegroundColor Green
    }
    else {
        Write-Host "Compress Fail!!" -ForegroundColor Red
        return $false
    }
    return $true
}

function Expand-ZipFile {
    param (
        [string] $Source,
        [string] $Destination
    )
    #& "C:\Program Files\7-Zip\7z.exe" x $Source "-o$(Split-Path $Source)"
    #& "C:\Program Files\7-Zip\7z.exe" x $Source "-o$Destination"
    Expand-Archive $Source $Destination
}

function Show-PressAnyKey {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Show-Message ([string]$Msg, [string]$Title) {
    # Messagebox
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.Popup($Msg, 0, $Title, 0 + 64)
}


function Copy-Reversion ([System.IO.FileSystemInfo]$Source, [string]$Destination, [int]$MaxRevs = 3) {
    if (-not [string]::IsNullOrEmpty($Destination)) {
        $same_files = Get-ChildItem -Path $Destination -Filter "$($Source.Name)*"
        $same_count = $same_files.Count
        $old_file = (Get-ChildItem -Path $Destination).Where({ $_.Name -eq $Source.Name }, 'First', 1)
        # Rename file when different last write time
        if ($old_file.LastWriteTime -ne $Source.LastWriteTime) {
            # Max version files default is ,3
            if ($same_count -gt $MaxRevs) {
                $index = 1
                # Delete oldest file(,1)
                $same_files.Where({ $_.Name -ne $Source.Name }) | Sort-Object Name | Select-Object -First ($same_count - $MaxRevs) | Remove-Item -Recurse -Confirm:$false
                # Refresh the list of files after deletion
                $same_files = Get-ChildItem -Path $Destination -Filter "$($Source.Name)*"
                # Rename files
                foreach ($file in $same_files.Where({ $_.Name -ne $Source.Name }) | Sort-Object Name) {
                    $new_name = ($file.Name -split ",")[0] + ",$index"
                    $file | Rename-Item -NewName $new_name
                    $index++
                }

                $same_count = $index
            }

            $old_file | Rename-Item -NewName "$old_file,$same_count"
        }

        Write-Host "# Copying $Source to $Destination"
        # -PassThru to get copied item
        $copied = Copy-Item $Source.FullName $Destination -Force -Recurse -PassThru
        # to remove ReadOnly flag when Attributes has ReadOnly flag
        foreach ($copy in $copied) {
            if ($copy.Attributes.HasFlag([System.IO.FileAttributes]::ReadOnly)) {
                $copy.Attributes -= [System.IO.FileAttributes]::ReadOnly
            }
        }
    }
}
<#
$path = "$dir\$global:skill_path\allegro.ilinit"
$dest = "$dest_dir\$ver"
Copy-WithProgress -Source $path -Destination $dest

$sourceFolder = "\\tpint60002\ECAD\Library-Special"
$sourceFolder = "W:\Library-Flex"
$destinationDir = "C:\Users\tpiwiche\Documents\TempBackup\Backup_Flex_PAD\2025-04-15_094723"
Copy-WithProgress -Source $sourceFolder -Destination $destinationDir
#>
function Copy-WithProgress {
    param (
        [string] $Source,
        [string] $Destination
    )

    # Create the destination directory if it doesn't exist
    if (-not (Test-Path $Destination -PathType Container)) {
        New-Item -Path $Destination -ItemType Directory -Force | Out-Null
    }

    Write-Host "# Copying $Source to $Destination"
    if (Test-Path $Source -PathType Leaf) {
        $source_data = (Split-Path $Source)
        $file_nm = (Split-Path -Leaf $Source)
        robocopy $source_data $Destination $file_nm /NJH /NJS | % {
            $data = $_.Split([char]9)
            if ("$($data[4])" -ne "") {
                $file = "$($data[4])"
            }
            Write-Progress "Percentage $($data[0])" -Activity "Robocopy $Source" -CurrentOperation $file -ErrorAction SilentlyContinue
        }
    }
    elseif (Test-Path $Source -PathType Container) {
        $dest_dir = Join-Path -Path $Destination -ChildPath (Split-Path -Leaf $Source)
        Copy-FolderWithProgress -Source $Source -Destination $dest_dir
    }
    else {
        Write-Host "Source file not found: $Source" -ForegroundColor Red
        return $false
    }

    Write-Host "Copy complete!" -ForegroundColor Green
    return $true
}

function Copy-FolderWithProgress {
    # NOTES : https://stackoverflow.com/a/21209726
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
    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | ForEach-Object { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
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
    while (-not $Robocopy.HasExited) {
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
        Source      = $Source;
        Destination = $Destination;
        BytesCopied = $BytesCopied;
        FilesCopied = $CopiedFileCount;
    };
    #endregion Function output
}

function Select-Folder([string]$Directory = "", [string]$Description = "Select a folder") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.rootfolder = "MyComputer"
    $dialog.SelectedPath = $Directory
    if ($dialog.ShowDialog() -eq "OK") {
        $folder += $dialog.SelectedPath
    }

    return $folder
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
$dir = [System.Environment]::CurrentDirectory + "\Pad_backup"
Out-EncryptedFile "$dir\register_backup_task.ps1"
#>
function Out-EncryptedFile ([string]$Path) {
    $Key = Get-Content $key_file
    $code = Get-Content $Path -Raw
    $code_secure_string = ConvertTo-SecureString $code -AsPlainText -Force
    $encrypted = ConvertFrom-SecureString -key $Key -SecureString $code_secure_string
    $item = Get-Item $Path
    $new_name = $item.BaseName -replace "_"
    $new_path = $item.DirectoryName + "\" + $new_name.Replace($new_name[0], $new_name[0].ToString().ToUpper()) + $item.Extension
    $encrypted | Out-File -FilePath $new_path
}

$key_file = "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Modules\FlexAES.key"
function Out-AESKey {
    $Key = New-Object Byte[] 16   # You can use 16 (128-bit), 24 (192-bit), or 32 (256-bit) for AES
    [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
    $Key | Out-File $key_file
}

function Set-AccessToDecrypt ([string]$Path, [Array]$Key) {
    $User = $Env:UserName
    New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $Path | ConvertTo-SecureString -Key $Key)
}
#endregion