# user
$user_id = [Environment]::UserName
Write-Host "Hello $user_id"
$w_dir = "\\tpint035\ECAD\footprint_building_aid_skill\Backup"
if (!(Test-Path -Path $w_dir)) {
    Write-Host "Can't connect to W:\, Please check your connection!" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

$dir = "C:\Users\" + $user_id
if (!(Test-Path -Path $dir)) {
    Write-Host "Directory does not exist $dir" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

#region Features
function Backup-Item ([string]$path, [string]$dest) {
    if (!(Test-Path -Path $dest)) {
        New-Item -ItemType directory -Path $dest
    }
    
    if ((Test-Path -Path $path) -or (Test-Path -Path $path -PathType Leaf)) {
        write-host "# Backing up $path to $dest"
        Copy-Item -Path $path -Destination $dest -Recurse -Force
    }
    else {
        Write-Warning "$path not found"
    }
}

function Compress-Folder ([string]$dir) {
    # Compress
    write-host "# Compressing $dir"
    $zip = "$dir.zip"
    Compress-Archive -Path $dir -DestinationPath $zip

    # Delete folder
    Remove-Item $dir -Recurse -Force
    return $zip
}

function Show-Message ([string]$msg, [string]$title) {
    # Messagebox
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.Popup($msg, 0, $title, 0 + 64)
}

#endregion

# backup destination
$dest_home = "C:\Users\" + $user_id + "\Box\Backup_Allegro_setting\"
$dest_dir = $dest_home + (Get-Date).ToString("yyyy-MM-dd_HHmmss")

$allegro_home = (Get-ChildItem -Path Env:\HOME).Value + "\"

# PCBENV / ENV / SCRIPT / VIEW / allegro.ilinit(user)
$pcbenv_dir = $allegro_home + "pcbenv\"
Backup-Item -path $pcbenv_dir -dest $dest_dir

# versions
$version_array = 'SPB_16.6', 'SPB_17.2', 'SPB_17.4'

$cuimenus_path = "share\pcb\text\cuimenus\"
$capture_path = "tools\capture\"
$nclegend_path = "share\pcb\text\nclegend\"
$pcb_path = "share\local\pcb\"
$skill_path = $pcb_path + "skill\"

foreach ($ver in $version_array) {
    $dir = $allegro_home + $ver + "\"
    $addr_array = @()
    $addr_array += $dir + $cuimenus_path + "allegro.men"
    $addr_array += $dir + $cuimenus_path + "pcb_symbol.men"
    $addr_array += $dir + $capture_path + "allegro.cfg"
    $addr_array += $dir + $capture_path + "CAPTURE.INI"
    $addr_array += $dir + $nclegend_path + "default-mil.dlt"
    $addr_array += $dir + $pcb_path + "license_packages_Allegro.txt"
    $addr_array += $dir + $skill_path + "allegro.ilinit"

    $dest = $dest_dir + "\" + $ver
    foreach ($addr in $addr_array) {
        Backup-Item -path $addr -dest $dest
    }
}

$computer_info = Get-ComputerInfo
$computer_info | Out-File -FilePath $dest_dir"\computer_info.txt"

$ip_config = ipconfig /all
$ip_config | Out-File -FilePath $dest_dir"\ip_config.txt"

$zip = Compress-Folder -dir $dest_dir
if ((Test-Path -Path $zip -PathType Leaf)) {
    Write-Host "Allegro setting backup complete!!`n$zip" -ForegroundColor Green
}
else {
    Write-Host "Allegro setting backup Fail!!" -ForegroundColor Red
}

Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
#Invoke-ps2exe C:\Users\tpiwiche\Documents\BackupProgram\BackupAllegroSetting.ps1 C:\Users\tpiwiche\Documents\BackupProgram\BackupAllegroSetting.exe