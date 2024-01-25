Write-Host "Hello $Env:UserName"
Write-Host "Starting backup Allegro settings..."
$w_dir = "\\tpint035\ECAD"

#region Check connection
$test_result = $true
if (!(Test-Path -Path $w_dir)) {
    Write-Host "Can't connect to W:\, Please check your connection!" -ForegroundColor Red
    $test_result = $false
}

$box_dir = "C:\Users\$Env:UserName\Box"
if (!(Test-Path -Path $box_dir)) {
    Write-Host "Directory not found! $box_dir" -ForegroundColor Red
    $test_result = $false
}

if (!$test_result) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}
#endregion

# check Allegro not running
if (Get-Process allegro -ErrorAction SilentlyContinue) {
    Write-Host "Please close Allegro before backup settings!" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# backup destination
$backup_dir = "$box_dir\Backup_Allegro_setting\"
$dest_dir = $backup_dir + (Get-Date).ToString("yyyy-MM-dd_HHmmss")
if (!(Test-Path -Path $dest_dir)) {
    New-Item -ItemType directory -Path $dest_dir
}

#region Computer information
$computer_info = Get-ComputerInfo
$computer_info | Out-File -FilePath "$dest_dir\computer_info.txt"

$ip_config = ipconfig /all
$ip_config | Out-File -FilePath "$dest_dir\ip_config.txt"
#endregion

#region Create README
$read_me = "env: @HOME/pcbenv`r`n
C:\Cadence\
@version: SPB_16.6, SPB_17.2, SPB_17.4`r`n
license_packages_Allegro.txt: @version\share\local\pcb
allegro.ilinit: @version\share\local\pcb\skill
allegro.men: @version\share\pcb\text\cuimenus
pcb_symbol.men: @version\share\pcb\text\cuimenus
CAPTURE.INI: @version\tools\capture
allegro.cfg: @version\tools\capture
default-mil.dlt: @version\share\pcb\text\nclegend"
$read_me | Out-File -FilePath "$dest_dir\README.txt"
#endregion

# PCBENV / ENV / SCRIPT / VIEW / allegro.ilinit(user)
$pcbenv_dir = "$Env:HOME\pcbenv\"
Copy-Files -FilePath $pcbenv_dir -Destination $dest_dir

# versions
$version_array = 'SPB_16.6', 'SPB_17.2', 'SPB_17.4'

$cuimenus_path = "share\pcb\text\cuimenus"
$capture_path = "tools\capture"
$nclegend_path = "share\pcb\text\nclegend"
$pcb_path = "share\local\pcb"
$skill_path = "$pcb_path\skill"

$cadance_dir = "C:\Cadence"
foreach ($ver in $version_array) {
    $dir = "$cadance_dir\$ver"
    $addr_array = @()
    $addr_array += "$dir\$cuimenus_path\allegro.men"
    $addr_array += "$dir\$cuimenus_path\pcb_symbol.men"
    $addr_array += "$dir\$capture_path\allegro.cfg"
    $addr_array += "$dir\$capture_path\CAPTURE.INI"
    $addr_array += "$dir\$nclegend_path\default-mil.dlt"
    $addr_array += "$dir\$pcb_path\license_packages_Allegro.txt"
    $addr_array += "$dir\$skill_path\allegro.ilinit"

    $dest = "$dest_dir\$ver"
    Copy-Files -FilePath $addr_array -Destination $dest
}

Compress-Folder -dir $dest_dir


<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting_backup"
Invoke-ps2exe -version 1.0.0.2 "$dir\backup_allegro_setting.ps1" "$dir\BackupAllegroSetting.exe"
#>