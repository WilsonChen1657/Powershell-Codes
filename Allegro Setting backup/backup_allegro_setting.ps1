# user
$user_id = [Environment]::UserName
Write-Host "Hello $user_id"
$w_dir = "\\tpint035\ECAD"
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

Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# backup destination
$dest_home = "C:\Users\" + $user_id + "\Box\Backup_Allegro_setting\"
$dest_dir = $dest_home + (Get-Date).ToString("yyyy-MM-dd_HHmmss")
if (!(Test-Path -Path $dest_dir)) {
    New-Item -ItemType directory -Path $dest_dir
}

# computer information
$computer_info = Get-ComputerInfo
$computer_info | Out-File -FilePath "$dest_dir\computer_info.txt"

$ip_config = ipconfig /all
$ip_config | Out-File -FilePath "$dest_dir\ip_config.txt"

$allegro_home = (Get-ChildItem -Path Env:\HOME).Value + "\"

# PCBENV / ENV / SCRIPT / VIEW / allegro.ilinit(user)
$pcbenv_dir = $allegro_home + "pcbenv\"
Copy-Files -path $pcbenv_dir -dest $dest_dir

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
    Copy-Files -path $addr_array -dest $dest
}

Compress-Folder -dir $dest_dir
#Invoke-ps2exe -version 1.0.0.0 "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Allegro Setting backup\backup_allegro_setting.ps1" "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Allegro Setting backup\BackupAllegroSetting.exe"