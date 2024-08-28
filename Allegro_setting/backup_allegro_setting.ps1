Write-Host "Hello $Env:UserName"
Write-Host "Starting backup Allegro settings..."

#region Check connection
$box_dir = "C:\Users\$Env:UserName\Box"
UtilityProgram\Test-PathExist $box_dir
UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist $global:pcbenv_path
UtilityProgram\Test-Allegro
#endregion

# backup destination
$backup_dir = "$box_dir\@Backup_Allegro_setting\$Env:COMPUTERNAME"
$dest_dir = UtilityProgram\Get-NowFolder $backup_dir

# Computer information
$computer_info = Get-ComputerInfo
$computer_info | Out-File -FilePath "$dest_dir\computer_info.txt"
$ip_config = ipconfig /all
$ip_config | Out-File -FilePath "$dest_dir\ip_config.txt"

# PCBENV / ENV / SCRIPT / VIEW / allegro.ilinit(user)
UtilityProgram\Copy-Files -FilePath $global:pcbenv_path -Destination $dest_dir
# CIS
UtilityProgram\Copy-Files -FilePath $global:cdssetup_path -Destination $dest_dir

foreach ($ver in $global:ver_array) {
    $dir = "$global:cadance_dir\$ver"
    $path_array = @()
    $path = "$dir\$global:menus_path\allegro.men"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $path = "$dir\$global:menus_path\pcb_symbol.men"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $path = "$dir\$global:capture_path\allegro.cfg"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $path = "$dir\$global:nclegend_path\default-mil.dlt"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $path = "$dir\$global:pcb_path\license_packages_Allegro.txt"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $path = "$dir\$global:skill_path\allegro.ilinit"
    if (Test-Path $path -PathType Leaf) {
        $path_array += $path
    }
    $dest = "$dest_dir\$ver"
    UtilityProgram\Copy-Files -FilePath $path_array -Destination $dest
}

#region Create README
$read_me = "@home:$Env:HOME
env: @home\pcbenv
CAPTURE.INI: @home\cdssetup\OrCAD_Capture\@version`r`n
C:\Cadence\
@version: SPB_16.6, SPB_17.2, SPB_17.4`r`n
license_packages_Allegro.txt: @version\$global:pcb_path
allegro.ilinit: @version\$global:skill_path
allegro.men: @version\$global:menus_path
pcb_symbol.men: @version\$global:menus_path
allegro.cfg: @version\$global:capture_path
default-mil.dlt: @version\$global:nclegend_path
"
$read_me | Out-File "$dest_dir\README.txt"
#endregion

if ((UtilityProgram\Compress-Folder $dest_dir) -eq $true) {
    Write-Host "Backup Successfully!!" -ForegroundColor Green
}
Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.2 "$dir\backup_allegro_setting.ps1" "$dir\BackupAllegroSetting.exe"

Out-EncryptedFile "$dir\backup_allegro_setting.ps1"
#>