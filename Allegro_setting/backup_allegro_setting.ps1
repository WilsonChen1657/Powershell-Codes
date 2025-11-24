Write-Host "Hello $Env:UserName"
Write-Host "Starting backup Allegro settings..."

#region Check connection
$user_dir = "C:\Users\$Env:UserName"
UtilityProgram\Test-PathExist "$user_dir\Box"
UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist $global:pcbenv_path
UtilityProgram\Test-Allegro
#endregion

# backup destination
$temp_backup_dir = "$user_dir\Documents\TempBackup\Backup_Allegro_setting\"
$backup_dir = "$user_dir\Box\@Backup_Allegro_setting\$Env:COMPUTERNAME"
$dest_dir = UtilityProgram\Get-NowFolder $temp_backup_dir

# Computer information
$computer_info = Get-ComputerInfo
$computer_info | Out-File -FilePath "$dest_dir\computer_info.txt"
$ip_config = ipconfig /all
$ip_config | Out-File -FilePath "$dest_dir\ip_config.txt"

# PCBENV / ENV / SCRIPT / VIEW / allegro.ilinit(user)
# CIS
foreach ($source in $global:pcbenv_path, $global:cdssetup_path) {
    UtilityProgram\Copy-WithProgress -Source $source -Destination $dest_dir
}

foreach ($ver in $global:ver_array) {
    $dir = "$global:cadance_dir\$ver"
    $source_path_array = @()
    foreach ( $path in "allegro.men", "pcb_symbol.men") {
        $source_path_array += "$dir\$global:menus_path\$path"
    }

    $source_path_array += "$dir\$global:capture_path\allegro.cfg"
    $source_path_array += "$dir\$global:nclegend_path\default-mil.dlt"
    $source_path_array += "$dir\$global:pcb_path\license_packages_Allegro.txt"
    $source_path_array += "$dir\$global:skill_path\allegro.ilinit"
    $path_array = @()
    foreach ( $path in $source_path_array) {
        if (Test-Path $path -PathType Leaf) {
            $path_array += $path
        }
    }
    
    foreach ($path in $path_array) {
        UtilityProgram\Copy-WithProgress -Source $path -Destination "$dest_dir\$ver"
    }
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
    $zip_file = "$dest_dir.zip"
    # Copy to box
    UtilityProgram\Copy-WithProgress -Source $zip_file -Destination $backup_dir
    # Delete temp file
    Get-ChildItem -Path $zip_file | Remove-Item
    Write-Host "Backup Successfully!!" -ForegroundColor Green
}
UtilityProgram\Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.2 "$dir\backup_allegro_setting.ps1" "$dir\BackupAllegroSetting.exe"

UtilityProgram\Out-EncryptedFile "$dir\backup_allegro_setting.ps1"
#>