Write-Host "Hello $Env:UserName"
Write-Host "Starting recover Allegro settings..."

#region Check connection
UtilityProgram\Test-PathExist $global:w_dir
$user_dir = "C:\Users\$Env:UserName"
UtilityProgram\Test-PathExist "$user_dir\Box"
UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist $Env:HOME
UtilityProgram\Test-Allegro
#endregion

$backup_dir = "$user_dir\Box\@Backup_Allegro_setting"
#region Create select menu
$options = @()
$cn_folders = Get-ChildItem $backup_dir -Directory
$i = 1
$def_key = "x"
foreach ($folder in $cn_folders) {
    $option = "[$i] $folder"
    if ($folder.Name -eq $Env:COMPUTERNAME) {
        $option = "$option (current)"
        $def_key = $i
    }
    $options += $option
    $i++
}
$cn_menu = UtilityProgram\New-SelectMenu 'Recover from computer backup' $options $def_key
#endregion

if (!(Test-Path $backup_dir -PathType Container)) {
    Write-Host "$backup_dir not found!!" -ForegroundColor Red
    UtilityProgram\Show-PressAnyKey
}
else {
    $folder = $cn_folders[0]
    # When there are more than one dir
    if ( $cn_menu.Options.Count -gt 1) {
        do {
            $input_key = $cn_menu.Show()
        }
        until((($input_key -match "^\d+$") -and ($input_key -le $cn_menu.Options.Count)) -or ($input_key -eq "x"))

        if ($input_key -eq "x") {
            return
        }
        $folder = $cn_folders[[int]$input_key - 1]
    }
    $backup_dir = "$backup_dir\$folder"
    $latest_backup_zip = Get-ChildItem $backup_dir *.zip | Sort-Object LastWriteTime | Select-Object -Last 1
    if (!($latest_backup_zip -is [System.IO.FileSystemInfo])) {
        Write-Host "$backup_dir is empty!!" -ForegroundColor Red
        UtilityProgram\Show-PressAnyKey
    }
    else {
        $temp_dir = "$user_dir\Documents\TempBackup\Backup_Allegro_setting\"
        # Copy zip file to temp dir
        UtilityProgram\Copy-WithProgress -Source $latest_backup_zip.FullName -Destination $temp_dir
        $copy_zip = $temp_dir + $latest_backup_zip.Name
        # Expand zip file in temp dir
        UtilityProgram\Expand-ZipFile $copy_zip $temp_dir
        $backup_folder = $temp_dir + $latest_backup_zip.BaseName
        $backup_files = Get-ChildItem $backup_folder
        # pcbenv & CIS
        $folder_array = "pcbenv", "cdssetup"
        foreach ($folder_name in $folder_array) {
            $file_info = ($backup_files).Where({ $_.Name -eq $folder_name }) | Select-Object -First 1
            if ($file_info -is [System.IO.DirectoryInfo]) {
                UtilityProgram\Copy-Reversion -Source $file_info -Destination $Env:HOME -MaxRevs 10
            }
        }
        # Version files
        foreach ($ver in $global:ver_array) {
            $ver_folder = ($backup_files).Where({ $_.Name -eq $ver })
            $ver_files = Get-ChildItem $ver_folder.FullName
            foreach ($ver_file in $ver_files) {
                $dest_path = ""
                switch ($ver_file.Name) {
                    { ($_ -eq "allegro.men") -or ($_ -eq "pcb_symbol.men") } {
                        $dest_path = $global:menus_path
                    }
                    "allegro.cfg" {
                        $dest_path = $global:capture_path
                    }
                    "default-mil.dlt" {
                        $dest_path = $global:nclegend_path
                    }
                    "license_packages_Allegro.txt" {
                        $dest_path = $global:pcb_path
                    }
                    "allegro.ilinit" {
                        $dest_path = $global:skill_path
                    }
                }
                if (![string]::IsNullOrEmpty($dest_path)) {
                    $dest_path = "$cadance_dir\$ver\$dest_path"
                    UtilityProgram\Copy-Reversion -Source $ver_file -Destination $dest_path
                }
            }
        }
        # Remove zip file and temporary extracted backup folder
        Remove-Item $backup_folder -Recurse -Force
        Remove-Item $copy_zip -Recurse -Force

        Write-Host "Recovery Successfully!! $latest_backup_zip" -ForegroundColor Green
        UtilityProgram\Show-PressAnyKey
    }
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.0 "$dir\recover_allegro_setting.ps1" "$dir\RecoverAllegroSetting.exe"

UtilityProgram\Out-EncryptedFile "$dir\recover_allegro_setting.ps1"
#>