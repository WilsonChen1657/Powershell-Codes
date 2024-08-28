Write-Host "Hello $Env:UserName"
Write-Host "Starting recover Allegro settings..."

#region Check connection
UtilityProgram\Test-PathExist $global:w_dir
$box_dir = "C:\Users\$Env:UserName\Box"
UtilityProgram\Test-PathExist $box_dir
UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist "$Env:HOME\pcbenv"
UtilityProgram\Test-Allegro
#endregion

$backup_dir = "$box_dir\@Backup_Allegro_setting"
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
    Show-PressAnyKey
}
else {
    $folder = $cn_folders[0]
    # when there are more than one dir
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
        Show-PressAnyKey
    }
    else {
        Expand-Archive -LiteralPath $latest_backup_zip.FullName -DestinationPath $backup_dir -Force
        $backup_folder = Get-ChildItem $backup_dir -Directory | Sort-Object LastWriteTime | Select-Object -Last 1
        $backup_files = Get-ChildItem $backup_folder.FullName
        # pcbenv & CIS
        $folder_list = "pcbenv", "cdssetup"
        foreach ($folder_name in $folder_list) {
            $file_info = ($backup_files).Where({ $_.Name -eq $folder_name }) | Select-Object -First 1
            if ($file_info -is [System.IO.DirectoryInfo]) {
                UtilityProgram\Copy-Reversion -Source $file_info -Destination $Env:HOME -MaxRevs 100
            }
        }
        # version files
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
                    #$dest_path = "C:\Users\tpiwiche\Desktop\restore_test\$ver\$dest_path"
                    UtilityProgram\Copy-Reversion -Source $ver_file -Destination $dest_path
                }
            }
        }
        # delete Expand-Archive folder
        Get-ChildItem $backup_dir $backup_folder.Name -Directory | Remove-Item -Recurse -Confirm:$false
        Write-Host "Recovery Successfully!! $latest_backup_zip" -ForegroundColor Green
        Show-PressAnyKey
    }
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.0 "$dir\recover_allegro_setting.ps1" "$dir\RecoverAllegroSetting.exe"

Out-EncryptedFile "$dir\recover_allegro_setting.ps1"
#>