Write-Host "Hello $Env:UserName"
Write-Host "Starting initialize Allegro settings..."

UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist "$Env:HOME\pcbenv"
UtilityProgram\Test-Allegro
# User select folder
$def_pkg_dir = UtilityProgram\Select-Folder "W:\footprint_building_aid_skill\Backup\Default_packages\" "Select a default package folder"
$def_files = Get-ChildItem $def_pkg_dir

foreach ($file in $def_files) {
    if ($global:ver_array.Contains( $file.Name)) {
        $ver_files = Get-ChildItem $file.FullName
        foreach ($ver_file in $ver_files) {
            $dest_path = ""
            switch ($ver_file.Name) {
                { ($_ -eq "allegro.men") -or ($_ -eq "pcb_symbol.men") } {
                    $dest_path = $global:menus_path
                }
                { ($_ -eq "allegro.cfg") } {
                    $dest_path = $global:capture_path
                }
                "default-mil.dlt" {
                    $dest_path = $global:nclegend_path
                }
            }
            if (![string]::IsNullOrEmpty($dest_path)) {
                $dest_path = "$global:cadance_dir\$file\$dest_path"
                UtilityProgram\Copy-Reversion -Source $ver_file -Destination $dest_path
            }
        }
    }
    else {
        $dest_path = ""
        switch ($file.Name) {
            "license_packages_Allegro.txt" {
                $dest_path = $global:pcb_path
                $max_revs = 3
            }
            "allegro.ilinit" {
                $dest_path = $global:skill_path
                $max_revs = 3
            }
            "env" {
                #default is 17.4
                $dest_path = $global:pcbenv_path
                $max_revs = 10
            }
            "cdssetup" {
                $dest_path = $Env:HOME
                $max_revs = 10
            }
        }
        if (![string]::IsNullOrEmpty($dest_path)) {
            if (!$dest_path.ToLower().Contains($Env:HOME.ToLower())) {
                foreach ($ver in $global:ver_array) {
                    $dest = "$global:cadance_dir\$ver\$dest_path"
                    UtilityProgram\Copy-Reversion -Source $file -Destination $dest -MaxRevs $max_revs
                }
            }
            else {
                UtilityProgram\Copy-Reversion -Source $file -Destination $dest_path -MaxRevs $max_revs
            }
        }
    }
}

Write-Host "Initialize Successfully!!" -ForegroundColor Green
UtilityProgram\Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.0 "$dir\init_allegro_setting.ps1" "$dir\InitAllegroSetting.exe"

UtilityProgram\Out-EncryptedFile "$dir\init_allegro_setting.ps1"
#>