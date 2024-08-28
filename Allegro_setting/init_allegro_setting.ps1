Write-Host "Hello $Env:UserName"
Write-Host "Starting initialize Allegro settings..."

UtilityProgram\Test-PathExist $global:cadance_dir
UtilityProgram\Test-PathExist "$Env:HOME\pcbenv"
UtilityProgram\Test-Allegro
# User select folder
$def_pkg_dir = Select-Folder "W:\footprint_building_aid_skill\Backup\Default_packages\" "Select a default package folder"
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
                #$dest_path = "C:\Users\tpiwiche\Desktop\restore_test\$file\$dest_path"
                UtilityProgram\Copy-Reversion -Source $ver_file -Destination $dest_path -MaxRevs 100
            }
        }
    }
    else {
        $dest_path = ""
        switch ($file.Name) {
            "license_packages_Allegro.txt" {
                $dest_path = $global:pcb_path
            }
            "allegro.ilinit" {
                $dest_path = $global:skill_path
            }
            "env" {
                #default is 17.4
                $dest_path = $global:pcbenv_path
                $max_revs = 100
            }
            "cdssetup" {
                $dest_path = $Env:HOME
                $max_revs = 100
            }
        }
        if (![string]::IsNullOrEmpty($dest_path)) {
            if (!$dest_path.ToLower().Contains($Env:HOME.ToLower())) {
                foreach ($ver in $global:ver_array) {
                    $dest = "$global:cadance_dir\$ver\$dest_path"
                    #$dest_path = "C:\Users\tpiwiche\Desktop\restore_test\$ver\$dest_path"
                    UtilityProgram\Copy-Reversion -Source $file -Destination $dest -MaxRevs 100
                }
            }
            else {
                UtilityProgram\Copy-Reversion -Source $file -Destination $dest_path -MaxRevs 100
            }
            
        }
    }
}

Write-Host "Initialize Successfully!!" -ForegroundColor Green
Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe -version 1.0.0.0 "$dir\init_allegro_setting.ps1" "$dir\InitAllegroSetting.exe"

Out-EncryptedFile "$dir\init_allegro_setting.ps1"
#>