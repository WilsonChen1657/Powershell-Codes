Write-Host "Hello $Env:UserName"
Write-Host "Starting recover Allegro settings..."

$w_dir = "\\tpint035\ECAD"
#$w_dir = "\\10.208\ECAD"

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
if (Get-Process allegro) {
    Write-Host "Please close Allegro before recover settings!" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"
#Import-Module "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Modules\UtilityProgram.psm1"

#source directory
$source_dir = "$box_dir\Backup_Allegro_setting\"
$latest_backup_zip = Get-ChildItem $source_dir *.zip | Sort-Object LastWriteTime | Select-Object -Last 1

Expand-Archive -LiteralPath $latest_backup_zip.FullName -DestinationPath $source_dir -Force

$backup_folder = Get-ChildItem $source_dir -Directory | Sort-Object LastWriteTime | Select-Object -Last 1
$backup_files = Get-ChildItem $backup_folder.FullName

#region pcbenv
$source = ($backup_files).Where({ $_.Name -eq "pcbenv" })
$same_count = (Get-ChildItem -Path $Env:HOME pcbenv* -Directory).Count
if ($same_count -gt 0) {
    Rename-Item -Path "$Env:HOME\pcbenv" -NewName "pcbenv,$same_count"
}

Copy-Files -FilePath $source.FullName -Destination $Env:HOME
#endregion

# versions
$version_array = 'SPB_16.6', 'SPB_17.2', 'SPB_17.4'
$cuimenus_path = "share\pcb\text\cuimenus\"
$capture_path = "tools\capture\"
$nclegend_path = "share\pcb\text\nclegend\"
$pcb_path = "share\local\pcb\"
$skill_path = $pcb_path + "skill\"

foreach ($ver in $version_array) {
    $ver_folder = ($backup_files).Where({ $_.Name -eq $ver })
    $file_list = Get-ChildItem -Path $ver_folder.FullName
    foreach ($filename in $file_list) {
        switch ($filename.Name) {
            { ($_ -eq "allegro.men") -or ($_ -eq "pcb_symbol.men") } {
                $file_path = $cuimenus_path
            }
            { ($_ -eq "allegro.cfg") -or ($_ -eq "CAPTURE.INI") } {
                $file_path = $capture_path
            }
            "default-mil.dlt" {
                $file_path = $nclegend_path
            }
            "license_packages_Allegro.txt" {
                $file_path = $pcb_path
            }
            "allegro.ilinit" {
                $file_path = $skill_path
            }
        }

        if (![string]::IsNullOrEmpty($file_path)) {
            $dest = "$Env:HOME\$ver\$file_path"
            #$dest = "C:\Users\tpiwiche\Desktop\restore_test\$ver\$file_path"
            $same_count = (Get-ChildItem -Path $dest $filename*).Count
            $old_file = (Get-ChildItem -Path $dest).Where({ $_.Name -eq $filename.Name }, 'First', 1)
            # rename file when different last write time
            if ($old_file.LastWriteTime -eq $filename.LastWriteTime) {
                Copy-Files -FilePath $filename.FullName -Destination $dest -Force
            }
            else {
                $old_file | Rename-Item -NewName "$filename,$same_count"
                Copy-Files -FilePath $filename.FullName -Destination $dest
            }
        }
    }
}

Write-Host "Recovery complete with $latest_backup_zip" -ForegroundColor Green

Get-ChildItem $source_dir $backup_folder.Name | Remove-Item -Recurse -Confirm:$false

Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
Exit

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro Setting backup"
Invoke-ps2exe -version 1.0.0.0 "$dir\recover_allegro_setting.ps1" "$dir\RecoverAllegroSetting.exe"
#>