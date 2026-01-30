Write-Host "Hello $Env:UserName"
$w_dir = "W:" # "\\tpint60002\ECAD"
#region 檢查W槽連線
$test_result = $true
if (!(Test-Path -Path $w_dir)) {
    Write-Host "Can't connect to W:\, Please check your connection!" -ForegroundColor Red
    $test_result = $false
}

$user_dir = "C:\Users\$Env:UserName"
$dest_dir = "$user_dir\OneDrive - Flex"
if (!(Test-Path -Path $dest_dir)) {
    Write-Host "Directory not found! $dest_dir" -ForegroundColor Red
    $test_result = $false
}

if (!$test_result) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}
#endregion

# import module in parent dir's Modules folder
Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# 要備份文件的路徑
$source = "$w_dir\footprint_building_aid_skill"

# 備份目標路徑
$backup_dir = "$dest_dir\Backup\Flex_Skill"
$temp_backup_dir = UtilityProgram\Get-NowFolder "$user_dir\Documents\TempBackup\Flex_Skill\"

UtilityProgram\Copy-WithProgress $source $temp_backup_dir

if ((UtilityProgram\Compress-Folder $temp_backup_dir) -eq $true) {
    $zip_file = "$temp_backup_dir.zip"
    # Copy to Cloud drive
    UtilityProgram\Copy-WithProgress -Source $zip_file -Destination $backup_dir
    # Delete temp file
    Get-ChildItem -Path $zip_file | Remove-Item
    Write-Host "Backup Successfully!!" -ForegroundColor Green
}
UtilityProgram\Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory
Invoke-ps2exe -version 1.0.0.1 "$dir\backup_skill.ps1" "$dir\BackupFlexPAD.exe"
#>