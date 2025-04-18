Write-Host "Hello $Env:UserName"
$w_dir = "W:"
#$w_dir = "\\tpint60002\ECAD"
#region 檢查W槽連線
$test_result = $true
if (-not (UtilityProgram\Test-Path -Path $w_dir)) {
    Write-Host "Can't connect to W:\, Please check your connection!" -ForegroundColor Red
    $test_result = $false
}

$box_dir = "C:\Users\$Env:UserName\Box"
if (-not (UtilityProgram\Test-Path -Path $box_dir)) {
    Write-Host "Directory not found! $box_dir" -ForegroundColor Red
    $test_result = $false
}

if (-not $test_result) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}
#endregion

# import module in parent dir's Modules folder
Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# 要備份文件的路徑
$path = "footprint_building_aid_skill"

# 備份目標路徑
$backup_dir = "$box_dir\Backup_skill"
$dest_dir = UtilityProgram\Get-NowFolder $backup_dir

UtilityProgram\Copy-WithProgress -Source "$w_dir\$path" -Destination $dest_dir

if ((UtilityProgram\Compress-Folder $dest_dir) -eq $true) {
    Write-Host "Backup Successfully!!" -ForegroundColor Green
}
UtilityProgram\Show-PressAnyKey

<#
$dir = [System.Environment]::CurrentDirectory
Invoke-ps2exe -version 1.0.0.1 "$dir\backup_skill.ps1" "$dir\BackupFlexPAD.exe"
#>