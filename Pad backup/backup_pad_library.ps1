# user
Write-Host "Hello $Env:UserName"
$w_dir = "\\tpint035\ECAD"
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

# import module in parent dir's Modules folder
#Import-Module ".\Modules\UtilityProgram.psm1"
Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# backup destination
$backup_dir = "$box_dir\Backup_Flex_PAD\"
$dest_dir = $backup_dir + (Get-Date).ToString("yyyy-MM-dd_HHmmss")
if (!(Test-Path -Path $dest_dir)) {
    New-Item -ItemType directory -Path $dest_dir
}

#data source
<# 批次複製 & 進度條
$path_array = @( "Library-Intel", "Library-Nvida", "Library-Special", "Library-Flex")
foreach ( $path in $path_array) {
    Copy-WithProgress -Source "$w_dir\$path" -Destination "$dest_dir\$path"
}
#>

$path_array = New-Object System.Collections.ArrayList
foreach( $path in @( "Library-Intel", "Library-Nvida", "Library-Special", "Library-Flex")){
    $path_array.Add("$w_dir\$path")
}

Copy-Files -FilePath $path_array -Destination $dest_dir

Compress-Folder -dir $dest_dir

# Delete oldest file when > 3 month
Get-ChildItem -Path $backup_dir -Filter "*.zip" | Where-Object { ($_.LastWriteTime -lt (Get-Date).AddMonths(-3)) } | Remove-Item

<#
$dir = [System.Environment]::CurrentDirectory + "\Pad backup"
Invoke-ps2exe -version 1.0.0.0 "$dir\backup_pad_library.ps1" "$dir\BackupPADLibrary.exe"
#>