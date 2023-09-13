# user
$user_id = [Environment]::UserName
Write-Host "Hello $user_id"
$w_dir = "\\tpint035\ECAD"
if (!(Test-Path -Path $w_dir)) {
    Write-Host "Can't connect to W:\, Please check your connection!" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

$dir = "C:\Users\" + $user_id
if (!(Test-Path -Path $dir)) {
    Write-Host "Directory does not exist $dir" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}
# import module in parent dir's Modules folder
#Import-Module ".\Modules\UtilityProgram.psm1"
Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

# backup destination
$dest_home = "C:\Users\" + $user_id + "\Box\Backup_Flex_PAD\"
$dest_dir = $dest_home + (Get-Date).ToString("yyyy-MM-dd_HHmmss")
if (!(Test-Path -Path $dest_dir)) {
    New-Item -ItemType directory -Path $dest_dir
}

#data source
$path_array = @()
$path_array += "Library-Intel"
$path_array += "Library-Nvida"
$path_array += "Library-Special"
$path_array += "Library-Flex"

foreach ( $path in $path_array) {
    Copy-WithProgress -Source "$w_dir\$path" -Destination "$dest_dir\$path"
}

#Copy-Files -path_array $path_array -dest $dest_dir

Compress-Folder -dir $dest_dir

# Delete oldest file when > 3 month
Get-ChildItem -Path $dest_home -Filter "*.zip" | Where-Object { ($_.LastWriteTime -lt (Get-Date).AddMonths(-3)) } | Remove-Item

<#
$loaction = [System.Environment]::CurrentDirectory + "\Pad backup"
Invoke-ps2exe -version 1.0.0.0 "$loaction\backup_pad_library.ps1" "$loaction\BackupPADLibrary.exe"
#>