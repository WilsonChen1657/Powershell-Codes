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
Import-Module "$w_dir\footprint_building_aid_skill\Backup\Modules\UtilityProgram.psm1"

#data source
$path_array = @()
$path_array += $w_dir + "\Library-Flex"
$path_array += $w_dir + "\Library-Intel"
$path_array += $w_dir + "\Library-Nvida"
$path_array += $w_dir + "\Library-Special"

# backup destination
$dest_home = "C:\Users\" + $user_id + "\Box\Backup_Flex_PAD\"
$dest_dir = $dest_home + (Get-Date).ToString("yyyy-MM-dd_HHmmss")

Copy-Files -path_array $path_array -dest $dest_dir

Compress-Folder -dir $dest_dir
#Invoke-ps2exe -version 1.0.0.0 "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Pad backup\backup_pad_library.ps1" "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Pad backup\BackupPADLibrary.exe"