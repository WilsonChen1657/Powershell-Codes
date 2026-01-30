# user
Write-Host "Hello $Env:UserName"

#region Check connection
UtilityProgram\Test-PathExist $global:w_dir
$user_dir = "C:\Users\$Env:UserName"
$dest_dir = "$user_dir\OneDrive - Flex"
UtilityProgram\Test-PathExist $dest_dir
#endregion

# Backup destination
$backup_dir = "$dest_dir\Backup\Flex_PAD\"
$temp_backup_dir = UtilityProgram\Get-NowFolder "$user_dir\Documents\TempBackup\Flex_PAD\"

# 批次複製 & 進度條
$excluded_folders = @("footprint_building_aid_skill", "Library-Checking")
$path_list = Get-ChildItem -Path $global:w_dir | 
Where-Object { $_.Name -notin $excluded_folders } | 
Select-Object -ExpandProperty FullName
foreach ( $path in $path_list) {
    UtilityProgram\Copy-WithProgress -Source $path -Destination $temp_backup_dir
}

if ((UtilityProgram\Compress-Folder $temp_backup_dir) -eq $true) {
    $zip_file = "$temp_backup_dir.zip"
    # Copy to Cloud drive
    UtilityProgram\Copy-WithProgress -Source $zip_file -Destination $backup_dir
    # Delete temp file
    Get-ChildItem -Path $zip_file | Remove-Item
    # Delete oldest file when > 3 month
    Get-ChildItem -Path $backup_dir -Filter "*.zip" | Where-Object { ($_.LastWriteTime -lt (Get-Date).AddMonths(-3)) } | Remove-Item
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Pad_backup"
Invoke-ps2exe -version 1.0.0.0 "$dir\backup_pad_library.ps1" "$dir\BackupPADLibrary.exe"
UtilityProgram\Out-EncryptedFile "$dir\backup_pad_library.ps1"
#>