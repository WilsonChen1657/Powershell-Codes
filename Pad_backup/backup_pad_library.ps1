# user
Write-Host "Hello $Env:UserName"

#region Check connection
UtilityProgram\Test-PathExist $global:w_dir
$user_dir = "C:\Users\$Env:UserName"
UtilityProgram\Test-PathExist "$user_dir\Box"
#endregion

# Backup destination
$temp_backup_dir = "$user_dir\Documents\TempBackup\Backup_Flex_PAD\"
$backup_dir = "$user_dir\Box\Backup_Flex_PAD\"
$dest_dir = UtilityProgram\Get-NowFolder $temp_backup_dir

# 批次複製 & 進度條
$excluded_folders = @("footprint_building_aid_skill", "Library-Checking")
$path_list = Get-ChildItem -Path $global:w_dir | 
Where-Object { $_.Name -notin $excluded_folders } | 
Select-Object -ExpandProperty FullName
foreach ( $path in $path_list) {
    UtilityProgram\Copy-WithProgress -Source $path -Destination $dest_dir
}

if ((UtilityProgram\Compress-Folder $dest_dir) -eq $true) {
    $zip_file = "$dest_dir.zip"
    # Copy to box
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