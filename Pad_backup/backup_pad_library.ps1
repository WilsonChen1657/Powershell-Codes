# user
Write-Host "Hello $Env:UserName"

#region Check connection
UtilityProgram\Test-PathExist $global:w_dir
$box_dir = "C:\Users\$Env:UserName\Box"
UtilityProgram\Test-PathExist $box_dir
#endregion

# backup destination
$temp_backup_dir = "C:\Users\$Env:UserName\Documents\TempBackup\Backup_Flex_PAD\"
$backup_dir = "$box_dir\Backup_Flex_PAD\"
$dest_dir = UtilityProgram\Get-NowFolder $temp_backup_dir

#data source

# 批次複製 & 進度條
$excluded_folders = @("footprint_building_aid_skill", "Library-Checking")
$path_list = Get-ChildItem -Path $global:w_dir | 
Where-Object { $_.Name -notin $excluded_folders } | 
Select-Object -ExpandProperty FullName
foreach ( $path in $path_list) {
    UtilityProgram\Copy-WithPorgress -Source $path -Destination $dest_dir
}

if ((UtilityProgram\Compress-Folder $dest_dir) -eq $true) {
    $zip_file = "$dest_dir.zip"
    # Copy to box
    UtilityProgram\Copy-WithPorgress -Source $zip_file -Destination $backup_dir
    # Delete temp file
    Get-ChildItem -Path $zip_file | Remove-Item
    # Delete oldest file when > 3 month
    Get-ChildItem -Path $backup_dir -Filter "*.zip" | Where-Object { ($_.LastWriteTime -lt (Get-Date).AddMonths(-3)) } | Remove-Item
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Pad_backup"
Invoke-ps2exe -version 1.0.0.0 "$dir\backup_pad_library.ps1" "$dir\BackupPADLibrary.exe"
Out-EncryptedFile "$dir\backup_pad_library.ps1"
#>