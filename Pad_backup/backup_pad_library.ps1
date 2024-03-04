# user
Write-Host "Hello $Env:UserName"

#region Check connection
Test-PathExist $global:w_dir
$box_dir = "C:\Users\$Env:UserName\Box"
Test-PathExist $box_dir
#endregion

# backup destination
$backup_dir = "$box_dir\Backup_Flex_PAD\"
$dest_dir = UtilityProgram\Get-NowFolder $backup_dir

#data source
<# 批次複製 & 進度條
$path_array = @( "Library-Intel", "Library-Nvida", "Library-Special", "Library-Flex")
foreach ( $path in $path_array) {
    Copy-WithProgress -Source "$w_dir\$path" -Destination "$dest_dir\$path"
}
#>

$path_array = New-Object System.Collections.ArrayList
foreach ( $path in @( "Library-Intel", "Library-Nvida", "Library-Special", "Library-Flex")) {
    $path_array.Add("$global:w_dir\$path")
}

UtilityProgram\Copy-Files -FilePath $path_array -Destination $dest_dir

Compress-Folder $dest_dir

# Delete oldest file when > 3 month
Get-ChildItem -Path $backup_dir -Filter "*.zip" | Where-Object { ($_.LastWriteTime -lt (Get-Date).AddMonths(-3)) } | Remove-Item

<#
$dir = [System.Environment]::CurrentDirectory + "\Pad_backup"
Invoke-ps2exe -version 1.0.0.0 "$dir\backup_pad_library.ps1" "$dir\BackupPADLibrary.exe"
Out-EncryptedFile "$dir\backup_pad_library.ps1"
#>