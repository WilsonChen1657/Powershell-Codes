# 要備份文件的路徑
$source_path="\\tpint035\ECAD\footprint_building_aid_skill"
#$source_path="\\10.60.4.10\CloudECAD\footprint_building_aid_skill"

# 備份目標路徑
$destination_home="C:\Users\tpiwiche\Box\Wilson_Allegro_data\footprint_building_aid_skill_backup"
$destination_path=$destination_home+"\"+(Get-Date).ToString("yyyy-MM-dd_hhmmss")

write-host "# Proceed backup"
write-host "Source path: $source_path"
write-host "Target path: $destination_path"

# check path exist
if(!(Test-Path -Path $destination_path))
{
    New-Item -ItemType directory -Path $destination_path
}

write-host "# Begin backup"
# start copy
foreach($path in $source_path)
{
    if(Test-Path -Path $path)
    {
        # to seperate two source's
        #$path_array=$path.Trim("\\").Split("\")
        #$dir=$path_array[0]+"_"+$path_array[1]
        write-host "# Backing up $path"
        Copy-Item -Path $path -Destination $destination_path -Recurse -Force
    }
}

# 壓縮
write-host "# Compressing $destination_path"
$zip="$destination_path.zip"
Compress-Archive -Path $destination_path -DestinationPath $zip

#Add-Type -AssemblyName System.IO.Compression
#[System.IO.Compression.ZipFile]::CreateFromDirectory(".\source", ".\src-dotnet.zip", [System.IO.Compression.CompressionLevel]::Optimal, $true)

# 刪除資料夾
Remove-Item $destination_path -Recurse -Force

# Messagebox
$wsh = New-Object -ComObject WScript.Shell
$msg = "Backup complete !!`n$zip"
$wsh.Popup($msg, 0, "Skill backup", 0 + 64)