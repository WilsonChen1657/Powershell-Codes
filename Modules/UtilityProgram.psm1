#region Features
function Copy-Files ([Array]$path_array, [string]$dest) {
    if (!(Test-Path -Path $dest)) {
        New-Item -ItemType directory -Path $dest
    }
    
    foreach ( $path in $path_array) {
        if ((Test-Path -Path $path) -or (Test-Path -Path $path -PathType Leaf)) {
            write-host "# Backing up $path to $dest"
            Copy-Item -Path $path -Destination $dest -Recurse -Force
        }
        else {
            Write-Warning "$path not found"
        }
    }
} 

function Compress-Folder ([string]$dir) {
    # Compress
    write-host "# Compressing $dir"
    $zip_file = "$dir.zip"
    Compress-Archive -Path $dir -DestinationPath $zip_file

    # Delete folder
    Remove-Item $dir -Recurse -Force

    if ((Test-Path -Path $zip_file -PathType Leaf)) {
        Write-Host "Compress complete !!`n$zip_file" -ForegroundColor Green
    }
    else {
        Write-Host "Compress Fail !!" -ForegroundColor Red
    }
    
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Show-Message ([string]$msg, [string]$title) {
    # Messagebox
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.Popup($msg, 0, $title, 0 + 64)
}
#endregion