param(
    [string]$input_key
)


$w_dir = "\\tpint035\ECAD"
if (!(Test-Path $w_dir)) {
    throw "Can't connect to W:\, Please check your VPN connection!"
}

function Show-Menu([string]$Title = 'Allegro setting') {
    Clear-Host
    $banner_line = $(('=') * 16)
    "$banner_line $Title $banner_line"
    '[1] Init Allegro Setting'
    '[2] Register backup task'
    '[3] Backup Allegro Setting'
    '[4] Recover Allegro Setting'
    '[X] Exit  (default is "X")'
}

$w_backup_dir = "$w_dir\footprint_building_aid_skill\Backup"
$default_key = "x"
$temp_key = ""
do {
    if ([string]::IsNullOrEmpty($input_key)) {
        Show-Menu
        if (!($input_key = Read-Host "Which program do you want to run?")) { $input_key = $default_key }
    }
    $program_dir = "$w_backup_dir\Programs"
    #$program_dir = "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Allegro_setting"
    $program_path = ""
    switch ($input_key) {
        "1" { $program_path = "$program_dir\Initallegrosetting.ps1" }
        "2" { $program_path = "$program_dir\Registerbackuptask.ps1" }
        "3" { $program_path = "$program_dir\Backupallegrosetting.ps1" }
        "4" { $program_path = "$program_dir\Recoverallegrosetting.ps1" }
        "pad_library" { $program_path = "$program_dir\Backuppadlibrary.ps1" }
    }

    if (![string]::IsNullOrEmpty($program_path)) {
        # check before import module
        # RemoteSigned : 1 | Restricted : 3
        if ((Get-ExecutionPolicy CurrentUser).value__ -ne 1) {
            Write-Host "Setting ExecutionPolicy to RemoteSigned..."
            Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        }
        if (!(((Get-Module).Where({ $_.Name -eq "UtilityProgram" }) | Select-Object -First 1) -is [PSModuleInfo])) {
            Write-Host "Importing UtilityProgram..."
            Import-Module "$w_backup_dir\Modules\UtilityProgram.psm1"
            #Import-Module "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Modules\UtilityProgram.psm1"
        }
        if (Test-Path $program_path -PathType Leaf) {
            $aes_key = Get-Content "$w_backup_dir\Modules\FlexAES.key"
            UtilityProgram\Set-AccessToDecrypt $program_path $aes_key

            $content = Get-Content $program_path
            $decrypt = $content | ConvertTo-SecureString -Key $aes_key
            $code = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($decrypt))
            Invoke-Expression $code
        }
    }
    $temp_key = $input_key
    $input_key = ""
}
until($temp_key.ToLower() -eq "x")