$exe_file = "\\tpint035\ECAD\footprint_building_aid_skill\Backup\programs\BackupAllegroSetting.exe"
if (!(Test-Path -Path $exe_file -PathType Leaf)) {
    Write-Host "$exe_file not found!!" -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Exit
}

$taskName = "備份Allegro設定"
$taskPath = '備份作業'
$ErrorActionPreference = "STOP"
$chkExist = Get-ScheduledTask | Where-Object { $_.TaskName -eq $taskName -and $_.TaskPath -eq "\$taskPath\" }
if ($chkExist) {
    if ($(Read-Host "[$taskName] 已存在，是否刪除? (Y/N)").ToUpper() -eq 'Y') {
        Unregister-ScheduledTask $taskName -Confirm:$false 
    }
    else {
        Write-Host "放棄新增排程作業" -ForegroundColor Yellow
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Exit
    }
}

$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
<#
$action = New-ScheduledTaskAction -Execute $exe_file
$trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Friday -At 5pm
Register-ScheduledTask $taskName -TaskPath $taskPath -Action $action -Trigger $trigger -User $user
#Register-ScheduledTask $taskName -TaskPath $taskPath -Action $action -Trigger $trigger
#>

#New-ScheduledTaskTrigger 不支援 Monthly，改用 schtasks
# 用 PowerShell 陣列轉字串做出以空白間隔的參數字串，含空白的參數值要加雙引號
$taskParams = @(
    "/Create",
    "/TN", "`"$taskPath\$taskName`"",
    "/SC", "monthly",
    "/D", "1", #每個月幾號
    "/ST", "17:00", #開始時間
    "/TR", $exe_file,
    "/RU", $user
)
schtasks.exe $taskParams
Write-Host "新增排程作業完成" -ForegroundColor Green
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
<#
Invoke-ps2exe C:\Users\tpiwiche\Documents\BackupProgram\RegisterBackupScheduledTask.ps1 C:\Users\tpiwiche\Documents\BackupProgram\RegisterBackupScheduledTask.exe
#>