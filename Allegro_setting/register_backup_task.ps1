Write-Host "Hello $Env:UserName"
Write-Host "Starting register Allegro setting backup task..."

UtilityProgram\Test-PathExist $global:w_dir

$backup_file = "$global:w_dir\footprint_building_aid_skill\Backup\Programs\AllegroSettings.ps1"
#$backup_file = "C:\Users\tpiwiche\Documents\Git\Powershell-Codes\Allegro_setting\AllegroSettings.ps1"
if (!(Test-Path -Path $backup_file -PathType Leaf)) {
    Write-Host "$backup_file not found!!" -ForegroundColor Red
    Show-PressAnyKey
}
else {
    $task_name = "Backup Allegro Settings"
    $task_path = 'Flex Backup'
    $ErrorActionPreference = "STOP"
    $is_task_exist = Get-ScheduledTask | Where-Object { $_.TaskName -eq $task_name -and $_.TaskPath -eq "\$task_path\" }
    if ($is_task_exist) {
        if ((Read-Host "Task [$task_name] already exist, do you want to overwrite? (Y/N)").ToUpper() -eq 'Y') {
            Unregister-ScheduledTask $task_name -Confirm:$false 
            $is_task_exist = $false
        }
        else {
            Write-Host "Exit register task schedule" -ForegroundColor Yellow
            Show-PressAnyKey
        }
    }

    if (!$is_task_exist) {
        <#
        $action = New-ScheduledTaskAction -Execute "Powershell.exe -ExecutionPolicy RemoteSigned -File $backup_file 3"
        $trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Friday -At 5pm
        Register-ScheduledTask $task_name -TaskPath $task_path -Action $action -Trigger $trigger -User $user
        #Register-ScheduledTask $task_name -TaskPath $task_path -Action $action -Trigger $trigger
        #>

        #New-ScheduledTaskTrigger 不支援 Monthly，改用 schtasks
        # 用 PowerShell 陣列轉字串做出以空白間隔的參數字串，含空白的參數值要加雙引號
        $task_params = @(
            "/Create",
            "/TN", "`"$task_path\$task_name`"", #task name
            "/SC", "monthly", #schedule type
            "/D", "1", #day
            "/ST", "14:00", #Start time
            "/TR", "Powershell.exe -ExecutionPolicy RemoteSigned -File $backup_file 3", #task run
            "/RU", [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        )
        schtasks.exe $task_params

        Write-Host "Register task schedule Successfully" -ForegroundColor Green
        Show-PressAnyKey
    }
}

<#
$dir = [System.Environment]::CurrentDirectory + "\Allegro_setting"
Invoke-ps2exe $dir\register_backup_task.ps1 $dir\RegisterBackupTask.exe

Out-EncryptedFile "$dir\register_backup_task.ps1"
#>