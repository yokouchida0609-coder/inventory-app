# ============================================================
# タスクスケジューラーに株価アラートを登録するスクリプト
# （管理者権限不要）
# ============================================================

$scriptPath = "$PSScriptRoot\stock_alert.ps1"
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

# チェック時刻: 9:10 / 12:00 / 16:00（平日のみ）
$triggers = @(
    $(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "09:10"),
    $(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "12:00"),
    $(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "16:00")
)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable

# 既存のタスクを削除してから登録
Unregister-ScheduledTask -TaskName "株価アラート" -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName "株価アラート" `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "日本株高配当ポートフォリオの株価下落アラート（9:10 / 12:00 / 16:00）" `
    -RunLevel Limited

Write-Host ""
Write-Host "✅ タスクスケジューラーへの登録が完了しました！"
Write-Host "   平日の 9:10 / 12:00 / 16:00 に自動チェックします"
Write-Host ""
Write-Host "手動でテスト実行したい場合:"
Write-Host "  powershell -ExecutionPolicy Bypass -File `"$scriptPath`""
