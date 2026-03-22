# 此腳本會將 run_analysis.bat 註冊到 Windows 工作排程器中，每天下午 17:30 自動執行。
# 請以「系統管理員身分 (Run as Administrator)」執行此腳本。

$TaskName = "TaiwanStockAnalysis_5MA"
$Description = "每天下午 17:30 執行台灣股市分析並寄送 Email"
$BatPath = Join-Path -Path $PSScriptRoot -ChildPath "run_analysis.bat"

# 建立執行動作
$Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$BatPath`""

# 建立觸發條件 (每天 17:30)
$Trigger = New-ScheduledTaskTrigger -Daily -At "5:30PM"

# 設定
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries

try {
    # 註冊工作排程
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description $Description -Force
    Write-Host "✅ 成功建立工作排程：$TaskName 於每天下午 17:30 執行。" -ForegroundColor Green
} catch {
    Write-Host "❌ 建立工作排程失敗，請確認您已經使用系統管理員身分執行此 PowerShell。" -ForegroundColor Red
    Write-Host $_.Exception.Message
}
pause
