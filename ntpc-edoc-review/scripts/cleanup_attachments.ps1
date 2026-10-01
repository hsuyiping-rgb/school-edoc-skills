# 核閱完成後，把本次下載的公文附件移到資源回收筒（可救回，不做永久刪除）。
# 用法：
#   .\cleanup_attachments.ps1 -Since "2026-09-29 11:20"          # 只列出，不動檔案（預設）
#   .\cleanup_attachments.ps1 -Since "2026-09-29 11:20" -Recycle # 確認清單無誤後才加 -Recycle
# -Since 填本次開始下載附件前的時間。先列出清單核對，確認全部是本次附件再執行；
# 清單裡若混入校長自己下載的其他檔案，改用 -Names 逐一指定檔名。
param(
  [Parameter(Mandatory = $true)][datetime]$Since,
  [string[]]$Names,
  [switch]$Recycle
)
Add-Type -AssemblyName Microsoft.VisualBasic
$dl = Join-Path $HOME 'Downloads'
$files = Get-ChildItem $dl -File | Where-Object { $_.LastWriteTime -ge $Since -and $_.Extension -notin '.crdownload', '.tmp' }
if ($Names) { $files = $files | Where-Object { $Names -contains $_.Name } }
$files | Sort-Object LastWriteTime | Format-Table LastWriteTime, Length, Name -AutoSize | Out-String -Width 250
if (-not $Recycle) { "（僅列出，共 $($files.Count) 個；確認後加 -Recycle 執行）"; return }
foreach ($f in $files) {
  [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
}
"已移到資源回收筒：$($files.Count) 個"
