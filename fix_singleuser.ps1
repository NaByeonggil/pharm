$ErrorActionPreference='Continue'
Write-Host "=== SQL 단일사용자(-m) 모드 해제 ==="
$paths = @(
 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\MSSQL12.PM2014\MSSQLServer\Parameters',
 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.PM2014\MSSQLServer\Parameters'
)
$removed=$false
foreach($pk in $paths){
  if(Test-Path $pk){
    Write-Host ("검사: {0}" -f $pk)
    $p=Get-ItemProperty $pk
    foreach($pr in ($p.PSObject.Properties | Where-Object { $_.Name -like 'SQLArg*' })){
      Write-Host ("  {0} = {1}" -f $pr.Name,$pr.Value)
      if($pr.Value -like '-m*'){
        Remove-ItemProperty -Path $pk -Name $pr.Name -ErrorAction SilentlyContinue
        Write-Host ("  >> 제거함: {0}" -f $pr.Name)
        $removed=$true
      }
    }
  }
}
if(-not $removed){ Write-Host "레지스트리 -m 없음 (런타임 재시작으로 해제 시도)" }
Write-Host "서비스 재시작 중..."
Stop-Service 'MSSQL$PM2014' -Force -ErrorAction SilentlyContinue
Start-Sleep 2
Start-Service 'MSSQL$PM2014'
Start-Service 'SQLAgent$PM2014' -ErrorAction SilentlyContinue
Start-Sleep 3
Write-Host "--- 확인 (single_user = 0 이면 정상) ---"
sqlcmd -S .\PM2014 -E -Q "SELECT SERVERPROPERTY('IsSingleUser') AS single_user, SUSER_NAME() AS me"
Write-Host ""
Write-Host "single_user = 0 이면 성공입니다. 이제 '유효기간조회' 아이콘을 실행하세요."
Read-Host "닫으려면 Enter"