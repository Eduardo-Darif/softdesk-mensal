$nome = "SoftdeskMensal - Diario"

# Troca o gatilho para semanal, de segunda a sexta, as 10:00 (descarta sab/dom).
$gatilho = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "10:00"
Set-ScheduledTask -TaskName $nome -Trigger $gatilho | Out-Null

# Garante que "acordar o PC" e "recuperar execucao perdida" seguem ligados.
$t = Get-ScheduledTask -TaskName $nome
$s = $t.Settings
$s.WakeToRun = $true
$s.StartWhenAvailable = $true
Set-ScheduledTask -TaskName $nome -Settings $s | Out-Null

Write-Host "OK: '$nome' agora roda SEG-SEX as 10:00 (sabado e domingo descartados)."
Write-Host ("Proxima execucao: " + (Get-ScheduledTaskInfo -TaskName $nome).NextRunTime)
Write-Host ""
Write-Host "A tarefa de Fechamento (dia 1 do mes) continua igual, de proposito."
