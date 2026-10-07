$nomes = @("SoftdeskMensal - Diario", "SoftdeskMensal - Fechamento")
foreach ($n in $nomes) {
    $t = Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue
    if ($null -eq $t) { Write-Host "NAO ENCONTRADA: $n"; continue }
    $s = $t.Settings
    $s.WakeToRun = $true
    $s.StartWhenAvailable = $true
    Set-ScheduledTask -TaskName $n -Settings $s | Out-Null
    Write-Host "OK: $n  ->  WakeToRun + StartWhenAvailable ligados"
}
Write-Host ""
Write-Host "Pronto. Confira no Agendador de Tarefas: a tarefa -> aba Condicoes (Ativar/Reativar) e Configuracoes."
