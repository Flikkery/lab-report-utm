param(
    [Parameter(Mandatory = $true)][string]$In,
    [Parameter(Mandatory = $true)][string]$Out
)
# Рендерит .docx в PDF через Microsoft Word (COM). Пагинация — ровно такая, какую увидит преподаватель в Word.
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($In, $false, $true)
    $doc.Repaginate()
    $doc.ExportAsFixedFormat($Out, 17)
    Write-Output ('Pages: ' + $doc.ComputeStatistics(2))
}
finally {
    if ($null -ne $doc) { $doc.Close(0) }
    if ($null -ne $word) { $word.Quit() }
}
