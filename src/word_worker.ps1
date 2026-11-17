[Console]::OutputEncoding=[Text.Encoding]::UTF8
function B([string]$s){if($null -eq $s){return ''};[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s))}
# Record existing Word processes before COM activation. Hidden Word instances
# may report Application.Hwnd=0, so this snapshot helps identify an
# explicitly new /Automation or -Embedding process without guessing a PID.
$wordIdsBefore=@{}
try {
  foreach($p in @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)){
    $wordIdsBefore[[int]$p.Id]=$true
  }
} catch {}
$word=$null
$doc=$null
function Release-Doc {
  if($null -ne $script:doc){
    try{$script:doc.Close(0)}catch{}
    try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:doc)|Out-Null}catch{}
    $script:doc=$null
  }
}
try {
  $word=New-Object -ComObject Word.Application
  if($null -eq $word){throw 'Cannot create Word.Application.'}
  $ownedPid=0
  # Only a single NEW Word process launched during this COM activation may
  # be tracked. Validate its automation command line before allowing the
  # GUI to terminate it, even if the HWND is present. Never track an existing
  # user Word process only because Word.Application returned its window.
  try {
    $newProcesses=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object {
      -not $wordIdsBefore.ContainsKey([int]$_.Id)
    })
    if($newProcesses.Count -eq 1){
      $candidate=[int]$newProcesses[0].Id
      $proc=Get-CimInstance Win32_Process -Filter ("ProcessId="+$candidate) -ErrorAction Stop
      if(($null -ne $proc) -and ($proc.Name -ieq 'WINWORD.EXE') -and
         ([string]$proc.CommandLine -match '(?i)[/-](?:automation|embedding)(?:\s|$)')){
        $ownedPid=$candidate
      }
    }
  } catch {}
  if($ownedPid -gt 0){Write-Output ('OWNED|0|'+[string]$ownedPid);[Console]::Out.Flush()}
  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.AutomationSecurity=3}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  try{$word.Options.ConfirmConversions=$false}catch{}
  try{$word.Options.CheckSpellingAsYouType=$false}catch{}
  try{$word.Options.CheckGrammarAsYouType=$false}catch{}
  try{$word.Options.BackgroundSave=$false}catch{}
  try{$word.Options.UpdateLinksAtOpen=$false}catch{}
  try{$word.Options.UpdateFieldsAtPrint=$false}catch{}
  try{$word.Options.UpdateLinksAtPrint=$false}catch{}
  try{$word.Options.PrintBackground=$false}catch{}
  if([string]::IsNullOrWhiteSpace([string]$word.Version)){throw 'Word does not respond.'}
  $wordHwnd=0
  try { $wordHwnd=[long]$word.Hwnd } catch {}
  Write-Output ('READY|'+[string]$wordHwnd+'|'+[string]$ownedPid)
  [Console]::Out.Flush()
  while($true){
    $line=[Console]::In.ReadLine()
    if($null -eq $line){break}
    $line=$line.Trim()
    if($line -eq 'QUIT'){break}
    if([string]::IsNullOrWhiteSpace($line)){continue}
    $pair=$line.Split('|')
    if($pair.Length -ne 2){throw 'Invalid input message.'}
    $path=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($pair[0]))
    $pdf=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($pair[1]))
    $name=[IO.Path]::GetFileName($path)
    $script:doc=$null
    $om=0.0;$em=0.0;$cm=0.0
    $stage='PREPARE'
    try {
      Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue
      $stage='OPEN'
      $sw=[Diagnostics.Stopwatch]::StartNew()
      $script:doc=$word.Documents.Open($path,$false,$true,$false)
      $sw.Stop();$om=$sw.Elapsed.TotalMilliseconds
      if($null -eq $script:doc){throw 'Word returned an empty document.'}
      $stage='EXPORT'
      $sw.Restart()
      $script:doc.ExportAsFixedFormat($pdf,17,$false,0,0,1,1,0,$false,$false,0,$false,$true,$false)
      $sw.Stop();$em=$sw.Elapsed.TotalMilliseconds
      if(-not (Test-Path -LiteralPath $pdf)){throw 'Word did not create PDF.'}
      $stage='CLOSE'
      # This Office build can invalidate the original document proxy during export.
      # Reacquire only the exact staged document from this isolated Word instance.
      $currentName=$null
      try{$currentName=[string]$script:doc.FullName}catch{}
      if($currentName -ne $path){
        $replacement=$null
        foreach($candidateDoc in $word.Documents){
          if([string]$candidateDoc.FullName -eq $path){$replacement=$candidateDoc;break}
        }
        if($null -eq $replacement){throw 'Cannot safely reacquire the exported document.'}
        try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:doc)|Out-Null}catch{}
        $script:doc=$replacement
        [Console]::Error.WriteLine('RECOVERED_DOCUMENT_PROXY')
      }
      $sw.Restart()
      $script:doc.Close(0)
      $sw.Stop();$cm=$sw.Elapsed.TotalMilliseconds
      try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:doc)|Out-Null}catch{}
      $script:doc=$null
      Write-Output ('R|'+(B $name)+'|1|'+$om.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$em.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$cm.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|')
      [Console]::Out.Flush()
    } catch {
      $msg='Stage '+$stage+': '+$_.Exception.Message+' ['+$_.Exception.GetType().FullName+']'
      Release-Doc
      Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue
      Write-Output ('R|'+(B $name)+'|0|'+$om.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$em.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$cm.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+(B $msg))
      [Console]::Out.Flush()
      $alive=$true
      try{$v=$word.Version;if([string]::IsNullOrWhiteSpace([string]$v)){$alive=$false}}catch{$alive=$false}
      if(-not $alive){throw 'Word COM connection was lost.'}
    }
  }
} catch {
  Write-Output ('F|'+(B ($_.Exception.Message+' ['+$_.Exception.GetType().FullName+']')))
  [Console]::Out.Flush()
  exit 2
} finally {
  Release-Doc
  if($null -ne $word){
    try{$word.Quit(0)}catch{}
    try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)|Out-Null}catch{}
  }
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

