[Console]::OutputEncoding=[Text.Encoding]::UTF8
function B([string]$s){if($null -eq $s){return ''};[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s))}
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
  # PowerShell is started in STA mode. Word is a SingleUse COM server, so each
  # CreateObject request gets a separate WINWORD instance.
  $word=New-Object -ComObject Word.Application
  if($null -eq $word){throw 'Не удалось создать Word.Application.'}
  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.AutomationSecurity=3}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  try{$word.Options.ConfirmConversions=$false}catch{}
  try{$word.Options.CheckSpellingAsYouType=$false}catch{}
  try{$word.Options.CheckGrammarAsYouType=$false}catch{}
  try{$word.Options.SuggestSpellingCorrections=$false}catch{}
  try{$word.Options.BackgroundSave=$false}catch{}
  try{$word.Options.UpdateLinksAtOpen=$false}catch{}
  try{$word.Options.UpdateFieldsAtPrint=$false}catch{}
  try{$word.Options.UpdateLinksAtPrint=$false}catch{}
  try{$word.Options.PrintBackground=$false}catch{}
  $v=$word.Version
  if([string]::IsNullOrWhiteSpace([string]$v)){throw 'Word создан, но не отвечает.'}
  Write-Output 'READY'
  [Console]::Out.Flush()

  while($true){
    $line=[Console]::In.ReadLine()
    if($null -eq $line){break}
    $line=$line.Trim()
    if($line -eq 'QUIT'){break}
    if([string]::IsNullOrWhiteSpace($line)){continue}

    $path=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line))
    $name=[IO.Path]::GetFileName($path)
    $pdf=[IO.Path]::ChangeExtension($path,'.pdf')
    $script:doc=$null
    $om=0.0;$em=0.0;$cm=0.0
    $stage='PREPARE'
    try {
      Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue

      $stage='OPEN'
      $sw=[Diagnostics.Stopwatch]::StartNew()
      # Short late-bound Open call: this is the stable form used by the
      # successful single-Word build.
      $script:doc=$word.Documents.Open($path,$false,$true,$false)
      $sw.Stop();$om=$sw.Elapsed.TotalMilliseconds
      if($null -eq $script:doc){throw 'Word вернул пустую ссылку на документ.'}

      $stage='EXPORT'
      $sw.Restart()
      $script:doc.ExportAsFixedFormat(
        $pdf,
        17,
        $false,
        0,
        0,
        1, 1,
        0,
        $false,
        $false,
        0,
        $false,
        $true,
        $false
      )
      $sw.Stop();$em=$sw.Elapsed.TotalMilliseconds
      if(-not (Test-Path -LiteralPath $pdf)){throw 'Word завершил экспорт без создания PDF.'}

      $stage='CLOSE'
      $sw.Restart()
      $script:doc.Close(0)
      $sw.Stop();$cm=$sw.Elapsed.TotalMilliseconds
      try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:doc)|Out-Null}catch{}
      $script:doc=$null

      Write-Output ('R|'+(B $name)+'|1|'+$om.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$em.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$cm.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|')
      [Console]::Out.Flush()
    } catch {
      $etype=$_.Exception.GetType().FullName
      $hr=''
      try{$hr=('0x{0:X8}' -f ($_.Exception.HResult -band 0xFFFFFFFF))}catch{}
      $msg='Этап '+$stage+': '+$_.Exception.Message
      if(-not [string]::IsNullOrWhiteSpace($etype)){$msg=$msg+' ['+$etype+']'}
      if(-not [string]::IsNullOrWhiteSpace($hr)){$msg=$msg+' HR='+$hr}
      Release-Doc
      Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue
      Write-Output ('R|'+(B $name)+'|0|'+$om.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$em.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+$cm.ToString('F3',[Globalization.CultureInfo]::InvariantCulture)+'|'+(B $msg))
      [Console]::Out.Flush()

      # If the COM server itself is gone, terminate this worker cleanly. The
      # parent will continue the remaining queue on the surviving Word worker.
      $alive=$true
      try{$vv=$word.Version;if([string]::IsNullOrWhiteSpace([string]$vv)){$alive=$false}}catch{$alive=$false}
      if(-not $alive){throw 'Экземпляр Word потерял COM-соединение.'}
    }
  }
} catch {
  $msg=$_.Exception.Message
  $etype=$_.Exception.GetType().FullName
  if(-not [string]::IsNullOrWhiteSpace($etype)){$msg=$msg+' ['+$etype+']'}
  Write-Output ('F|'+(B $msg))
  [Console]::Out.Flush()
  exit 2
} finally {
  Release-Doc
  if($null -ne $word){
    try{$word.Quit(0)}catch{}
    try{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)|Out-Null}catch{}
    $word=$null
  }
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
