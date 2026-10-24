param([string]$Source,[string]$OutRoot)
$ErrorActionPreference='Stop'
$sourcePath=(Resolve-Path -LiteralPath $Source).Path
if(Test-Path -LiteralPath $OutRoot){throw 'Choose a new output directory'}
$null=New-Item -ItemType Directory -Path $OutRoot
$before=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$record=[ordered]@{status='FAIL';scope='single prepared cover template, headings, lists, fields and native WPS save';source=$sourcePath;existing_session_preserved=$true}
$app=$null;$doc=$null
try{
 $app=New-Object -ComObject KWPS.Application
 # The user already has WPS documents open. Keep its visibility and application
 # lifetime intact; only this test-owned document is opened and closed.
 $record.documents_before=$app.Documents.Count
 $doc=$app.Documents.Open($sourcePath,$false,$true)
 $record.wps_version=[string]$app.Version
 $scope=@($doc.Bookmarks | Where-Object {$_.Name.StartsWith('HBD_')})
 if($scope.Count -ne 1){throw 'Expected one generated body range'}
 $record.body_scope=[string]$scope[0].Name
 $record.field_update_first=$scope[0].Range.Fields.Update()
 $record.field_update_second=$scope[0].Range.Fields.Update()
 $record.toc_count=$doc.TablesOfContents.Count
 for($i=1;$i -le $doc.TablesOfContents.Count;$i++){$null=$doc.TablesOfContents.Item($i).Update()}
 $labels=[ordered]@{}
 foreach($p in $doc.Paragraphs){
  $text=([string]$p.Range.Text).TrimEnd("`r")
  if($text -in @('槽位回归报告','试验目的','细节','结论','列表甲','列表乙','新列表')){$labels[$text]=[string]$p.Range.ListFormat.ListString}
 }
 $record.labels=$labels
 if($labels['槽位回归报告'] -ne '' -or $labels['试验目的'] -ne '1' -or $labels['细节'] -ne '1.1' -or $labels['结论'] -ne '2'){throw 'Heading numbering differs'}
 if($labels['列表甲'] -ne '1.' -or $labels['列表乙'] -ne '2.' -or $labels['新列表'] -ne '1.'){throw 'List restart differs'}
 $fields=@()
 foreach($field in $doc.Fields){$code=[string]$field.Code.Text;if($code.Trim().StartsWith('SEQ ') -or $code.Trim().StartsWith('REF ')){$fields+=@{code=$code;value=[string]$field.Result.Text}}}
 $record.fields=$fields
 if(-not ($fields | Where-Object {$_.code -match 'SEQ 表' -and $_.value.Trim() -eq '1'})){throw 'Native table SEQ not resolved'}
 if(-not ($fields | Where-Object {$_.code -match 'REF HN_' -and $_.value.Trim() -eq '1'})){throw 'Native REF not resolved'}
 $record.pages=$doc.ComputeStatistics(2)
 $target=Join-Path $OutRoot 'cover-wps.docx';$pdf=Join-Path $OutRoot 'cover-wps.pdf'
 $null=$doc.SaveAs($target,12);$null=$doc.ExportAsFixedFormat($pdf,17)
 $record.output=$target;$record.pdf=$pdf;$record.status='PASS'
}catch{$record.error=$_.Exception.Message;throw}
finally{
 if($null -ne $doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}
 if($null -ne $app){$record.documents_after=$app.Documents.Count;[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
 $record.source_unchanged=((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq $before)
 $record | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $OutRoot 'RESULTS.json') -Encoding utf8
 $record | ConvertTo-Json -Depth 12
}
