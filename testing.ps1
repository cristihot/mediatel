Add-Type -AssemblyName System.Windows.Forms

Clear-Host

$input = @()
#$input = import-csv -Path $fileName -Delimiter ';'

$input = import-csv -Path "D:\Kris\PowerShell\!! PROIECTE !!\mediatel\raport_transferate2.csv" -Delimiter ';'
#$input = import-csv -Path "c:\Kris\PowerShell\mediatel\raport_transferate2.csv" -Delimiter ';'

# Write-Host -NoNewline -ForegroundColor Yellow 'Procesare fisier CSV.....'

$nrColoaneAgent    = 0
$nrColoaneDomeniu  = 0
$nrColoaneTalkTime = 0
$nrColoaneCallCode = 0

######## Verifc nr de intrari din array pentru fiecare coloana
foreach ($inputrow in $input) {
  $inputAgent    = $inputrow.Agents   -split ','
  $inputDomeniu  = $inputrow.Domeniu  -split ','
  $inputTalkTime = $inputrow.TalkTime -split ','
  $inputCallCode = $inputrow.Callcode -split ','

  $nrAgent    = $inputAgent.Count
  $nrDomeniu  = $inputDomeniu.Count
  $nrTalkTime = $inputTalkTime.Count
  $nrCallCode = $inputCallCode.Count

  if($nrAgent -gt $nrColoaneAgent)       {$nrColoaneAgent = $nrAgent}
  if($nrDomeniu -gt $nrColoaneDomeniu)   {$nrColoaneDomeniu = $nrDomeniu}
  if($nrTalkTime -gt $nrColoaneTalkTime) {$nrColoaneTalkTime = $nrTalkTime}
  if($nrCallCode -gt $nrColoaneCallCode) {$nrColoaneCallCode = $nrCallCode}

}

Write-Host "`n"

$arraySplit = @();

foreach ($rand in $input) {
  $agenti = $rand.Agents
  $talkTime = $rand.TalkTime

  $dataTable = New-Object -TypeName PSObject
  $dataTable | Add-Member -MemberType NoteProperty -Name "Agenti" -Value $agenti
  $dataTable | Add-Member -MemberType NoteProperty -Name "TalkTime" -Value $talkTime

  $arraySplit += $dataTable
}

#$arraySplit | Format-Table -AutoSize

# for($x = 0; $x -lt $arraySplit.Count; $x++){
#   $splitAgenti = $arraySplit.Agenti[$x] -split ','
#   $splitTime   = $arraySplit.TalkTime[$x] -split ','

#   for($y = 0; $y -lt $splitAgenti.Count; $y++) {
#     if($splitTime[$y].TRIM() -ne 0) {
#       $coloane = New-Object -TypeName PSObject
#       $coloane | Add-Member -MemberType NoteProperty -Name "Linia" -Value $x
#       Write-Host $splitAgenti[$y].TRIM() "-->" $splitTime[$y].TRIM()
#       $arrayFinal += $coloane

#     }
#   }

#   # foreach($inputrow in $splitTime) {
#   #   if($inputrow.TRIM() -ne 0) {
#   #     Write-Host $inputrow.TRIM()
#   #   }
#   # }
#   Write-Host "*******************************"
  
# }

# $arrayFinal | Format-Table -AutoSize


$arrayGugu = @()

for($x=0; $x -lt $arraySplit.Count; $x++) {
  $splitAgenti = $arraySplit.Agenti[$x] -split ','
  $splitTime   = $arraySplit.TalkTime[$x] -split ','
  #Write-Host "Linia$($x) --> $($splitAgenti)"
  for($y=0; $y -lt $splitAgenti[$x].Count; $y++) {
    $agent = $splitAgenti[$y].TRIM()
    #Write-Host $agent
    $coloane = New-Object -TypeName PSObject
    $coloane | Add-Member -MemberType NoteProperty -Name "Coloana$($y)" -Value $agent
    
  }
  $arrayGugu += $coloane
}

$arrayGugu | Format-Table -AutoSize

Write-Host "Done"



break




for($x=0; $x -lt $arraySplit.Count; $x++) {
  $splitAgenti = $arraySplit.Agenti[$x] -split ','
  $splitTime   = $arraySplit.TalkTime[$x] -split ','
  for($y=0; $y -lt $splitAgenti[$x].Count; $y++) {
    $sender   = $splitAgenti[$y].TRIM()
    #$reciever = $splitTime[$y].TRIM()
    $coloane  = New-Object -TypeName PSObject
    $coloane | Add-Member -MemberType NoteProperty -Name "Sender($y)" -Value $sender
  }
  $arrayFinal += $coloane
}

$arrayFinal | Format-Table -AutoSize

break

for($x=0; $x -lt $arraySplit.Count; $x++) {
  $splitAgenti = $arraySplit.Agenti[$x] -split ','
  $splitTime   = $arraySplit.TalkTime[$x] -split ','
  for($y=0; $y -lt $splitAgenti.Count; $y++) {
    if($splitTime[$y].TRIM() -ne 0) {
      $sender   = $splitAgenti[$y].TRIM()
      $reciever = $splitTime[$y].TRIM()
      $coloane  = New-Object -TypeName PSObject
      $coloane | Add-Member -MemberType NoteProperty -Name "Linia" -Value $x
      $coloane | Add-Member -MemberType NoteProperty -Name "Sender" -Value $sender
      $coloane | Add-Member -MemberType NoteProperty -Name "Time" -Value $reciever
      $arrayFinal += $coloane
    }
  }
}


$arrayFinal | Format-Table -AutoSize


