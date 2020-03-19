Add-Type -AssemblyName System.Windows.Forms

Clear-Host

#$FileBrowser = New-Object -TypeName System.Windows.Forms.OpenFileDialog -Property @{
#    InitialDirectory = ('D:\Kris\mediatel\') # pentru selectie default la Desktop = [Environment]::GetFolderPath('Desktop')
#    Filter = 'Fisier tip (*.csv)|*.csv'
#}

#$null = $FileBrowser.ShowDialog()

#$fileName = $FileBrowser.FileName
#$initialFolder = $FileBrowser.InitialDirectory

#if($fileName -eq $null) {
#    Write-Host "Nu ai selectat fisierul de procesat!" -ForegroundColor Red
#    #$raspuns = Read-Host
#    break
#}


$input = @()
#$input = import-csv -Path $fileName -Delimiter ';'

$input = import-csv -Path "c:\Kris\PowerShell\mediatel\raport_transferate2.csv" -Delimiter ';'

Write-Host -NoNewline -ForegroundColor Yellow 'Procesare fisier CSV.....'

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

#Write-Host "Numar max col agent: $nrColoaneAgent"
#Write-Host "Numar max col domeniu: $nrColoaneDomeniu"
#Write-Host "Numar max col TalkTime: $nrColoaneTalkTime"
#Write-Host "Numar max col CallCode: $nrColoaneCallCode"

$arraySplit = @();

foreach ($rand in $input) {
  $agenti = $rand.Agents #-split ','
  $talkTime = $rand.TalkTime #-split ','

  $dataTable = New-Object -TypeName PSObject
  $dataTable | Add-Member -MemberType NoteProperty -Name "Agenti" -Value $agenti
  $dataTable | Add-Member -MemberType NoteProperty -Name "TalkTime" -Value $talkTime

  $arraySplit += $dataTable
}

$arrayFinal = @()

for($x=0; $x -lt $arraySplit.Count; $x++) {
  $splitAgenti = $arraySplit.Agenti[$x] -split ','
  $splitTime   = $arraySplit.TalkTime[$x] -split ','
  for($y=0; $y -lt $splitAgenti.Count; $y++) {
    if($splitTime[$y].TRIM() -ne 0) {
      $sender   = $splitAgenti[$y].TRIM()
      $reciever = $splitTime[$y].TRIM()
      #Write-Host "Linia$($x): $($splitAgenti[$y].TRIM())-->$($splitTime[$y].TRIM())"
      $coloane = New-Object -TypeName PSObject
      $coloane | Add-Member -MemberType NoteProperty -Name "Linia" -Value $x
      $coloane | Add-Member -MemberType NoteProperty -Name "Sender" -Value $sender
      $coloane | Add-Member -MemberType NoteProperty -Name "Time" -Value $reciever
      $arrayFinal += $coloane
    }
  }
  #Write-Host $arraySplit.Agenti[$x]
}

$arrayFinal | Format-Table -AutoSize




break



output = @()

foreach ($inputrow in $input) {
  $C1 = $inputrow.Calltraceid
  $C2 = $inputrow.DNIS
  $C3 = $inputrow.ANI
  $C4 = $inputrow.Date
  $C5 = $inputrow.Time
  $C6 = $inputrow.Calltype
  $C7 = $inputrow.ClientName
  $C8 = $inputrow.ClientID
  $inputAgent    = $inputrow.Agents   -split ','
  $inputDomeniu  = $inputrow.Domeniu  -split ','
  $inputTalkTime = $inputrow.TalkTime -split ','
  $inputCallCode = $inputrow.Callcode -split ','

  $coloane = New-Object -TypeName PSObject

  $coloane | Add-Member -MemberType NoteProperty -Name "Calltraceid" -Value $C1
  $coloane | Add-Member -MemberType NoteProperty -Name "DNIS" -Value $C2
  $coloane | Add-Member -MemberType NoteProperty -Name "ANI" -Value $C3
  $coloane | Add-Member -MemberType NoteProperty -Name "Date" -Value $C4
  $coloane | Add-Member -MemberType NoteProperty -Name "Time" -Value $C5
  $coloane | Add-Member -MemberType NoteProperty -Name "Calltype" -Value $C6
  $coloane | Add-Member -MemberType NoteProperty -Name "ClientName" -Value $C7
  $coloane | Add-Member -MemberType NoteProperty -Name "ClientID" -Value $C8

  ######## Agent ########
  For ($i=0; $i -lt $nrColoaneAgent; $i++){
    $coloane | Add-Member -MemberType NoteProperty -Name "Agent$($i+1)" -Value $inputAgent[$i]
  }

  ######## Domeniu ########
  For ($i=0; $i -lt $nrColoaneDomeniu; $i++){
    $coloane | Add-Member -MemberType NoteProperty -Name "Domeniu$($i+1)" -Value $inputDomeniu[$i]
  }

  ######## TalkTime ########
  For ($i=0; $i -lt $nrColoaneTalkTime; $i++){
    $coloane | Add-Member -MemberType NoteProperty -Name "TalkTime$($i+1)" -Value $inputTalkTime[$i]
  }

  ######## CallCode ########
  For ($i=0; $i -lt $nrColoaneCallCode; $i++){
    $coloane | Add-Member -MemberType NoteProperty -Name "CallCode$($i+1)" -Value $inputCallCode[$i]
  }

  $output += $coloane


}

Write-Host -ForegroundColor Green 'OK'

$FileSave = New-Object -TypeName System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = ($initialFolder) # pentru selectie default la Desktop = [Environment]::GetFolderPath('Desktop')
    Filter = 'Fisier tip (*.csv)|*.csv'
}

$null = $FileSave.ShowDialog()

#$output | Export-Csv -NoTypeInformation -Path "$initialFolder\export.csv"

$output | Export-Csv -NoTypeInformation -Path $FileSave.FileName
