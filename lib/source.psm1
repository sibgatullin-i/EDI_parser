function Connect-Mail {
  param(
    [Parameter(Mandatory)][string]$Server,
    [Parameter(Mandatory)][string]$Port,
    [Parameter(Mandatory)][string]$Username,
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory=$false)][string]$enableSSL = $true
  )
  $pop3Client = New-Object OpenPop.Pop3.Pop3Client
  $pop3Client.connect( $server, $port, $enableSSL )
  if ( !$pop3Client.connected ) {
      throw "Unable to create POP3 client. Connection failed with server $server"
  }
  try { $pop3Client.authenticate( $username, $password ) }
  catch { return $false }
  return $pop3Client
}

function Check-Mail {
  param(
    [Parameter(Mandatory)][OpenPop.Pop3.Pop3Client]$pop3Client,
    [Parameter(Mandatory)]$From
  )
  $messageCount = $pop3Client.getMessageCount()
  $targetMessages = @()
  for ($currentIndex = $messageCount; $currentIndex -gt 0; $currentIndex--){
    $messageFrom = $pop3Client.getMessage($currentIndex).Headers.From.Address
    $messageAttachment = $pop3Client.GetMessage($currentIndex).FindAllAttachments().count
    if ($From.Contains($messageFrom) -and $messageAttachment -gt 0) {
      $targetMessages += [pscustomobject]@{index = $currentIndex; target = $true}
    } else {
      $targetMessages += [pscustomobject]@{index = $currentIndex; target = $false}
    }
  }
  write-host "$messageCount total messages"
  write-host "$(($targetMessages | where-Object {$_.target -eq $true}).count ) from $From with attachments"
  return $targetMessages
}

function saveAttachment {
   Param
      (
      [System.Net.Mail.Attachment] $attachment,
      [string] $Path
      )
   New-Item -Path $Path -ItemType "File" -Force | Out-Null
   $outStream = New-Object IO.FileStream $Path, "Create"
   $attachment.contentStream.copyTo( $outStream )
   $outStream.close()
  }

function FetchAndSave-Attachment {
  Param (
    [OpenPop.Pop3.Pop3Client] $pop3Client,
    [string] $Folder,
    [int]$messageIndex
  )
  $uid = $pop3Client.getMessageUid( $messageIndex )
  $incomingMessage = $pop3Client.getMessage( $messageIndex ).toMailMessage()
  foreach ($attachment in $incomingMessage.Attachments) {
    $attachmentURL = Join-Path -Path $Folder -ChildPath "$(get-date -Format 'yyyyMMdd_hhMMssffff')-_-$($attachment.name)"
    Write-Host "`tSaving attachment to:" $attachmentURL
    saveAttachment $attachment $attachmentURL
  }
}

function Parse-HTML {
  param( [Parameter(Mandatory)][string]$Path )
  if (!(Test-Path $path)) {write-warning "File $Path not found"; return $false}
  $HTML = (Get-Content $Path | ConvertFrom-Html)
  $HTML = $HTML.SelectNodes('//table') | where-object {$_.InnerText -like "event*"}
  $HTML = $HTML.SelectNodes('tr')
  $HTMLheader = $HTML.selectnodes('//table')[0].selectnodes('tr')[0].innertext
  $HTMLdate = $HTMLheader.Substring($HTMLheader.length - 10)
  $tableHeader = $HTML[0].SelectNodes('td').InnerText
  $HTML = $HTML | where-object {$_.InnerText -notlike "event*"}
  $array = @()
  foreach ($line in $HTML) {
    $array += [pscustomobject]@{
      $tableHeader[0] = $line.SelectNodes('td')[0].InnerText
      $tableHeader[1] = $line.SelectNodes('td')[1].InnerText
      $tableHeader[2] = $line.SelectNodes('td')[2].InnerText
      $tableHeader[3] = $line.SelectNodes('td')[3].InnerText
      $tableHeader[4] = $line.SelectNodes('td')[4].InnerText
      $tableHeader[5] = $line.SelectNodes('td')[5].InnerText
      $tableHeader[6] = $line.SelectNodes('td')[6].InnerText
      $tableHeader[7] = $line.SelectNodes('td')[7].InnerText
      $tableHeader[8] = $line.SelectNodes('td')[8].InnerText
      Url = ($line.InnerHtml -split "'" | Where-Object {$_ -like "http*"})
    }
  }
  Return $HTMLheader,$HTMLdate,$array
}

function Download-Pages {
  param(
    [Parameter(Mandatory)]$sourceData,
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$Prefix,
    [Parameter(Mandatory)][string]$HTMLdate,
    [Parameter(Mandatory)][string]$HTMLheader
  )
  foreach ($item in $sourceData) {
    $page = ''
    $newName = $Prefix + '-' + $HTMLdate + (get-date -Format '_hhmmssffff') + '.html'
    $newPath = Join-Path -Path $Folder -ChildPath $newName
    write-host "Downloading $newName..."
    try { $page = (Invoke-WebRequest -UseBasicParsing $item.Url).Content }
    catch { write-warning "$newName download failed" }
    $page = $page -replace "(?s)<script.+?</script>", ""
    $page = $page -replace "(?s)<style.+?</style>", "<style type='text/css'>@import url('./style.css');</style>"
    Set-Content -Path $newPath -Value $page
    $item.EventID = "<a href='./" + $newName + "'>" + $item.EventID + "</a>" 
    $item.Url = "<a href='./" + $newName + "'></a>"
  }
  $sourceData | select-object -Property * -ExcludeProperty Url| convertto-html -CssUri "./style.css" -PreContent "<h2>$HTMLheader</h2>" -Title $HTMLheader |
    ForEach-Object {$_ -replace "&#39;","'" -replace '&lt;','<' -replace '&gt;','>'} |
    ForEach-Object {$_ -replace '<link rel="stylesheet" type="text/css" href="./style.css" />',"<style type='text/css'>@import url('./style.css');</style>"} |
    Out-File (Join-Path -Path $folder -ChildPath ("index_" + $Prefix + '-' + $HTMLdate + (get-date -Format '_hhMMssffff') + '.html'))
}

function Send-TelegramMessage{
    param(
      [Parameter(Mandatory)]$chatId,
      [Parameter(Mandatory)]$tgToken,
      [Parameter(Mandatory)]$text
      )
    $URL = "https://api.telegram.org/bot$tgToken/sendMessage"
    $ht = @{
        text = $text
        parse_mode = "HTML"
        chat_id = $chatID
        }
    $json = $ht | ConvertTo-Json
    $e = $true
    while ($e){
        $e = $false
        Write-Host "sending message..."
        try {Invoke-RestMethod $URL -Method Post -ContentType 'application/json; charset=utf-8' -Body $json | Out-Null}
        catch {write-warning "failed! retry in five seconds" ; $e = $true ; Start-Sleep 5}
        }
    }