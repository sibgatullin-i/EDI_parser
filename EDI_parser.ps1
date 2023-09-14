#read settings
$settings = (get-content (Join-Path -Path $PSScriptRoot -ChildPath "settings.json") | ConvertFrom-json)

add-type -path ((join-path -path (join-path -Path $PSScriptRoot -ChildPath "lib") -ChildPath "OpenPop.dll"))
Import-Module Transferetto -Force -DisablenameChecking
Import-Module ((join-path -path (join-path -Path $PSScriptRoot -ChildPath "lib") -ChildPath "source.psm1")) -DisablenameChecking

$msg = "EDI parser status" # will collect error messages here

# first we connect to POP3 and store it to variable
$pop3Client = Connect-Mail -Server $settings.mailServer -Port $settings.mailPort -Username $settings.mailUsername -Password $settings.mailPassword
if (!$pop3Client) {
  $msg = '¯\_(ツ)_/¯`r`nEDI parser failed to connect to`r`n$($settings.mailServer)`r`nwith username`r`n$($settings.mailUsername)'
  Send-TelegramMessage -tgToken $settings.tgToken -chatId $settings.chatId -text $msg
  write-warning $msg
  start-sleep 10
  exit 1  
  }

# now we check to see if there is any mail for us
$mails = Check-Mail $pop3Client -From $settings.mailTargetfrom

# if none - exit, if >= 1 - ask to confirm
$targetMails = ($mails | Where-Object {$_.target -eq $true} | measure-object).count
if ( $targetMails -eq 0 ) {
  write-host "No matching emails found. Please choose:" 
  $response = read-host "type YES to parse existing inbox files"
  if ($response -like "yes") {
    # parse existing inbox files, remove all outbox files
    Get-ChildItem $settings.outboxFolder | Remove-Item -Force -Recurse
  } else {
    # exit
    write-host "See ya!`r`nWill exit in 10 seconds..."
    $pop3Client.disconnect()
    start-sleep 10
    exit
  }
} elseif ($targetMails -eq 1) {
  write-host "Found one mail. Going to delete old inbox files and proceed new ones."
  write-warning "Ctrl-C to abort"
  start-sleep 10
  Get-ChildItem $settings.inboxFolder | Remove-Item -Force -Recurse
  Get-ChildItem $settings.outboxFolder | Remove-Item -Force -Recurse
} elseif ($targetMails -gt 1) {
  write-warning "There are several mails found.`r`nCheck mailbox and delete (if any) duplicates."
  $response = read-host "type YES to continue"
  if ($response -like "yes") {
    Get-ChildItem $settings.inboxFolder | Remove-Item -Force -Recurse
    Get-ChildItem $settings.outboxFolder | Remove-Item -Force -Recurse
  } else {
    write-host "See ya!`r`nWill exit in 10 seconds..."
    $pop3Client.disconnect()
    start-sleep 10
    exit
  }
} else { # seems like connection error or something
  Write-warning "Cannot count mails. Check network connection"
  write-host "See ya!`r`nWill exit in 10 seconds..."
  $pop3Client.disconnect()
  start-sleep 10
  exit
}

# proceed mails, if $_.target -eq $true - download and mark for deletion, else mark for deletion
foreach ($mail in $mails) { 
  if ($mail.target -eq $true) {
    FetchAndSave-Attachment -pop3Client $pop3Client -Folder $settings.inboxFolder -messageIndex $mail.index  
    $pop3Client.DeleteMessage($mail.index)
  } else {
    $pop3Client.DeleteMessage($mail.index)
  }
}
# done with mailbox
$pop3Client.disconnect()

$incomingFiles = (Get-ChildItem $settings.inboxFolder)
$msg += "`r`nfiles saved: $((Get-ChildItem $settings.inboxFolder).count)"

### ignore ssl errors ###
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
#########################

# proceed files - move, parse, download and shit
$currentFilePosition = 0
foreach ($incomingFile in $incomingFiles) {
  $currentFilePosition += 1
  $currentProgress = [Math]::Round(($currentFilePosition * 100) / $incomingFiles.Count)
  Write-Progress -Id 1 -Activity "Processing incoming files" -Status "$currentFilePosition / $($incomingFiles.Count): $($incomingFile.Name)..." -PercentComplete $currentProgress
  $incomingFileBaseName = ($incomingFile.BaseName -split "-_-")[1]
  #$incomingFileTimeStamp = ($incomingFile.BaseName -split "-_-")[0]
  $source = (Parse-HTML $incomingFile.FullName)
  $HTMLheader = $source[0]
  $HTMLdate = $source[1]
  $sourceData = $source[2]
  $folder = ( (Join-Path -Path $settings.outboxFolder -ChildPath "$incomingFileBasename-$HTMLdate")  )
  mkdir $folder
  Download-Pages -sourceData $sourceData -Folder $folder -Prefix $incomingFileBaseName -HTMLdate $HTMLdate -HTMLheader $HTMLheader -ProgressParentId 1
  Copy-Item $settings.cssFile $folder
} 

# connect to SFTP
$sftpClient = Connect-SFTP -Server $settings.sftpServer -Port $settings.sftpPort -Verbose -Username $settings.sftpUsername -Password $settings.sftpPassword

# Create folders and upload files
$currentFolderPosition = 0
$folders = (Get-ChildItem $settings.outboxFolder)
foreach ($folder in $folders) {
  $currentFolderPosition += 1
  $currentProgress = [Math]::Round(($currentFolderPosition * 100) / $folders.Count)
  Write-Progress -Id 2 -Activity "Uploading folders" -Status "$currentFolderPosition / $($folders.Count): $($folder.Name)..." -PercentComplete $currentProgress
  $sftpFolder = $settings.sftpParentFolder + $folder.BaseName
  $sftpClient.CreateDirectory("$sftpFolder")
  $currentFilePosition = 0
  $files = (Get-ChildItem $folder.FullName)
  foreach ($file in $files){
    $currentFilePosition += 1
    $currentProgress = [Math]::Round(($currentFilePosition * 100) / $files.Count)
    Write-Progress -Id 201 -Activity "Uploading files" -Status "$currentFilePosition / $($files.Count): $($file.Name)..." -PercentComplete $currentProgress -ParentId 2
    write-host "Uploading file $($file.Name)..."
    $result = Send-SFTPFile -SftpClient $sftpClient -LocalPath $file.FullName -RemotePath "$sftpFolder/$($file.name)" -AllowOverride
    if (!$result.Status) { write-warning "failed"}
  }
}

Disconnect-SFTP $sftpClient

if ($msg.length -gt 0) { Send-TelegramMessage -tgToken $settings.tgToken -chatId $settings.chatId -text $msg }

Remove-Module source 

write-host "will exit in 10 seconds..."
start-sleep 10
