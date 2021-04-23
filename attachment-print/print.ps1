#Script is intended to be stored in C:\attachment-print
#If you prefer another directory, feel free to modify the code below :)
#to execute task silently in background, use C:\attachment-print\task.vbs



#Load ImapX.dll (IMAP client, https://github.com/azanov/imapx )
[Reflection.Assembly]::LoadFile("C:\attachment-print\ImapX.dll")



#Clean downloaded attachment folder, in case some garbage is left there.
Remove-Item C:\attachment-print\attachments\*.* -Force

$Username = "example@example.com"
$Password = "ExamplePass"

$client = New-Object ImapX.ImapClient
$client.Behavior.MessageFetchMode = "Full"
$client.Host = "mail.example.com"
$client.Port = 993
$client.UseSsl = $true
$client.SslProtocol = [Net.SecurityProtocolType]::Tls12
$client.Connect()
$client.Login($Username, $Password)


#select inbox folder
$res = $client.folders| where { $_.path -eq 'Inbox' }


# fetch last 100 messages
$numberOfMessagesLimit = 100
$messages = $res.search("ALL", $client.Behavior.MessageFetchMode, $numberOfMessagesLimit)

# Display the messages in a formatted table
#$messages | ft *

#download attachements from each unread message ( including inline attachements )
foreach($m in $messages) {
  if($m.Seen) {
  } else {
    $m.Subject
    foreach($r in $m.Attachments) {
        $r.Download()
        $r.Save('C:\attachment-print\attachments\')
    }
    foreach($r in $m.EmbeddedResources) {
        $r.Download()
        $r.Save('C:\attachment-print\attachments\')
    }
     
  }
#mark messages as read
  $m.Flags.Add("\SEEN")
  
}
#wait 20 seconds, in some cases files appear to late.
timeout 20


#delete empty attechements and files smaller than 1kb
Get-ChildItem "C:\attachment-print\attachments\" -Filter *.stat -recurse |?{$_.PSIsContainer -eq $false -and $_.length -lt 1000}|?{Remove-Item $_.fullname -WhatIf}
Get-ChildItem -Path "C:\attachment-print\attachments\" -Recurse -Force | Where-Object { $_.PSIsContainer -eq $false -and $_.Length -eq 0 } | remove-item
#delete emoticons
Remove-Item C:\attachment-print\attachments\*Emoticon* -Force
Remove-Item C:\attachment-print\attachments\*emoticon* -Force
Remove-Item C:\attachment-print\attachments\*smiley* -Force

#Print PDF files
$FilesToPrint = Get-ChildItem "C:\attachment-print\attachments\*" -Recurse -Include *.pdf,*.PDF
foreach($File in $FilesToPrint) {
    Start-Process -WindowStyle Hidden -FilePath $File.FullName -Verb Print -PassThru | %{ sleep 15;$_ } | kill
    Start-Sleep -Seconds 10
}

#Print images
$FilesToPrint = Get-ChildItem "C:\attachment-print\attachments\*" -Include *.jpg,*.jpeg,*.JPG,*.JPEG,*.PNG,*.png
foreach($File in $FilesToPrint) {
    mspaint /pt $File.FullName
    Start-Sleep -Seconds 10
}

#move printed files to printed-attachments folder
Get-ChildItem "C:\attachment-print\attachments\*" -Include *.jpg,*.jpeg,*.JPG,*.JPEG,*.PNG,*.png,*.pdf,*.PDF | Move-Item -Force -Destination "C:\attachment-print\printed-attachments\"
#izdzest visu kas paliek pari
#delete garbage
Remove-Item C:\attachment-print\attachments\*.* -Force

#delete printed-attachments older than 60 days
$limit = (Get-Date).AddDays(-60)
Get-ChildItem -Path "C:\attachment-print\printed-attachments\" -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
