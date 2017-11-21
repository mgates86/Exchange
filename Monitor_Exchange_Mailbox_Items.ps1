$emailTo = "##ENTER-EMAIL##"
$emailFrom = "##ENTER-EMAIL##"
$subject  = "ALERT: EV is not Archiving. Mailbox: ##MONITORED MAILBOX## "
$smtpServer  = "##ENTER-SMTP-SERVER##"
$MessageCount = "3000"
$EWSURL = "https://##ENTER-SERVER##/ews/exchange.asmx"
$MAILBOXMONITORED = "##ENTER-EMAIL##@domain.com" 

$dllPath =  "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

$EWS = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)
$EWS.Url = New-Object System.Uri($EWSURL)

$objInboxFolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MAILBOXMONITORED)
$objInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS, $objInboxFolderID)
#Write-Host $objInbox.UnreadCount

$counter = $objInbox.UnreadCount

if($counter -ge $MessageCount) 

{
$body = "Active Batch Trigger --- Number of Messages Stuck in EV Mailbox: " + $counter


send-mailmessage -To $emailTo -From $emailFrom -Subject $subject -body $body -smtpserver $smtpServer

$display = "Over"
$display

}



else

{ 

$display = "Under"
$display

}