

#Use the Where to filter out objects that lack email addresses and are not enabled accounts.
#$Email = (Get-ADuser -Filter {MemberOf -like "CN=YourMomsGroup,OU=DistributionLists,OU=Exchange,OU=WhereUAt,DC=Company,DC=com"}).Where({ ($_.enabled -eq $true) -and ($_.UserPrincipalName -notin $null,'') }).UserPrincipalName 

#$Date = Get-Date -UFormat %B-%d-%Y
$Subject = "CAPS IS SHOUTING"

$MailSender = "" # Email address

$MailServer = "" # yourexchangser@YourDomain.Com

$Msg = @"
Hi Guys,

Just running a test for <Some Guy>

-Coolish IT Guy

"@

 function Send-Mail {
   <#
     .SYNOPSIS
         Send emails
     .DESCRIPTION
         Automate the dispersement of official emails
     .PARAMETER EmailAddresses
         The full email addresses needed (UserPrincipalName)
         Send a list of users.
     .PARAMETER MessageBody
         The message that is going to be sent
     .PARAMETER ExchangeServer
         The server that will be handling the mail service
     .PARAMETER Subject
         The topic of the email
     .PARAMETER Sender
         Who is sending the message

      .EXAMPLE
         
         #$Email = (Get-ADuser -Filter {MemberOf -like "CN=YourMomsGroup,OU=DistributionLists,OU=Exchange,OU=WhereUAt,DC=Company,DC=com"}).Where({ ($_.enabled -eq $true) -and ($_.UserPrincipalName -notin $null,'') }).UserPrincipalName 

         Send-Mail -EmailAddresses $Email -MessageBody $Msg -Subject $Subject -Sender $MailSender
         (!!!This might not be there long term!!!)
         From                        : ITSecurity@YourMom.com
         Sender                      : 
         ReplyTo                     : 
         ReplyToList                 : {}
         To                          : {Gary.Guy@YourMom.com, SoMuch.Steve@YourMom.com, Tom.AndJerry@YourMom.com, Rich.IeRich@YourMom.com...}
         Bcc                         : {}
         CC                          : {}
         Priority                    : Normal
         DeliveryNotificationOptions : None
         Subject                     : 
         SubjectEncoding             : 
         Headers                     : {}
         HeadersEncoding             : 
         Body                        : 
         BodyEncoding                : 
         BodyTransferEncoding        : Unknown
         IsBodyHtml                  : False
         Attachments                 : {}
         AlternateViews              : {}

     .INPUTS
        Objects
     .OUTPUTS
        String
     .NOTES 
        Last Updated: 2019-05-10 11:36:04Z

        ========== HISTORY ==========
        Author: Kevin Van Bogart
        Created: 2019-05-10 11:36:04Z
    
        Updated By: 
        Updated On: 

   #>
    # Write-Host "Sending Email"

   param (
      <#
      [Parameter(ValueFromPipelineByPropertyName=$true)]
      [string]$ExchangeServer = "Your normal email server",
      #>
      [Parameter(ValueFromPipelineByPropertyName=$true)]
      [string]$ExchangeServer,

      [Parameter(Mandatory=$True,
         ValueFromPipelineByPropertyName=$true)]
      [string[]]$EmailAddresses,

      [Parameter(ValueFromPipelineByPropertyName=$true)]
      [string]$Subject,

      [Parameter(Mandatory=$True,
      ValueFromPipelineByPropertyName=$true)]
      [string]$MailSender,

      [Parameter(Mandatory=$True,
      ValueFromPipelineByPropertyName=$true)]
      [string]$MessageBody
   )
   begin {
      #Ended up not needing this
   }
   Process {

      #Create objects
      try {
         #Creating a Mail object
         $msg = new-object Net.Mail.MailMessage
         #Creating SMTP server object
         $smtp = new-object Net.Mail.SmtpClient($ExchangeServer)   
      }
      catch {"Failed to create mail object[s]. Error: $($_.Exception.Message)"}

      #Email structure 
      try {$msg.From = $MailSender}
      catch {"Failed to set 'From' variable. Error: $($_.Exception.Message)"}

      #Add the recipients
      try {
         #Oddly, this isn't a ';'
         $msg.To.Add($EmailAddresses -join ',')
      }
      catch {"Failed to add recipients. Error: $($_.Exception.Message)"}

      #$msg.Bcc.Add
      try {$msg.subject = $Subject}
      catch {"Failed to set subject. Error: $($_.Exception.Message)"}

      try {$msg.body = $MessageBody}
      catch {"Failed to set message body. Error: $($_.Exception.Message)"}

      #Sending email 
      #$msg # for tests

      try {$smtp.Send($msg)}
      catch {"Failed to send message. Error: $($_.Exception.Message)"}
   }
}

#Calling function
$Message = Send-Mail -EmailAddresses $Email -MessageBody $Msg -Subject $Subject -Sender $MailSender
$Message