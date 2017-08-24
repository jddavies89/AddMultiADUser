<#
.Synopsis
   This Module sends Office365 emails for logging.
.DESCRIPTION
   - There is two functions for this module, StoreCred and send0365Mail.
   - The StoreCreds checks to see if the folder and files exist for authenticating to Office365 and if it doesnt exist, then stores the credentials to C:\o365\.
   - The SendO354Mail function retrives the files which were created with StoreCreds function and sends the Email to Office365 email address.
   - Nothing has to be changed for this to work except under Send-MailMessage, the -To and the -From.
.Notes
   Author: Joe Richards
   Date: 02/Feb/2017
.LINK
  https://github.com/joer89/Logging
#>
function checkFolder{

    #Check the accessed time on the folder.
    $lastwritetime = (Get-Item "C:\O365").LastWriteTime
    #Adds one day.
    $timeSpan = New-TimeSpan -Day 1
    #Gets the current time.
    $currentTime = Get-Date
    
    #If the date is within the last day of the creation time of the folder sends the email otherwise prompts for the credentials.
    if(!($currentTime -le ($lastwritetime + $timeSpan))){
        StoreCreds
    }
}

#Checks to see if the file exists if not it prompts for the password and creates a file with the encrypted password.
function StoreCreds{

        if((-Not (Test-Path -LiteralPath "C:\O365\O365User.txt")) -or (-Not (Test-Path -LiteralPath "C:\O365\O365Pass.txt"))){   
        
            #Checks to see if C:\O365 folder is there otherwise creates it.        
            if(-Not (Test-Path -LiteralPath "C:\O365")){
                #Creates the folder
                New-Item -Path "C:\" -Name "O365" -ItemType directory
            }

            #Gets the credentials.
            $Credentials = Get-Credential -Message "Office 365 Authentication for SMTP & the sender."
           
            #Stores the office365 credentials to a text file.
            $Credentials.UserName | Set-Content "C:\O365\O365User.txt"
            #Stores the office365 credentials to a text file.
            $Credentials.Password | ConvertFrom-SecureString | Set-Content "C:\O365\O365Pass.txt"

            #Sends the email.
            send0365Mail
        }
   
}#end function.
function sendO365Mail{

            #Sends the email.
            $User = Get-Content "C:\O365\O365User.txt"
            Send-MailMessage –From $user –To itsupport@SMTPEmail.com -Cc CCEMail@SMTPEmail.com –Subject “New User” –Body $body -SmtpServer mail.messaging.microsoft.com
            #Writes out to the screen.
            Write-Host "Message Sent."
}