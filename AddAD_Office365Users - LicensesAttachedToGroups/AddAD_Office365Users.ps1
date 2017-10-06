<#
.SYNOPSIS
    This script Creates user in Active Director and Office365 with licenses attached to the groups instead of individually.
.DESCRIPTION    


    Files:
        The following files get added to the current working directory;
            userlog.log: Logs each step that the script has taken.
            usermaillist.csv: Logs each email address and date of added to a csv file, format of firstname.surname@Office365tenancy.com,datetime


    Functions to edit before running the script for the first time:
        configureEmail1; outputs username@
        configureEmail2; outputs firstname.surname@
        configureHomeDir; specify your file servers, the $intake is designed for a folder structure such as;
            \\Server\Share\Intake17\ = For Year sevens as of 2017
            \\Server\Share\Intake16\ = For Year eight as of 2017
        configureOUPath; speciffy your OU for storing faculty and students.
        configureAttribute; edits attribute editor details of each user.
        configureGroups; adds the security and distribution groups for each users.
        ProxyAddresses; Change the proxy addresses to suit your tenancy.
        O365Licenses; 
    For a list of parameters in runtime;
        get-help AddNewUser -Parameter *

    Modules:
        MSOnline
        ActiveDirectory
        Advanced module logToFile logger; https://github.com/joer89/Logging.git

.PARAMETER -Connect
    Connects to Office365 and prompts for the password, this has to be with -O365User.
.PARAMETER -O365User jon.doe@office365tenancy.com
    Stores the username of Office365 logon tenancy.
.PARAMETER -O354Password Password_Of_Account
    Stores the Password for OFfice365 logon tenancy.
.PARAMETER -connectToO365Spec
    Connects to Office365 tenancy with specific username and password from -O365User and -O365Password.
.PARAMETER -disconnect
    Disconnects all PSSessions with *.outlook

.PARAMETER -CSVPath
    Imports the csv file.


.EXAMPLE
    AddNewUser -CSVPath C:\AddUsers\users.csv -connectToO365Spec -O365User Office365User@Office365tenancy.co.uk -O365Password Office365Password
            Connects to Office365 with specific credentials, imports each user from the csv file and adds them to Active Directory and runs a delta sync to add them to Office365.
.Notes
   Author: Joe Richards
   Date:   6/10/2017

.LINK
  https://github.com/joer89/AddMultiADUser.git
#>

param(

    $Username,
    $GivenName,
    $Surname,
    $intake,
    $Area,
    $Department,
    $StudentYr

)

#region advanced function

function AddNewUser {

[CmdletBinding(SupportsShouldProcess=$true)]        


    param (
        [Parameter(Mandatory=$false)]
        [switch]$Connect,

        [parameter(Mandatory=$false)]
        [switch]$connectToO365Spec,

        [parameter(Mandatory=$false)]
        [switch]$disconnect,
       
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365User,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365Password,

        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$CSVPath
        
        )

        Begin{
            #Login to Microsoft Office365.
            checkMSOnlineConnection
        }#End Begin
        process{
            #Imports the email address of the users from a csv file to edit their licenses.
            if($CSVPath){
               createUsers($CSVPath)
            }
        }#End Process
        end{
             if($disconnect){
                #Disconnect all Outlook PSSessions.
                disconnect
            }#End if 
        }#End end
}


#endregion


#region Microsoft Office

#Checks to see if the connection to Office365 is open.
function checkMSOnlineConnection{
#Checks to see if the MSOnline module is loaded.
           $MSOnlineMod = Get-Module -Name MSOnline | Format-Table Name -HideTableHeaders
           if($MSOnlineMod -ne "MSOnline"){            
                impModules
                if($Connect){
                    if($O365User -ne $null -or $O365User -ne ""){ 
                        #Connects to Office365 with manual credentials.                      
                        connectToO365($O365User)
                    }
                }
                elseif($connectToO365Spec){
                    if($O365User -ne $null -or $O365User -ne "" -and $O365Password -ne $null -or $O365Password -ne ""){
                        #Connects to Office365 with user's input Username and Password with no prompt.
                        connectToO365SpecifyUsrPwd($O365User)($O365Password)
                    }#End if
                }#End if
            }#End if
}
#Imports all the modules needed to run the script.
function impModules{
    #Gets the running directory.
    $curDir = (Get-Item -Path ".\" -Verbose).FullName
    #Imports the MSOnline module for Office365.
    Import-Module MsOnline -ErrorAction Stop
    Write-Host "Imported MSOnline." -ForegroundColor Magenta
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Imported Active Directory." -ForegroundColor Magenta
    Import-Module "$($curDir)\AdvancedLoggerModule.psm1" -ErrorAction Stop
     Write-Host "Imported Logger module." -ForegroundColor Magenta
}
#Connect to Office365 with specific usernmae and password.
function connectToO365SpecifyUsrPwd($O365User, $O365Password){
    #Specified Credentials.
    #This is so you don't have to enter the password each time.
        try{
            $pass = ConvertTo-SecureString -AsPlainText $O365Password -Force
            $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $O365User,$pass
       
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Cred -Authentication Basic -AllowRedirection
            Import-PSSession $session -AllowClobber
            Connect-MsolService -Credential $Cred
            Write-Host "Connected to Office365"
    }
    catch{
            Write-Host "Connection failed."
    }
}
#Connects to Office365 and ask for credentials.
function connectToO365($O365User){
    try{
        $cred = Get-Credential $O365User
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
        Import-PSSession $session -AllowClobber
        Connect-MsolService -Credential $cred
        Write-Host "Connected to Office365" -ForegroundColor Magenta
    }
    catch{
        Write-Host "Connection failed."
    }
}
#Disconnects all PSSessions with *.outlook.com.
function disconnect{
    #Disconnects the PSSEssions with *.outlook
    Get-PSSession | ? {$_.ComputerName -like ".outlook.com"} | Remove-PSSession
    Write-Host "Removed all PSSessions with *.outlook.com" -ForegroundColor Magenta
}

#endregion


#region Active Directory.


#Created the Active Directory Users.
function createUsers($csvFilePath){

   try{
        Import-Csv $csvFilePath | ForEach-Object{
    
            $Username = $_.Username
            $GivenName = $_.GivenName
            $Surname = $_.Surname

            $intake = $_.intake
            $Area = $_.Area
            $Department = $_.Department
            $StudentYr = $_.StudentYear
            $O365License = $_.O365License


            Write-Host "`n`n`nAdding user $($Username) to Active Directory." -ForegroundColor Magenta

            if(($_.Enabled.ToLower()) -eq "true"){
                $Enabled = $True
            }
            else{
                $Enabled = $False 
            }

             #User's email address.
             $email1 = configureEmail1($Username)
             Write-Host $email1 -ForegroundColor Magenta
             #Logs to .\userlog.log with text.
             Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $email1
        
             $email2str = configureEmail2($GivenName)($Surname)
            
             #User's email address.
             $email2 = configureEmail2($GivenName)($Surname)
             Write-Host $email2 -ForegroundColor Magenta                         
             #Logs to .\userlog.log with text.
             Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $email2
                 
            
             Write-Host "Adding $($email2) to useremaillist.csv format 'email,datetime'" -ForegroundColor Magenta           
             #Logs to .\useremaillist.log with text.
             Log-ToFile -Path .\ -fileName useremaillist.csv -SimpleLogging -Text "$($email2),"


             #User's description, configured for 'STUDENT' or 'FACULTY' in the csv file.
             $Description = getDescription($Area)
             Write-Host $Description -ForegroundColor Magenta
             #Logs to .\userlog.log with text.
             Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $Description


             #User's home directory.
             $HomeDir =  configureHomeDir($Area)($intake)($Username)
             Write-Host $HomeDir -ForegroundColor Magenta
             #Logs to .\userlog.log with text.
             Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $HomeDir

             #User's Organizational Units path.
             $OUPath = configureOUPath($Area)
             Write-Host $OUPath -ForegroundColor Magenta
             #Logs to .\userlog.log with text.
             Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $OUPath


             #Add the user in Active Directory.
             New-ADUser -Name $Username `
                -Enabled $Enabled `
                -SamAccountName $Username `
                -GivenName $GivenName `
                -Surname $Surname `
                -UserPrincipalName $email2str `
                -AccountPassword (ConvertTo-SecureString "School1" -AsPlainText -force) `
                -ChangePasswordAtLogon $True `
                -CannotChangePassword $False `
                -PasswordNeverExpires $False `
                -PasswordNotRequired $False `
                -Description $Description `
                -EmailAddress $email2str `
                -HomeDirectory $HomeDir `
                -HomeDrive "U:" `
                -Path $OUPath
            
            #Sets the following parameters for the user account.
            #Extention attribute three and eleven.
            $userAttributes = configureAttribute($Area)($Username)
            Write-Host $userAttributes -ForegroundColor Magenta
            #Logs to .\userlog.log with text.
            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $userAttributes

            #Add the Users the the correct Groups
            $Groups = configureGroups($Area)($Username)($Department)($StudentYr)
            Write-Host $Groups -ForegroundColor Magenta    
            #Logs to .\userlog.log with text.
            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $Groups
            
            #Set the proxy address for the user.
            $ProxyAddr = setProxyAddress($Username)($GivenName)($Surname) 
            Write-Host $ProxyAddr -ForegroundColor Magenta    
            #Logs to .\userlog.log with text.
            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $ProxyAddr.ToString() 
            
            #Adding the O365 license security group.
            $O365LicensesResult = O365Licenses($O365License)($Username)
            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $O365LicensesResult

            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "***Completed adding $($email2).***"



        }

        Write-Host "Finished importing CSV file: $($csvFilePath)" -ForegroundColor Magenta

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Running Delta Sync."
        #Run Delta Sync to sync the new user from Active Directory to Office365.          
        runDeltaSync
        
        Write-Host "There is $($emailArray.Count) new users added to Active Directory, adding the licenses for each user." -ForegroundColor Magenta
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "`nThere is $($emailArray.Count) new users added to Active Directory, adding the licenses for each user."
   }
   catch{
        $ErrorMsg = $PSItem.Exception
        Write-Host "Failed: $($ErrorMsg)." -ForegroundColor Red
   }
}

#Start function, configures the description for each user.
function getDescription($Area){
    if($Area.ToLower() -eq "FACULTY"){
        #If Area is FACULTY then ULT - FACULTY - Date of Day, Month and year.
        $Description = "Faculty - " + (Get-Date -Format "%d-%M-%y")
    }
    elseif($Area.ToLower() -eq "STUDENT"){
        #If Area is FACULTY then ULT - STUDENT - Date of Day, Month and year.
        $Description = "Student - " + (Get-Date -Format "%d-%M-%y")
    }   
    #Returns the description.
    return $Description
}#End function.

#Start function, configures the email account.
function configureEmail1($Username){
    #Configures the email account to %username%@Office365tenancy.co.uk.
    $email = "$($Username)@Office365tenancy.co.uk"
    #Returns the user's email.
    return $email
}#End function.

#Start function, configures the email account.
function configureEmail2($GivenName, $Surname){
    #Configures the email account to Office365tenancy.co.uk.
    $email2 = "$($GivenName).$($Surname)@Office365tenancy.co.uk"
    #Returns the user's email.
    return $email2
}#End function.

#Start function, configures the user's home directory. 
function configureHomeDir($Area, $intake, $Username){
    if($Area.ToLower() -eq "FACULTY"){
        $HomeDirectory = "\\FileSherver\StaffShare$\$($Username)"
    }
    elseif($Area.ToLower() -eq "STUDENT"){  
        #$($intake) is designed to seperate each year of students to each folder.        
        $HomeDirectory =  "\\FileServer\StudentShare$\$($intake)\$($Username)"
    }
    #Returns the home directory.
    return $HomeDirectory
 }#End function.

#Start function, puts the user in the correct Organizational Unit in Active Directory.
function configureOUPath($Area){
    if($Area -match "FACULTY"){
        #If Area is a FACULTY, the users created in the FACULTY Organizational Unit if it's FACULTY.
        $Path = "OU=TempStore,OU=Staff,OU=Users,OU=Company,DC=Office365tenancy,DC=co,DC=uk"
    }
    elseif($Area -match "STUDENT"){
        #If Area is a STUDENT, the users created in the STUDENT tempstore Organizational Unit.
        $Path = "OU=TempStor,OU=Student,OU=Users,OU=Company,DC=Office365tenancy,DC=co,DC=uk"
    }
    #Returns the Organizational Unit's path.
    return $Path
 }#End function.
 
#Start function, add extention attributes three and eleven to the users, This is for the exchange server.
function configureAttribute($Area, $Username){
        if($Area -match "FACULTY"){
            #If Area is FACULTY, add the following extention attributes to the user.
            Set-ADUser -Identity $Username -Add @{extensionAttribute3="EA3"}
            Set-ADUser -Identity $Username -Add @{extensionAttribute11="EA11"}
        }
        elseif ($Area -match "STUDENT"){
            #If Area is STUDENT, add the following extention attributes to the user.
            Set-ADUser -Identity $Username -Add @{extensionAttribute3="EA3"}
            Set-ADUser -Identity $Username -Add @{extensionAttribute11="EA11"}           
        }    
}#End function.

#Start function, configure each group the user has to be in (Not Office365 Licenses groupo in AD.)
function configureGroups($Area, $Username, $Department, $StudentYr){

    #Checks if its FACULTY or a STUDENT account.
    if($Area -match "FACULTY" -and $Department -match "MATH"){
             #Maths faculty groups.
             add-adgroupmember -identity 'Security groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
             add-adgroupmember -identity 'Distribution groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
             $Activity = "Added Maths Security and distribution groups."
             return $Activity
     }
     if($Area -match "FACULTY" -and $Department -match "ENGLISH"){
            #English faculty groups.
            add-adgroupmember -identity 'Security groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
            add-adgroupmember -identity 'Distribution groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
            $Activity = "Added English Security and distribution groups."
            return $Activity
     }
     if($Area -match "FACULTY" -and $Department -match "SCIENCE"){}
     if($Area -match "FACULTY" -and $Department -match "PE"){ }
     if($Area -match "FACULTY" -and $Department -match "ART"){ }
     if($Area -match "FACULTY" -and $Department -match "FINANCE"){}
     if($Area -match "FACULTY" -and $Department -match "ADMIN"){}
     if ($Area -match "STUDENT" -and $StudentYr -match  "7"){
            #STUDENT year seven groups.
            add-adgroupmember -identity 'Security groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
            add-adgroupmember -identity 'Distribution groupA' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
            $Activity = "Added Year 7 Security and distribution groups."
            return $Activity
     }
     if ($Area -match "STUDENT" -and $StudentYr -match  "8"){}  
     if ($Area -match "STUDENT" -and $StudentYr -match  "9"){}  
     if ($Area -match "STUDENT" -and $StudentYr -match  "10"){} 
     if ($Area -match "STUDENT" -and $StudentYr -match  "11"){}   
     if ($Area -match "STUDENT" -and $StudentYr -match  "12"){}   
     if ($Area -match "STUDENT" -and $StudentYr -match  "13"){}

}#End Function

#Set the proxy addresses.
function setProxyAddress($Username, $GivenName, $Surname){

    $proxy1 = "SMTP:$($GivenName).$($Surname)@@Office365tenancy.co.uk"
    $proxy2 = "smtp:$($GivenName).$($Surname)@@Office365tenancy.mail.onmicrosoft.com"
    $proxy3 = "smtp:$($Surname)$($GivenName.SubString(0,1))@Office365tenancy.co.uk"

    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy1" }
    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy2" }
    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy3" }

    return "Added proxy addresses for $($Username)."
}


#endregion


#region Microsoft Office365 Licenses.
#Adding Office365
function O365Licenses($O365License, $Username){

    #Adding the Office365 security group for Azure adding licenses within office365.
    if($O365License -match "ADMIN"){
        #Maths faculty groups.
        add-adgroupmember -identity 'Contoso Administrator 365 Licence group' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
        $Activity = "Added admin Office365 Scurity Groups."
        return $Activity
    }
    elseif($O365License -match "STAFF"){
        #Maths faculty groups.
        add-adgroupmember -identity 'Contoso Staff 365 Licence group' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
        $Activity = "Added Staff Office365 Scurity Groups."
        return $Activity
    }
    elseif($O365License -match "STUDENT"){
        #Maths faculty groups.
        add-adgroupmember -identity 'Contoso Student 365 Licence group' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
        $Activity = "Added Student Office365 Scurity Groups."
        return $Activity
    }
    else{
        return "No office365 security groups added for $($Username)."
    }
}
#Delta Sync
function runDeltaSync{
    write-host "Running 'DELTA' synch, which is the 6 different stages..."
    $pass = ConvertTo-SecureString -AsPlainText AzurePassword -Force
    $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList contoso\admin,$pass
    Invoke-Command -ComputerName AzureServer -Credential $cred -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
    write-host "...DELTA synch will take approximately 1 minute."
    Start-Sleep -s 60
}#End function


#endregion
