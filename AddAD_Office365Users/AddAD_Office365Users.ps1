<#
.SYNOPSIS
    This is an advanced function script designed to add users from a csv file to Active Directory and Office365.
.DESCRIPTION    

    I haven't put logging in within this script because i have covered lots of different logging methods here; https://github.com/joer89/Logging.git

    Functions to edit before running the script for the first time:
        addSellectedProPlusFacultyUsr; Here 'Office 365 ProPlus for faculty' get enabled with only 'SWAY' license disabled.
        addSellectedProPlusStudentsUsr; Here 'Office 365 ProPlus for Students' gets enabled with only 'SWAY' and 'DFFICE_FORMS_PLAN_2' lcenses disabled.
        addSellectedEduFacultyUser; Here 'Office 365 Education for faculty' gets enabled with School Data Sync (plan1), Stream for Office365, Microsoft Teams, PowerApps for Office365, Azure Rights Management licenses disabled.
        addSellectedEduStudentUsr; Here 'Office365 for student education' gets enabled with Microsoft Forms (plan 2), Microsoft Planner, SWAY, Office Online for Education, Sharepoint 1 for EDU, Exchange Online (plan 1)

        configureEmail1; outputs username@
        configureEmail2; outputs firstname.surname@
        configureHomeDir; specify your file servers, the $intake is designed for a folder structure such as;
            \\Server\Share\Intake17\ = For Year sevens as of 2017
            \\Server\Share\Intake16\ = For Year eight as of 2017
        configureOUPath; speciffy your OU for storing faculty and students.
        configureAttribute; edits attribute editor details of each user.
        configureGroups; adds the security and distribution groups for each users.
        ProxyAddresses; Change the proxy addresses to suit your tenancy.
        Delta Sync; Change the Invoke-command credentials for Delta Sync.
        
    For a list of parameters in runtime;
        get-help AddNewUser -Parameter *

    Modules:
        MSOnline

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
    This parameter has to be implemented with one or more of the following parameters if you want Office365 emails:

    .PARAMETER -addSellectedProPlusFacultyUsr
    .PARAMETER -addSellectedProPlusStudentsUsr
    .PARAMETER -addSellectedEduFacultyUsr
    .PARAMETER -addSellectedEduStudentUsr


.EXAMPLE
    AddNewUser -connectToO365Spec -O365User adminUser@office365tenancy.co.uk -O365Password Password
            Connects to office365 with specific redentials and then disconnects from Office365.
.EXAMPLE
    AddNewUser -CSVPath C:\AddUsers\users.csv -addSellectedProPlusFacultyUser -addSellectedEdcuFacultyUser -connectToO365Spec -O365User Office365User@Office365tenancy.co.uk -O365Password Office365Password
            Connects to Office365 with specific credentials,
.EXAMPLE
    AddNewUser -CSVPath C:\AddUsers\users.csv -addSellectedProPlusFacultyUser -addSellectedEdcuFacultyUser
            Once connected to office365, adds all the users from the CSV file to Active Directory and adds Office365 ProPlus for faculty and Education for faculty.
.EXAMPLE
     AddNewUser -CSVPath C:\AddUsers\users.csv -connectToO365Spec -O365User Office365Admin@tenancy.com -O365Password Office365Password -addSellectedProPlusStudentsUser -addSellectedEdcuStudentUser -disconnect
            Connects to Office365 with specified credentials, adds the users to Active Directory from C:\AddUsers\users.csv and adds Office365 ProPlus for students and Education for Students, once completed then disconnects.
            
.Notes
   Author: Joe Richards
   Date:   28/09/2017

.LINK
  https://github.com/joer89/AddMultiADUser.git

.COPYRIGHT
    This file is part of AddAD_Office365Users.

    AddAD_Office365Users is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    AddAD_Office365Users is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.
    Please visit the GNU General Public License for more details: http://www.gnu.org/licenses/.

#>

param(

    $Username,
    $GivenName,
    $Surname,
    $intake,
    $Area,
    $Department,
    $StudentYr,
    $emailArray =@()

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
        [string]$CSVPath,


        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [switch]$addSellectedProPlusFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [switch]$addSellectedProPlusStudentsUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [switch]$addSellectedEdcuFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [switch]$addSellectedEdcuStudentUser
        
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
    #Imports the MSOnline module for Office365.
    Import-Module MsOnline -ErrorAction Stop
    Write-Host "Imported MSOnline." -ForegroundColor Magenta
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Imported Active Directory." -ForegroundColor Magenta
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

            Write-Host "Adding user $($Username) to Active Directory." -ForegroundColor Magenta

            if(($_.Enabled.ToLower()) -eq "true"){
                $Enabled = $True
            }
            else{
                $Enabled = $False 
            }

             #User's email address.
             $email1 = configureEmail1($Username)
             Write-Host $email1 -ForegroundColor Magenta

             #User's email address.
             $email2 = configureEmail2($GivenName)($Surname)
             #Adds th email address to the array, to be used for adding the Office365 licenses.
             $emailArray = $email2

             #Used for 'New-ADUser' as it requires a string not an array.
             $email2str = configureEmail2($GivenName)($Surname)
             Write-Host $email2str -ForegroundColor Magenta

             #User's description, configured for 'STUDENT' or 'FACULTY' in the csv file.
             $Description = getDescription($Area)
             Write-Host $Description -ForegroundColor Magenta


             #User's home directory.
             $HomeDir =  configureHomeDir($Area)($intake)($Username)
             Write-Host $HomeDir -ForegroundColor Magenta

             #User's Organizational Units path.
             $OUPath = configureOUPath($Area)
             Write-Host $OUPath -ForegroundColor Magenta

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

            #Add the Users the the correct Groups
            $Groups = configureGroups($Area)($Username)($Department)($StudentYr)
            Write-Host $Groups -ForegroundColor Magenta  
            
            #Set the proxy address for the user.
            $ProxyAddr = setProxyAddress($Username)($GivenName)($Surname) 
            Write-Host $ProxyAddr -ForegroundColor Magenta    
            #Logs to .\userlog.log with text.
            Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text $ProxyAddr.ToString()               
        }

        Write-Host "Finished importing CSV file: $($csvFilePath)" -ForegroundColor Magenta

        #Run Delta Sync to sync the new user from Active Directory to Office365.          
        runDeltaSync

        Write-Host "There is $($emailArray.Count) new users added to Active Directory, adding the licenses for each user."
        
        #Loop through the $emailArray and add the licenses to each user.
        foreach($email in $emailArray){
            Write-Host "Editing $($email)'s licenses" -ForegroundColor Magenta
                  
            #Add sellected licenses under the SKU of Office365 ProPlus for faculty.
            if($addSellectedProPlusFacultyUser){
                addSellectedProPlusFacultyUsr($email)
            }
            #Add sellected licenses under the SKU of Office365 ProPlus for students.
            if($addSellectedProPlusStudentsUser){
               addSellectedProPlusStudentsUsr($email)
            }
            #Add sellected licenses under the SKU of Office 365 Education for faculty.
            if($addSellectedEdcuFacultyUser){
                addSellectedEduFacultyUsr($email)
            }
            #Add sellected licenses under the SKU of Office 365 Education for students.
            if($addSellectedEdcuStudentUser){
                addSellectedEdcuStudentUsr($email)
            }
            Write-Host "Finished editing $($email)'s licenses" -ForegroundColor Magenta
        }

    }
    catch{
        Write-Host "Failed: Could not import CSV file from $($csvFilePath)" -ForegroundColor Red
    }
}

#Start function, configures the description for each user.
function getDescription($Area){
    if($Area.ToLower() -eq "FACULTY"){
        #If Area is FACULTY then FACULTY - Date of Day, Month and year.
        $Description = "Faculty - " + (Get-Date -Format "%d-%M-%y")
    }
    elseif($Area.ToLower() -eq "STUDENT"){
        #If Area is FACULTY then STUDENT - Date of Day, Month and year.
        $Description = "Student - " + (Get-Date -Format "%d-%M-%y")
    }   
    #Returns the description.
    return $Description
}#End function.

#Start function, configures the email account.
function configureEmail1($Username){
    #Configures the email account to %username%@Office365Tenancy.co.uk.
    $email = "$($Username)@Office365Tenancy.co.uk"
    #Returns the user's email.
    return $email
}#End function.

#Start function, configures the email account.
function configureEmail2($GivenName, $Surname){
    #Configures the email account to firstname.surname@Office365Tenancy.co.uk.
    $email2 = "$($GivenName).$($Surname)@Office365Tenancy.co.uk"
    #Returns the user's email.
    return $email2
}#End function.

#Start function, configures the user's home directory. 
function configureHomeDir($Area, $intake, $Username){
    if($Area.ToLower() -eq "FACULTY"){        
        $HomeDirectory = "\\Server\Faculty$\$($Username)"
    }
    elseif($Area.ToLower() -eq "STUDENT"){          
        $HomeDirectory =  "\\Server\student$\$($intake)\$($Username)\Documents"
    }
    #Returns the home directory.
    return $HomeDirectory
 }#End function.

#Start function, puts the user in the correct Organizational Unit in Active Directory.
function configureOUPath($Area){
    if($Area.ToLower() -eq "FACULTY"){
        #If Area is a FACULTY, the users created in the FACULTY Organizational Unit if it's FACULTY.
        $Path = "OU=FACULTY,OU=Users,OU=company,DC=contoso,DC=Org,DC=UK"
    }
    elseif($Area.ToLower() -eq "STUDENT"){
        #If Area is a STUDENT, the users created in the STUDENT tempstore Organizational Unit.
        $Path = "OU=TempStor,OU=Students,OU=Users,OU=company,DC=contoso,DC=Org,DC=UK"
    }
    #Returns the Organizational Unit's path.
    return $Path
 }#End function.
 
#Start function, add extention attributes three and eleven to the users.
function configureAttribute($Area, $Username){
        if($Area -match "FACULTY"){
            #If Area is FACULTY, add the following extention attributes to the user.
            Set-ADUser -Identity $Username -Add @{extensionAttribute3="AA3"}
            Set-ADUser -Identity $Username -Add @{extensionAttribute11="AA11"}
        }
        elseif ($Area -match "STUDENT"){
            #If Area is STUDENT, add the following extention attributes to the user.
            Set-ADUser -Identity $Username -Add @{extensionAttribute3="AA3"}
            Set-ADUser -Identity $Username -Add @{extensionAttribute11="AA11"}           
        }    
}#End function.

#Start function, configure which groups the user needs to be in.
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

    $proxy1 = "SMTP:$($GivenName).$($Surname)@tenancy.co.uk"
    $proxy2 = "smtp:$($GivenName).$($Surname)@tenancy.mail.onmicrosoft.com"
    $proxy3 = "smtp:$($GivenName.SubString(0,1))$($Surname)@tenancy.co.uk"

    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy1" }
    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy2" }
    Get-ADUser $Username | set-aduser -Add @{Proxyaddresses="$proxy3" }

    return "Added proxy addresses for $($Username)."
}

#endregion


#region Microsoft Office365 Licenses.


#Add sellected licenses under the SKU of Office365 ProPlus for faculty.
function addSellectedProPlusFacultyUsr($user){
    try{
       Write-Host "Adding sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB  
       $O365AllFac = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_FACULTY -DisabledPlans SWAY
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_FACULTY -LicenseOptions $O365AllFac  
       Write-Host "Completed, added sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: Adding sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Red
   }
}#End function
#Add sellected licenses under the SKU of Office365 ProPlus for students.
function addSellectedProPlusStudentsUsr($user){
    try{
       Write-Host "Adding sellected plans from 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $addO365AllStu = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_STUDENT -DisabledPlans SWAY, OFFICE_FORMS_PLAN_2
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_STUDENT -LicenseOptions $addO365AllStu
       Write-Host "Completed, adding selected plans for 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add sellected plans for 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Red
   }
}#End function
#Add sellected licenses under the Office365 for faculty education.
function addSellectedEduFacultyUsr($user){
  try{
        Write-Host "Adding all of 'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
        Set-MsolUser -UserPrincipalName $user -UsageLocation GB
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_FACULTY -DisabledPlans SCHOOL_DATA_SYNC_P1, STREAM_O365_E3, TEAMS1, POWERAPPS_O365_P2, RMS_S_ENTERPRISE, YAMMER_EDU
        Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:STANDARDWOFFPACK_FACULTY -LicenseOptions $O365All
        Write-Host "Completed, added all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}#End function
#Add sellected licenses under the Office365 for student education.
function addSellectedEdcuStudentUsr($user){
    try{
       Write-Host "Adding all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_STUDENT -DisabledPlans AAD_BASIC_EDU, SCHOOL_DATA_SYNC_P1, STREAM_O365_E3, TEAMS1, Deskless, FLOW_O365_P2, POWERAPPS_O365_P2, RMS_S_ENTERPRISE, YAMMER_EDU, MCOSTANDARD           
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses  regisschool:STANDARDWOFFPACK_STUDENT -LicenseOptions $O365All 
       Write-Host "Completed, added all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}#End function
#Delta Sync
function runDeltaSync{
    write-host "Running 'DELTA' synch, which is the 6 different stages..."
    $pass = ConvertTo-SecureString -AsPlainText Passw0rd -Force
    $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList domain\admin,$pass
    Invoke-Command -ComputerName AzureServer -Credential $cred -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
    write-host "...DELTA synch will take approximately 1 minute."
    Start-Sleep -s 60
}#End function

#endregion
