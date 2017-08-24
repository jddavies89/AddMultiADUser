<#
.Synopsis
   This script creates multiple users in Active Directory and sets their licences for Office 365.
.DESCRIPTION    
    This script pulls all the users in Users.csv and adds them to Active Directory and Office365, while adding the licences and mailbox for each user.

    Modules:
        AdvancedLoggerModule is used for logging userlog.log and EmailBody.log in current working directory.
        O365Email is used for sending the EmailBody.log via email.
        MSOnline is used for connecting to Office365.
        ActiveDirectory is used for connecting to ActiveDirectroy.

    Steps:
        1. Creates userlog.log and emailBody.log in current directory if it doesn't exist, Function createFiles.
        2. Imports all the modules, function ImpModule.
        3. Connects to Office365, function connectToO365.
        4. Creates the licence file (CSV File) in user folder, function createLicenceFile.
        5. Reads Users.csv and adds the users to Active Directory and Office365, function ADCreateUser.
        6. Under ADCreateUser: Edits attributes three and eleven of each user (internal use), function attributeEditing
        7. Under ADCreateUser: Create security permission on user's home drive folder (I cant get this working, only seems to work if you use %username% in the user's home drive attribute in AD.), function createSecurityPermissions.
        8. Under ADCreateUser: Sets the SMTP and domain name inthe user's attributes proxy, function setProxy.
        9. Under ADCreateUser: Adds the Security Groups to the user in AD, function addUserSecurityGroups.
        10.Under ADCreateUser: Connects to the Azure server with built in credentials, (The rerason why its built in is because otherwise, you will be prompted for the credentials each line of Users.csv), function connectToAzureLocalServer
        11.Under ADCreateUser: Adds the User Principal Name (UPN) to the csv file under user, function createO365licenseFile
        12.Under ADCreateUser: imports the csv file from user folder and adds the Office365 licences for that user, function addO365LicensesForUser.
        13.Under ADCreateUser: imports the csv file from user folder and sets the mailbox for the user, function setMailboxForUser.
        14.Under ADCreateUser: Appends the text for the user to be stored in EMailBody.log, function storeUsersForEmail.
        15.Under ADCreateUser: At the end of this function, the function diconnectPSSession is called to diconnect all PS Sessions.
        16.At the end of the script, the file EmailBody.log gets read and emailed, see O365EMail.psm1 module in running directory for emailing or visit my github; https://github.com/joer89/Logging/Office365EmailLogger for the standalone module.
        
.EXAMPLE
    For examples on Log-ToFile, please visit ToFileLogger on my github; https://github.com/joer89/Logging/ToFileLogger/
    For examples on sendO365Mail, please visit Office365EmailLogger on my Github; https://github.com/joer89/Logging/Office365EmailLogger/
.Notes
   Author: Joe Richards
   Date:   23/08/2017
.LINK
  https://github.com/joer89/AddMultiADUser/
#>

param(
    #Retrives the execution directory and Stores the content of $body.
    $path = (split-path $script:MyInvocation.MyCommand.Path -parent),
    $file = (get-date -Format yyyy-MM-ddTHH-mm-ss)
)#end paramter.

#Creates the log file for the advanced logger.
function createFiles(){
    if (!(test-Path -Path ".\userlog.log")){        
        New-Item .\userlog.log -ItemType file
        New-Item .\EmailBody.log -ItemType file        
        write-host "Created .\userlog.log." -foregroundcolor magenta
    }
}

#Imports the Logging module.
function ImpModule{
    #Imports the 'AdvancedModuleLogToFile' Module located at https://github.com/joer89/Logging/AdvancedModuleLogToFile.
    Import-Module $path\AdvancedLoggerModule.psm1 -ErrorAction Stop -DisableNameChecking
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "importing the modules."
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished importing the AdvancedLoggerModule"#Logs to running directory\logger.log with text.

    #Imports the O365Logger module.
    Import-Module  $path\O365EMail.psm1 -WarningAction Stop
     #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished importing O364Email module."

    #Imports the MSOnline module for Office365.
    Import-Module MsOnline -ErrorAction Stop
     #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished importing MSOnline module."

    #Imports the Active Directory module for conencting to AD.
    Import-Module ActiveDirectory -ErrorAction Stop
     #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished importing the Active Directory module."
    write-host "Imported all the modules." -foregroundcolor magenta

}#end function.

#Connects to Office365 
function connectToO365(){
    # #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Connecting to Office365."
    #Set-ExecutionPolicy RemoteSigned
    #Connect up to cloud and Import various Modules
    $UserCredential = Get-Credential 'AdminEMail@Office365Domain.com'
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection
    #Set-ExecutionPolicy RemoteSigned
    Import-PSSession $Session -AllowClobber
    Connect-MsolService -Credential $UserCredential
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished connecting to Office365."
    write-host "Connected to Office365." -foregroundcolor magenta
}

#Create temporary file for Licence stage later
function createLicenceFile(){ 
    #Logs to running directory\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Creating the licence file."
    "UserprincipalName,Username" | out-file -FilePath ".\user\$file.csv" -Append
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished."
    write-host "Created the licence file called ($($file).csv" -foregroundcolor magenta
}

#Appends the text for the user to be stored in EMailBody.log.
function storeUsersForEmail(){
    #Current logged on user.
    $Admin = Get-WMIObject -class Win32_ComputerSystem | select username

     $body = "
        New student account created at $(Get-Date -format "dd-MMM-yyyy HH:mm").
    
        Creator is $Admin</h3>
        Username is $Username.
        Fullname is $firstn $surname.
        Display name is $DisplayName.
        The SamAccountName is $samacctname.
        The UserPrincipalName is $userPrincipal.
        ChangePasswordAtLogon is true.  
        Account Enabled is true.
        Description is $Description.
        Profile path is $ppath.
        homeDirectory is $HomeDirectory and has to be created manually!!!!!!!!!!!
        Company is $school
        Department is $dept.
        The HomeDrive is U:.


    "
 
    write-host "Stored $($Username) for sending email." -foregroundcolor magenta
    #Stores the content of the email in the text file.
    Log-ToFile -Path .\ -fileName EmailBody.log -SimpleLogging -Text "$($body)"
    #Displays the attributes of the AD user on screen.
    Write-Host "Attributes of $($Username): $($body)"
}

# Set Attributes 3 & 11 for internal use (optional)
function attributeEditing(){ 
    Set-ADUser -Identity $Username -Add  @{extensionAttribute3="Student"}
    Set-ADUser -identity $Username -Add @{extensionAttribute11="TRS"}
    write-host "Added the ExtensionAttributes to $($Username)." -foregroundcolor magenta
}

#Create Security Permissions On Home Directory($Username)($dept)
function createSecurityPermissions(){

    if(Test-Path $folder){
            #Stores the Name from excel to $Group.
            $Group = $Username
            #Gets the access control list of the HomeDirectory from excel.
            $acl = Get-Acl $folder
            #Sets the Modify rule to $rule.
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$Group", "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
            #Sets the Modify rule.
            $acl.AddAccessRule($rule)
            #Sets the Read rule to $rule.
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$Group", "Read", "ContainerInherit, ObjectInherit", "None", "Allow")
            #Sets the Read rule.
            $acl.AddAccessRule($rule)
            #Sets the Writey rule to $rule.
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$Group", "Write", "ContainerInherit, ObjectInherit", "None", "Allow")
            #Adds the write rule.
            $acl.AddAccessRule($rule)
            #Sets the rules to the folder.
            Set-Acl $folder $acl
            $Activity = "Created security on $folder"
        }
        else{
            #Writes activity.
            $Activity = "failed to create security on $folder"
        } 
        write-host "Creating security permissions for $($Username)" -foregroundcolor magenta
}

#Next 3 lines swaps Proxy Address round so SMTP = our domain name
function setProxy(){
    $Newuser = Get-ADUser -Identity $username -Properties ProxyAddresses
    $Newuser.ProxyAddresses.add($proxy1)
    $Newuser.ProxyAddresses.add($proxy2)
    write-host "Set the proxy settings for $($Username)." -foregroundcolor magenta
}

#Next section adds user to Security Groups
function addUserSecurityGroups(){    
    Set-ADUser -instance $Username # $Newuser
    add-adgroupmember -identity 'ADSecurityGroup1' -members (Get-ADUser -filter "SamAccountName -eq '$Username'");
    add-adgroupmember -identity 'ADSecurityGroup2' -members (Get-ADUser -filter "SamAccountName -eq '$Username'")
    write-host "Adding the security groups for $($Username)." -foregroundcolor magenta
}

# Next line just prints the UPN to a date-sensitive csv file for the next phase
function createO365licenseFile{   
    $userPrincipal+","+$username | out-file -FilePath ".\user\$file.csv" -Append
    write-host "Added the UPN and username to $($file).csv for $($Username)." -foregroundcolor magenta
}

#Connect to Azure DC to run Delta synch
function connectToAzureLocalServer(){
    #Specify the credentials for Azure Server.
    #This is so you don't have to enter the password each time a user is created.
    $Username = 'domain\AzureAccount'
    $Password = 'AzurePassword'
    $pass = ConvertTo-SecureString -AsPlainText $Password -Force
    $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass 

    write-host "Connecting to the Azure Server and running DeltaSync in order to apply licences." -foregroundcolor magenta
    write-host "...this window will stay open for 60 seconds on a timer while syncing takes place." -foregroundcolor magenta
    Invoke-Command -ComputerName trs-dc2012-2 -Credential $Cred -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
    Start-Sleep -s 60
}

#Adds the Office365 licences for the user.
function addO365LicensesForUser(){
    ####################################### Import csv file of newly created accounts to apply students licences
    $filePath = ".\user\$file.csv"

    #####################  '-Header UserPrincipalName' has been removed as it's adding an extra user...
    Import-Csv $filePath | ForEach {
        $usr = $_.UserPrincipalName
        $usrname = $_.Username
        Set-MsolUser -UserPrincipalName $usr -UsageLocation GB
        Start-Sleep -s 2
        write-host ""
        # Apply Subscription Licences
        $O365PP = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_STUDENT -DisabledPlans OFFICE_FORMS_PLAN_2,INTUNE_O365
        Set-MsolUserLicense -UserPrincipalName $usr -AddLicenses regisschool:OFFICESUBSCRIPTION_STUDENT -LicenseOptions $O365PP
        Start-Sleep -s 3
        write-host ""

        # Apply Office pack Licences
        $O365PP2 = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_STUDENT -DisabledPlans TEAMS1,Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,RMS_S_ENTERPRISE,OFFICE_FORMS_PLAN_2,PROJECTWORKMANAGEMENT,INTUNE_O365,YAMMER_EDU,MCOSTANDARD,SHAREPOINTSTANDARD_EDU
        #Start-Sleep -s 2
        write-host "assigning license for $usr..." -foregroundcolor magenta
        Set-MsolUserLicense -UserPrincipalName $usr -AddLicenses regisschool:STANDARDWOFFPACK_STUDENT -LicenseOptions $O365PP2
        Start-Sleep -s 3
        write-host ""
        write-host "completed for $usr ..." -foregroundcolor magenta

        # Apply Policies
        Write-Host ""
        #Write-Host "Each entry will need at least a minute for the Cloud to create the mailbox before the Address Book Policies can apply...."
        #Start-Sleep -s 60

        #Set-Mailbox -identity $usrname -AddressBookPolicy "TRS Student ABP" -RoleAssignmentPolicy "Student Role Assignment Policy" -SharingPolicy "TRS Student Sharing Policy"    
        #$a = Read-Host "Please enter 'Y' to continue once all Mailboxes are created....Please check Cloud before continuing..."
        write-host "Finished adding the mailbox for $($Username)." -foregroundcolor magenta
    }
}

#Sets the Office365 mail box permissions.
function setMailboxForUser(){
    ####################################### Import csv file of newly created accounts to apply students licences *** use csv file created from cstudents.ps1 ***
    $filePath = ".\user\$file.csv"
    #####################  '-Header UserPrincipalName' has been removed as it's adding an extra user...
    Import-Csv $filePath | ForEach {
    $usr = $_.UserPrincipalName
    $usrname = $_.Username
    Set-Mailbox -identity $usrname -AddressBookPolicy "TRS Student ABP" -RoleAssignmentPolicy "Student Role Assignment Policy" -SharingPolicy "TRS Student Sharing Policy"
    Write-Host ""
    Write-Host "Address book policy has been set for >>>>>>   $usr" -ForegroundColor Magenta
    Write-Host ""
    }
}

#Disconnect PSSessions
function diconnectPSSession(){
    #Get all PSSessions and disconnect.
    Get-PSSession | ? {$_.computername -like “*.outlook.com”} | remove-pssession
    write-host "Disconnected from Office365." -foregroundcolor magenta
}

#Start the import of the users to create accounts in AD and assign various variables.
function ADCreateUser(){
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Importing .\Users.csv" 
    Import-Csv ".\Users.csv" | ForEach-Object {
        #$userPrincipal = $_.Username + "@theregisschool.co.uk"
        $path = "OU=TempStor,OU=Students,OU=Users,OU=TRS,DC=regis,DC=ult,DC=Org,DC=UK"
        $username = $_.Username
        $firstn = $_.Givenname
        $surname = $_.Surname
        $intake = $_.intake
        $dept = $_.Area
        $DisplayName = $firstn + " " + $surname
        $samacctname = $_.Username
        $userPrincipal = $firstn + "." + $surname + "@Domain.com"
        $Description = "UL - Student - " + (Get-Date -Format "%d-%M-%y")
        $HomeDirectory = "\\FileServer\Share$\$($intake)\$($Username)"
        $school = "Company"
        $dom = "@domain.mail.onmicrosoft.com"
        $proxy1 = "SMTP:" + $userPrincipal
        $proxy2 = "smtp:" + $firstn + "." + $surname + $dom
        $folder = $HomeDirectory
        $ppath = "\\FileServer\Share$\Mandatory"
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Creating the user $($Username)."
        New-ADUser -Name $Username `
            -Path $path `
            -SamAccountName  $samacctname `
            -UserPrincipalName  $userPrincipal `
            -AccountPassword (ConvertTo-SecureString "School1" -AsPlainText -Force) `
            -ChangePasswordAtLogon $true  `
            -Enabled $true `
            -Description $Description `
            -surname $surname `
            -givenName $firstn `
            -displayName $DisplayName `
            -profilePath $ppath `
            -homeDirectory $HomeDirectory `
            -Company $school `
            -Department $dept `
            -HomeDrive "U:" `
            -EmailAddress $userPrincipal
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished creating the user $($Username)." 
      
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished importing all the users."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Editing the attributes for user $($usr)."
        # Set Attributes 3 & 11 for internal use (optional)
        attributeEditing($Username)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished editing the attributes for $($usr)."
    
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Creating the security permissions for $($usr)."
        #Create Security Permissions On Home Directory($Username)($dept)
        createSecurityPermissions($username)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished creating the security permissions for $($usr)."
      
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Settng proxy addresses for $($usr)."
        #Configure SMTP and Domain name for user.
        setProxy($Username)($proxy1)($proxy2)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished setting proxy addresses for $($usr)."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Adding the user security groups to user $($usr)."
        #Next section adds user to Security Groups
        addUserSecurityGroups($Username)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished adding the security groups for user $($usr)."
        
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Running Delta Sync."
        #Connect to Azure DC to run Delta synch
        connectToAzureLocalServer    
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished running Delta Sync."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Creating the License file for Ofice365 for user $($usr)."
        #Next line just prints the UPN to a date-sensitive csv file for the next phase
        createO365licenseFile($userPrincipal)($username)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished creating the office365 license file for $($usr)."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Adding the Office 365 licences for $($usr)."
        ##Adds the Office365 licences for the user.
        addO365LicensesForUser($file)($usr)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished adding the office365 licences for $($usr)."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Setting the mailbox for user $($usr)."
        setMailboxForUser($file)($usr)
        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished setting up the mailbox for the user $($usr)."

        #Logs to .\userlog.log with text.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Storing $($usr) for emailing at the end of the program."
        #Appends the text for the user to be stored in EMailBody.log.
        storeUsersForEmail($Username)($surname)($firstn)($DisplayName)($samacctname)($userPrincipal)($Description)($ppath)($HomeDirectory)($school)($dept)
        #Logs to .\userlog.log with text and adds three new lines.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished storing the user details for $($Username)."
        #Draws -------- at the end of each user being added in the userlog.log file.
        Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "------------------------------------------------------------------------------------------------------------"
    }
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Disconnecting all PS Sessions."
    #Disconnects all PS Sessions.
    diconnectPSSession
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Finished disconnecting all PS Sessions."
}

#Creates the userlog.log and emailBody.log in current directory if it doesn't exist.
createFiles
#Imports all the modules.
ImpModule
#Connects to Office365.
connectToO365
#Creates the licence file (text file).
createLicenceFile($file)
#Adds the users to Active Directory and Office365.
ADCreateUser

#For testing, Just need to send the email! email keeps failing.

#Send the email for logging.
try{
    #Read EMailBody.log
    $EmailBodyOfAllUsrs = Get-Content -Path .\EMailBody.log
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Emailing the user's details."
    #Sends the email with the content of EmailBody.log.
    sendO365Mail $EmailBodyOfAllUsrs
}
catch {
    #Logs to .\userlog.log with text.
    Log-ToFile -Path .\ -fileName userlog.log -SimpleLogging -Text "Failed to email the users details, please send the EMailBody.log content manually."
    Write-Host "Failed to email the users details, please send the EMailBody.log content manually." -ForegroundColor Magenta
}