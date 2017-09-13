<#
.Synopsis
   This script creates multiple users in Active Directory.
.DESCRIPTION    
    This script pulls all the users in Users.csv and adds them to Active Directory.

    Modules:
       log.psm1; Module for logging the activity of the main script, for a more detailed of the logging module, please view the description of the module.
    
    Things to do before using the script:
        1.	Under 'CreateSecurityPermissionsOnHomeDirectory' function, you need to change the $HomeDirectory to your UNC share.
        2.	Under 'CreateHomeDirectory' function, you need to change the $HomeDirectory to your UNC share.
        3.	Under 'configureGroups' function, you need to add your groups to which the user your creating is a part of.
        4.	Under 'configureAttribute' function, you need to specify want you want in extention attributes three and eleven if any.
        5.	Under 'configureOUPath' function, you need to specifiy where the users are going to be in Active Directory.
        6.	Under 'configureHomeDir' function, you need to change the $HomeDirectory to your UNC share.
        7.	Under 'configureEmail' function, the email address is in theformat of username@contoso.com
        8.	Under 'configureUPN' function, the UPN is in the format of firstname.surname@contoso.com
        9.	Under 'importModules' function, Both the modules Active directory and the loggin module are loaded. the module log.psm1 need to be in the directory of main.ps1
        10.	Under 'main' function, there is a foreach-object that imports all the users, this excel document needs to be a csv comma seperated file and the headers of GivenName, Surname, Username and Enabled(TRUE or FALSE) 
.EXAMPLE
.Notes
   Author: Joe Richards
   Date:   2015
.LINK
  https://github.com/joer89/AddMultiADUser.git
#>

#Start function.
#Creates the user's home directory. 
function CreateSecurityPermissionsOnHomeDirectory(){
    if($Area.ToLower() -eq "staff"){
        #Stores the Home directory for staff in \\Server\Staff$\%Username%.
        $HomeDirectory = "\\Server\Staff$\$($Username)" 
    }
    elseif ($Area.ToLower() -eq "student"){
         #If Area is student then the Home directory is  \\Server\student$\intake$($intake)\%Username%.        
         if ((get-date).Month -gt 07){     
                #If the month is greater than 07th month, format is intake current year.           
                $intake = (Get-Date -uformat %y)                
                $HomeDirectory = "\\Server\student$\intake$($intake)\$($Username)"
         }
         else{
                #If the month is less than 07th month, format is intake last year.      
                $intake = (get-date).AddYears(-1).ToString("%y")
                $HomeDirectory = "\\Server\students$\intake$($intake)\$($Username)"
         }
    }
    #Adding the rights to the user's home directory folder.
    #Stores the home directory.             
    $folder = $HomeDirectory
    #If the directory exists give the user permissions.
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
      #Returns the activity.
      return $Activity
 }
#End function.

#Start function.
#Creates the user's home directory. 
function CreateHomeDirectory(){
    if($Area.ToLower() -eq "staff"){
        #Stores the Home directory for staff in "\\trs-admin\Home$\Staff\%Username%.
        $HomeDirectory = "\\Server\Staff\$($Username)" 
    }
    elseif ($Area.ToLower() -eq "student"){
         #If Area is student then the Home directory is "\\brcc-media\intake$($intake)\%Username%.        
         if ((get-date).Month -gt 07){     
                #If the month is greater than 07th month, format is intake current year.           
                $intake = (Get-Date -uformat %y)                
                $HomeDirectory = "\\Server\students$\intake$($intake)\$($Username)"
         }
         else{
                #If the month is less than 07th month, format is intake last year.      
                $intake = (get-date).AddYears(-1).ToString("%y")
                $HomeDirectory = "\\Server\students$\intake$($intake)\$($Username)"
         }
    }
    #If the path does not exist. 
    if(!(Test-Path $HomeDirectory)){             
            #Creates the directory path.
            New-Item -ItemType Directory -Path $HomeDirectory
            #Stores activity.
            $Activity = "Home drive created for $($Username)."
    }   
    else{
           $Activity = "Home drive already exists for $($Username)."
    }
    #Returns the activity.
    return $Activity
 }
#End function.

#Start function
#Configure which groups the user needs to be in.
function configureGroups(){
    if($Area.ToLower() -eq "staff"){
         #If $Area is staff, add the user to the groups, 4051Staff, TRS All Staff Redirected Profiles and TRS - All Staff.
         add-adgroupmember -identity 'GroupName1' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")
         add-adgroupmember -identity 'Group Name2' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")
         add-adgroupmember -identity 'Group name3' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")
         $Activity = "Added GroupName1, Group Name2 and Group Name3."
     }
     elseif ($Area.ToLower() -eq "student"){
         #If $Area is student, add the user to the groups, 4051Student, TRS All Students and TRS - All Students.
         add-adgroupmember -identity 'Group name4' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")
         add-adgroupmember -identity 'Group Name5' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")             
         add-adgroupmember -identity 'Group name6' -member (Get-ADUser -filter "SamAccountName -eq '$Username'")             
         $Activity = "Added Group name4, Group Name5 and Group name6."
     }  
     return $Activity
}
#End Function

#Start function
#Add extention attributes three and eleven to the users, This is for the exchange server.
function configureAttribute(){
        if($Area.ToLower() -eq "staff"){
            #If Area is staff, add the following extention attributes to the user.
            #Staff and TRS. 
            Set-ADUser -Identity $Username -Add  @{extensionAttribute3="attribute3"}
            Set-ADUser -identity $Username -Add @{extensionAttribute11="attribute11"} 
        }
        elseif ($Area.ToLower() -eq "student"){
            #If Area is student, add the following extention attributes to the user.
            #Student and TRS.
            Set-ADUser -Identity $Username -Add  @{extensionAttribute3="attribute3"}
            Set-ADUser -identity $Username -Add @{extensionAttribute11="attribute11"}             
        }        
}
#End function.

#Start function
#Configures the description for each user.
function getDescription{
    if($Area.ToLower() -eq "staff"){
        #If Area is staff then ULT - Staff - Date of Day, Month and year.
        $Description = "Staff - " + (Get-Date -Format "%d-%M-%y")
    }
    elseif($Area.ToLower() -eq "student"){
        #If Area is staff then ULT - Student - Date of Day, Month and year.
        $Description = "Student - " + (Get-Date -Format "%d-%M-%y")
    }   
    #Returns the description.
    return $Description
}
#End function.

#Start function.
#Puts the user in the correct Organizational Unit in Active Directory.
function configureOUPath(){
    if($Area.ToLower() -eq "staff"){
        #If Area is a staff, the users created in the Staff Organizational Unit if it's staff.
        $Path = "OU=Staff,OU=Users,DC=contoso,DC=com"
    }
    elseif($Area.ToLower() -eq "student"){
        #If Area is a student, the users created in the Student Organizational Unit if it's student.
        $Path = "OU=Students,OU=Users,DC=contoso,DC=com"
    }
    #Returns the Organizational Unit's path.
    return $Path
 }
 #End function.
 
#Start function.
#Configures the user's home directory. 
function configureHomeDir(){
    if($Area.ToLower() -eq "staff"){
        #If Area is staff then the Home directory is \\Server\Staff$\%Username%.
        $HomeDirectory = "\\Server\Staff$\$($Username)"
    }
    elseif($Area.ToLower() -eq "student"){ 
         #If Area is student then the Home directory is "\\brcc-media\intake$($intake)\%Username%.        
         if ((get-date).Month -gt 07){     
                #If the month is greater than 07th month, format is intake current year.           
                $intake = (Get-Date -uformat %y)
                $HomeDirectory =  "\\Server\students$\intake$($intake)\$($Username)"
         }
         else{
                #If the month is less than 07th month, format is intake last year.      
                $intake = (get-date).AddYears(-1).ToString("%y")
                $HomeDirectory =  "\\Server\students$\intake$($intake)\$($Username)"
         }
    }
    #Returns the home directory.
    return $HomeDirectory
 }
#End function.

#Start function.
#Configures the email account.
function configureEmail($Username){
    #Configures the email account to %username%@theregisschool.co.uk.
    $email = "$($Username)@contoso.com"
    #Returns the user's email.
    return $email  
}
#End function.

#Start function.
#Configures the users UPN name.
function configureUPN($GivenName,$Surname){
    #Configures the users username to firstname.sirname@theregisschool.co.uk.
    $upn = "$($GivenName).$($Surname)@contoso.com"
    #Returns tghe UPN name.
    return $upn   
}
#End function.

#start function
#Imports the AD and logging module.
function importModules{
    #Imports the logging module.
    Import-Module .\log.psm1 -ErrorAction Stop
    #Imports the ActiveDirectory module.
    Import-Module ActiveDirectory -ErrorAction Stop    
}
#End function

#Start function.
#Retrieves all the users details from the CSV file and configures the paramters of the user.
function main(){ 
    #Imports the modules.
    importModules
    addlog "Modules loaded."
    
    #Imports the CSV File of users and creates the user accounts.
    Import-Csv -Path .\Users.csv | foreach-object {
        #For each object do the following.    
        try{          
            #Stores the GivenName, surname, username, area and enabled which is a booleon.
            $GivenName = $_.GivenName  
            $Surname = $_.Surname
            $Username = $_.Username
            $Area = $_.Area
            if(($_.Enabled.ToLower()) -eq "true"){
                    $Enabled = $True
            }
            else{
                   $Enabled = $False 
            }
            addlog ("Creating user $($GivenName) at " + (Get-Date -Format "%d-%M-%y %h:%m:%s"))
            addlog "Configuring the User principal name."
            #Configures the following parameters for the user account.
            #User's Principal name.
            $upn = configureUPN($GivenName)($Surname)
            addlog $upn 
            addlog "Finished configuring the User principal name."
            
            addlog "Configuring the User email address."
            #User's email address.
            $email = configureEmail($Username)
            addlog $email
            addlog "Finished configuring the User email address."
            
            addlog "Configuring the user's home directory."
            #User's home directory.
            $HomeDir =  configureHomeDir($Area)($Username)
            addlog $HomeDir
            addlog "Finished configuring the users home directory."
            
            addlog "Configuring the user's OU location."
            #User's Organizational Units path.
            $OUPath = configureOUPath($Area)
            addlog $OUPath
            addlog "Finsihed configuring the user's OU location."
            
            addlog "Configuring the user's description."
            #User's description.
            $getDescription = getDescription($Area)
            addlog $getDescription 
            addlog "Finished configuring the user's description."
            
            addlog "Creating the user account in Active Directory."  
            #Create the user with the parameters from the csv file and functions above.
            New-ADUser -Name $Username -Enabled $Enabled -SamAccountName $Username -GivenName $GivenName -Surname $Surname -UserPrincipalName $upn -AccountPassword (ConvertTo-SecureString $Username -AsPlainText -force) -ChangePasswordAtLogon $True -CannotChangePassword $False -PasswordNeverExpires $False -PasswordNotRequired $False -Description $getDescription -EmailAddress $email -HomeDirectory $HomeDir -HomeDrive "U:" -Path $OUPath
            addlog "Finished crating the user account in Active Directory."
            
            addlog "Configuring the user's attributes"            
            #Sets the following parameters for the user account.
            #Extention attribute three and eleven.
            $ConfAttributes = configureAttribute($Area)($Username)
            addlog "Finished configuring the user's attributes."
            
            addlog "Adding the user to the specified groups."
            #Sets the user to the correct groups.
            $ConfGroup = configureGroups($Area)($Username) 
            addlog "Finished adding the user to the specified groups."
            
            addlog "Creating the user's home directory."
            #Creates the home directory for the user.
            $CreateHomeDir = CreateHomeDirectory($Username)($Area)
            addlog $CreateHomeDir
            addlog "Finished creating the user's home directory."
            
            addlog "Creating the security permissions on the user's home directory."
            #Creates the security permissions on the users home directory.
            $CreateSecurityPermissions = CreateSecurityPermissionsOnHomeDirectory($Username)($Area)
            addlog $CreateSecurityPermissions
            addlog "Finished creating the security permissions on the user's home directory."
            addlog ("Finished creating the user $($GivenName) at " + (Get-Date -Format "%d-%M-%y %h:%m:%s"))
             
            }
            catch [Exception]{
                #Write the error to the log.
                addlog $_.ToString()
                #Writes out the exception if there is one.
                Write-Host $_
            }
            #End try catch.
    }  
    #End for each loop.
}
#End function.

#Clears the screen.
cls
#Calls the function.
main
