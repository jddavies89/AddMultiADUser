Things to do before using the script.
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

When you run the script, it creates a log in the directory of main.ps1 from the log.psm1 of the activity

Modules loaded.
Creating user user at 27-4-15 11:2:18
Configuring the User principal name.
user.name@theregisschool.co.uk
Finished configuring the User principal name.
Configuring the User email address.
username1@contoso.com
Finished configuring the User email address.
Configuring the user's home directory.
\\Server\students$\intake14\username1
Finished configuring the users home directory.
Configuring the user's OU location.
OU=Students,OU=Users,DC=contoso,DC=com
Finsihed configuring the user's OU location.
Configuring the user's description.
Student - 27-4-15
Finished configuring the user's description.
Creating the user account in Active Directory.
Finished crating the user account in Active Directory.
Configuring the user's attributes
Finished configuring the user's attributes.
Adding the user to the specified groups.
Finished adding the user to the specified groups.
Creating the user's home directory.
\\Server\students$\intake14\username1
Home drive created for username1.
Finished creating the user's home directory.
Creating the security permissions on the user's home directory.
Created security on \\Server\students$\intake14\username1
Finished creating the security permissions on the user's home directory.
Finished creating the user user at 27-4-15 11:2:18