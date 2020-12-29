# Author  : Aleksandr Khoroshiy
# Date    : 01/01/2020
# Purpose : Automate account creation in AD DS based on information in SharePoint table
# Version : 1.0.0.0
# Details : Automate account creation in AD DS based on information in SharePoint table
# Notes   : 

Start-Transcript "C:\Scripts\logs\SP-newEmployee-$(get-date -Format hh-mm---dd-MM-yyyy).txt"

#-----------------------------------------------------------[Settings]------------------------------------------------------------
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue 
$autoapprove_members = "adm-user"
$AD_groups = "GPO-Homefolder"
$dc = "dc01-cl"
$WebURL = "https://portal.company.local/IT"
$listName = "Employees" 
$web = Get-SPWeb -identity $WebURL  
#-----------------------------------------------------------[Settings]------------------------------------------------------------


#-----------------------------------------------------------[Functions]------------------------------------------------------------
#Retrive objects
Function Retrieve {
    $list = $web.Lists[$listName] 
    return $list
}

#Mail
function mail {
    $from = "notification@company.local"
    $smtpserver = "Edge.office.local"

    #---mail
    Send-MailMessage  -to $mailto -from $from -body $bodyemail -subject $subject -smtpserver $smtpserver -DeliveryNotificationOption onfailure
}

Function GetUseraccount($UserValue) {
    #Uservalue: E.g: "1;#user name";
    $arr = $UserValue.Split(";#");
    $UserID = $arr[0];
    $user = $web.SiteUsers.GetById($UserId);    
    #the above line returns: SPUser Object
    return ($user.LoginName -split '\\')[1]
}

Function GetUsermail($UserValue) {
    #Uservalue: E.g: "1;#user name";
    $arr = $UserValue.Split(";#");
    $UserID = $arr[0];
    $user = $web.SiteUsers.GetById($UserId);
    $user_name = ($user."LoginName" -split "\\")[1]    
    #the above line returns: SPUser Object
    return (    Get-Recipient $user_name   ).primarysmtpaddress
}

Function Maildomain ($n) {
    switch ($n) {
        { $n -eq "brand1" } { "brand1.email" }
        { $n -eq "brand2" } { "brand2.email" }
        { $n -eq "brand3" } { "brand3.email" }
        { $n -eq "brand4" } { "brand4.email" }
        { $n -eq "brand5" } { "brand5.online" }
        default { "company.local" }
    }
}


Function User-OU ($n) {
    switch ($n) {
        # Ukraine
        { $n["Department"] -eq "Sales EN" -and $n["Office"] -eq "Ukraine" } { "OU=Sales,OU=Users,OU=Ukraine,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "Sales RU" -and $n["Office"] -eq "Ukraine" } { "OU=Sales,OU=Users,OU=Ukraine,OU=Offices,OU=Company,DC=company,DC=local"; continue }
    
        # Israel
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Israel" } { "OU=BackOffice,OU=Users,OU=Israel,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Israel" } { "OU=BackOffice,OU=Users,OU=Israel,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Israel" } { "OU=BackOffice,OU=Users,OU=Israel,OU=Offices,OU=Company,DC=company,DC=local"; continue } 

        # Spain
        { $n["Department"] -eq "Finance" -and $n["Office"] -eq "Spain" } { "OU=Finance,OU=Users,OU=Israel,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -like "Retention*" -and $n["Office"] -eq "Spain" } { "OU=Retention,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -like "Sales*" -and $n["Office"] -eq "Spain" } { "OU=Sales,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -eq "administration" -and $n["Office"] -eq "Spain" } { "OU=Administration,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -eq "management" -and $n["Office"] -eq "Spain" } { "OU=Management,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -eq "hr" -and $n["Office"] -eq "Spain" } { "OU=HR,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -like "Sales*" -and $n["Office"] -eq "Spain" } { "OU=Sales,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 
        { $n["Department"] -like "Sales*" -and $n["Office"] -eq "Spain" } { "OU=Sales,OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local"; continue } 


        # Albania
        { $n["Department"] -like "Sales*" -and $n["Office"] -eq "Albania" } { "OU=Sales,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "IT" -and $n["Office"] -eq "Albania" } { "OU=IT,OU=Sales,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Albania" } { "OU=Administration,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "Management" -and $n["Office"] -eq "Albania" } { "OU=Management,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "Quality Control" -and $n["Office"] -eq "Albania" } { "OU=Quality Control,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -like "Retention*" -and $n["Office"] -eq "Albania" } { "OU=Retention,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }
        { $n["Department"] -eq "IT" -and $n["Office"] -eq "Albania" } { "OU=IT,OU=Users,OU=Albania,OU=Offices,OU=Company,DC=company,DC=local"; continue }

        # Rest
        { $n["Brand"] -eq "Trade" -and $n["Department"] -eq "Sales RU" -and $n["Office"] -eq "Ukraine" } { "OU=RU,OU=Trade,OU=Sales,OU=Users,OU=Ukraine,OU=Offices,OU=Company,DC=company,DC=local"; continue }



        { $n["Office"] -eq "Ukraine" } { "OU=Users,OU=Ukraine,OU=Offices,OU=Company,DC=company,DC=local" }
        { $n["Office"] -eq "Bulgaria" } { "OU=Users,OU=Bulgaria,OU=Offices,OU=Company,DC=company,DC=local" }
        { $n["Office"] -eq "Cyprus" } { "OU=Users,OU=Cyprus,OU=Offices,OU=Company,DC=company,DC=local" }
        { $n["Office"] -eq "Israel" } { "OU=Users,OU=Israel,OU=Offices,OU=Company,DC=company,DC=local" }
        { $n["Office"] -eq "Romania" } { "OU=Users,OU=Romania,OU=Offices,OU=Company,DC=company,DC=local" } 
        { $n["Office"] -eq "Spain" } { "OU=Users,OU=Spain,OU=Offices,OU=Company,DC=company,DC=local" }
        { $n["Office"] -eq "Albania" } { "OU=Albania,OU=Offices,OU=Company,DC=company,DC=local" }
        default { "CN=Users,DC=company,DC=local" }
    }
}


Function mail-prefix ($n) {
    switch ($n) {
        { $n["Brand"] -eq "Fincenter" -and $n["Department"] -eq "Presale" } { "$($n["FirstName"]).$( ($n["LastName"]  -split " ")[0])" }
        { $n["Brand"] -eq "Trade" -and ($n["Department"] -like "Sales*" -or $n["Department"] -like "Retention*") } { "$($n["FirstName"]).$( ($n["LastName"]  -split " ")[0])" }
    }
}

Function mail-access ($n) {
    switch ($n) {
        { $n["Department"] -eq "IT" -or $n["Department"] -eq "Help Desk" } { $true, $true, $false, $false, $true }
        default { $false, $true, $false, $false, $true }
    }
}
    
Function Mailbox-quota ($n) {
    switch ($n) {
        { $n -eq "IT" -or $n -eq "Help Desk" -or $n -eq "Management" -or $n -eq "Finance" -or $n -eq "HR" -or $n -eq "Marketing" } { 4GB, 4.5GB, 5GB }
        default { 1GB, 1.5GB, 2GB }
    }
}

Function Mailbox-recipient ($n) {
    switch ($n) {
        { $n -eq "IT" -or $n -eq "Help Desk" -or $n -eq "Management" -or $n -eq "Finance" -or $n -eq "HR" -or $n -eq "Marketing" } { "500" }
        default { "10" }
    }
}

Function Trainers ($n) {
    switch ($n) {
        { $n["Office"] -eq "Ukraine" -and $n["Department"] -eq "Sales EN" -and ($n["Brand"] -ne "Trade1" -and $n["Brand"] -ne "brand1") } { "Matt.Besler@brand12.email" }
        { $n["Office"] -eq "Ukraine" -and $n["Department"] -eq "Sales RU" -and ($n["Brand"] -ne "Trade1" -and $n["Brand"] -ne "brand1") } { "brand1@brand1.online"  }
        { $n["Office"] -eq "Ukraine" -and $n["Department"] -eq "Retention RU" -and ($n["Brand"] -ne "Trade1" -and $n["Brand"] -ne "brand1") } { "user@brand1.online" }
        { $n["Office"] -eq "Ukraine" -and $n["Department"] -eq "Retention EN" -and ($n["Brand"] -ne "Trade1" -and $n["Brand"] -ne "brand1") } { "user@brand1.email" }
        { $n["Brand"] -eq "Trade1" } { "Michael.Emerson@brand12.email" }
        #default {"CN=Users,DC=company,DC=local"}
    }
}

Function SIP-script ($n) {
    switch ($n) {
        { $n["Brand"] -eq "Trade1" -or $n["Brand"] -eq "Trade" -or $n["Department"] -like "Sales*" } { "logon_2.bat"; continue }
        { $n["Brand"] -eq "brand14" } { "OfficePhones_1.vbs"; continue }
        default { "logon.bat" }
    }
}

Function HRgroup ($n) {
    switch ($n) {
        { $n["Office"] -eq "Bulgaria" } { "UA_Office_HR@company.local" }
        { $n["Office"] -eq "Cyprus" } { "hr@brand1.com" }
        { $n["Office"] -eq "Israel" -and $n["Brand"] -eq "brand16" } { "IL-HR-brand1@company.local" }
        { $n["Office"] -eq "Israel" -and ($n["Brand"] -eq "Capital" -and $n["Brand"] -eq "RPC" ) } { "IL-HR@company.local" }
        { $n["Office"] -eq "Israel" -and ($n["Brand"] -ne "Capital" -and $n["Brand"] -ne "RPC" -and $n["Brand"] -ne "brand16" ) } { "HR_Israel@company.local" }
        { $n["Office"] -eq "Romania" } { "Brindusa.Bira@brand1.ro" } 
        { $n["Office"] -eq "Spain" } { "HR_S@company.local", "UA@company.local" }
        #default {"CN=Users,DC=company,DC=local"}
    }
}
        
Function HDgroup ($n) {
    switch ($n) {
        { $n["Office"] -eq "Ukraine" } { "UA_HD@company.local" }
        { $n["Office"] -eq "Israel" } { "IL_HD@company.local" } 
        { $n["Office"] -eq "Spain" } { "HelpDesk-Spain@company.local" }
        { $n["Office"] -eq "Albania" } { "HelpDesk-Spain@company.local" }
        default { "HD_GLOBAL@company.local" }
    }
}


Function user-folder ($n) {
    switch ($n) {
        default { "\\files.office.local\users\" }
    }
}


Function Team-group ($n) {
    switch ($n) {
        #Ukraine office
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Ukraine" } { "Team-administration-UA" }
        { $n["Department"] -eq "Back office" -and $n["Office"] -eq "Ukraine" } { "Team-administration-UA" }
        { $n["Department"] -eq "Customer Support and Compliance" -and $n["Office"] -eq "Ukraine" } { "Team-Customer Support and Compliance" }
        { $n["Department"] -eq "Dealing Room" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Dealing Room" }
        { $n["Department"] -eq "Finance" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Finance" }
        { $n["Department"] -eq "Fin-center" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Fin-center" }
        { $n["Department"] -eq "HR" -and $n["Office"] -eq "Ukraine" } { "Team-UA-HR" }
        { $n["Department"] -eq "IT" -and $n["Office"] -eq "Ukraine" } { "Team-IT" }
        { $n["Department"] -eq "Marketing" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Marketing" }
        { $n["Department"] -eq "Onboarding" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Onboarding" }
        { $n["Department"] -eq "Operating personnel" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Operating personnel" }
        { $n["Department"] -eq "Operations" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Operations" }
        { $n["Department"] -eq "Payments" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Payments" }
        { $n["Department"] -eq "Product" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Product" }
        { $n["Department"] -eq "Presale" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Presale" }
        { $n["Department"] -eq "Retention" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Retention" }
        { $n["Department"] -eq "Retention RCP" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Retention RCP" }
        { $n["Department"] -eq "Retention EN" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Retention EN" }
        { $n["Department"] -eq "Retention RU" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Retention RU" }
        { $n["Department"] -eq "Sales AR" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Sales AR" }
        { $n["Department"] -eq "Sales EN" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Sales EN" }
        { $n["Department"] -eq "Sales ES" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Sales ES" }
        { $n["Department"] -eq "Sales RCP" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Sales RCP" }
        { $n["Department"] -eq "Sales RU" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Sales RU" }
        { $n["Department"] -eq "Service" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Service" }
        { $n["Department"] -eq "Training" -and $n["Office"] -eq "Ukraine" } { "Team-UA-Training" }
        { $n["Department"] -eq "Innovecs" -and $n["Office"] -eq "Ukraine" } { "Team-Innovecs" }
        { $n["Department"] -eq "Sales German" -and $n["Office"] -eq "Ukraine" } { "Team-UA-SalesGerman" }
        { $n["Department"] -eq "Quality Control" -and $n["Office"] -eq "Ukraine" } { "Team-Ua-QC" }
        { $n["Company"] -eq "Credit Sense" -and $n["Office"] -eq "Ukraine" } { "Team-UA-CreditSense" }
        { $n["Company"] -eq "brand16" -and $n["Office"] -eq "Ukraine" } { "Team-UA-brand16" }

        #Israel office
        { $n["Department"] -eq "HR" -and $n["Office"] -eq "Israel" } { "Team-IL-HR" }
        { $n["Department"] -eq "Administration" -and $n["Office"] -eq "Israel" } { "Team-IL-Administration" }
        { $n["Department"] -eq "Help Desk" -and $n["Office"] -eq "Israel" } { "Team-IL-Help Desk" }
        { $n["Department"] -eq "Finance" -and $n["Office"] -eq "Israel" } { "Team-IL-Finance" }
        { $n["Department"] -eq "Dealing Room" -and $n["Office"] -eq "Israel" } { "Team-IL-Dealing room" }
        { $n["Department"] -eq "Management" -and $n["Office"] -eq "Israel" } { "Team-IL-Management" }
        { $n["Department"] -eq "Retention RCP" -and $n["Office"] -eq "Israel" } { "Team-IL-Retention RCP" }
        { $n["Department"] -eq "Marketing" -and $n["Office"] -eq "Israel" } { "Team-IL-Marketing" }
        { $n["Department"] -eq "Sales AR" -and $n["Office"] -eq "Israel" } { "Team-IL-Sales AR" }
        { $n["Department"] -eq "Sales RCP" -and $n["Office"] -eq "Israel" } { "Team-IL-Sales RCP" }
        { $n["Department"] -eq "Retention" -and $n["Office"] -eq "Israel" } { "Team-IL-Retention" }
                                                                                                                                                                           
        #Spain office
        { $n["Department"] -eq "Operations" -and $n["Office"] -eq "Spain" } { "Team-SP-Operations" }
        { $n["Department"] -eq "Back Office" -and $n["Office"] -eq "Spain" } { "Team-SP-Back Office" }
        { $n["Department"] -eq "Help Desk" -and $n["Office"] -eq "Spain" } { "Team-SP-Help Desk" }
        { $n["Department"] -eq "Finance" -and $n["Office"] -eq "Spain" } { "Team-SP-Finance" }
        { $n["Department"] -eq "QA/QC" -and $n["Office"] -eq "Spain" } { "Team-SP-QA/QC" }
        { $n["Department"] -eq "Training" -and $n["Office"] -eq "Spain" } { "Team-SP-Training" }
        { $n["Department"] -eq "Customer Service" -and $n["Office"] -eq "Spain" } { "Team-SP-Customer Service" }
        { $n["Department"] -eq "Retention" -and $n["Office"] -eq "Spain" } { "Team-SP-Retention" }
        { $n["Department"] -eq "HR" -and $n["Office"] -eq "Spain" } { "Team-SP-HR" }
        { $n["Department"] -eq "Management" -and $n["Office"] -eq "Spain" } { "Team-SP-Management" }
        { $n["Department"] -eq "Sales & Retention" -and $n["Office"] -eq "Spain" } { "Team-SP-Sales & Retention" }
        { $n["Department"] -eq "Sales ES" -and $n["Office"] -eq "Spain" } { "Team-SP-Sales ES" }

        #all offices 
        { $n["Department"] -eq "Help Desk" } { "Team-HelpDesk" }  
        default { "Team-unsorted" }
    }
}


Function user-personalfolder ($N) {
    New-Item -name $N["Login"] -ItemType Directory -Path "\\files\Users"
    Disable-NTFSAccessInheritance (Get-Item "\\files\Users\$($N["Login"])")
    Get-NTFSAccess -Path (Get-Item "\\files\Users\$($N["Login"])") -Account "OFFICE\Domain Users" | Remove-NTFSAccess
    add-NTFSAccess -AccessRights FullControl -Account "OFFICE\$($N["Login"])" -Path (Get-Item "\\files\Users\$($N["Login"])")
    add-NTFSAccess -AccessRights    GenericRead, GenericExecute -Account "OFFICE\fs-permissions-r" -Path (Get-Item "\\files\Users\$($N["Login"])")

    #Owner
    $acct1 = New-Object System.Security.Principal.NTAccount((Get-remADUser $N["Login"]).userprincipalname)
    $profilefolder = (Get-Item "\\files\Users\$($N["Login"])")
    $acl1 = $profilefolder.GetAccessControl()
    $acl1.SetOwner($acct1)
    set-acl -aclobject $acl1 -path $profilefolder
}



Function user-personalfolder-notification {
    $mailto = (GetUsermail $item["Manager"]), "$(mail-prefix ($item))@$($maildomain)"
    $subject = "Personal drive - Created successfully"

    $bodyemail = "Hi $($fullname)!

To improve security regarding storing company files, you have been attached your 'Personal network drive', where you should store your work files. Storing files in network 'Personal network drive' - gives possibility to centrally backup your files and protect them against damage, viruses\encryptors, accidental\forcible remove of data or any other cases\types of data loss.
You'll find your personal drive as:
1. Q drive in your system, in 'My Computer' section. 
2. Via shortcut 'Personal folder' on your desktop.
3. You can access it via link '\\files\users\<username>'. 

Automatically connecting Q drive to your computer - require reboot of your PC. So please reboot your PC.
Only you personally have access to your information inside your personal Q drive. The quota for personal files - 2 GB. In case of demand – this quota will be extend.
Make sure you store all your work files in this drive and migrated your existing files.

In case of questions please contact to Help Desk team.

"
    Mail
}
#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Logic]------------------------------------------------------------
#-------------------------[Retrieve items]-------------------------
$items = (Retrieve).items

#auto approve for changes
$items_changeded = $items | ? { $_["Approval Status"] -eq 2 -and $_["Modified"] -gt (get-date).AddMinutes(-15) -and $_["Status"] -eq "Active" -and $_["Created"] -lt (get-date).AddMinutes(-15) }
$items_changeded.name
$items_changeded | % { $item = $_; if ( $autoapprove_members -contains (GetUseraccount $item["Modified By"]) ) { $item.ModerationInformation.Status = "Approved"; $item.update() } }

#Pending for approve
$items_pending_create = $items | ? { $_["Approval Status"] -eq 2 -and $_["Modified"] -gt (get-date).AddMinutes(-5) -and $_["Status"] -eq "Create" }
$items_pending_create.name

#Unapproved pending reminder
$items_pending_unapproved = $items | ? { $_["Approval Status"] -eq 2 -and ( (get-date).hour -eq 8 -or (get-date).hour -eq 10 -or (get-date).hour -eq 12 -or (get-date).hour -eq 14 -or (get-date).hour -eq 16 -or (get-date).hour -eq 18 -or (get-date).hour -eq 20) -and (get-date).minute -lt 5 }
$items_pending_unapproved.name

#Approved for creation
$items_approved = $items | ? { $_["Approval Status"] -eq 0 -and $_["Modified"] -gt (get-date).AddMinutes(-100) -and $_["Status"] -eq "Create" }
$items_approved.name

#Approved for suspending
$items_suspend = $items | ? { $_["Approval Status"] -eq 0 -and $_["Modified"] -gt (get-date).Adddays(-90) -and $_["Status"] -eq "Suspend" }
$items_suspend.name

#Approved for change
$items_changeded = $items | ? { $_["Approval Status"] -eq 0 -and $_["Modified"] -gt (get-date).AddMinutes(-5) -and $_["Status"] -eq "Active" -and $_["Created"] -lt (get-date).AddMinutes(-15) }
$items_changeded.name

#Pending for suspend
$items_pending_suspend = $items | ? { $_["Approval Status"] -eq 2 -and $_["Modified"] -gt (get-date).AddMinutes(-5) -and $_["Status"] -eq "Suspend" }
$items_pending_suspend.name
#-------------------------[Retrieve items]-------------------------

#-------------------------[Import modules]-------------------------
if ($items_approved -or $items_suspend -or $items_changeded -or $items_pending_suspend -or $items_pending_create) {
    #Active directory
    $s = New-PSSession -ComputerName $dc
    Invoke-Command -ScriptBlock { Import-Module activedirectory } -Session $s
    Import-PSSession -Session $s -Module activedirectory -Prefix rem

    #Exchange
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch01-fr.office.local/PowerShell/ 
    Import-PSSession $Session

}
#-------------------------[Import modules]-------------------------

#-------------------------[Create - Pending approve]-------------------------
foreach ($item in $items_pending_create) {
    ######MAIL##############################
    $mailto = "ITSystem@company.local", (HDgroup $item)
    $subject = "Employee request - Pending create"
    $bodyemail = "
There is new pending request for new employee:
https://portal.office.local/IT/Lists/Employees/DispForm.aspx?ID=$($item.id)&e=Pk9kyY

Employee: $($item["FirstName"]+" "+ $item["LastName"])
Login: $($item["Login"])
Account type: $($item["Account"])

Title: $($item["Position"])
Department: $($item["Department"])
Brand: $($item["Brand"])
Office: $($item["Office"])
Manager: $( $manager )


Accesses: $(($item["Accesses"]) -split(";#"))

Requested by: $((Get-remADUser (GetUseraccount $item["Created By"])).name)

"

    #Send mail
    mail

    ######MAIL##############################
}
#-------------------------[Create - Pending approve]-------------------------


#-------------------------[Create - Approved]-------------------------
foreach ($item in $items_approved) {
    #Settings#
    $fullname = ($item["FirstName"] + " " + $item["LastName"]).trim()
    $crm_account = $item["Login"] + (Get-Random -Minimum 11 -Maximum 99)
    $password_string = "!Password" + (Get-Random -Minimum 1000 -Maximum 9999)
    $Password = ConvertTo-SecureString ($password_string) -AsPlainText -Force 
    $upnsuffix = "office.local"
    $maildomain = Maildomain $item["Brand"]
    $UPN = "$(mail-prefix ($item))@$($maildomain)"
    ${manager-identity} = (GetUseraccount $item["Manager"])
    $manager = if (get-remaduser -Filter "samaccountname -eq '$manager'"  ) { (get-remaduser -Filter "samaccountname -eq '$manager'").name } else { ${manager-identity} }
    


    #CRM
    if ($item["Accesses"] -split ";#" -contains "CRM - Proftit") { $CRM = $true; $Password_crm = Get-Random (1000..9999) }
 
    #Trillian
    if ($item["Accesses"] -split ";#" -contains "Trillian") { $Trillian = $true }
    if ($item["Accesses"] -split ";#" -contains "Mailbox") { $Access_Mailbox = $true }


    #Creating AD accounts
    try {
        new-remaduser -name $fullname  -DisplayName $fullname -GivenName $item["FirstName"] -Surname $item["LastName"] `
            -UserPrincipalName $UPN -SamAccountName  $item["Login"]  -AccountPassword $Password -path (User-OU $item) `
            -OtherAttributes @{title = $item["Position"]; Department = $item["Department"] } -Enabled:$true -ChangePasswordAtLogon $true -Manager $( ( get-remaduser (GetUseraccount $item["Manager"])).samaccountname) -Company $item["Brand"] -Server $dc -ScriptPath (SIP-script $item) -HomeDirectory "$(user-folder $item)$($item["Login"])" -HomeDrive "Q:"
    }
    catch {
        Write-Error "Account for $($item["Login"]) hasn't been created dur error:
    $($_.Exception.Message)"
        return
    }

    #membership
    #if($AD_groups){$AD_groups|%{Add-remADGroupMember $_ -Members $item["Login"] -Server $dc}}

    Set-remADUser $item["Login"] -EmployeeID $item["Employee ID"] -Server $dc
    
    #Set Password policy
    Add-remADFineGrainedPasswordPolicySubject "Password policy - 12 symbols" -Subjects $item["Login"] -Server $dc
 
    #Set-remADUser -Identity $item["Login"] -UserPrincipalName $upnsuffix -Server $dc
  
    #Secondary account
    if ($item["Account"] -eq "Secondary") { Set-remADUser $item["Login"] -Description "This is secondary account for $(($items|?{$_["Employee ID"] -eq $item["Employee ID"] -and $_["Account"] -eq "Primary" })["Login"])" }

    #Exchange mailbox
    If ($Access_Mailbox) {
        try {
            "Enabling Exchange"
            Enable-Mailbox -Identity $item["Login"] -PrimarySmtpAddress ("$(mail-prefix ($item))@$($maildomain)") -DomainController $dc
            Set-Mailbox  -Identity $item["Login"] -RecipientLimits (Mailbox-recipient $item["Department"]) -IssueWarningQuota (Mailbox-quota $item["Department"])[0] -ProhibitSendQuota (Mailbox-quota $item["Department"])[1] -ProhibitSendReceiveQuota (Mailbox-quota $item["Department"])[2] -UseDatabaseQuotaDefaults $false -DomainController $dc -SingleItemRecoveryEnabled $true -RetainDeletedItemsFor 120 -EmailAddressPolicyEnabled $false -PrimarySmtpAddress ("$(mail-prefix ($item))@$($maildomain)")
            Set-CASMailbox $item["Login"] -ActiveSyncEnabled (mail-access $item)[0] -OWAEnabled (mail-access $item)[1] -PopEnabled (mail-access $item)[2] -ImapEnabled (mail-access $item)[3] -MAPIEnabled (mail-access $item)[4] -DomainController $dc
        }
        catch {
            Write-Error "Mailbox for $($item["Login"]) hasn't been created due error:
    $($_.Exception.Message)"
            return
        }
    }


    #Trillian
    if ($Trillian -eq $true -and (user-trillian $item)) { Add-remADGroupMember  (user-trillian $item) -Members $item["Login"] }

    #create personal folder
    try {
        "Creating personal folder for $($item["Login"])"
        user-personalfolder $item
    }
    catch {
        Write-Error "Personal folder for $($item["Login"]) hasn't been created due error:
$($_.Exception.Message)"
    }

    #Create mail user for Clicklogic users
    if ($item["Brand"] -eq "brand16") { Enable-MailUser -Identity $item["Login"] -PrimarySmtpAddress ("$(mail-prefix ($item))@$($maildomain)") -ExternalEmailAddress ("$(mail-prefix ($item))@$($maildomain)") -DomainController $dc }

    #Team group
    if (Team-group $item) { Add-remADGroupMember  (Team-group $item) -Members $item["Login"] }

    ######MAIL##############################
    $mailto = "ITSystem@company.local", (HDgroup $item), "$(mail-prefix ($item))@$($maildomain)"
    $mailto += (HRgroup $item)
    if ((GetUsermail $item["Manager"])) { $mailto += (GetUsermail $item["Manager"]) }
    if (trainers $item) { $mailto += (trainers $item) }

    $subject = "Employee request - Created successfully"
    $bodyemail = "
New employee account has been created:

Employee: $fullname
Login: $($item["Login"])
Login as UPN: $UPN
$(if ($access_mailbox){"Email address: $(mail-prefix ($item))@$($maildomain) "})
$(if (-not $access_mailbox -and $item["Brand"] -eq "brand16"){"Mail contact address: $(mail-prefix ($item))@$($maildomain) "})
Password: $password_string
Account type: $($item["Account"])

Title: $($item["Position"])
Department: $($item["Department"])
Brand: $($item["Brand"])
Office: $($item["Office"])
Manager: $( $manager )
"
    if ($CRM) {
        $bodyemail += "
Proftit CRM Details
CRM Login: $($crm_account)
CRM Password: $($item["Login"])$($Password_crm)
*This credentials will be activated in ~30 minutes.
"
    }

    $bodyemail += "

Requested by: $((Get-remADUser (GetUseraccount $item["Created By"])).name)
Approved by: $((Get-remADUser (GetUseraccount $item["Modified By"])).name)

Start: $($item["Start"])"

    #Send mail
    mail


    $bodyemail >> "C:\Scripts\logs\SP-newEmployee-$(get-date -Format hh-mm---dd-MM-yyyy).txt"

    #Personal drive notification
    user-personalfolder-notification

    ######MAIL##############################


    #Set status on "Active"
    $item["Status"] = "Active"
    $item.ModerationInformation.Status = "Approved"  
    $item.update()
}
#-------------------------[Create - Approved]-------------------------




#-------------------------[Suspend - Pending approve]-------------------------
foreach ($item in $items_pending_suspend) {
    ${manager-identity} = (GetUseraccount $item["Manager"])
    $manager = if (get-remaduser -Filter "samaccountname -eq '$manager'"  ) { (get-remaduser -Filter "samaccountname -eq '$manager'").name } else { ${manager-identity} }
    

    ######MAIL##############################
    $mailto = "ITSystem@company.local", (HDgroup $item)
    $mailto += (HRgroup $item)
    $subject = "Employee request - Pending suspend"
    $bodyemail = "
There is new pending request for suspend of employee:

https://portal.office.local/IT/Lists/Employees/DispForm.aspx?ID=$($item.id)&e=Pk9kyY
Employee: $($item["FirstName"]+" "+ $item["LastName"])
Login: $($item["Login"])
Account type: $($item["Account"])

Title: $($item["Position"])
Department: $($item["Department"])
Brand: $($item["Brand"])
Office: $($item["Office"])
Manager: $( $manager )

Last user log in: $($item["User log in"])
Last mailbox log in: $($item["Mailbox log in"])s

Resign date: $($item["Resign"])

Description: $($item["Description"])

Requested by: $((Get-remADUser (GetUseraccount $item["Modified By"])).name)
"

    #Send mail
    mail

    ######MAIL##############################
}
#-------------------------[Suspend - Pending approve]-------------------------

#-------------------------[Suspend - Approved]-------------------------
foreach ($item in $items_suspend) {
    $leave = $item["Resign"]
    $name = $item.name
    $identity = (Get-remADUser $item["Login"]).distinguishedname 
    $login = $item["Login"]
    ${manager-identity} = (GetUseraccount $item["Manager"])
    $manager = if (get-remaduser -Filter "samaccountname -eq '${manager-identity}'"  ) { (get-remaduser -Filter "samaccountname -eq '${manager-identity}'").name } else { ${manager-identity} }
    


   
    #Suspend logic
    try {
        set-remaduser $item["Login"] -replace @{accountExpires = ($item["Resign"]).ToFileTimeUTC() } 
        Move-remADObject (Get-remADUser $item["Login"]).distinguishedname   -TargetPath "OU=Users,OU=For remove,OU=Company,DC=company,DC=local"
    }
    catch {
        Write-host "Account for $($item["Login"]) hasn't been suspended dur error:
    $($_.Exception.Message)"
        return
    }




    #Microsoft Exchange - mailbox resign actions
    if (Get-Mailbox -Filter { samaccountname -eq  $login }) {
        Set-Mailbox $identity -HiddenFromAddressListsEnabled $true
        #Remove activesync devices
        Get-MobileDevice -Mailbox $identity | Remove-MobileDevice -Confirm:$false
        $primarysmtpaddress = $((Get-Mailbox $item["Login"]).primarysmtpaddress)
    }

    #Remove membership
    #$Groups=(Get-remADPrincipalGroupMembership $item["Login"]|?{ $_.distinguishedname -notin (Get-remADUser $item["Login"] -Properties *).PrimaryGroup }).name
    $Groups = (Get-remADUser $item["Login"] -Properties *).memberof #|?{ $_ -notin (Get-remADUser $item["Login"] -Properties *).PrimaryGroup }))
    $Groups | % { Remove-remADGroupMember $_ -Members $item["Login"] -Confirm:$false }

    #Rename users home folder
    if ((test-path "\\fs02-nl\g$\Users\$($item["Login"])") -eq $true) {
        Rename-Item "\\fs02-nl\g$\Users\$($item["Login"])" -NewName "\\fs02-nl\g$\Users\resigned-user-$($item["Login"])"
        Move-Item "\\fs02-nl\g$\Users\resigned-user-$($item["Login"])" -Destination "\\fs01-nl.office.local\File Archive\Users"
    }

    <#
    if (get-item "\\colo-storage\h$\Users\$($item["Login"])"  -ErrorAction SilentlyContinue) {
        Rename-Item "\\colo-storage\h$\Users\$($item["Login"])" -NewName "\\colo-storage\h$\users\resigned-$($item["Login"])"
        Move-Item "\\colo-storage\h$\Users\resigned-$($item["Login"])" -Destination "\\fs01-nl.office.local\File Archive\Users"
    }
    #>

    if ((test-path "\\fs02-nl\j$\Homefolder\$($item["Login"])") -eq $true) {
        Rename-Item "\\fs02-nl\j$\Homefolder\$($item["Login"])" -NewName "\\fs02-nl\j$\Homefolder\resigned-home-$($item["Login"])"
        Move-Item "\\fs02-nl\j$\Homefolder\resigned-home-$($item["Login"])" -Destination "\\fs01-nl.office.local\File Archive\Users"
    }


    ######MAIL##############################
    $mailto = "ITSystem@company.local", (HDgroup $item), "AlexanderV@WRdirect.online"
    $mailto += (HRgroup $item)
    #$mailto="aleksandrkh@wrdirect.online"
    $subject = "Employee request - Suspended successfully"
    $bodyemail = "
Following account will be suspended:

Employee: $($item["FirstName"]+" "+ $item["LastName"])
Login: $($item["Login"])
Email address: $($primarysmtpaddress)
Suspend date: $($item["Resign"])
Account type: $($item["Account"])

Title: $($item["Position"])
Department: $($item["Department"])
Brand: $($item["Brand"])
Office: $($item["Office"])
Manager: $( $manager )

Description: $($item["Description"])

Approved by: $((Get-remADUser (GetUseraccount $item["Modified By"])).name)
"
    if ($items | ? { $_["Employee ID"] -eq $item["Employee ID"] -and $_["Status"] -eq "Active" -and $item["Account"] -eq "Primary" }) {
        $bodyemail += "
Please additionally check another accounts related to this employee:
$(($items|?{$_["Employee ID"] -eq $item["Employee ID"] -and $_["Status"] -eq "Active" }).name)
"
    }

    #Send mail
    mail

    ######MAIL##############################


    #Set status on "Active"
    $item["Status"] = "Disabled"
    $item.ModerationInformation.Status = "Approved"  
    $item.update()
}
#-------------------------[Suspend - Approved]-------------------------






#-------------------------[Unapproved]-------------------------
if ($items_pending_unapproved) {

    $Items = $items_pending_unapproved | select name, @{Name = "Office"; Expression = { $_["Office"] } }, @{Name = "Status"; Expression = { $_["Status"] } }, @{Name = "Request date"; Expression = { $_["Modified"] } }, @{Name = "Start date"; Expression = { $_["Start"] } } , @{Name = "Resign date"; Expression = { $_["Resign"] } }

    if ($Items | ? { $_.Office -eq "Ukraine" }) {
        #$mailto="aleksandrkh@wrdirect.online"
        $subject = "Employee request - Unapproved pending requests - UA"
        $mailto = "ITSystem@company.local", "UA_HD@company.local"
        Send-MailMessage -From notification@company.local -to $mailto `
            -BodyAsHtml ($Items | ? { $_.Office -eq "Ukraine" } | ConvertTo-Html -Head ("There are unapproved pending requests. Please check them and approve:
                 https://portal.office.local/IT/Lists/Employees/Unapproved%20items.aspx
                 ") -Title ("List of devices remove with this session " + $Deleted.Count) | Out-String)  `
            -Subject ($subject) `
            -SmtpServer  "edge.office.local" `
            -ErrorAction Continue
    }


    if ($Items | ? { $_.Office -eq "Israel" }) {
        $mailto = "ITSystem@company.local", "IL_HD@company.local"
        #$mailto="aleksandrkh@wrdirect.online"
        $subject = "Employee request - Unapproved pending requests - IL"
        Send-MailMessage -From notification@company.local -to $mailto `
            -BodyAsHtml ($Items | ? { $_.Office -eq "Israel" } | ConvertTo-Html -Head ("There are unapproved pending requests. Please check them and approve:
                 https://portal.office.local/IT/Lists/Employees/Unapproved%20items.aspx
                 ") -Title ("List of devices remove with this session " + $Deleted.Count) | Out-String)  `
            -Subject ($subject) `
            -SmtpServer  "edge.office.local" `
            -ErrorAction Continue
    }

    if ($Items | ? { $_.Office -eq "Spain" }) {
        $mailto = "ITSystem@company.local", "HelpDesk-Spain@company.local"
        #$mailto="aleksandrkh@wrdirect.online"
        $subject = "Employee request - Unapproved pending requests - SP"
        Send-MailMessage -From notification@company.local -to $mailto `
            -BodyAsHtml ($Items | ? { $_.Office -eq "Spain" } | ConvertTo-Html -Head ("There are unapproved pending requests. Please check them and approve:
                 https://portal.office.local/IT/Lists/Employees/Unapproved%20items.aspx
                 ") -Title ("List of devices remove with this session " + $Deleted.Count) | Out-String)  `
            -Subject ($subject) `
            -SmtpServer  "edge.office.local" `
            -ErrorAction Continue
    }


    if ($Items | ? { $_.Office -ne "Israel" -and $_.Office -ne "Colombia" -and $_.Office -ne "Ukraine" -and $_.Office -ne "Spain" }) {
        $mailto = "ITSystem@company.local", "HD_GLOBAL@company.local"
        #$mailto="aleksandrkh@wrdirect.online"
        $subject = "Employee request - Unapproved pending requests"
        Send-MailMessage -From notification@company.local -to $mailto `
            -BodyAsHtml ($Items | ? { $_.Office -ne "Israel" -and $_.Office -ne "Colombia" -and $_.Office -ne "Ukraine" -and $_.Office -ne "Spain" } | ConvertTo-Html -Head ("There are unapproved pending requests. Please check them and approve:
                 https://portal.office.local/IT/Lists/Employees/Unapproved%20items.aspx
                 ") -Title ("List of devices remove with this session " + $Deleted.Count) | Out-String)  `
            -Subject ($subject) `
            -SmtpServer  "edge.office.local" `
            -ErrorAction Continue
    }


    ######MAIL##############################
}
#-------------------------[Unapproved]-------------------------


#-------------------------[Apply changes]-------------------------
foreach ($item in $items_changeded) {
    Set-remADUser -Identity $item["Login"] -GivenName $item["FirstName"] -Surname $item["LastName"] -Title $item["Position"] -Department $item["Department"] -Company $item["Brand"] -EmployeeID $item["Employee ID"] -AccountExpirationDate $null
    if ($item["Manager"]) { Set-remADUser $item["Login"]   -Manager $( ( get-remaduser (GetUseraccount $item["Manager"])).samaccountname) }
}
#-------------------------[Apply changes]-------------------------


#-------------------------[Update default value for Employee ID]-------------------------
$id = (Retrieve).Fields | ? { $_.title -eq "Employee ID" }
$id.Defaultvalue = (( (Retrieve).Items | % { $_["Employee ID"] } | sort -Descending)[0] + 1)
$id.Update()
#-------------------------[Update default value for Employee ID]-------------------------



#-------------------------[Update information in sharepoint based on AD DS information]-------------------------
if ((get-date).hour -eq 2 -and (get-date).minute -lt 10 -and ( (get-date).DayOfWeek -eq "Tuesday" -or (get-date).DayOfWeek -eq "Thursday"  ) ) {
    $s = New-PSSession -ComputerName $dc
    Invoke-Command -ScriptBlock { Import-Module activedirectory } -Session $s
    Import-PSSession -Session $s -Module activedirectory -Prefix rem

    #Exchange
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch01-nl.office.local/PowerShell/ 
    Import-PSSession $Session



    $items = $items | ? { $_["Status"] -eq "Active" -and $_["Approval Status"] -eq 0 }

    $items | % {
        $object = $_
        $mailbox = $null
        $user = $null
        $user = Get-remADUser $object["Login"] -Properties lastlogondate, ipphone, company, department, title -ErrorAction silentlycontinue
        $user.samaccountname
        if ($user) {

            $mailbox = Get-Mailbox $user.UserPrincipalName -ErrorAction silentlycontinue
            if ($mailbox) {
                $object["Mailbox log in"] = ($mailbox | Get-MailboxStatistics).LastLogonTime

                $object["Primary SMTP"] = ($mailbox).Primarysmtpaddress
                #Mailbox sharing
                if ((Get-MailboxPermission $user.UserPrincipalName | ? { $_.IsInherited -eq $false } | ? { $_.user -like "office\*" } )) { $object["Mailbox sharing"] = (Get-MailboxPermission $user.UserPrincipalName | ? { $_.IsInherited -eq $false } | ? { $_.user -like "office\*" }).user }

                #Forwarding
                if ( ($mailbox).ForwardingAddress ) { $object["Forwarding"] = (Get-Recipient ( ($mailbox).ForwardingAddress) ).PrimarySmtpAddress }
            }

            $object["User log in"] = $user.lastlogondate
            if ($user.Department) { $object["Department"] = $user.Department }
            if ($user.Company) { $object["Brand"] = $user.Company }
            #if($user.title)      {$_["Title"]=$user.title}
            if ($user.description -eq "No such user in Active Directory") { $_["Description"] = $null }
            #$_["Forwarding"]=(Get-Recipient ( (Get-Mailbox $user.samaccountname).ForwardingAddress) ).PrimarySmtpAddress

            if ($user.ipphone) { $object["SIP Extension"] = $user.ipphone }
            $object.ModerationInformation.Status = "Approved"  
            $object.Update()
        }
    

        if (-not $user) {
            #$_["User log in"]=$user.lastlogondate
            $object["Description"] = "No such user in Active Directory"
            $object.ModerationInformation.Status = "Approved"  
            $object.Update()
        }
    }
}
#-------------------------[Update information in sharepoint based on AD DS information]-------------------------




#-------------------------[Update information in sharepoint based on HR system]-------------------------
<#
if ((get-date).hour -eq 2 -and (get-date).minute -lt 10 -and ( (get-date).DayOfWeek -eq "Tuesday" -or (get-date).DayOfWeek -eq "Thursday"  ) ) {
$Connection=[MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString='server=35.190.198.111;uid=it.sharepoint;pwd=Erdl343edNidek34~!2;database=BI;Pooling=false'}
$Connection.Open()
$MYSQLCommand = New-Object MySql.Data.MySqlClient.MySqlCommand
$MYSQLDataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
$MYSQLDataSet = New-Object System.Data.DataSet
$MYSQLCommand.Connection=$Connection
$MYSQLCommand.CommandText='SELECT * FROM BI.`IT_SP.Employees`;'
$MYSQLDataAdapter.SelectCommand=$MYSQLCommand
$NumberOfDataSets=$MYSQLDataAdapter.Fill($MYSQLDataSet, "data")

#  $MYSQLDataSet.tables[0]
$items = $items | ? { $_["Status"] -eq "Active" -and $_["Approval Status"] -eq 0 }
$items[21..30] | % {$object=$_
    ${user-change}=$MYSQLDataSet.tables[0] |?{$_.EMP_EMAIL -eq $object["Primary SMTP"] } 
    ${user-change}
    if (${user-change}.EMP_DEPARTMENT) { $object["Department"] = ${user-change}.EMP_DEPARTMENT}
    if (${user-change}.EMP_DEPARTMENT) { $object["Position"] = ${user-change}.EMP_ROLE}
 #   if (${user-change}.EMP_BRAND) { $object["Brand"] = ${user-change}.EMP_Brand}
    $object.ModerationInformation.Status = "Approved"  
    $object.Update()
}
$Connection.Close()
}
#>
#-------------------------[Update information in sharepoint based on HR system]-------------------------
#-----------------------------------------------------------[Logic]------------------------------------------------------------


#-------------------------[Error notifications]-------------------------
if ($error) {
    $error | % { [string]$bodyerror += $_.Exception }
    function mailerror {
        $from = "notification@company.local"
        $subject = "Script error -  in script <SP-NewEmployee> on $Env:COMPUTERNAME"
        $smtpserver = "edge.office.local"
        $mailto = "user@company.online"
        #---mail
        Send-MailMessage  -to $mailto -from $from -body $bodyerror -subject $subject -smtpserver $smtpserver -DeliveryNotificationOption onfailure
    }

    mailerror
}
#-------------------------[Error notifications]-------------------------

Stop-Transcript
