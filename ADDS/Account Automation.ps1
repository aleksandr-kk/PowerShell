# Author  : Aleksandr Khoroshiy
# Date    : 07/07/2020
# Purpose : Automate account management in AD DS\AzureAD based on information in Bamboo HR system
# 
# 
# Version : 1.0.0.0
# Details : 
# * 
# * 
# Notes   : 

[CmdletBinding()]
param (
    $configpath,
    $apiKey,
    $configfunction
)
  
#-----------------------------------------------------------[Transcript]------------------------------------------------------------
Start-Transcript "C:\Scripts\logs\Account Automation-$(get-date -Format hh-mm---dd-MM-yyyy).txt"
 
#-----------------------------------------------------------[Settings]------------------------------------------------------------
# Primary config
if ($configpath) {
    $configpath
    "pipe"
    # For Pipeline
    $Config = Get-Content $configpath | ConvertFrom-Json

    # Build a BambooHR credential object using the provided API key
    $apiPassword = ConvertTo-SecureString 'x' -AsPlainText -Force

    ${API-Auth} = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $apiKey, $apipassword
    #${API-Auth} = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Config.'Systems configs'.'HR system'.'API-Key', $apipassword
    
    #import functions
    .$configfunction

}
else {
    "local"
    # For Local environemnt
    $Config = Get-Content 'C:\Store\companyname\Repository\Identity Governance\Account Automation - Config.json' | ConvertFrom-Json
    # Build a BambooHR credential object using the provided API key
    $apiPassword = ConvertTo-SecureString 'x' -AsPlainText -Force

    #${API-Auth} = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $apiKey, $apipassword
    ${API-Auth} = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Config.'Systems configs'.'HR system'.'API-Key', $apipassword

    # Import functions
    ."C:\Store\companyname\Repository\Identity Governance\AccountAutomation-Functions.ps1"
}

# Force use of TLS1.2 for compatibility with BambooHR's API server. Powershell on Windows defaults to 1.1, which is unsupported
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Get list of fields which should be retrived from HR system
[array]$fields = $Config.'Systems configs'.'HR system'.'API-fields'.Values


# Photo folder provisioning
# Create paths if they dont exist
if (!(test-path $(Join-Path -Path $config."systems configs"."photo path" -ChildPath "\Lists"))) {
    New-Item -ItemType Directory -Force -Path $(Join-Path -Path $config."systems configs"."photo path"-ChildPath "\Lists")
}
if (!(test-path $(Join-Path -Path $config."systems configs"."photo path" -ChildPath "\Photos"))) {
    New-Item -ItemType Directory -Force -Path $(Join-Path -Path $config."systems configs"."photo path" -ChildPath "\Photos")
}
#-----------------------------------------------------------[Settings]------------------------------------------------------------


#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function get-user-type {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$location)
    switch ($location) {
        { $location -eq "Connected-Organization-HUK" } { $true }
        default { $false }
    }
}


function Get-HRusers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]${Api-URL},
        [Parameter(Mandatory = $true)]
        [string]$Company,
        [Parameter(Mandatory = $true)]
        [PSCredential]$Credentials,
        [Parameter(Mandatory = $true)]
        [string]$Query
    )
    Begin {}
    Process {
        Write-Verbose "Processing:
        API URL: $(${Api-URL} -f $Company )
        Query $($query)
        "
        $data = Invoke-WebRequest (${Api-URL} -f $Company )  -method POST -Credential $credentials -body $query -UseBasicParsing
        $count = (($data.Content | ConvertFrom-Json).employees).count
        $data
        Write-Verbose "Processed  $count objects"
    }
    End {}
}

function New-APIquery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$Fields
    )
    Begin {
    }
    Process {
        #Verbose message
        Write-Verbose "Processing:
        Fileds: $($fields )
        "
        # Create a new blank array to work with
        $fieldsArray = @()

        # For each field provided, create the XML required
        foreach ($field in $fields) {
            $item = '<field id="{0}" />' -f $field
            $fieldsArray += $item
        }

        # Join the array to create a single string
        $fields = $fieldsArray -join ''

        # Join the above array to create a string
        $query = $query -join ''

        # Construct a query string to use for the employee directory report
        $query = @(
            '<report>'
            '<title>Bamboozled Employee Directory</title>'
            $sinceXML
            '<fields>'
            $fields
            '<field id="status" />'
            '</fields>'
            '</report>'
        )
        # Join the above array to create a string
        $query = $query -join ''
        $query

        #Verbose message
        Write-Verbose "Output:
        query: $($query )
        " 
    }
    End {

    }
}

Function get-adOU {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$location)
    switch ($location) {
        { $location -eq "Bermuda" } { "OU=HBM,OU=Users,OU=companynameResources,DC=companybm,DC=com" }
        { $location -eq "Dubai" -or $location -eq "London" -or $location -eq "Dublin" } { "OU=HGM Users London,DC=companyuk,DC=com" }
        { $location -eq "Los Angeles" -or $location -eq "Miami" -or $location -eq "New York" -or $location -eq "Princeton, NJ" -or $location -eq "US Remote" } { "OU=HUS,OU=Users,OU=companynameResources,DC=companybm,DC=com" }
        { $location -eq "Connected-Organization-HUK" } { "OU=External Organizations,OU=Users,OU=companynameResources,DC=companyuk,DC=com" }
        default { "OU=HGM Users London,DC=companyuk,DC=com" }
    }
}

# Get AD DS settings for account based on location
Function get-adsettings {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$location)
    switch ($location) {
        { $location -eq "Bermuda" -or $location -eq "US Remote" -or $location -eq "Los Angeles" -or $location -eq "Miami" -or $location -eq "New York" } { @{DC = $Config.'Systems configs'.'ADDS System'.dc."dc-bm"; UPN = "companybm.com" ; DCsecondary = $Config.'Systems configs'.'ADDS System'.dc."dc-uk" } }
        { $location -eq "Dubai" -or $location -eq "London" -or $location -eq "Dublin" } { @{DC = $Config.'Systems configs'.'ADDS System'.dc."dc-UK"; UPN = "companyuk.com" ; DCsecondary = $Config.'Systems configs'.'ADDS System'.dc."dc-bm" } } 
        default { @{DC = $Config.'Systems configs'.'ADDS System'.dc."dc-UK"; UPN = "companyuk.com" ; DCsecondary = $Config.'Systems configs'.'ADDS System'.dc."dc-bm" } }
    }
}

# Get manager account from AD DS for employee
Function get-manager {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$manager)
    begin {}
    process {
        $data = (($manager).replace(", ", " ")).split(" ") 
        @{
            Login        = ($data[1].remove(1) + $data[0]);
            Emailaddress = (${users-HR} | ? { $_.firstname -eq $data[1] -and $_.lastname -eq $data[0] }).workEmail
        }
    }
    end {}
}

# Get naming policy for account based on employee metadata - is employee contractor or not.
Function get-ADNamingPolicy {
    param (
        [Parameter(Mandatory = $true)]    
        $user)
    switch ($location) {
       
        { $user.division -eq "Consultant/Contractors" } { "cc_$($user.firstname.remove(1) + $(($user.lastname -split " ")[0]))" }
        default { ($user.firstname.remove(1) + ($user.lastname -split " ")[0]) }
    }
}

# Set Office 365 license for account based on employee metadata - is employee contractor or not.
Function set-AzureLicenses {
    param (
        [Parameter(Mandatory = $true)]    
        $user,
        [Parameter(Mandatory = $true)]    
        $login,
        [Parameter(Mandatory = $true)]    
        $server
    )
    switch ($location) {
        { $user.division -eq "Consultant/Contractors" } { Get-ADUser $login -Server $settings.DC  | % { add-adgroupmember "Office365-Licenses-E1"  -members $_ -server $Config.'Systems configs'.'ADDS System'.dc."dc-uk" } }
        default { Get-ADUser $login -server $server  | % { add-adgroupmember "Office365-Licenses-E3"  -members $_ -server $Config.'Systems configs'.'ADDS System'.dc."dc-uk" } }
    }
}

function mail-notification {
    param (
        [Parameter(Mandatory = $true)] 
        $bodyemail,
        [Parameter(Mandatory = $true)] 
        $mailto,
        [Parameter(Mandatory = $true)] 
        $subject
    )
    $smtp = New-Object System.Net.Mail.SmtpClient
    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = $Config.'Systems configs'.'Exception SMTP notification'.from
    $mailto | % { $msg.to.Add($_) }
    $msg.subject = $subject
    $msg.body = $bodyemail
    $smtp.host = $Config.'Systems configs'.'Exception SMTP notification'.smtpserver
    $smtp.send($msg)
}


function mail-notification-exception {
    param (
        [Parameter(Mandatory = $true)] 
        $bodyemail
    )
    $smtp = New-Object System.Net.Mail.SmtpClient
    $to = New-Object System.Net.Mail.MailAddress($Config.'Systems configs'.'Exception SMTP notification'.mailto)
    $from = $Config.'Systems configs'.'Exception SMTP notification'.from
    $msg = New-Object System.Net.Mail.MailMessage($from, $to)
    $msg.subject = $Config.'Systems configs'.'Exception SMTP notification'.subject
    $msg.body = $bodyemail
    $smtp.host = $Config.'Systems configs'.'Exception SMTP notification'.smtpserver
    $smtp.send($msg)
}


Function get-HRusers-misconfiguration {
    param (
        [Parameter(Mandatory = $false)]    
        [array]$exclude,
        [Parameter(Mandatory = $true)]    
        [array]$fields
    )
    begin {}
    process {
        if ($exclude) { $fields = $fields | ? { $exclude -notcontains $_ } }
        $brokenobjects = $fields | % { $field = $_ ; ${users-HR-active} | ? { $_.hireDate -ne "0000-00-00" } | ? { $_.$field -eq $null } } 
        $brokenobjects = ($brokenobjects | select workemail -unique).workemail
        $brokenobjects | % { $field = $_ ; ${users-HR-active}  | ? { $_.workemail -eq $field } }
    }
    end {}
}
#-----------------------------------------------------------[Functions]------------------------------------------------------------





#-----------------------------------------------------------[Logic]------------------------------------------------------------
#----Azure AD----#
login-AzureAD
#----Azure AD----#

#----HR users----#
#Prepare query string
$query = new-APIquery -Fields $Config.'Systems configs'.'HR system'.'API-fields'.Values -Verbose

# Attempt to connect to the BambooHR API Service
try {
    # Perform the API query
    #$bambooHRDirectory = Invoke-WebRequest ($config.'Systems configs'.'HR system'.'API-Url' -f $config.'Systems configs'.'HR system'.company )  -method POST -Credential $bambooHRAuth -body $query -UseBasicParsing
    ${users-HR} = Get-HRusers -Api-URL $config.'Systems configs'.'HR system'.'API-Url' -Company $config.'Systems configs'.'HR system'.'API-Company' -credentials ${API-Auth} -query $query -Verbose
    # Convert the output to a PowerShell object and get array of users from HR system
    ${users-HR} = ${users-HR}.Content | ConvertFrom-Json
    ${users-HR} = ${users-HR}.employees
}
catch {
    throw "Directory download failed."
}

${users-HR} = ${users-HR} | ? { $_.hiredate -ne "0000-00-00" }
${users-HR-active} = ${users-HR} | ? { $_.status -eq "Active" }
${users-HR-disabled} = ${users-HR} | ? { $_.status -eq "inactive" -and $_.workemail }
#----HR users----#



#----Get AD users----#
${users-AD-BM} = $clear
${users-AD-UK} = $clear
$Config.'Systems configs'.'ADDS System'.'Base DN'.'dc-BM' | % { [array]${users-AD-BM} += Get-aduser -filter *  -properties emailaddress, title, department, office, l, manager -server $Config.'Systems configs'.'ADDS System'.dc."dc-bm" -searchbase $_ } #| ? { $_.enabled -eq $true } }
$Config.'Systems configs'.'ADDS System'.'Base DN'."dc-uk" | % { [array]${users-AD-UK} += Get-aduser -filter *  -properties emailaddress, title, department, office, l, manager -server $Config.'Systems configs'.'ADDS System'.dc."dc-uk" -searchbase  $_ } #| ? { $_.enabled -eq $true }
[array]${Users-AD} = ${users-AD-BM} + ${users-AD-UK}

#Get candidates for action - Create
${users-delta-create} = (compare (${users-HR-active} | ? { $_.workemail }).workemail (${Users-AD} | ? { $_.emailaddress }).emailaddress | ? { $_.sideindicator -eq "<=" }).InputObject
${users-delta-create} = ${users-delta-create} | % { $user = $_; ${users-HR-active} | ? { $_.workemail -eq $user } }

#Get candidates for action - disable
${users-delta-disable} = (compare (${users-HR-disabled}).workemail (${Users-AD} | ? { $_.emailaddress -and $_.enabled -eq $true }).emailaddress -IncludeEqual | ? { $_.sideindicator -eq "==" }).InputObject
${users-delta-disable} = ${users-delta-disable} | % { $user = $_; ${users-HR-disabled} | ? { $_.workemail -eq $user } }
#----Get AD users----#




#-------------------------[Exclude exceptions]-------------------------
${users-delta-create} = ${users-delta-create} | ? { $config.exceptions -notcontains $_.workemail -and $config."Exceptions-create-external" -notcontains $_.workemail }

#${users-delta-create} = ${users-delta-create} | ? { ([datetime]::ParseExact($_.hiredate, "yyyy-MM-dd", $null)) -lt (get-date) }
${users-delta-disable} = ${users-delta-disable}  | ? { $config.exceptions -notcontains $_.workemail }

# Filter accounts which are candidates for changes 
${users-HR-active-change} = ${users-HR-active} | ? { ([datetime]::ParseExact($_.hiredate, "yyyy-MM-dd", $null)) -lt (get-date) } | ? { $config.'Exceptions-change' -notcontains $_.workemail }


if (${users-delta-create} ) { write-host "Accounts for provisioning - $(${users-delta-create}.workemail)" }
if (${users-delta-disable} ) { write-host "Accounts for disable - $(${users-delta-disable}.workemail)" }
#-------------------------[Exclude exceptions]-------------------------


${users-delta-create} | % { if (get-user-type -location $_.location ) { [array]${users-delta-create-external} += $_ }  else { [array]${users-delta-create-internal} += $_ } }



#-------------------------[Automation part - Create]-------------------------
if ($config."Script logic blocks status"."Automation - Create" -eq "Enabled") {
    foreach ($user in ${users-delta-create-internal}) {

        write-host "Processing accounts - internal -  $($user.workemail) "
        # Generate password
        $password_string = "!Password" + (Get-Random -Minimum 1000 -Maximum 9999)
        $Password = ConvertTo-SecureString ($password_string) -AsPlainText -Force 
        $login = get-ADNamingPolicy -user $user
        if ($user.supervisor) { ${user-manager} = (get-manager $user.supervisor).Login }
        #${user-external} = get-user-type -location $user.location
        # ${user-emailaddress} = if (${user-external} ) { "$(($user.workEmail -split "@")[0])@company.com" } else { $user.workEmail }
        $settings = get-adsettings -location $user.location


        # Validate account existence
        if (( ${Users-AD} | ? { $_.samaccountname -eq $login }) ) {
            #----Email notification----#
            [array]$mailto = "user1@company.com" , "user2@company.com" , "user3@company.com" , "user4@company.com"
            $subject = "Account automation - Provisioning - Unsuccessful"
            $bodyemail = "
New employee account hasn't been provisoined. Following account have one of the possible problems:
  * Doesn't have proper value of email address in EmailAddress property. Please add value to EmailAddress even if employee doesn't have mailbox.

Account details:
Employee: $($user.firstName) $($user.lastName)
Login: $($login)
Login as UPN: $($user.firstName.remove(1) + $user.lastname + "@" + $settings.UPN)
Email address: $($user.workEmail)
"
            mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
            Department group membership

            continue
        }

        # Provisioning of new account
        New-ADUser -Name $user.displayName -GivenName $user.firstName -Surname $user.lastName -SamAccountName  $login  -Company "Find out value" -PasswordNeverExpires $false -UserPrincipalName ($login + "@" + $settings.UPN) `
            -Server $settings.dc -AccountPassword $Password  -DisplayName ($user.firstName + $user.lastName) -Office $user.location -Department $user.department -Title $user.jobTitle -Enabled $true -Path (get-adOU -location $user.location) -emailaddress $user.workemail -EmployeeID $user.id

        # Set manager
        if ($user.supervisor -and (  Get-aduser -filter { samaccountname -eq ${user-manager} } -Server $settings.DC  )) { set-aduser  $login -manager (get-manager $user.supervisor).login -Server $settings.DC }

        #----Department group membership----#
        # Department group membership
        # Local AD DS group
        <#
        try {
            add-adgroupmember "Department-$($user.department)-$($user.location)" -members $login -server $settings.dc
        }
        catch {
            write-host "There is error during adding to Department-$($user.department)-$($user.location) on AD Forest $($settings.dc):
            $_.Exception.Message"   
        }
        #>
        # Set LOB attribute
        if ($user.customLOB) { set-aduser $login  -add @{extensionAttribute4 = $user.customLOB } -Server $settings.dc }       

        # Azure AD group
        try {
            ${AAD-group-id} = Get-AzureADGroup -SearchString "Department-$($user.department)-$($user.location)" | ? { $_.DirSyncEnabled -ne "True" }
            if (get-azureaduser -searchstring ($login + "@" + $settings.UPN)) {
                Add-AzureADGroupMember -ObjectId  ${AAD-group-id}.objectid -RefObjectId (get-azureaduser -searchstring ($login + "@" + $settings.UPN) ).objectid
            }
        }
        catch {
            write-host "There is error during adding to Department-$($user.department)-$($user.location) on AD Forest $($settings.dc):
            $_.Exception.Message"   
        }
        #----Department group membership----#

        # Assign licenses
        set-AzureLicenses -user $user -login $login -server $settings.dc 

        # Assign tag for Fulltime or contractor employee
        if ($user.employmentHistoryStatus -eq "Consultant/Contractor" -or $user.employmentHistoryStatus -eq "Non-Executive Director") { set-aduser $login  -add @{extensionAttribute1 = "Contractor" } -Server $settings.dc } else { set-aduser $login  -add @{extensionAttribute1 = "Fulltime" } -server $settings.dc }
       
        #----Email notification----#
        [array]$mailto = "user1@company.com" , "user2@company.com" , "user3@company.com" , "user4@company.com"
        if ((get-manager -manager $user.supervisor).Emailaddress ) { $mailto += (get-manager -manager $user.supervisor).Emailaddress }
        $subject = "Account automation - Created successfully"
        $bodyemail = "
New employee account has been created:

Employee: $($user.firstName) $($user.lastName)
Login: $($login)
Login as UPN: $($user.firstName.remove(1) + $user.lastname + "@" + $settings.UPN)
Email address: $($user.workEmail)
Password: $password_string

Title: $($user.jobtitle)
Department: $($user.department)
Office: $($user.location)
Manager: $($user.supervisor)
Start date: $($user.hiredate)
"
        mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
        #----Email notification----#

    }


    foreach ($user in ${users-delta-create-external}) {
        write-host "Processing accounts - external -  $($user.workemail)"
        # Generate password
        $password_string = "!Password" + (Get-Random -Minimum 1000 -Maximum 9999)
        $Password = ConvertTo-SecureString ($password_string) -AsPlainText -Force 
        $login = get-ADNamingPolicy -user $user
        #${user-external} = get-user-type -location $user.location
        # ${user-emailaddress} = if (${user-external} ) { "$(($user.workEmail -split "@")[0])@company.com" } else { $user.workEmail }
        $settings = get-adsettings -location $user.location

        # Validate account existence
        if (( ${Users-AD} | ? { $_.samaccountname -eq $login }) ) {
            #----Email notification----#
            [array]$mailto = "user1@company.com" , "user2@company.com" , "user3@company.com" , "user4@company.com"
            $subject = "Account automation - Provisioning - Unsuccessful"
            $bodyemail = "
New employee account hasn't been provisoined. Following account have one of the possible problems:
  * Doesn't have proper value of email address in EmailAddress property. Please add value to EmailAddress even if employee doesn't have mailbox.

Account details:
Employee: $($user.firstName) $($user.lastName)
Login: $($login)
Login as UPN: $($user.firstName.remove(1) + $user.lastname + "@" + $settings.UPN)
Email address: $($user.workEmail)
"
            mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
            #----Email notification----#

            continue
        }

        New-ADUser -Name $user.displayName -GivenName $user.firstName -Surname $user.lastName -SamAccountName  $login  -Company "Find out value" -PasswordNeverExpires $false -UserPrincipalName ($login + "@" + $settings.UPN) `
            -Server $settings.dc -AccountPassword $Password  -DisplayName ($user.firstName + $user.lastName) -Office $user.location -Department $user.department -Title $user.jobTitle -Enabled $true -Path (get-adOU -location $user.location) -emailaddress $user.workemail -EmployeeID $user.id

        # Set manager
        if ($user.supervisor) {
            ${user-manager} = (get-manager $user.supervisor).Login
            if ($user.supervisor -and (  Get-aduser -filter { samaccountname -eq ${user-manager} } -Server $settings.DC  )) { set-aduser  $login -manager (get-manager $user.supervisor).login -Server $settings.DC }
        }

        # Assign tag for Fulltime\Contractor\External employee
        set-aduser $login  -replace @{extensionAttribute1 = "External" } -server $settings.dc 
        
        # Set original email address
        set-aduser $login  -replace @{extensionAttribute2 = $user.homeEmail } -server $settings.dc 


        #----Email notification----#
        [array]$mailto = "user1@company.com" , "user2@company.com" , "user3@company.com", "user4@company.com" , $user.homeEmail
        if ((get-manager -manager $user.supervisor).Emailaddress ) { $mailto += (get-manager -manager $user.supervisor).Emailaddress }
        $subject = "Account automation - Created successfully"
        $bodyemail = "
New account for external employee has been created:

Employee: $($user.firstName) $($user.lastName)
Login: $($login)
Login as UPN: $($user.firstName.remove(1) + $user.lastname + "@" + $settings.UPN)
Email address: $($user.homeEmail)
Password: $password_string

Title: $($user.jobtitle)
Department: $($user.department)
Office: $($user.location)
Manager: $($user.supervisor)
"
        mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
        #----Email notification----#

    }

}
#-------------------------[Automation part - Create]-------------------------

#-------------------------[Automation part - Disable]-------------------------
if ($config."Script logic blocks status"."Automation - Disable" -eq "Enabled") {
    foreach ($user in ${users-delta-disable}) {
        

        if (${Users-AD} | ? { $_.emailaddress -eq $user.workemail }) {
            #Get settings
            $settings = get-adsettings -location $user.location
            $login = (${Users-AD} | ? { $_.emailaddress -eq $user.workemail }).samaccountname
           
            # Write inforamtion about detected settings
            write-host "User detected for disabling:
            $user
            Login: $login
            AD DC server: $($settings.dc)"


            try { Set-ADUser -Identity $login  -Enabled $false  -Verbose -Server $settings.dc }
            catch {
                Write-host "There is an issue with disabling following account $login on $($settings.dc). Trying disable account on secondary AD Forest"
                try { Set-ADUser -Identity $login  -Enabled $false  -Verbose -Server $settings.DCsecondary }
                catch { Write-host "There is an issue with disabling following account $login on $($settings.DCsecondary)." }
            }

            # Delete User from Groups
            $groups = (Get-ADUser $login -Properties *).memberof #|?{ $_ -notin (Get-remADUser $item["Login"] -Properties *).PrimaryGroup }))
            $groups | % { Remove-ADGroupMember $_ -Members $login -Confirm:$false }
            $groups = $groups | % { (Get-ADGroup $_).samaccountname }

           
            
            #----Email notification----#
            $mailto = "user1@company.com" , "user2@company.com", "user3@company.com" , "user4@company.com"
            if ($user.supervisor) { if (   (get-manager -manager $user.supervisor).Emailaddress   ) { $mailto += (get-manager -manager $user.supervisor).Emailaddress } }
            $subject = "Account automation - Disabled successfully"
            $bodyemail = "
                    Following employee's account has been disabled:

Employee: $($user.firstName) $($user.lastName)
Login: $($login)
Login as UPN: $($user.firstName.remove(1) + $user.lastname + "@" + $settings.UPN)

Email address: $($user.workemail)
Title: $($user.jobtitle)
Department: $($user.department)
Office: $($user.location)
Manager: $($user.supervisor)
"
            #Start date: not avaliable
            mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
            #----Email notification----#

        }
        #else { "No such user" }
    }
}
#-------------------------[Automation part - Disable]-------------------------

#-------------------------[Data validation]-------------------------
if ($config."Script logic blocks status"."Automation - Data Validation" -eq "Enabled") {
    ${users-HR-validate} = get-HRusers-misconfiguration -fields $fields -exclude "workphone", "supervisor", "hireDate", "terminationDate", "preferredName", "mobilePhone", "homePhone", "employmentHistoryStatus", "homeemail"  | ? { $_.hiredate -ne "0000-00-00" } | ? { ([datetime]::ParseExact($_.hiredate, "yyyy-MM-dd", $null)) -lt (get-date) } 
    if (${users-HR-validate}) {
        #----Email notification----#
        $mailto = "user1@company.com" , "user2@company.com", "georgina.harrison@company.com" , "user4@company.com"
        $subject = "Account automation - Data to validate"
        $bodyemail = "
        The records below are missing critical fields in BambooHR, Please address at your earliest convenience:
$(${users-HR-validate}|fl workemail,location,department,division,displayname,lastname,firstname|out-string)
"
        #Start date: not avaliable
        Write-host "Sending email notification about broken accounts in HR system"
        mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
        #----Email notification----#
    }



    ### Get list of users who exist and active in HR system, but don't exist in Azure AD
    foreach ($user in ${users-HR-active-change}) {
        $login = get-ADNamingPolicy -user $user
        $settings = get-adsettings -location $user.location

        if (-not (Get-AzureADGroup -SearchString "Department-$($user.department)-$($user.location)")) {
            write-host "Following active account in HR system - doesn't exist in Azure AD - $($login)@$($settings.UPN)"
            [array]${$groups-AAD-missed} += "Department-$($user.department)-$($user.location)"
        }

        if (-not (get-azureaduser -searchstring ($login + "@" + $settings.UPN) ).objectid ) {
            write-host "Following active account in HR system - doesn't exist in Azure AD - $($login)@$($settings.UPN)"
            [array]${$users-AAD-missed} += ($login + "@" + $settings.UPN)
        }
    }

    $mailto = "user1@company.com", "user4@company.com", "user2@company.com"
    $subject = "Account automation - Azure AD Users - validation error"
    $bodyemail = "
Following Azure AD user is active in HR system and isn't active(or doesn't exist) in Azure AD:
$((${$users-AAD-missed})|fl|out-string)

Following Azure AD user is active in HR system and isn't active(or doesn't exist) in Azure AD:
$((${$groups-AAD-missed}|group).name|fl|out-string)
"
    Write-host "Sending email notification about broken Azure AD secutiry groups for departments"
    mail-notification -bodyemail $bodyemail -mailto $mailto -subject $subject
}
#-------------------------[Data validation]-------------------------


#-------------------------[Automation part - Change]-------------------------
foreach ($user in ${users-HR-active-change}) {
    
    
    ${user-AD} = ${Users-AD} | ? { $_.EmailAddress -eq $user.workEmail } 

    #Avoid cases when accounts duplicated in both AD forests
    if ((${user-AD}).count -ge 2) { continue ; }

    if (${user-AD} -and $user.location) {
        #----Settings----#
        $login = get-ADNamingPolicy -user $user
        $settings = get-adsettings -location $user.location
        #----Settings----#
        
        ######################################
        #   UPDATE ACCOUNT DETAILS - START   #
        ######################################


        # POPULATE PARAMETERS - for Set-ADUser command
        [hashtable]$setUserParams = @{}
        $setUserParams.Add('EmployeeID', $user.id)
        $setUserParams.Add('Surname', $user.lastName)
        $setUserParams.Add('Division', $user.division)
        $setUserParams.Add('OfficePhone', $user.workPhone)
        $setUserParams.Add('MobilePhone', $user.mobilePhone)
        $setUserParams.Add('HomePhone', $user.homePhone)
        $setUserParams.Add('Department', $user.department)
        $setUserParams.add('office', $user.location)
     
    

        # POPULATE PARAMETER - DisplayName, If preferred name is completed use that instead of first name for DisplayName.
        if (-not ([string]::IsNullOrWhiteSpace($($user.preferredName)))) {
            $setUserParams.Add('DisplayName', "$($user.preferredName) $($user.lastName)")
            $setUserParams.Add('GivenName', $user.preferredName)
            write-host "Preferred name was used"
        }
        else {
            $setUserParams.Add('DisplayName', "$($user.firstName) $($user.lastName)")
            $setUserParams.Add('GivenName', $user.FirstName)
        }

        # POPULATE PARAMETER - Title. If the  Title exceeds 128 characters dont update it, but add it to a list of exceptions.
        if ($($user.jobTitle.Length) -gt 128) {

            # Add to user to bamBooHrUserTitleExceeds128CharsList
            $bamBooHrUserTitleExceeds128CharsList.add([PSCustomObject]@{
                    FirstName = $user.firstName
                    LastName  = $user.lastName
                    WorkEmail = $user.workEmail
                    Title     = $user.jobTitle
                })
        }
        else {
            $setUserParams.Add('Title', $user.jobTitle)
        }

        # Get manager
        if ($user.supervisor) {
            $manager = (get-manager $user.supervisor).Login
            # Set manager
            if ($user.supervisor -and (Get-aduser -filter { samaccountname -eq $manager } -Server $settings.DC  )) {
                "Settings for $(${user-AD}.samaccountname) manager $((get-manager $user.supervisor).login) "
                set-aduser  ${user-AD}.samaccountname -manager (get-manager $user.supervisor).login -Server $settings.DC 
            }
        }

        # Update User AD attributes
        try {
           
            ${user-AD} | ft samaccountname, office, emailaddress
            $setUserParams
            try {
                Write-host "Updating AD attributes for user : $(${user-AD}.samaccountname)" -foregroundcolor green
                Set-ADUser -identity ${user-AD}.samaccountname  @setUserParams  -Server $settings.dc  
            }
            catch { 
                write-host "Updating AD attributes for user : $(${user-AD}.samaccountname) on secondary DC" -foregroundcolor yellow
                Set-ADUser -identity ${user-AD}.samaccountname  @setUserParams  -Server $settings.DCsecondary 
            }
        }
        catch {
            Write-host "ERROR --- Unable to update AD attributes for user : $(${user-AD}.samaccountname). The error message was $($_.Exception.Message)"
        }
        
        # Assign tag for Fulltime or contractor employee
        if ($user.employmentHistoryStatus -eq "Consultant/Contractor") { 
            Write-host "Set value Contractor for $(${user-AD}.samaccountname)"
            try { set-aduser ${user-AD}.samaccountname  -replace @{extensionAttribute1 = "Contractor" } -Server $settings.dc } 
            catch {
                Write-host "Set value Contractor for $(${user-AD}.samaccountname) with secondary DC"
                set-aduser ${user-AD}.samaccountname  -replace @{extensionAttribute1 = "Contractor" } -Server $settings.DCsecondary 
            }
        }
        else { 
            Write-host "Set value Fulltime for $(${user-AD}.samaccountname)"
            try { set-aduser ${user-AD}.samaccountname  -replace @{extensionAttribute1 = "Fulltime" } -Server $settings.dc } 
            catch { set-aduser ${user-AD}.samaccountname  -replace @{extensionAttribute1 = "Fulltime" } -Server $settings.DCsecondary }
        }

        # Set LOB attribute
        if ($user.customLOB) { set-aduser ${user-AD}.samaccountname  -add @{extensionAttribute4 = $user.customLOB } -Server $settings.dc }


        #----Department group membership----#
        # Department group membership
        # Local AD DS group
        <#
        try {
            add-adgroupmember "Department-$($user.department)-$($user.location)" -members $login -server $settings.dc
        }
        catch {
            write-host "There is error during adding to Department-$($user.department)-$($user.location) on AD Forest $($settings.dc):
            $_.Exception.Message"   
        }
        #>

        # Azure AD group
        try {
            ${AAD-group-id} = Get-AzureADGroup -SearchString "Department-$($user.department)-$($user.location)" | ? { $_.DirSyncEnabled -ne "True" }
            if (-not (get-AzureADGroupMember -ObjectId  ${AAD-group-id}.objectid | ? { $_.userprincipalname -eq ($login + "@" + $settings.UPN) })) {
                Write-host "Processing adding user $($user.displayName) to Azure AD group $("Department-$($user.department)-$($user.location)")"
                Add-AzureADGroupMember -ObjectId  ${AAD-group-id}.objectid -RefObjectId (get-azureaduser -searchstring ($login + "@" + $settings.UPN) ).objectid
            }
        }
        catch {
            write-host "There is error during adding to Department-$($user.department)-$($user.location) on AD Forest $($settings.dc):
            $_.Exception.Message"   
        }
        #----Department group membership----#
        


        #----Automation part - Photo----#
        #####################################
        #   UPDATE PHOTO FOR USER - START   #
        #####################################

        # Update Photo in AD if bambooUser.photoUploaded set to true
        if ($($user.photoUploaded) -eq $true) {

            $photoPath = $(Join-Path -Path $config."systems configs"."photo path" -ChildPath "\Photos\$($user.id).jpg")
            
            try {
                write-host  "------ AD Photo - Download photo from BambooHR for user : $(${user-AD}.samaccountname)"
                Invoke-WebRequest $($user.photoUrl) -OutFile $photoPath
                try {
                    write-host "------ AD Photo - Change AD Photo for user : $(${user-AD}.samaccountname)"
                    $adPhoto = [byte[]](Get-Content $photoPath -Encoding byte)
                    #applying of photo from HR system to bamboo system
                    #${user-AD}| Set-ADUser -Replace @{thumbnailphoto = ($adPhoto) } -Server $settings.dc -verbose
                }
                catch {
                    write-host "ERROR --- Unable to Change AD Photo for user  : $(${user-AD}.samaccountname). The error message was $($_.Exception.Message)"
                }
            }
            catch {
                write-host "ERROR --- Unable to Download photo from BambooHR for user  : $(${user-AD}.samaccountname). The error message was $($_.Exception.Message)"
            }
        }
    }
    else {
        write-host "User with email address $($user.workEmail) is not located at any of the following $(${user-AD}.samaccountname)"
    }
    ###################################
    #   UPDATE PHOTO FOR USER - END   #
    ###################################
    #----Automation part - Photo----#

    write-host "#####################################"
        
}
#-------------------------[Automation part - Change]-------------------------
#-----------------------------------------------------------[Logic]------------------------------------------------------------





#-----------------------------------------------------------[Error notifications]------------------------------------------------------------
if ($Config.'Systems configs'.'Exception SMTP notification'.Enabled -eq "true" -and $error) {
    $Exceptions_SMTP = ($Config.'Systems configs'.'Exception SMTP notification' ) 
    #Import and execute function for exception notification by SMTP
    $error | % { [string]$bodyerror += $_.Exception }
    mail-notification-exception -bodyemail $bodyerror   
}
#-----------------------------------------------------------[Error notifications]------------------------------------------------------------
#-----------------------------------------------------------[Transcript]------------------------------------------------------------

Stop-Transcript 

