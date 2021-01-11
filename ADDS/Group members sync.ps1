#-----------------------------------------------------------[Transcript]------------------------------------------------------------
Start-Transcript "C:\Scripts\logs\Azure\Azure-Subscription-Usage Quota-$(get-date -Format hh-mm---dd-MM-yyyy).txt"

#-----------------------------------------------------------[Settings]------------------------------------------------------------
$groups = @(("Team-RO-HR", "SP-Portal-Read"), 
    ("Team-RO-HR", "SP-Portal-read"),
    ("Team-IL-HR", "SP-Portal-Read"),
    ("Team-IL-HR", "sp-portal-IT-employee-edit"),
    ("Team-SP-HR", "SP-Portal-Read"),
    ("Team-SP-HR", "sp-portal-IT-employee-edit"),
    ("Team-UA-HR", "SP-Portal-Read"),
    ("Team-UA-HR", "sp-portal-IT-employee-edit"),
    ("Team-HelpDesk", "SP-Portal-Read"),
    ("Team-HelpDesk", "sp-portal-IT-employee-edit"),
    ("Team-HelpDesk", "sp-portal-hd-edit"),
    ("Team-HelpDesk", "sp-portal-hd-read"),
    ("Team-HelpDesk", "sp-portal-hd-read"),
    ("Team-HelpDesk", "sp-portal-it-read"),
    ("Team-HelpDesk", "SP-Portal-Tier1-read")
)
#-----------------------------------------------------------[Settings]------------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function sync-adgroups {
    [CmdletBinding()]
    param (
        ${group-source} , ${group-destination}
    )
    ${members-source} = get-ADGroupMember ${group-source}
    ${members-destination} = get-ADGroupMember -identity ${group-destination}
    $changes = (Compare-Object   ${members-destination}.SamAccountName ${members-source}.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
    if ($changes) {
        foreach ($change in $changes) {
            Add-ADGroupMember -identity ${group-destination} -members $change -verbose 
            write-host "Synced $($change) to $(${group-destination})"
        }
    }
    else { write-host "No members detected for sync for $( ${group-destination})" }
}
#-----------------------------------------------------------[Functions]------------------------------------------------------------

#-----------------------------------------------------------[Logic]------------------------------------------------------------
$users = Get-ADUser -Properties department, city, title, company -Filter * -SearchBase "OU=Offices,OU=Company,DC=office,DC=local"

#all CCS
$employees = $users | ? { $_.company -eq "DRP" }
$members = get-ADGroupMember -identity "SP-Portal-BackOffice-R"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "SP-Portal-BackOffice-R" -members $_ -verbose } }


#all CCS
$employees = $users | ? { $_.City -eq "Ukraine" }
$members = get-ADGroupMember -identity "SP-Portal-Read"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "SP-Portal-Read" -members $_ -verbose } }


#all computers
$employees = Get-ADComputer -Filter * -SearchBase "OU=Offices,OU=Company,DC=office,DC=local"
$members = get-ADGroupMember -identity "Computers-PC"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "Computers-PC" -members $_ -verbose } }


#all UA computers withour creditcubbe
$employees = Get-ADComputer -Filter * -SearchBase "OU=Ukraine,OU=Offices,OU=Company,DC=office,DC=local"
$employees = $employees | ? { $_.distinguishedname -notlike "*OU=CreditCube,OU=Computers,OU=Ukraine,OU=Offices,OU=Company,DC=office,DC=local" }
$members = get-ADGroupMember -identity "Computers-UA-PC"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "Computers-UA-PC" -members $_ -verbose } }

#all Spain computers
$employees = Get-ADComputer -Filter * -SearchBase "OU=Spain,OU=Offices,OU=Company,DC=office,DC=local"
$members = get-ADGroupMember -identity "Computers-SP-PC"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "Computers-SP-PC" -members $_ -verbose } }

#all Israel computers
$employees = Get-ADComputer -Filter * -SearchBase "OU=Israel,OU=Offices,OU=Company,DC=office,DC=local"
$members = get-ADGroupMember -identity "Computers-IL-PC"
$changes = (Compare-Object   $members.SamAccountName $employees.SamAccountName | ? { $_.SideIndicator -eq "=>" }).InputObject
if ($changes) { $changes | % { Add-ADGroupMember -identity "Computers-IL-PC" -members $_ -verbose } }



$groups | % {
    sync-adgroups -group-source $_[0] -group-destination $_[1]
}
#-----------------------------------------------------------[Logic]------------------------------------------------------------


#----------------Error notifications----------------#
if ($error) {
    $error | % { [string]$bodyerror += $_.Exception }
    function mailerror {
        $from = "notification@company.local"
        $subject = "Error in script about password expiration in AD  $($MyInvocation.ScriptName)"
        $smtpserver = "edge.company.local"
        $mailto = "mailaddress"
        #---mail
        Send-MailMessage  -to $mailto -from $from -body $bodyerror -subject $subject -smtpserver $smtpserver -DeliveryNotificationOption onfailure
    }

    mailerror
}
#----------------Error notifications----------------#


#-----------------------------------------------------------[Transcript]------------------------------------------------------------
Stop-Transcript 
