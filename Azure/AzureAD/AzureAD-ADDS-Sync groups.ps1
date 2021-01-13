# Author  : Aleksandr Khoroshiy
# Date    : 08/01/2020
# Purpose : Script which syncronize members of AD DS secutiry groups to Office 365 groups
# Version : 1.0.0.0
# Details : Script which syncronize members of AD DS secutiry groups to Office 365 groups
# Notes   : 

#-----------------------------------------------------------[Transcript]------------------------------------------------------------
Start-Transcript "C:\Script\logs\Azure\Azure-AD-SyncGroups-$(get-date -Format hh-mm---dd-MM-yyyy).txt"
 

#-----------------------------------------------------------[Settings]------------------------------------------------------------
#Groups list for sync
$Groups = @()
#global groups
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Help Desk"; "GroupAD" = "Team-HelpDesk" })
#UA
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Finance - UA"; "GroupAD" = "Team-UA-Finance" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Administration - UA"; "GroupAD" = "Team-Office Management-UA" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Customer Support - UA"; "GroupAD" = "Team-Customer Support and Compliance" })
#RO
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Marketing - RO"; "GroupAD" = "Team-RO-Marketing" })
#IL
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Retention - HQ"; "GroupAD" = "Team-Retention-hq" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Retention - Haifa"; "GroupAD" = "Team-Retention-haifa" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Sales - Haifa"; "GroupAD" = "Team-Sales-haifa" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Sales - HQ"; "GroupAD" = "Team-Sales-hq" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Retention - RU"; "GroupAD" = "Team-UA-Retention RU" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Retention - EN"; "GroupAD" = "Team-UA-Retention EN" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Human Resources"; "GroupAD" = "Team-UA-HR" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Quality control"; "GroupAD" = "Team-Quality Control" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Marketing - IL"; "GroupAD" = "Team-IL-Marketing" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Human Resources"; "GroupAD" = "Team-IL-HR" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Finance - IL"; "GroupAD" = "Team-IL-Finance" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Team-IL-DealingRoom"; "GroupAD" = "Team-IL-Dealing room" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Backoffice - IL"; "GroupAD" = "Team-IL-Backoffice" })
#SP
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Finance - SP"; "GroupAD" = "Team-SP-Finance" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Human Resources"; "GroupAD" = "Team-SP-HR" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Quality control - SP"; "GroupAD" = "Team-SP-QC" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Retention - SP"; "GroupAD" = "Team-SP-Retention" })
$Groups += New-Object -TypeName PSObject -Property (@{"GroupOffice" = "Sales - SP"; "GroupAD" = "Team-SP-Sales" })
$Parameters = (Get-Content "C:\script\Azure\Azure-Config.json") | ConvertFrom-Json
#-----------------------------------------------------------[Settings]------------------------------------------------------------


#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function Login-Azure { Login-AzAccount -CertificateThumbprint $Parameters.Parameters.Thumbprint -ApplicationId $Parameters.Parameters.ApplicationId -TenantId $Parameters.Parameters.TenantId -ServicePrincipal }


function syncgroup {
    param($GroupAD, $GroupOffice)
    write-host "Processing groups $($GroupAD) and $($GroupOffice)" -ForegroundColor green
    $groupA = (Get-ADGroupMember $GroupAD).samaccountname | % { (Get-ADUser $_).userprincipalname }
    $groupO = (Get-AzADGroupMember -GroupDisplayName $GroupOffice).UserPrincipalName
    compare $groupO $groupA | ? { $_.SideIndicator -eq "=>" } | % { Add-AzADGroupMember -TargetGroupDisplayName $GroupOffice -MemberUserPrincipalName $_.InputObject }
}

#-----------------------------------------------------------[Functions]------------------------------------------------------------


#-----------------------------------------------------------[Logic]------------------------------------------------------------
Login-Azure
$Groups | % { syncgroup -GroupAD $_.Groupad -GroupOffice $_.Groupoffice }
#-----------------------------------------------------------[Logic]------------------------------------------------------------




#-----------------------------------------------------------[Error notifications]------------------------------------------------------------
if ($error) {
    $error | % { [string]$bodyerror += $_.Exception }
    function mailerror {
        $from = "notification@office.local"
        $subject = "Error in script Azure AD Sync on $Env:COMPUTERNAME"
        $smtpserver = "edge.office.local"
        $mailto = "aleksandrkh@wrdirect.online"
        #---mail
        Send-MailMessage  -to $mailto -from $from -body $bodyerror -subject $subject -smtpserver $smtpserver -DeliveryNotificationOption onfailure
    }

    mailerror
}
#-----------------------------------------------------------[Error notifications]------------------------------------------------------------

#-----------------------------------------------------------[Transcript]------------------------------------------------------------
Stop-Transcript

