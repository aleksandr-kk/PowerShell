# Author  : Aleksandr Khoroshiy
# Date    : 12/02/2020
# Purpose : Start and stop Az VMs based on schedule specified in Azure tags
# 
# 
# Version : 1.0.0.0
# Details : 
# * 
# * 
# Notes   : 

#----------------Transcript----------------#
Start-Transcript "C:\Script\logs\Azure\Azure-VM-PowerPlans-$(get-date -Format hh-mm---dd-MM-yyyy).txt"
#----------------Transcript----------------#
get-date
#----------------Settings----------------#
${vms-all} = $clear
#Exception
${exceptions-subscriptions} = "Prod UK (Production)"


#----------------Settings----------------#
#----------------Functions----------------#
function mail-notification {
    param (
        [Parameter(Mandatory = $true)] 
        $bodyemail,
        [Parameter(Mandatory = $true)] 
        $mailto,
        [Parameter(Mandatory = $true)] 
        $subject,
        $attachments,
        [switch]$bodyashtml
    )
    $smtp = New-Object System.Net.Mail.SmtpClient
    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = "notification@maildomain.com"
    $mailto | % { $msg.to.Add($_) }
    $msg.subject = $subject
    $msg.body = $bodyemail
    if ($bodyashtml) { $msg.IsBodyHtml = $true }
    if ($attachments) { $msg.Attachments.add($attachments) }
    $smtp.host = "mail.maildomain.com"
    $smtp.send($msg)
}
#----------------Functions----------------#




#----------------Logic----------------#
#Scope of subscription context
$subscriptions = Get-AzSubscription

#Exclude Subsricriptions as exceptions based on config file
$subscriptions = $subscriptions | ? { ${exceptions-subscriptions} -notcontains $_.Name }

# Get all resources in all subscriptions. For all resources filters resources with tag Environment:Production. 
$subscriptions | % {
    $subscription = $_
    Set-AzContext -Subscription $subscription.name | Out-Null
    [array]${vms-all} += Get-AzVM -Status
}

# Filter Az VMs 
${vms-candidates} = ${vms-all} | ? { $_.tags."Power management policy" -and ( ($_.tags."Power management policy" -split ";")[0] -eq "enabled" -or ($_.tags."Power management policy" -split ";")[0] -eq "shutdown"  ) }
#$vms = $vms | ? { $_.name -like "hg-ne-tst-oleksadr*" -or $_.name -eq "dev-gup-w10-ne3" }


# Apply power policy for Az VMs
foreach ($vm in ${vms-candidates} ) {
    ${power-status} = ($vm.tags."Power management policy").split(";")[0]
    $poweroff = ($vm.tags."Power management policy").split(";")[1] -split " "
    $poweron = ($vm.tags."Power management policy").split(";")[2] -split " "
    ${power-timezone} = ($vm.tags."Power management policy").split(";")[3]
    ${power-notifications} = ($vm.tags."Power management policy").split(";")[4] -split ","
    
    # Power off
    if (($vm.tags."Power management policy").split(";")[1] -and ${power-status} -eq "enabled") {

        if ($poweroff[0] -eq "daily" -or ($poweroff[0] -eq "working" -and (get-date).DayOfWeek -in "monday", "tuesday", "wednesday", "thursday", "friday")  ) {
            if ( (get-date $poweroff[1]) -lt ((Get-Date).ToUniversalTime()).addhours(${power-timezone} ) -and (get-date $poweron[1]) -lt ((Get-Date).ToUniversalTime()).addhours(${power-timezone} ) -and $vm.PowerState -eq "VM running" ) { 
                stop-azvm -id $vm.Id -Confirm:$false -Force
                write-host "VM $($vm.name) has been shutdown according to power policy" -foreground green
                mail-notification -bodyemail "Following VM has been shutdown - $($vm.name)" -mailto  ${power-notifications} -subject "Azure - VM - Power Policy" 
            }
            else { write-host "VM $($vm.name) hasn't been shutdown since shutdown time hasn't come" -foreground yellow }
        }
    }

    # Power on
    if (($vm.tags."Power management policy").split(";")[2] -and ${power-status} -eq "enabled") {
        if ($poweroff[0] -eq "daily" -or ($poweroff[0] -eq "working" -and (get-date).DayOfWeek -in "monday", "tuesday", "wednesday", "thursday", "friday")  ) {
            if ( (get-date $poweron[1]) -lt ((Get-Date).ToUniversalTime()).addhours(${power-timezone} ) -and (get-date $poweroff[1]) -gt ((Get-Date).ToUniversalTime()).addhours(${power-timezone} ) -and $vm.PowerState -ne "VM running" ) { 
                start-azvm -id $vm.Id  
                write-host  "VM $($vm.name) has been started according to power policy"  -foreground green
                mail-notification -bodyemail "Following VM has been started - $($vm.name)" -mailto  ${power-notifications} -subject "Azure - VM - Power Policy" 
            }
            else { write-host  "VM $($vm.name) hasn't been started since start time hasn't come" -foreground yellow }
        }
    }

    # Shutdown
    if ( ${power-status} -eq "shutdown" -and $vm.PowerState -eq "VM running") {
        stop-azvm -id $vm.Id -Confirm:$false -Force
        write-host "VM $($vm.name) has been shutdown according to power policy status shutdown" -foreground green
        mail-notification -bodyemail "Following VM has been shutdown and set to shutdown state - $($vm.name)" -mailto  ${power-notifications} -subject "Azure - VM - Power Policy" 
    }

    # Notification about upcoming shutdown of Az VMs
    if (($vm.tags."Power management policy").split(";")[1] -and ${power-status} -eq "enabled" -and (get-date $poweroff[1]).addhours(-1) -lt (get-date) -and (get-date) -lt (get-date $poweroff[1])   ) {
        if ($poweroff[0] -eq "daily" -or ($poweroff[0] -eq "working" -and (get-date).DayOfWeek -in "monday", "tuesday", "wednesday", "thursday", "friday")  ) {
            write-host "VM $($vm.name) will be  shutdown according to power policy less than in 1 hour" -foreground green
            mail-notification -bodyemail "Following VM will be shutdown according to power policy less than in 1 hour - $($vm.name)" -mailto  ${power-notifications} -subject "Azure - VM - Power Policy"   
        }
    }
}
#----------------Logic----------------#

#----------------Transcript----------------#
Stop-Transcript
#----------------Transcript----------------#
