Start-Transcript "C:\Script\logs\exchange\exchange-cleanlogs-$(get-date -Format dd-MM-yy).txt"
#-----------------------------------------------------------[Settings]------------------------------------------------------------
${Servers-exchange} = "exch01-nl.office.local", "exch02-nl.office.local", "exch03-nl.office.local", "exch04-nl.office.local", "exch01-fr.office.local", "exch02-fr.office.local", "exch03-fr.office.local", "exch04-fr.office.local", "exch05-fr.office.local", "exch05-nl.office.local"
#$servers="exch04-nl.office.local"
${servers-edge} = "edge02-fr.office.local", "edge01-fr.office.local", "edge02-nl.office.local", "edge01-nl.office.local"
${servers-edge-queueclean} = $false

$destionation = "D:\Exchange logs\"

$retention_period = ((get-date).adddays(-5))
$Store_Period = -45
#-----------------------------------------------------------[Settings]------------------------------------------------------------

$folders = '{ "Folders": [
    ["OABGeneratorLog" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\OABGeneratorLog" ],
    ["Search" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\Search" ],
    ["ConversationAggregationLog" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\ConversationAggregationLog" ],
    ["Ews" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\Ews" ],
    ["MapiHttp" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\MapiHttp\\Mailbox" ],
    ["mapi" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\HttpProxy\\Mapi" ],
    ["NotificationBroker" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\NotificationBroker\\Client"],
    ["MailboxAssistantsLog" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\MailboxAssistantsLog" ],
    ["DailyPerformanceLogs" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\Diagnostics\\DailyPerformanceLogs" ],
    ["Pop3" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\Pop3" ],
    ["imap4" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\imap4" ],
    ["BodyFragmentExtractorLog" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\BodyFragmentExtractorLog" ],
    ["ADDriver" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\ADDriver" ],
    ["AuthZ" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\CmdletInfra\\Powershell-Proxy\\AuthZ" ],
    ["Cmdlet" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\CmdletInfra\\Powershell-Proxy\\Cmdlet" ],
    ["Powershell-Proxy-Http" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\CmdletInfra\\Powershell-Proxy\\Http" ],
    ["HttpProxy-PowerShell" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\HttpProxy\\PowerShell"],
    ["W3SVC2" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\RpcHttp\\W3SVC2" ],
    ["W3SVC1" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\RpcHttp\\W3SVC1" ],
    ["BigFunnelRetryFeederTimeBasedAssistant" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\BigFunnelRetryFeederTimeBasedAssistant" ],
    ["RPC Client Access" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\RPC Client Access" ],
    ["ContactChangeLogging" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\ContactChangeLogging" ],
    ["Query" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\Query" ],
    ["lodctr_backups" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\lodctr_backups" ],
    ["eas" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\HttpProxy\\eas" ],
    ["owa" , "C$\\Program Files\\Microsoft\\Exchange Server\\V15\\Logging\\HttpProxy\\owa" ],
    ["W3SVC1" , "C$\\inetpub\\logs\\LogFiles\\W3SVC1" ],
    ["W3SVC2" , "C$\\inetpub\\logs\\LogFiles\\W3SVC2" ],
    ["W3SVC3" , "C$\\inetpub\\logs\\LogFiles\\W3SVC3" ],
    ["W3SVC4" , "C$\\inetpub\\logs\\LogFiles\\W3SVC4" ],
    ["W3SVC5" , "C$\\inetpub\\logs\\LogFiles\\W3SVC5" ],
    ["W3SVC1213907421" , "C$\\inetpub\\logs\\LogFiles\\W3SVC1213907421" ]
    ]

}' | convertfrom-json  

#-----------------------------------------------------------[Logic]------------------------------------------------------------
#-------------------------[Exchange Servers]-------------------------
foreach ($server in ${Servers-exchange}) {
    foreach ($file in $folders.Folders) {

        #Create folder for logs
        if (-not (test-path "\\files.office.local\system\Exchange Logs\" ) ) { new-item -ItemType Directory -name "Exchange logs" -Path "\\files.office.local\system\" }

        #Create folder for specific server 
        if (-not (test-path "\\files.office.local\system\Exchange Logs\$server" ) ) { new-item -ItemType Directory -name $server -Path "\\files.office.local\system\Exchange Logs\" }
        
        #Create folder for specific server 
        if (-not (test-path "\\files.office.local\system\Exchange Logs\$server\$($file[0])" ) ) { new-item -ItemType Directory -name $file[0] -Path "\\files.office.local\system\Exchange Logs\$server\" }

        # Move logs
        get-childitem -path "\\$server\$($file[1])" | ? { $_.lastwritetime -lt $retention_period } | move-item -destination "\\files.office.local\system\Exchange Logs\$server\$($file[0])" -verbose
    }

}
#-------------------------[Exchange Servers]-------------------------


#-------------------------[EDGE Servers]-------------------------
foreach ($server in ${servers-edge}) {

    get-childitem "\\$($server)\c$\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\DailyPerformanceLogs" | ? { $_.LastWriteTime -lt (get-date).adddays(-1) } | Remove-Item -Force -Confirm:$false
    
    
    if (${servers-edge-queueclean} -eq $true) {
        Invoke-command $server  -scriptblock { get-service *transport* | Stop-Service
            remove-item "C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\data\Queue" -recurse -confirm:$false
            get-service *transport* | Start-Service
        }
    }

}
#-------------------------[EDGE Servers]-------------------------
#-----------------------------------------------------------[Logic]------------------------------------------------------------

#-----------------------------------------------------------[Error notifications]------------------------------------------------------------
if ($error) {
    $error | % { [string]$bodyerror += $_.Exception }
    ######MAIL##############################
    function mail {
        $mailto = "aleksandrkh@wrdirect.online"
        $mailsubject = "Exchange - Error in Exchange-cleanlogs script on $Env:COMPUTERNAME"
        $smtp = New-Object System.Net.Mail.SmtpClient
        $to = New-Object System.Net.Mail.MailAddress($mailto)
        $from = New-Object System.Net.Mail.MailAddress("notification@office.local")
        $msg = New-Object System.Net.Mail.MailMessage($from, $to)
        #$msg.To.Add("aleksandr.k@dreamscapenetworks.com")  #add additional recipients
        $msg.subject = $mailsubject
        $msg.body = $bodyerror
        $smtp.host = "edge.office.local"
        $smtp.send($msg)
    }
    ######MAIL##############################

    mail
}
#-----------------------------------------------------------[Error notifications]------------------------------------------------------------
Stop-Transcript
