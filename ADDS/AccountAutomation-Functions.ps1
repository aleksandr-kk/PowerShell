#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Login-Azure { 
    $passwd = ConvertTo-SecureString ${SP-secret} -AsPlainText -Force
    $pscredential = New-Object System.Management.Automation.PSCredential($Config.'Systems configs'.Azure.ApplicationId, $passwd)
    Login-AzAccount  -TenantId $Config.'Systems configs'.Azure.TenantId -ServicePrincipal -Credential $pscredential 
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



Function login-AzureAD { 
    Connect-AzureAD -TenantId $Config.'Systems configs'.Azure.TenantId -ApplicationId $Config.'Systems configs'.Azure.ApplicationId -CertificateThumbprint $config.'Systems configs'.azure.Thumbprint
}




#Get owner of group, to which user belong
function get-manager-groupowner {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]${user-HR}
    )
    (Get-AzureADUserMembership -objectid  (${user-HR}).ObjectId) | ? { $_.displayname -ne "all users" -and $_.displayname -ne "HG-Location-London" -and $_.displayname -ne "SICS Cede Project" }|select -first 1 | % { (Get-AzureADGroupOwner -ObjectId $_.objectid).Mail }
}




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
    $msg.From = $Config.'Systems configs'.'Exception SMTP notification'.from
    $mailto | % { $msg.to.Add($_) }
    $msg.subject = $subject
    $msg.body = $bodyemail
    if ($bodyashtml) { $msg.IsBodyHtml = $true }
    if ($attachments) { $msg.Attachments.add($attachments) }
    $smtp.host = $Config.'Systems configs'.'Exception SMTP notification'.smtpserver
    $smtp.send($msg)
}
#-----------------------------------------------------------[Functions]------------------------------------------------------------
