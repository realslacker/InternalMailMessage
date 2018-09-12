# -----------------------------------------------------------------------------
# INIT
# -----------------------------------------------------------------------------


if ( ! $Script:Configuration ) {

    $Script:Configuration = Import-Configuration

}




# -----------------------------------------------------------------------------
# FUNCTIONS
# -----------------------------------------------------------------------------


<#
.SYNOPSIS
    Send and HTML formatted email.

.DESCRIPTION
    Send and HTML formatted email using either the SMTP server defined in the configuration, or your Outlook client.

.PARAMETER To
    Recipients of the message.

.PARAMETER CC
    Recipients who should be copied on the message.

.PARAMETER From
    Who the message is sent from.

.PARAMETER Subject
    Message subject line.

.PARAMETER Body
    The HTML formatted body.
#>
function Send-InternalMailMessage {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory)]
        [string[]]
        $To,

        [string[]]
        $CC,

        [string]
        $From = $Script:Configuration.SmtpFrom,
        
        [Parameter(Mandatory)]
        [string]
        $Subject,

        [Parameter(Mandatory)]
        [string]
        $Body

    )

    Write-Verbose "Inside Send-InternalMailMessage function"

    $MailMessage = @{
        SmtpServer = $Script:Configuration.SmtpServer
        Port       = $Script:Configuration.SmtpPort
        UseSSL     = $Script:Configuration.SmtpUseSSL
        BodyAsHtml = $true
        Subject    = $Subject
        From       = $From
        Body       = $Body
    }

    if ( $To -is [string] ) {

        $MailMessage['To'] = $To -split '[,;]'
    }

    if ( $To -is [array] ) {

        $MailMessage['To'] = $To
    }

    if ( $CC -is [string] ) {

        $MailMessage['CC'] = $CC -split '[,;]'
    }

    if ( $CC -is [array] ) {

        $MailMessage['CC'] = $CC
    }

    if ( $Script:Configuration.SmtpAuthRequired ) {

        $SmtpUser  = $Script:Configuration.SmtpUser
        $SmtpPass  = $Script:Configuration.SmtpPass | ConvertTo-SecureString -AsPlainText -Force
        
        $MailMessage['Credential'] = New-Object System.Management.Automation.PSCredential($SmtpUser, $SmtpPass)

    }
    
    if ( Test-NetTcpPortOpen -Destination $MailMessage['SmtpServer'] -Port $MailMessage['Port'] ) {

        Write-Verbose "Sending mail message via SMTP server $($MailMessage['SmtpServer']):$($MailMessage['Port'])"
        if ( $VerbosePreference ) { $MailMessage.Keys | Sort-Object | ForEach-Object { Write-Verbose ( '{0,-30} : {1}' -f $_, $MailMessage[$_] ) } }
        
        Send-MailMessage @MailMessage

    } else {

        Write-Verbose "Sending mail message via Outlook"

        try {

            $Outlook = New-Object -ComObject Outlook.Application

            $NewMail = $Outlook.CreateItem(0)
                
            if ( $MailMessage['To'] ) { $NewMail.To = $MailMessage['To'] }
            if ( $MailMessage['CC'] ) { $NewMail.CC = $MailMessage['CC'] }
                
            $NewMail.Subject  = $MailMessage['Subject']
            $NewMail.HTMLBody = $MailMessage['Body']

            $NewMail.Display()

        } catch {

            $DateStamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'

            $MailMessage.Body | Out-File -FilePath ".\$DateStamp-NewAccountEmail.html" -Encoding utf8

            Write-Warning "Outlook is not available, no confirmation email will be sent"
            Write-Warning "Mail message saved to: $DateStamp-NewAccountEmail.html"

        }

    }

}
