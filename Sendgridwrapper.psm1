<#PSSCRIPTInfo
.SYNOPSIS
    Handles data formatting to send an email through SendGrid. Supports adding a CSV attachment and basic 3-part TITLE/MAINTEXT/ENDTEXT HTML template or plain text one liner, for instance "this is an automated email message".
    Used for automated email messages associated with scripts. Requires an API_KEY which should be an Azure Runbook encrypted variable for Azure Automation, an Azure Keyvault Secret for Azure Functions, or a SecureString.

.DESCRIPTION
    Send an email through sendgrid's v3 SendMail API.
    Parameter Splatting is highly recommanded to increase readability of the script (About_Splatting)

.INPUTS
    System.Object[]

.OUTPUTS
    [System.Collections.Hashtable]

.NOTES
    Version         1.14
    Author          Marcellin Penicaud (mpe@openhost.io)
    Creation Date   19/11/2019
    This is an old script I did, which could use many many improvements. I learned a lot writing it and it saved me a great amount of time.

.PARAMETER APIKEY
    If the script runs from an Azure Runbook, the API key should be stored as an encrypted variable and called with the "Get-AutomationVariable -Name 'APIKEY-Sendgrid'" command.
    If the script runs from an Azure Function, the API Key should be stored in an Az Keyvault and called with the "(Get-AzKeyVaultSecret -SecretName APIKEY -VaultName <VaultName>).SecretValue" command.
    The System Managed Identity must have access to the keyvault and the read secret authorization.

.PARAMETER TEXT
-TEXT is only mandatory if not using HTML (-UseHTML $true).

.SYNTAX
Send-GridMailMessage -APIKEY <secure.string> -From <string> -To <array> -Subject <string> -Text <string> -ReplyToDisplayName <string> -ReplayToMailAddress <string> [-Bcc <array>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <secure.string> -From <string> -To <array> -Subject <string> -AttachmentFileName <string> -Text <string> -ReplyToDisplayName <string> -ReplayToMailAddress <string> [-Bcc <array>] [-AttachmentPath <Object>] [-UseHTML <bool>] [-HTMLTitle <string>] [-HTMLMainText <string>] [-HTMLEndText <string>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <secure.string> -From <string> -To <array> -Subject <string> -AttachmentFileName <string> -Text <string> -ReplyToDisplayName <string> -ReplayToMailAddress <string> [-Bcc <array>] [-AttachmentContent <Object>] [-UseHTML <bool>] [-HTMLTitle <string>] [-HTMLMainText <string>] [-HTMLEndText <string>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <secure.string> -From <string> -To <array> -Subject <string> -UseHTML <bool> -HTMLTitle <string> -HTMLMainText <string> -ReplyToDisplayName <string> -ReplayToMailAddress <string> [-Bcc <array>] [-HTMLEndText <string>] [<CommonParameters>]

.EXAMPLE
$SimpleParams = @{
    APIKEY = (Get-AutomationVariable -Name 'APIKEY')
    Subject = $Subject
    To = $To
    Bcc = $Bcc
    Text = $TextString
    From = $From
}
Send-GridMailMessage @SimpleParams

.EXAMPLE
$Arrayparams = @{
    APIKEY = (Get-AutomationVariable -Name 'APIKEY')
    Subject = $Subject
    To = [array]$To
    From = $From
    Bcc = [array]$Bcc
    Text = $Text

}
Send-GridMailMessage @Arrayparams

.EXAMPLE
$HTMLParams = @{
    APIKEY = (Get-AutomationVariable -Name 'APIKEY')
    Subject = $Subject
    To = $ToString
    Bcc = $BccString
    From = $From
    UseHTML = $true
    HTMLTitle = $HTMLTitle
    HTMLMaintext = $HTMLMaintext
    HTMLEndText = $HTMLEndText
}
Send-GridMailMessage @HTMLParams

.EXAMPLE
$AttachmentFileParams = @{
    APIKEY = $APIKEY
    To = $ToString
    Subject = $Subject
    Bcc = $BccString
    Text = $TextString
    From = $From
    AttachmentFile = $AttachmentFile
    AttachmentFileName = "Attachment-File-name"
}
Send-GridMailMessage @AttachmentFileParams

.EXAMPLE
$AttachmentContentParams = @{
    APIKEY = $APIKEY
    To = $ToString
    Subject = $Subject
    Bcc = $BccString
    Text = $TextString
    From = $From
    AttachmentContent = $AttachmentContent
    AttachmentFileName = "Attachment-File-name"
}
Send-GridMailMessage @AttachmentContentParams
#>
Function Send-GridMailMessage {
    [cmdletbinding(DefaultParameterSetName = 'none')]
    Param(
        [Parameter (Mandatory)]
        [System.Security.SecureString] $APIKEY,

        [Parameter (Mandatory)]
        [ValidateNotNullorEmpty()]
        [Validatescript( {
                if ($_ -like "*@*.*") {
                    $true
                }
                else {
                    throw "$_ is not a valid sender email address. -From Must match *@*.* pattern."
                }
            }
        )]
        [String] $From,

        [Parameter (Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 320)]
        [array] $To,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [array] $Bcc,

        [Parameter (Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $Subject,

        [Parameter(Mandatory = $false,
            ParameterSetName = 'Attachment')]
        [ValidateNotNullOrEmpty()]
        $AttachmentPath,

        [parameter(ParameterSetName = "Attachment_Content")]
        [ValidateNotNullOrEmpty()]
        $AttachmentContent,

        [Parameter(ParameterSetName = 'Attachment_Content',
            Mandatory)]
        [Parameter(ParameterSetName = 'Attachment',
            Mandatory)]
        [ValidateNotNullorEmpty()]
        [String] $AttachmentFileName,

        [Parameter(ParameterSetName = 'html',
            Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content')]
        [Parameter(ParameterSetName = 'Attachment')]
        [ValidateNotNullOrEmpty()]
        [bool] $UseHTML = $false,

        # Text is mandatory unless HTML is used, in which case it is forbidden.
        [Parameter(ParameterSetname = 'none',
            Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content',
            Mandatory = $false)]
        [Parameter(ParameterSetName = 'Attachment',
            Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [String] $Text,

        [Parameter(ParameterSetName = 'html',
            Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content')]
        [Parameter(ParameterSetName = 'Attachment')]
        [ValidateNotNullOrEmpty()]
        [String] $HTMLTitle,

        [Parameter(ParameterSetName = 'html',
            Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content')]
        [Parameter(ParameterSetName = 'Attachment')]
        [ValidateNotNullOrEmpty()]
        [String] $HTMLMainText,

        [Parameter(ParameterSetName = 'html',
            Mandatory = $false)]
        [Parameter(ParameterSetName = 'Attachment_Content')]
        [Parameter(ParameterSetName = 'Attachment')]
        [ValidateNotNullOrEmpty()]
        [String] $HTMLEndText = "This is an automated email message, please reply to ",

        [Parameter(ParameterSetName = 'html', Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content', Mandatory)]
        [Parameter(ParameterSetName = 'Attachment', Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ReplyToMailAddress,

        [parameter(ParameterSetName = 'none', Mandatory)]
        [Parameter(ParameterSetName = 'html', Mandatory)]
        [Parameter(ParameterSetName = 'Attachment_Content', Mandatory)]
        [Parameter(ParameterSetName = 'Attachment', Mandatory)]
        [ValidateNotNullOrEmpty()]

        [string] $ReplyToDisplayName
    )
    ### CRAFT HEADERS ###
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Bearer " + $APIKEY)
    $headers.Add("Content-Type", "application/json")

    ### Encode Attachment to Base64 ###
    # SendGrid API expects content as base64 encoded https://sendgrid.com/docs/API_Reference/Web_API_v3/Mail/index.html
    # This command encodes the FILE to a single base64 string

    # If attachment is a file, encode it to Base64. Else, convert the data to CSV in temp folder, then convert the content to base64.
    If ($AttachmentPath) {
        $base64String = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($AttachmentPath))
    }
    Else {
        if ($AttachmentFileName -notlike "*.csv") {
            $TempFilePath = (New-Item -type File -Path "$AttachmentFileName.csv").FullName
            $AttachmentContent | Export-CSV -Path $TempFilePath -Force -Delimiter ";" -NoTypeInformation -Encoding utf8
            $base64String = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($TempFilePath))
        }
        Else {
            $TempFilePath = (New-Item -type File -path "$AttachmentFileName").Fullname
            $AttachmentContent | Export-CSV -Path $TempFilePath -Force -Delimiter ";" -NoTypeInformation -Encoding utf8
            $base64String = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($TempFilePath))
        }
    }
    ### Transform parameters to array objects that will be used in the body ###

    # Craft the recipients
    $JSON_To = [array]($To | ForEach-Object {
            @{ email = $_ }
        })

    # Craft the BCC
    $JSON_Bcc = [array]($Bcc | ForEach-Object {
            @{ email = $_ }
        })

    If ($UseHTML -eq $false ) {
        [String]$EmailContentType = "text/plain"
    }
    if ($UseHTML -eq $true) {
        [String]$EmailContentType = "text/html"
    }
    ### HTML ###
    # Create a here-string containing a Standard, Simple looking HTML 3 lines text. Default Content should be TITLE + CONTENT + Default noreply message, but that's not mandatory.
    # Line breaks are supported in the form <br> in the $text parameter.
    If ($UseHTML -eq $true) {
        $HTMLContent = @"
        <html>

        <head></head>

        <body class="em_body" style="margin:0px; padding:0px;" bgcolor="#efefef">
            <table class="em_full_wrap" valign="top" width="100%" cellspacing="0" cellpadding="0" border="0" bgcolor="#efefef"
                align="center">
                <tbody>
                    <tr>
                        <td valign="top" align="center">
                            <table class="em_main_table" style="width:700px;" width="700" cellspacing="0" cellpadding="0"
                                border="0" align="center">
                                <tbody>
                                    <tr>
                                        <td style="padding:35px 70px 30px;" class="em_padd" valign="top" bgcolor="#0d1121"
                                            align="center">
                                            <table width="100%" cellspacing="0" cellpadding="0" border="0" align="center">
                                                <tbody>
                                                    <tr>
                                                        <td style="font-family:'Open Sans', Arial, sans-serif; font-size:18px; line-height:22px; color:#fbeb59; text-transform:uppercase; letter-spacing:2px; padding-bottom:12px;"
                                                            valign="top" align="center">$HTMLTitle</td>
                                                    </tr>
                                                    <tr>
                                                        <td style="font-size:0px; line-height:0px; height:15px;" height="15">
                                                            &nbsp;</td>
                                                    </tr>
                                                    <tr>
                                                        <td style="font-family:'Open Sans', Arial, sans-serif; font-size:18px; line-height:22px; color:#fbeb59; letter-spacing:2px; padding-bottom:12px;"
                                                            valign="top" align="center">$HTMLMainText
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td class="em_h20" style="font-size:0px; line-height:0px; height:25px;"
                                                            height="25">&nbsp;</td>
                                                    </tr>
                                                    <tr>
                                                        <td style="font-family:'Open Sans', Arial, sans-serif; font-size:18px; line-height:22px; color:#fbeb59; text-transform:uppercase; letter-spacing:2px; padding-bottom:12px;"
                                                            valign="top" align="center">$HTMLEndText
                                                        </td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                </tbody>
            </table>
            <div class="em_hide" style="white-space: nowrap; display: none; font-size:0px; line-height:0px;">&nbsp; &nbsp;
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</div>
        </body>

        </html>
"@
        #Remove special characters
        $HTMLContent = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($HTMLContent))
        #Encode $HTMLContent to UTF-8
        $HTMLContent = [Text.Encoding]::UTF8.GetString([text.encoding]::ASCII.GetBytes($HTMLContent))

        $text = $HTMLContent
    }
    # Craft the body.
    # Reply_to has been added in case someone tries to answer the email as end-users will receive them. Domain should match in $from and $reply_to to domains.
    $body = @{
        personalizations = @(
            @{
                to = $JSON_To
            }
        )
        from             = @{
            email = $from

        }
        reply_to         = @{
            email = $ReplyToMailAddress
            name  = $ReplyToDisplayName
        }
        subject          = $subject
        content          = @(
            @{
                type  = $EmailContentType
                value = $Text

            }
        )
    }
    # Add the BCC if it exists
    If ($null -ne $BCC) {
        $Body.personalizations[0].Add("Bcc", $JSON_Bcc)
    }
    # Add the attachment to the $Body if it exists
    if ($null -ne $AttachmentPath -or $null -ne $AttachmentContent) {
        $body["attachments"] = @(
            @{
                filename    = $AttachmentFileName
                content     = $base64String
                disposition = "attachment"
                type        = "text/csv"
            }
        )
    }
    # Converts the body to JSON
    $BodyJson = $body | ConvertTo-Json -Depth 4
    Write-Verbose $BodyJson
    # Sends the email through SendGrid API
    $Output = Invoke-RestMethod -Uri https://api.sendgrid.com/v3/mail/send -Method Post -Headers $headers -Body $bodyJson
    Write-Output $Output
}