# SendGrid basic API wrapper

A wrapper for the SenGrid v3 API that lets you send email as a Powershell command, similar to Send-MailMessage.

## Installation 

Use as an azure automation module, or import as a Powershell Module.
To use in an azure Function, copy the module files and publish it along with your code.

## Usage

Import-Module Sendgridwrapper

Send-gridmailmessage <parameters>

If you need to send an attachment, specify the content with -Attachmentcontent, or the path to the file with -AttachmentFilePath.

An HTML template is included with a standard title-body-signature template. Use -UseHTML $true.  

If the script runs from an Azure Runbook, the API key should be stored as an encrypted variable and called with the "Get-AutomationVariable -Name 'APIKEY-Sendgrid'" command.

If the script runs from an Azure Function, the API Key should be stored in an Az Keyvault and called with the "(Get-AzKeyVaultSecret -SecretName APIKEY -VaultName <VaultName>).SecretValueText" command.
    
The System Managed Identity must have access to the keyvault and the read secret authorization.

### Syntax
```powershell
Send-GridMailMessage -APIKEY <string> -From <string> -To <array> -Subject <string> -Text <string> [-Bcc <array>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <string> -From <string> -To <array> -Subject <string> -AttachmentFileName <string> -Text <string> [-Bcc <array>] [-AttachmentPath <Object>] [-UseHTML <bool>] [-HTMLTitle <string>] [-HTMLMainText <string>] [-HTMLEndText <string>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <string> -From <string> -To <array> -Subject <string> -AttachmentFileName <string> -Text <string> [-Bcc <array>] [-AttachmentContent <Object>] [-UseHTML <bool>] [-HTMLTitle <string>] [-HTMLMainText <string>] [-HTMLEndText <string>] [<CommonParameters>]

Send-GridMailMessage -APIKEY <string> -From <string> -To <array> -Subject <string> -UseHTML <bool> -HTMLTitle <string> -HTMLMainText <string> [-Bcc <array>] [-HTMLEndText <string>] [<CommonParameters>]
```
