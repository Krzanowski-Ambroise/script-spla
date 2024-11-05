# Export de la liste des utilisateurs connectés dans le mois
$NumDays = 30
$LogDir = ".\PVA-Last-Logon.csv"

$currentDate = [System.DateTime]::Now
$currentDateUtc = $currentDate.ToUniversalTime()
$lltstamplimit = $currentDateUtc.AddDays(- $NumDays)
$lltIntLimit = $lltstamplimit.ToFileTime()
$adobjroot = [adsi]''
$objstalesearcher = New-Object System.DirectoryServices.DirectorySearcher($adobjroot)
$objstalesearcher.filter = "(&(objectCategory=person)(objectClass=user)(!cn=Admin*)(!cn=Exploitation)(!cn=MSOL*)(!cn=compte*)(!cn=Progiteam)(!cn=SuperVision-Backup)(lastLogon>=" + $lltIntLimit + "))"

$users = $objstalesearcher.findall() | select `
@{e={$_.properties.cn};n='Display Name'},`
@{e={$_.properties.samaccountname};n='Username'},`
@{e={[datetime]::FromFileTimeUtc([int64]$_.properties.lastlogon[0])};n='Last Logon'},`
@{e={[string]$adspath=$_.properties.adspath;$account=[ADSI]$adspath;$account.psbase.invokeget('AccountDisabled')};n='Account Is Disabled'} 

$users | Export-Csv -Delimiter ";" -NoTypeInformation $LogDir

# Informations d'authentification OAuth2 a créer et mettre ici 
$tenantId = "579c1d1b-bc4b-444e-bc49-752299c25759"
$clientId = "e09e0b08-0f7d-4627-bf48-803977d662fd"
$clientSecret = "44YFveh6"
$userEmail = "sinheb-spla@sinergence.fr"

# Obtenir le jeton d'accès OAuth2
try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    } -ErrorAction Stop
}
catch {
    Write-Host "Erreur lors de l'obtention du jeton d'accès OAuth2 : $_"
    exit
}

# Préparer le message e-mail
$emailBody = @{
    message = @{
        subject = "PVA - SPLA"
        body = @{
            contentType = "Text"
            content = "Veuillez trouver ci-joint le rapport PVA-Last-Logon."
        }
        toRecipients = @(
            <#@{
                emailAddress = @{
                    address = "julie.vandenbussche@sinergence.fr"
                    name = "Julie"
                }
            },
            @{
                emailAddress = @{
                    address = "cedric.martin@sinergence.fr"
                    name = "Ced"
                }
            },#>
            @{
                emailAddress = @{
                    address = "ambroise.krzanowski@sinergence.fr"
                    name = "Ambroise"
                }
            }
        )
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = "PVA-Last-Logon.csv"
                contentBytes = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes("PVA-Last-Logon.csv"))
            }
        )
    }
}

# Envoyer l'e-mail via Microsoft Graph API
try {
    Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$userEmail/sendMail" -Headers @{
        Authorization = "Bearer $($tokenResponse.access_token)"
    } -Body ($emailBody | ConvertTo-Json -Depth 4) -ContentType "application/json" -ErrorAction Stop
    Write-Host "E-mail envoyé avec succès."
}
catch {
    Write-Host "Erreur lors de l'envoi de l'e-mail : $_"
}
