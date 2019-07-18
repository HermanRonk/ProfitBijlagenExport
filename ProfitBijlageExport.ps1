# Script voor het exporteren van abonementsbijlages bijlages
# 
# Functie: ProfitAuthorisatie
# - Autorisatie profit omgeving
# Functie: GetAbo
# - Geeft een object met alle abonnementen met bijlages. Administratiecode vereist
# Functie: GetBijlagen
# - Maak de lijst met beschikbare bijlagen
# Functie: DownloadBijlagen
# - Download aan de hand van de GetBijlagen output de daadwerkelijke files en schrijf deze lokaal weg
# Functie: ScriptUitvoeren
# - Centrale looper

# Geschreven door Herman Ronk
# Datum: 17-07-2019
# Contact: herman.ronk@detron.nl 



# Functie voor opbouwen Profit autorsatie
function ProfitAuthorisatie {
    # Bouw de autorisatie header voor AFAS Profit
    $token = '<token><version>1</version><data>***</data></token>'
    $encodedToken = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($token))
    $authValue = "AfasToken $encodedToken"
    $Headers = @{
        Authorization = $authValue
    }
    $Headers
    return
}

function GetAbo {
    param (
        [int32]$administratie
    )
    # Definieer de URL voor de Profit omgeving
    $url = 'https://***/profitrestservices/connectors/DossierBijlagen?filterfieldids=AdministratieAbo%3BAdministratieProject&filtervalues=' + $administratie + '%3B' + $administratie + '&operatortypes=1%3B1&skip=-1&take=200'
    # Connect and retrieve the table
    $auth = ProfitAuthorisatie
    ((Invoke-WebRequest -Uri $url -UseBasicParsing -Headers $auth).content | ConvertFrom-Json).Rows
}

function GetBijlagen {
    param (
        [int32]$DossieritemID
    )
    # Definieer de URL voor de Profit omgeving
    $url_dossier = 'https://***/profitrestservices/connectors/AbonnementenExport?filterfieldids=DossieritemID&filtervalues=' + $DossieritemID + '&operatortypes=1&skip=-1&take=-1'
    # Connect and retrieve the table
    $auth = ProfitAuthorisatie
    ((Invoke-WebRequest -Uri $url_dossier -UseBasicParsing -Headers $auth).content | ConvertFrom-Json).Rows
}

function DownloadBijlagen {
    param (
        [string]$filename,
        [string]$guid,
        [string]$map,
        [string]$abo
    )

    # Definieer de URL voor de Profit omgeving
    $url_download = 'https://***/profitrestservices/fileconnector/' + $guid + '/' + $filename 
    # Connect and retrieve the table
    $auth = ProfitAuthorisatie
    $fileout = $map + '\' + $abo + '-' + $filename 
    $download_error = 0

    # Voorkomen dat we lege files wegschrijven omdat de download niet uitgevoerd kan worden.
    try {
        $tempfile = Invoke-RestMethod -Uri $url_download -Method Get -Headers $auth 
    }
    catch {
        Write-Host $_ -fore Red
        $download_error = 1
    }
    
    if ($download_error -eq 0) {
        $bytes = [System.Convert]::FromBase64String($tempfile.filedata)
        [IO.File]::WriteAllBytes($fileout, $bytes)
        Write-Host $filename "- Succesvol weggeschreven in"  $fileout -ForegroundColor green
    } 
}

function ScriptUitvoeren {
    param (
        [int32]$administratie
    )
    
    # Eerst Getabo
    $records = GetAbo $administratie

    # Per Abo dossieritems ophalen
    foreach ($record in $records) {
        
        # Map bepalen voor het eventueel opslaan van bijlagen
        $map = "C:\Exports\" + $record.Abonnement  

        # Bijlagenlijst per dossieritem maken
        $bijlagenlijst = GetBijlagen $record.DossieritemID

        # Als de map voor het abonnement nog niet bestaat deze aanmaken
        If (!(test-path $map) -And $bijlagenlijst.count -gt 0) {
            New-Item -ItemType Directory -Force -Path $map
        }

        # Per bijlage file downloaden
        foreach ($bijlage in $bijlagenlijst) {
            DownloadBijlagen $bijlage.Naam $bijlage.Bijlage $map $record.Abonnement
        }
    }
}
