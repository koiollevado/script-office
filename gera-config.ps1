# Script PowerShell para configurar o arquivo XML de instalação do Office

# Função para gerar o conteúdo do arquivo XML com base nas escolhas do usuário
function Generate-OfficeXML {
    param (
        [string]$productID,
        [string]$bitVersion,
        [array]$selectedApps
    )

    # Cabeçalho do XML
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="$bitVersion" Channel="Current">
    <Product ID="$productID">
      <Language ID="pt-br" />
"@

    # Adiciona os aplicativos selecionados
    foreach ($app in @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneDrive", "Teams", "SkypeForBusiness", "Skype")) {
        if ($selectedApps -notcontains $app) {
            $xmlContent += "            <ExcludeApp ID=`"$app`" />`n"
        }
    }

    # Rodapé do XML
    $xmlContent += @"
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Updates Enabled="FALSE" />
  <RemoveMSI />
</Configuration>
"@

    return $xmlContent
}

# Menu para escolher a versão do Office
$officeOptions = @{
    1 = "O365ProPlusRetail"
    2 = "ProPlus2019Volume"
    3 = "ProPlus2021Volume"
    4 = "PerpetualVL2024"
    5 = "ProPlus2016"
}

Write-Host "Escolha a versão do Office:"
foreach ($key in $officeOptions.Keys) {
    Write-Host "$key. $($officeOptions[$key])"
}

$officeChoice = Read-Host "Digite o número correspondente à versão desejada"
$productID = $officeOptions[$officeChoice]

# Menu para escolher 32 ou 64 bits
$bitVersion = Read-Host "Digite a versão desejada (32 ou 64)"

# Exibir opções de aplicativos e coletar escolhas
$apps = @(
    "1 - Word", "2 - Excel", "3 - PowerPoint", "4 - Outlook", "5 - Access", 
    "6 - Publisher", "7 - OneDrive", "8 - Teams", "9 - SkypeForBusiness", "0 - Skype"
)

Write-Host "Escolha os aplicativos desejados (separados por vírgula):"
foreach ($app in $apps) {
    Write-Host $app
}

$appChoices = Read-Host "Digite os números dos aplicativos desejados"
$appIndexes = $appChoices -split "," | ForEach-Object { $_.Trim() }
$selectedApps = @()

foreach ($index in $appIndexes) {
    $selectedApps += $apps[$index - 1].Split(' - ')[1]
}

# Gerar o XML baseado nas escolhas
$xmlContent = Generate-OfficeXML -productID $productID -bitVersion $bitVersion -selectedApps $selectedApps

# Salvar o arquivo XML
$xmlPath = "config.xml"
$xmlContent | Out-File -Encoding UTF8 $xmlPath

Write-Host "Arquivo config.xml gerado com sucesso!"
