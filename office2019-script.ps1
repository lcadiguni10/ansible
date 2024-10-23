# ---------------------------------------> Office 2019

# Define o caminho para o diretório do Office
$officePath = "c:\temp"  # Altere o caminho conforme necessário
$configurationFile = "$officePath\configuration-Office2019Enterprise.xml"
$languagePackPath = "$officePath\officesetup.exe"  # Altere conforme o caminho do pacote de linguagem

# Executa a instalação do Office
Write-Host "Instalando o Office..."
Start-Process -FilePath "$officePath\setup.exe" -ArgumentList "/configure `"$configurationFile`"" -Wait

# Verifica se o pacote de linguagem existe e instala
if (Test-Path $languagePackPath) {
    Write-Host "Instalando o pacote de linguagem..."
    Start-Process -FilePath $languagePackPath -Wait
} else {
    Write-Host "Pacote de linguagem não encontrado em: $languagePackPath"
}

# Ativação do Office 2019 com chave
$officeKey = "NXJ7V-8JKDB-9XWMC-332XV-RGFXQ"  # Substitua com sua chave de produto do Office 2019

Write-Host "Ativando o Office..."
$osppPath = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"  # Caminho do ospp.vbs no sistema
if (Test-Path $osppPath) {
    & "cscript.exe" "$osppPath" /inpkey:$officeKey
    & "cscript.exe" "$osppPath" /act
} else {
    Write-Host "Arquivo ospp.vbs não encontrado em: $osppPath"
}

