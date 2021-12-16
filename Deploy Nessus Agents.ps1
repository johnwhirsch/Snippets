$group = $((Get-WmiObject win32_computersystem).Domain)
#Uncomment and edit the line below to manually override the group selection
#$group = "domain.local"
$Nessus_Key = "YOUR_NESSUS_KEY_HERE" 


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try{ Invoke-WebRequest -Uri "https://www.tenable.com/downloads/api/v1/public/pages/nessus-agents/downloads/15492/download?i_agree_to_tenable_license_agreement=true" -OutFile "$($env:APPDATA)\nessus.msi"}
catch{ Write-Error "Unable to download Nessus Agent from Tenable's site`n$error"; exit 660; }

if(Test-Path "$($env:APPDATA)\nessus.msi"){
     
    try{ Start-Process "msiexec.exe" -ArgumentList "/i `"$($env:APPDATA)\nessus.msi`" NESSUS_SERVER=`"sensor.cloud.tenable.com:443`" NESSUS_GROUPS=`"$($group)`" NESSUS_KEY=$($Nessus_Key) /qn" -Wait }
    catch{ Write-Error "Unable to install Nessus Agent"; exit 664; }
}
 
if((Get-Service -Name 'Tenable Nessus Agent').Status -eq "Running"){ Write-Output "Install Succeeeded"; exit 0; }
else{ Write-Error "Install failed"; Exit 666; }
