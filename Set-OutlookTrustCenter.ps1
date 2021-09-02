Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
Add-Type -AssemblyName "System.Runtime.Interopservices"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")

# the above somehow doesn't work...
Set-Location -Path "HKCU:\software\Microsoft\Office\16.0\Outlook\Security"
# https://answers.microsoft.com/en-us/msoffice/forum/all/registry-keys-for-outlook-trust-center/6ff7436f-4ab0-4e6e-8c37-01422d7ebbf6

# Prevention technique can be found here:
# https://activedirectorypro.com/disable-powershell-with-group-policy/