#The following variables should be changed:
#$file ? should be named with a .json extension
#$fromaddress
#$toaddress
#$smtpserver
#$Password
#$port

# Set output path to "C:\ProgramData\AssetInventory\inventory.json"
$Folder = $env:ALLUSERSPROFILE + "\\AssetInventory"
$ScriptPath = $Folder + "\\inventory.json"

New-Item -Path $Folder -ItemType Directory -Force | Out-Null
Set-Content -Path $ScriptPath -Value $Content -Force

$Argument = @"
-NoProfile -WindowStyle Hidden -File $ScriptPath
"@

# Make NuGet available
$Packager = Get-PackageProvider -ListAvailable | Where-Object {$_.Name -eq "NuGet"}

if ($Packager -eq $null)
{
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
}
else
{
    Import-PackageProvider -Name NuGet -Force -Confirm:$false
}


# Install/update AssetInventory module
$Module = Get-Module -ListAvailable | Where-Object {$_.Name -eq "AssetInventory"}

if ($Module -eq $null)
{
    Install-Module AssetInventory -Force -Confirm:$false
}
else
{
    Update-Module AssetInventory -Force -Confirm:$false
}

# Install/update RDM Helper module
$RdmModule = Get-Module -ListAvailable | Where-Object {$_.Name -eq "RdmHelper"}

if ($RdmModule -eq $null)
{
    Install-Module -Name RdmHelper -Force -Confirm:$false
}
else
{
    Update-Module RdmHelper -Force -Confirm:$false
}

# Install/update SnipeitPS module
$SnipeModule = Get-Module -ListAvailable | Where-Object {$_.Name -eq "SnipeitPS"}

if ($SnipeModule -eq $null)
{
    Install-Module -Name SnipeitPS -Force -Confirm:$false
}
else
{
    Update-Module SnipeitPS -Force -Confirm:$false
}

# Set execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

# Import modules
Import-Module AssetInventory
Import-Module RdmHelper
Import-Module SnipeitPS

# Install RDM Free
Install-RdmPackage -Edition 'Free' -Force

# Run inventory and output as JSON
Get-AssetInventory -AsJson | Out-File -FilePath $ScriptPath -Force

# Setup SnipeIT
$SnipeUri = "{[Snipe_Url]}" # eg 'https://asset.example.com'
$SnipeUri = "{[Snipe_Key]}"
Set-SnipeitInfo -URL '$SnipeUri' -apiKey '$SnipeKey'

#Send Email

$fromaddress = "{[From_Address]}"
$toaddress = "{[To_Address]}"
$Subject = "Asset inventory report for"
$body = Get-Content $ScriptPath
$smtpserver = "{[SMTP_Server]}" #for example, smtp.office365.com
$Password = "{[Password]}"
$port = {[SMTP_Port]} #for example, 587
 
$message = new-object System.Net.Mail.MailMessage
$message.IsBodyHTML = $true
$message.From = $fromaddress
$message.To.Add($toaddress)
$message.Subject = $Subject
$message.body = $body
$smtp = new-object Net.Mail.SmtpClient($smtpserver, $port)
$smtp.EnableSsl = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($fromaddress, $Password)
$smtp.Send($message)