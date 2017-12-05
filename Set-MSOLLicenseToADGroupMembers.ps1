######################################################################
## (C) 2017 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      Set-MSOLLicenseToADGroupMembers.ps1
##
## Version:       1.1
##
## Release:       Final
##
## Requirements:  -none-
##
## Description:   Assign Office 365 Licenses based on Active Directory
##                Groups.
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################
Set-PSDebug -Strict
Set-StrictMode -Version latest
  
 
function Set-MSOLLicenseToADGroupMembers {
    <#
    .SYNOPSIS
    Assigns Office 365 Licenses to Members of AD-Group
  
    .DESCRIPTION
    The Set-MSOLLicenseToADGroupMembers CMDlet gets all users from a
    specified AD-Group and assigns a specified Office 365 License to
    the corresponding Office 365 identities
  
    .PARAMETER GroupName
    Name of the Active Group
  
    .PARAMETER License
    Name of the License
 
    .PARAMETER UsageLocation
    Name of the Location
  
    .EXAMPLE
    Set-MSOLLicenseToADGroupMembers -GroupName "Office365_E3" -License "contoso:ENTERPRISEPACK"
 
    #>
      
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$GroupName,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$License,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$LicenseName,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]$UsageLocation="DE"
    )
 
    foreach ($User in (Get-ADGroupMember -Identity $GroupName)) {
    $User = Get-ADUser -Identity $User -Properties UserPrincipalName
    $MsOlUser =  Get-MsolUser -UserPrincipalName $User.UserPrincipalName -ErrorAction SilentlyContinue 
 
    if ($MsOlUser -ne $null) {
        Write-Host ("Found Office 365 User: " + $MsOlUser.UserPrincipalName)
            Set-MsolUser -UserPrincipalName $User.UserPrincipalName -UsageLocation $UsageLocation
 
        if ($MSOlUser.IsLicensed -eq $false) {
            Write-Host ("User not licensed - assigning license: " + $LicenseName)

            Set-MSOLUserLicense -UserPrincipalName $User.UserPrincipalName -AddLicenses $LicenseName
            Set-MsolUserLicense -UserPrincipalName $User.UserPrincipalName -LicenseOptions $License
        }
        else {
            foreach ($UserLicense in $MSOLUser.Licenses)
            {
                if ($UserLicense.AccountSkuId -eq $LicenseName)
                {
                    $UserLicensePresent = $true
                    break
                }
            }


            if ($UserLicensePresent -eq $false) {
                Write-Host ("User licensed, but not with correct license - assigning license: " + $LicenseName)

                Set-MSOLUserLicense -UserPrincipalName $User.UserPrincipalName -AddLicenses $LicenseName
                Set-MsolUserLicense -UserPrincipalName $User.UserPrincipalName -LicenseOptions $License
            }
            else {
                Write-Host ("User has already a correct license assigned: " + $LicenseName)
            }
        }
    }
    }
 
}
 
Import-Module MSOnline
 
$Username = "xxxxx"
$Password = "xxxxx"
 
# Convert the plain text password to a secure string
$SecurePassword=ConvertTo-SecureString –String $Password –AsPlainText –force
$Credential=New-object System.Management.Automation.PSCredential $Username,$SecurePassword
 
# Create new Office 365 license options
$LicenseOfficeProPlus = New-MsolLicenseOptions -AccountSkuId "SUBSCRIPTION_NAME:OFFICESUBSCRIPTION"
$LicenseE3withoutExchange = New-MsolLicenseOptions -AccountSkuId "SUBSCRIPTION_NAME:ENTERPRISEPACK" -DisabledPlans "EXCHANGE_S_ENTERPRISE"
 
# Connect to Microsoft Office 365 tenant
Connect-MsolService -Credential $credential
 
# Assign licenses based on Active Directory group membership
Set-MSOLLicenseToADGroupMembers -GroupName "O365_PROPLUS" -License $LicenseOfficeProPlus -LicenseName "SUBSCRIPTION_NAME:OFFICESUBSCRIPTION" -UsageLocation "DE"
Set-MSOLLicenseToADGroupMembers -GroupName "O365_E3" -License $LicenseE3withoutExchange -LicenseName "SUBSCRIPTION_NAME:ENTERPRISEPACK" -UsageLocation "DE"
 
