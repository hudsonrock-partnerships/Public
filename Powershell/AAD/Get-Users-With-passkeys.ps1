# Import Microsoft Graph modules
#Import-Module Microsoft.Graph.Users
#Import-Module Microsoft.Graph.Identity.SignIns

# Connect to Microsoft Graph (interactive or managed identity)
#Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All"

# Get all users - filter out guests and only include synced AD accounts
$users = Get-MgBetaUser -All -Property Id,DisplayName,UserPrincipalName,Department,CompanyName,UserType,OnPremisesSyncEnabled | Where-Object {
    $_.UserType -eq "Member" -and $_.OnPremisesSyncEnabled -eq $true
}
# For testing with single user:
# $users = Get-MgBetaUser -userid mteo@tbdvox.com -Property Id,DisplayName,UserPrincipalName,Department,CompanyName,UserType,OnPremisesSyncEnabled

# Prepare results array
$results = @()
$userCounter = 0

foreach ($user in $users) {
    $userCounter++
    Write-Progress -Activity "Processing Users" -Status "Processing $($user.DisplayName)" -PercentComplete (($userCounter / $users.Count) * 100)
    
    # Get authentication methods for the user
    $authMethods = Get-mgbetauserAuthenticationMethod -UserId $user.Id

    # Filter for passkeys (FIDO2 methods) - these typically have longer alphanumeric IDs or specific @odata.type
    $passkeys = $authMethods | Where-Object { 
        $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.fido2AuthenticationMethod'
    }

    # Extract passkey details including model
    $passkeyDetails = @()
    $passkeyModels = @()
    
    foreach ($passkey in $passkeys) {
        $model = $passkey.AdditionalProperties['model']
        $displayName = $passkey.AdditionalProperties['displayName']
        $aaGuid = $passkey.AdditionalProperties['aaGuid']
        
        # Determine device type based on aaGuid with comprehensive YubiKey detection
        $deviceType = switch ($aaGuid) {
            # Microsoft Authenticator
            'de1e552d-db1d-4423-a619-566b625cdc84' { 'Android Authenticator' }
            '90a3ccdf-635c-4729-a248-9b709135078f' { 'iOS Authenticator' }
            
            # YubiKey 5 Series
            'cb69481e-8ff7-4039-93ec-0a2729a154a8' { 'YubiKey 5 Series (FW 5.1)' }
            'ee882879-721c-4913-9775-3dfcce97072a' { 'YubiKey 5 Series (FW 5.2/5.4)' }
            'fa2b99dc-9e39-4257-8f92-4a30d23c4118' { 'YubiKey 5 NFC (FW 5.1)' }
            '2fc0579f-8113-47ea-b116-bb5a8db9202a' { 'YubiKey 5 NFC (FW 5.2/5.4)' }
            'a25342c0-3cdc-4414-8e46-f4807fca511c' { 'YubiKey 5 NFC (FW 5.7)' }
            'd7781e5d-e353-46aa-afe2-3ca49f13332a' { 'YubiKey 5 NFC (FW 5.7)' }
            '662ef48a-95e2-4aaa-a6c1-5b9c40375824' { 'YubiKey 5 NFC Enhanced PIN' }
            '19083c3d-8383-4b18-bc03-8f1c9ab2fd1b' { 'YubiKey 5 Nano (FW 5.7)' }
            'ff4dac45-ede8-4ec2-aced-cf66103f4335' { 'YubiKey 5 Nano (FW 5.7)' }
            'c5ef55ff-ad9a-4b9f-b580-adebafe026d0' { 'YubiKey 5Ci (FW 5.2/5.4)' }
            'a02167b9-ae71-4ac7-9a07-06432ebb6f1c' { 'YubiKey 5Ci (FW 5.7)' }
            '24673149-6c86-42e7-98d9-433fb5b73296' { 'YubiKey 5Ci (FW 5.7)' }
            
            # YubiKey 5 FIPS Series
            'c1f9a0bc-1dd2-404a-b27f-8e29047a43fd' { 'YubiKey 5 FIPS (FW 5.4)' }
            '73bb0cd4-e502-49b8-9c6f-b59445bf720b' { 'YubiKey 5 FIPS Nano/C (FW 5.4)' }
            '85203421-48f9-4355-9bc8-8a53846e5083' { 'YubiKey 5Ci FIPS (FW 5.4)' }
            'fcc0118f-cd45-435b-8da1-9782b2da0715' { 'YubiKey 5 FIPS RC NFC (FW 5.7)' }
            '57f7de54-c807-4eab-b1c6-1c9be7984e92' { 'YubiKey 5 FIPS RC Nano/C (FW 5.7)' }
            '7b96457d-e3cd-432b-9ceb-c9fdd7ef7432' { 'YubiKey 5Ci FIPS RC (FW 5.7)' }
            '79f3c8ba-9e35-484b-8f47-53a5a0f5c630' { 'YubiKey 5 FIPS Enterprise (FW 5.7)' }
            '905b4cb4-ed6f-4da9-92fc-45e0d4e9b5c7' { 'YubiKey 5 FIPS Enterprise Nano/C (FW 5.7)' }
            '3a662962-c6d4-4023-bebb-98ae92e78e20' { 'YubiKey 5Ci FIPS Enterprise (FW 5.7)' }
            
            # YubiKey Bio Series
            'd8522d9f-575b-4866-88a9-ba99fa02f35b' { 'YubiKey Bio FIDO (FW 5.5/5.6)' }
            'dd86a2da-86a0-4cbe-b462-4bd31f57bc6f' { 'YubiKey Bio FIDO (FW 5.7)' }
            '7409272d-1ff9-4e10-9fc9-ac0019c124fd' { 'YubiKey Bio FIDO (FW 5.7)' }
            '7d1351a6-e097-4852-b8bf-c9ac5c9ce4a3' { 'YubiKey Bio Multi-protocol (FW 5.6)' }
            '90636e1f-ef82-43bf-bdcf-5255f139d12f' { 'YubiKey Bio Multi-protocol (FW 5.7)' }
            '34744913-4f57-4e6e-a527-e9ec3c4b94e6' { 'YubiKey Bio Multi-protocol (FW 5.7)' }
            '83c47309-aabb-4108-8470-8be838b573cb' { 'YubiKey Bio Enterprise FIDO (FW 5.6)' }
            '8c39ee86-7f9a-4a95-9ba3-f6b097e5c2ee' { 'YubiKey Bio Enterprise FIDO (FW 5.7)' }
            'ad08c78a-4e41-49b9-86a2-ac15b06899e2' { 'YubiKey Bio Enterprise FIDO (FW 5.7)' }
            '97e6a830-c952-4740-95fc-7c78dc97ce47' { 'YubiKey Bio Enterprise Multi-protocol (FW 5.7)' }
            '6ec5cff2-a0f9-4169-945b-f33b563f7b99' { 'YubiKey Bio Enterprise Multi-protocol (FW 5.7)' }
            
            # Security Key Series
            'f8a011f3-8c0a-4d15-8006-17111f9edc7d' { 'Security Key by Yubico (FW 5.1)' }
            'b92c3f9a-c014-4056-887f-140a2501163b' { 'Security Key by Yubico (FW 5.2)' }
            '6d44ba9b-f6ec-2e49-b930-0c8fe920cb73' { 'Security Key NFC (FW 5.1)' }
            '149a2021-8ef6-4133-96b8-81f8d5b7f1f5' { 'Security Key NFC (FW 5.2/5.4)' }
            'a4e9fc6d-4cbe-4758-b8ba-37598bb5bbaa' { 'Security Key NFC Black (FW 5.4)' }
            'e77e3c64-05e3-428b-8824-0cbeb04b829d' { 'Security Key NFC Black (FW 5.7)' }
            'b7d3f68e-88a6-471e-9ecf-2df26d041ede' { 'Security Key NFC Black (FW 5.7)' }
            '0bb43545-fd2c-4185-87dd-feb0b2916ace' { 'Security Key NFC Enterprise (FW 5.4)' }
            '47ab2fb4-66ac-4184-9ae1-86be814012d5' { 'Security Key NFC Enterprise (FW 5.7)' }
            'ed042a3a-4b22-4455-bb69-a267b652ae7e' { 'Security Key NFC Enterprise (FW 5.7)' }
            '9ff4cc65-6154-4fff-ba09-9e2af7882ad2' { 'Security Key NFC Enterprise (FW 5.7)' }
            '72c6b72d-8512-4c66-8359-9d3d10d9222f' { 'Security Key NFC Enterprise (FW 5.7)' }
            
            # YubiKey Enterprise Series
            '1ac71f64-468d-4fe0-bef1-0e5f2f551f18' { 'YubiKey 5 NFC Enterprise (FW 5.7)' }
            '6ab56fad-881f-4a43-acb2-0be065924522' { 'YubiKey 5 NFC Enterprise (FW 5.7)' }
            'b2c1a50b-dad8-4dc7-ba4d-0ce9597904bc' { 'YubiKey 5 NFC Enterprise Enhanced PIN (FW 5.7)' }
            '20ac7a17-c814-4833-93fe-539f0d5e3389' { 'YubiKey 5 Nano Enterprise (FW 5.7)' }
            '4599062e-6926-4fe7-9566-9e8fb1aedaa0' { 'YubiKey 5 Nano Enterprise (FW 5.7)' }
            'b90e7dc1-316e-4fee-a25a-56a666a670fe' { 'YubiKey 5Ci Enterprise (FW 5.7)' }
            '3b24bf49-1d45-4484-a917-13175df0867b' { 'YubiKey 5Ci Enterprise (FW 5.7)' }
            
            default { 'Hardware Key (Unknown)' }
        }
        
        if ($model -and $displayName) {
            $passkeyDetails += "$displayName ($model) - $deviceType"
            $passkeyModels += $model
        } elseif ($model) {
            $passkeyDetails += "$model - $deviceType"
            $passkeyModels += $model
        } else {
            $passkeyDetails += "Passkey (ID: $($passkey.Id.Substring(0,10))...) - $deviceType"
            $passkeyModels += "Unknown Model"
        }
    }

    $results += [PSCustomObject]@{
        DisplayName      = $user.DisplayName
        UserPrincipalName= $user.UserPrincipalName
        Department       = $user.Department
        CompanyName      = $user.CompanyName
        PasskeyCount     = $passkeys.Count
        PasskeyDetails   = $passkeyDetails -join '; '
        PasskeyModels    = $passkeyModels -join '; '
    }
}

# Export to CSV with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$filename = "c:\temp\EntraUsersWithPasskeys_$timestamp.csv"
$results | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8 -Delimiter ";"

Write-Output "Export complete: $filename"
Write-Output "Found $($results.Count) users with $($results | Measure-Object PasskeyCount -Sum | Select-Object -ExpandProperty Sum) total passkeys"