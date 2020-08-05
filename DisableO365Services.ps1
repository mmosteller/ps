#################################################################################################################################
# I take no responsibility for use of this script breaking your setup. You are accountable for your own account administration. #
#################################################################################################################################

# Connect-MsolService must be run before the script will work

# Enable to viewing changes [to be made] after the script runs
# It's possible for a user to show up multiple times if they have multiple licenses that are parsed
$ViewChanges = $true
# clear data if need be
if($ViewChanges){ $OutData = @() }

# prevents actual changes from being made
$TestMode = $true

# Change to $false to make the script enable services instead
$Disable = $true

# Prevent issues later in the script
if(($Disable -ne $true) -and ($Disable -ne $false)){
    Write-Warning -Message "Disable must be true or false"
    exit
}

# Initialize users array
$Users = @()

# Include E3 users
#$Users += Get-MsolUser -all | Where-Object{ $_.Licenses.AccountSkuId -match "ENTERPRISEPACK" }

# Include E1 users
#$Users += Get-MsolUser -all | Where-Object{ $_.Licenses.AccountSkuId -match "STANDARDPACK" }

# Include Business Premium users
#$Users += Get-MsolUser -all | Where-Object{ $_.Licenses.AccountSkuId -match "O365_BUSINESS_PREMIUM" }

foreach ($User in $Users) {

    # Removes other liceses such as Visio from array
    $Licenses = $User.Licenses | where { ($_.AccountSkuId -match "ENTERPRISEPACK") -or ($_.AccountSkuId -match "STANDARDPACK") -or ($_.Licenses.AccountSkuId -match "O365_BUSINESS_PREMIUM") }

    foreach($License in $Licenses) {

        # Get a list of services already disabled.
        $DisabledServices = $License.ServiceStatus | Where-Object { $_.ProvisioningStatus -eq "Disabled" }

        # Swap to using this line if you wish to get services in the PendingActivation state as well.
        #$DisabledServices = $License.ServiceStatus | Where-Object { ($_.ProvisioningStatus -ne "Success") -and ($_.ServicePlan -notlike "INTUNE*") }

        # Convert to string array
        $DisabledServices = @($DisabledServices.ServicePlan.ServiceName)

        # Initialize variable
        $ModifyList = @()

        # Select E3 services to modify
        if($License.AccountSkuId -match "ENTERPRISEPACK"){
            
            # To-Do (Plan 2)
            #if($ModifyList -notcontains "BPOS_S_TODO_2"){ $ModifyList += "BPOS_S_TODO_2" }

            # Forms (Plan E3)
            #if($ModifyList -notcontains "FORMS_PLAN_E3"){ $ModifyList += "FORMS_PLAN_E3" }

            # Stream for O365
            #if($ModifyList -notcontains "STREAM_O365_E3"){ $ModifyList += "STREAM_O365_E3" }

            # StaffHub
            #if($ModifyList -notcontains "Deskless"){ $ModifyList += "Deskless" }

            # Flow for Office 365
            #if($ModifyList -notcontains "FLOW_O365_P2"){ $ModifyList += "FLOW_O365_P2" }
            
            # PowerApps for Office 365
            #if($ModifyList -notcontains "POWERAPPS_O365_P2"){ $ModifyList += "POWERAPPS_O365_P2" }
            
            # Teams
            #if($ModifyList -notcontains "TEAMS1"){ $ModifyList += "TEAMS1" }

            # Planner
            #if($ModifyList -notcontains "PROJECTWORKMANAGEMENT"){ $ModifyList += "PROJECTWORKMANAGEMENT" }

            # Sway
            #if($ModifyList -notcontains "SWAY"){ $ModifyList += "SWAY" }

            ##############################################################################
            # CANNOT Currently SET                                                       #
            # InTune / Mobile Device Management for Office 365                           #
            #if($ModifyList -notcontains "INTUNE_O365"){ $ModifyList += "INTUNE_O365" }  #
            ##############################################################################

            # Yammer Enterprise
            #if($ModifyList -notcontains "YAMMER_ENTERPRISE"){ $ModifyList += "YAMMER_ENTERPRISE" }

            # Azure Rights Management
            #if($ModifyList -notcontains "RMS_S_ENTERPRISE"){ $ModifyList += "RMS_S_ENTERPRISE" }

            # Office 365 ProPlus
            #if($ModifyList -notcontains "OFFICESUBSCRIPTION"){ $ModifyList += "OFFICESUBSCRIPTION" }

            # Skype for Business Online (Plan 2)
            #if($ModifyList -notcontains "MCOSTANDARD"){ $ModifyList += "MCOSTANDARD" }

            # Office Online
            #if($ModifyList -notcontains "SHAREPOINTWAC"){ $ModifyList += "SHAREPOINTWAC" }

            # SharePoint Online (Plan 2)
            #if($ModifyList -notcontains "SHAREPOINTENTERPRISE"){ $ModifyList += "SHAREPOINTENTERPRISE" }

            # Exchange Online (Plan 2)
            #if($ModifyList -notcontains "EXCHANGE_S_ENTERPRISE"){ $ModifyList += "EXCHANGE_S_ENTERPRISE" }
            
        }
        # Select E1 services to modify
        elseif($License.AccountSkuId -match "STANDARDPACK"){

            # Office Mobile Apps for Office 365
            #if($ModifyList -notcontains "OFFICEMOBILE_SUBSCRIPTION"){ $ModifyList += "OFFICEMOBILE_SUBSCRIPTION" }

            # To-Do (Plan 1)
            #if($ModifyList -notcontains "BPOS_S_TODO_1"){ $ModifyList += "BPOS_S_TODO_1" }

            # Forms (Plan E1)
            #if($ModifyList -notcontains "FORMS_PLAN_E1"){ $ModifyList += "FORMS_PLAN_E1" }

            # Stream for O365 E1 SKU
            #if($ModifyList -notcontains "STREAM_O365_E1"){ $ModifyList += "STREAM_O365_E1" }

            # StaffHub
            #if($ModifyList -notcontains "Deskless"){ $ModifyList += "Deskless" }

            # Flow for Office 365
            #if($ModifyList -notcontains "FLOW_O365_P1"){ $ModifyList += "FLOW_O365_P1" }
            
            # PowerApps for Office 365
            #if($ModifyList -notcontains "POWERAPPS_O365_P1"){ $ModifyList += "POWERAPPS_O365_P1" }
            
            # Teams
            #if($ModifyList -notcontains "TEAMS1"){ $ModifyList += "TEAMS1" }

            # Office Online
            #if($ModifyList -notcontains "SHAREPOINTWAC"){ $ModifyList += "SHAREPOINTWAC" }

            # Planner
            #if($ModifyList -notcontains "PROJECTWORKMANAGEMENT"){ $ModifyList += "PROJECTWORKMANAGEMENT" }

            # Sway
            #if($ModifyList -notcontains "SWAY"){ $ModifyList += "SWAY" }

            ##############################################################################
            # CANNOT Currently SET                                                       #
            # InTune / Mobile Device Management for Office 365                           #
            #if($ModifyList -notcontains "INTUNE_O365"){ $ModifyList += "INTUNE_O365" }#
            ##############################################################################

            # Yammer Enterprise
            #if($ModifyList -notcontains "YAMMER_ENTERPRISE"){ $ModifyList += "YAMMER_ENTERPRISE" }

            # Skype for Business Online (Plan 2)
            #if($ModifyList -notcontains "MCOSTANDARD"){ $ModifyList += "MCOSTANDARD" }

            # SharePoint Online (Plan 1)
            #if($ModifyList -notcontains "SHAREPOINTSTANDARD"){ $ModifyList += "SHAREPOINTSTANDARD" }

            # Exchange Online (Plan 1)
            #if($ModifyList -notcontains "EXCHANGE_S_STANDARD"){ $ModifyList += "EXCHANGE_S_STANDARD" }

        }
        # Select Business Premium services to modify
        elseif($License.AccountSkuID -match "O365_BUSINESS_PREMIUM"){

            # Stream
            #if($ModifyList -notcontains "STREAM_O365_SMB"){ $ModifyList += "STREAM_O365_SMB" }

            # StaffHub
            #if($ModifyList -notcontains "Deskless"){ $ModifyList += "Deskless" }
            
            # To-Do (Plan 1)
            #if($ModifyList -notcontains "BPOS_S_TODO_1"){ $ModifyList += "BPOS_S_TODO1" }

            # Bookings
            #if($ModifyList -notcontains "MICROSOFTBOOKINGS"){ $ModifyList += "MICROSOFTBOOKINGS" }

            # Forms (Plan E1)
            #if($ModifyList -notcontains "FORMS_PLAN_E1"){ $ModifyList += "FORMS_PLAN_E1" }

            # Flow for Office 365
            #if($ModifyList -notcontains "FLOW_O365_P1"){ $ModifyList += "FLOW_O365_P1" }

            # PowerApps for Office 365
            #if($ModifyList -notcontains "POWERAPPS_O365_P1"){ $ModifyList += "POWERAPPS_O365_P1" }

            # Outlook Customer Manager
            #if($ModifyList -notcontains "O365_SB_RelationshipManagement"){ $ModifyList += "O365_SB_Relationship_Management" }

            # Teams
            #if($ModifyList -notcontains "TEAMS1"){ $ModifyList += "TEAMS1" }

            # Planner
            #if($ModifyList -notcontains "PROJECTWORKMANAGEMENT"){ $ModifyList += "PROJECTWORKMANAGEMENT" }

            # Sway
            #if($ModifyList -notcontains "SWAY"){ $ModifyList += "SWAY" }

            ####################################################################
            # CANNOT Currently SET                                             #
            # InTune / Mobile Device Management for Office 365                 #
            #if($ModifyList -notcontains "INTUNE"){ $ModifyList += "INTUNE" }#
            ####################################################################

            # Office Online
            #if($ModifyList -notcontains "SHAREPOINTWAC"){ $ModifyList += "SHAREPOINTWAC" }

            # Office 365 Business Desktop Applications
            #if($ModifyList -notcontains "OFFICE_BUSINESS"){ $ModifyList += "OFFICE_BUSINESS" }

            # Yammer Enterprise
            #if($ModifyList -notcontains "YAMMER_ENTERPRISE"){ $ModifyList += "YAMMER_ENTERPRISE" }

            # Exchange Online (Plan 1)
            #if($ModifyList -notcontains "EXCHANGE_S_STANDARD"){ $ModifyList += "EXCHANGE_S_STANDARD" }

            # Skype for Business Online (Plan 2)
            #if($ModifyList -notcontains "MCOSTANDARD"){ $ModifyList += "MCOSTANDARD" }

            # SharePoint Online (Plan 1)
            #if($ModifyList -notcontains "SHAREPOINTSTANDARD"){ $ModifyList += "SHAREPOINTSTANDARD" }
        }

        # Modify the disabled services array to add/remove items based on disable/enable goal.
        if($Disable){            
            # Prevent visually listing items that are already disabled
            $ModifyList = $ModifyList | Where-Object{$DisabledServices -notcontains $_}

            ############################################
            #Leave one of the lines below commented out#
            ############################################
            # Add to the list of disabled items
            $DisPlan = $DisabledServices
            $DisPlan += $ModifyList | Where-Object{$DisabledServices -notcontains $_}
            # Reset all disabled services and only disable selected above
            #$DisPlan = $ModifyList
        } elseif($Disable -eq $false) {
            $ModifyList = $ModifyList | Where-Object{$DisabledServices -contains $_}
            $DisPlan = $DisabledServices | Where-Object{$ModifyList -notcontains $_}            
        }
        
        # Build changes array
        if($ViewChanges){
            $FormatEnumerationLimit=-1
            $EnabledServices = $License.ServiceStatus | where{$_.ProvisioningStatus -ne "Disabled"}
            $info = New-Object PSObject
            $info | Add-Member -Name "Display Name" -membertype NoteProperty -Value $User.DisplayName
            $info | Add-Member -name "UserPrincipalName" -MemberType NoteProperty -Value $user.UserPrincipalName
            $info | Add-Member -name "Account SKU Id" -MemberType NoteProperty -Value $License.AccountSkuId
            if($Disable){$modType = "Disable"}
            elseif($Disable -eq $false){$modType = "Enable"}
            $info | Add-Member -name "Service Modify Type" -MemberType NoteProperty -Value $modType
            $info | Add-Member -name "Modified Services" -MemberType NoteProperty -Value $ModifyList
            $info | Add-Member -Name "Disabled After Script" -MemberType NoteProperty -Value $DisPlan
            $outdata += $info
        }

        # Create new license options and apply to user account
        if($TestMode -eq $false){
            $LO = New-MsolLicenseOptions -AccountSkuId $License.AccountSkuId -DisabledPlans $DisPlan
            Set-MsolUserLicense -UserPrincipalName $User.UserPrincipalName -LicenseOptions $LO
        }

    }
    
}

if($ViewChanges){ $outdata | Sort-Object "Display Name" | Out-GridView -Title "Service Changes" }