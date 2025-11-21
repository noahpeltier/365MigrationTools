New-MigrationProject -Name "PmP to Spectrum tenant" -Path "$PWD/Projects"

$SourceScopes = @(
    'Directory.Read.All',
    'Policy.Read.All',
    'IdentityProvider.Read.All',
    'Organization.Read.All',
    'User.Read.All',
    'EntitlementManagement.Read.All',
    'UserAuthenticationMethod.Read.All',
    'IdentityUserFlow.Read.All',
    'APIConnectors.Read.All',
    'AccessReview.Read.All',
    'Agreement.Read.All',
    'Policy.Read.PermissionGrant',
    'RoleEligibilitySchedule.Read.Directory',
    'PrivilegedEligibilitySchedule.Read.AzureADGroup',
    'Application.Read.All',
    'OnPremDirectorySynchronization.Read.All',
    'Teamwork.Read.All', 
    'TeamworkAppSettings.Read.All', 
    'SharepointTenantSettings.Read.All',
    'Reports.Read.All',
    'RoleManagement.Read.All',
    'AuditLog.Read.All',
    'Device.Read.All'
)

$TargetScopes = @(
    'LicenseAssignment.ReadWrite.All',
    'Directory.Read.All',
    'Policy.Read.All',
    'IdentityProvider.Read.All',
    'Organization.Read.All',
    'User.ReadWrite.All',
    'Group.ReadWrite.All',
    'EntitlementManagement.Read.All',
    'UserAuthenticationMethod.Read.All',
    'IdentityUserFlow.Read.All',
    'APIConnectors.Read.All',
    'AccessReview.Read.All',
    'Agreement.Read.All',
    'Policy.Read.PermissionGrant',
    'RoleEligibilitySchedule.Read.Directory',
    'PrivilegedEligibilitySchedule.Read.AzureADGroup',
    'Application.Read.All',
    'OnPremDirectorySynchronization.Read.All',
    'Teamwork.Read.All', 
    'TeamworkAppSettings.ReadWrite.All', 
    'SharepointTenantSettings.Read.All',
    'RoleManagement.Read.All',
    'AuditLog.Read.All',
    'Reports.Read.All',
    'Device.Read.All'
)

Add-ProjectSourceTenant -TenantId 29aa692a-96ad-47b1-b073-9fd3d96cd3c2 -AppDisplayName SpetrumMigrationApp -Permissions $SourceScopes -EnableExchangeManageAsApp
Add-ProjectTargetTenant -TenantId a924863a-a67c-44b5-b1af-9735ccf85acb -AppDisplayName SpetrumMigrationApp -Permissions $TargetScopes -EnableExchangeManageAsApp