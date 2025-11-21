#Set-StrictMode -Version Latest
#$ErrorActionPreference = 'Stop'
#
# Default location where profile/config data is stored; overridable per project
$script:MgProfileDirectory = [Environment]::GetEnvironmentVariable('MG_PROFILE_DIRECTORY')
if (-not $script:MgProfileDirectory) {
	$script:MgProfileDirectory = [System.IO.Path]::Combine(
		[Environment]::GetFolderPath('ApplicationData'),
		'MgCachedProfiles'
	)
}
$script:MgConfigPath = [Environment]::GetEnvironmentVariable('MG_CONFIG_PATH')
if (-not $script:MgConfigPath) {
	$script:MgConfigPath = Join-Path $PSScriptRoot 'Config.json'
}
$script:ProductSheetDefaultPath = Join-Path $PSScriptRoot 'downloads/Product names and service plan identifiers for licensing.csv'
if (Test-Path -LiteralPath $script:ProductSheetDefaultPath) {
	$script:MicrosoftProductSheetPath = (Resolve-Path -LiteralPath $script:ProductSheetDefaultPath).Path
}
$script:MicrosoftProductSheetCache = $null

function Get-MicrosoftProductSheet {
	[CmdletBinding()]
	param(
		[string]$DestinationPath = $script:ProductSheetDefaultPath,
		[string]$Uri = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
	)

	$destDir = Split-Path -Parent $DestinationPath
	if (-not (Test-Path -LiteralPath $destDir)) {
		New-Item -ItemType Directory -Path $destDir -Force | Out-Null
	}
	Invoke-WebRequest -Uri $Uri -OutFile $DestinationPath -UseBasicParsing -ErrorAction Stop | Out-Null
	$resolved = (Resolve-Path -LiteralPath $DestinationPath).Path
	$script:MicrosoftProductSheetPath = $resolved
	$script:MicrosoftProductSheetCache = $null
	return $resolved
}

function Import-MicrosoftProductSheet {
	[CmdletBinding()]
	param(
		[string]$Path
	)

	if (-not $Path) {
		if ($script:MicrosoftProductSheetPath -and (Test-Path -LiteralPath $script:MicrosoftProductSheetPath)) {
			$Path = $script:MicrosoftProductSheetPath
		}
		elseif (Test-Path -LiteralPath $script:ProductSheetDefaultPath) {
			$Path = $script:ProductSheetDefaultPath
		}
		else {
			throw "Product sheet not found. Run Get-MicrosoftProductSheet first to download the CSV."
		}
	}
	if (-not (Test-Path -LiteralPath $Path)) {
		throw "Product sheet path '$Path' not found. Run Get-MicrosoftProductSheet to download it."
	}
	$resolved = (Resolve-Path -LiteralPath $Path).Path
	$data = Import-Csv -LiteralPath $resolved | ForEach-Object {
		$pdn = $_.Product_Display_Name
		if (-not $pdn -and $_.PSObject.Properties.Name -contains '﻿Product_Display_Name') {
			$pdn = $_.'﻿Product_Display_Name'
		}
		[pscustomobject]@{
			ProductDisplayName            = $pdn
			StringId                      = $_.String_Id
			SkuId                         = $_.GUID
			ServicePlanName               = $_.Service_Plan_Name
			ServicePlanId                 = $_.Service_Plan_Id
			ServicePlansIncludedFriendly  = $_.Service_Plans_Included_Friendly_Names
		}
	}
	$script:MicrosoftProductSheetPath = $resolved
	$script:MicrosoftProductSheetCache = $data
	return $data
}

function Get-MSProduct {
	[CmdletBinding()]
	param(
		[string]$SkuId,
		[string]$StringId,
		[string]$Name,
		[string]$ServicePlanId,
		[string]$ServicePlanName,
		[switch]$Reload
	)

	if ($Reload -or -not $script:MicrosoftProductSheetCache) {
		Import-MicrosoftProductSheet | Out-Null
	}
	$rows = $script:MicrosoftProductSheetCache

	if ($SkuId) {
		$sku = $SkuId.Trim().ToLower()
		$rows = $rows | Where-Object { ($_.SkuId ?? '').ToLower() -eq $sku }
	}
	if ($StringId) {
		$needle = $StringId.Trim()
		$rows = $rows | Where-Object { $_.StringId -like "*$needle*" }
	}
	if ($Name) {
		$needle = $Name.Trim()
		$rows = $rows | Where-Object {
			($_.ProductDisplayName -like "*$needle*") -or
			($_.ServicePlansIncludedFriendly -like "*$needle*")
		}
	}
	if ($ServicePlanId) {
		$plan = $ServicePlanId.Trim().ToLower()
		$rows = $rows | Where-Object { ($_.ServicePlanId ?? '').ToLower() -eq $plan }
	}
	if ($ServicePlanName) {
		$needle = $ServicePlanName.Trim()
		$rows = $rows | Where-Object { $_.ServicePlanName -like "*$needle*" }
	}
	return $rows
}

function Get-MicrosoftOfficeProduct {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$SkuId
	)

	return Get-MSProduct -SkuId $SkuId | Select-Object -First 1
}

function Set-MgProfileDirectory {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Path,
		[switch]$PersistToProcess
	)

	if (-not (Test-Path -LiteralPath $Path)) {
		New-Item -ItemType Directory -Path $Path -Force | Out-Null
	}
	$resolved = (Resolve-Path -LiteralPath $Path).Path
	$script:MgProfileDirectory = $resolved
	if ($PersistToProcess) {
		[Environment]::SetEnvironmentVariable('MG_PROFILE_DIRECTORY', $resolved, 'Process')
	}
	Write-Verbose "Profile directory set to $resolved"
}

function Set-MgConfigPath {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Path,
		[switch]$PersistToProcess
	)

	$created = $false
	if (-not (Test-Path -LiteralPath $Path)) {
		New-Item -ItemType File -Path $Path -Force | Out-Null
		$created = $true
	}
	$resolved = (Resolve-Path -LiteralPath $Path).Path
	if ($created -or [string]::IsNullOrWhiteSpace((Get-Content -LiteralPath $resolved -Raw))) {
		'{}' | Set-Content -Encoding UTF8 -Path $resolved
	}
	$script:MgConfigPath = $resolved
	if ($PersistToProcess) {
		[Environment]::SetEnvironmentVariable('MG_CONFIG_PATH', $resolved, 'Process')
	}
	Write-Verbose "Config path set to $resolved"
}

function Get-MigrationProjectContext {
	[CmdletBinding()]
	param(
		[string]$ProjectPath = (Get-Location).Path
	)

	if (-not (Test-Path -LiteralPath $ProjectPath)) {
		throw "Project path '$ProjectPath' does not exist."
	}
	$resolvedProject = (Resolve-Path -LiteralPath $ProjectPath).Path
	$configPath = Join-Path $resolvedProject 'Config.json'
	if (-not (Test-Path -LiteralPath $configPath)) {
		throw "Config.json not found at $resolvedProject. Run New-MigrationProject first."
	}
	$profilesPath = Join-Path $resolvedProject 'profiles'
	if (-not (Test-Path -LiteralPath $profilesPath)) {
		New-Item -ItemType Directory -Path $profilesPath -Force | Out-Null
	}
	$raw = Get-Content -LiteralPath $configPath -Raw
	if ([string]::IsNullOrWhiteSpace($raw)) {
		$config = [pscustomobject]@{}
	}
	else {
		$config = $raw | ConvertFrom-Json -Depth 10
	}
	$projectName = $config.projectName
	if (-not $projectName) {
		$projectName = Split-Path -Leaf $resolvedProject
	}
	[pscustomobject]@{
		ProjectPath  = $resolvedProject
		ConfigPath   = $configPath
		ProfilesPath = $profilesPath
		Config       = $config
		ProjectName  = $projectName
	}
}

function New-MgCustomApp {
	[CmdletBinding(SupportsShouldProcess)]
	[OutputType([pscustomobject])]
	param (
		[Parameter(Mandatory)]
		[string]$Name,
		[Parameter(Mandatory)]
		[ValidateSet('Certificate', 'ClientCredential')]
		[string]$AuthType,
		[Parameter(Mandatory)]
		[string[]]$Permissions,
		[string]$TenantId,
		[string]$CompanyName,
		[string]$DefaultDomain,
		[string]$ProjectDescription,
		[string]$ProfileName,
		[string]$ProfileDirectory,
		[string]$ConfigPath,
		[string]$ConfigTenantKey,
		[int]$CertYears = 2,
		[int]$SecretMonths = 12,
		[switch]$EnableExchangeManageAsApp
	)

	$previousErrorAction = $ErrorActionPreference
	$ErrorActionPreference = 'Stop'
	try {
		if (-not $ProfileDirectory) {
			$envProfileDir = [Environment]::GetEnvironmentVariable('MG_PROFILE_DIRECTORY')
			if ($envProfileDir) {
				$ProfileDirectory = $envProfileDir
			}
			else {
				$ProfileDirectory = [System.IO.Path]::Combine([Environment]::GetFolderPath('ApplicationData'), 'MgCachedProfiles')
			}
		}
		if (-not $ConfigPath) {
			$envConfigPath = [Environment]::GetEnvironmentVariable('MG_CONFIG_PATH')
			if ($envConfigPath) {
				$ConfigPath = $envConfigPath
			}
		}
		if (-not $ProfileName) {
			$ProfileName = $Name
		}

		$requiredScopes = @(
			'Application.ReadWrite.All',
			'AppRoleAssignment.ReadWrite.All',
			'RoleManagement.ReadWrite.Directory',
			'Organization.Read.All'
		)
		if ($TenantId) {
			Connect-MgGraph -TenantId $TenantId -Scopes $requiredScopes -NoWelcome | Out-Null
		}
		else {
			Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null
			$TenantId = (Get-MgContext).TenantId
		}

		$graphAppId = '00000003-0000-0000-c000-000000000000'
		$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property 'Id,AppRoles'
		if (-not $graphSp) {
			Write-Error "Microsoft Graph service principal not found in tenant $TenantId."
		}

		$graphAppRoles = @($graphSp.AppRoles | Where-Object { $_.IsEnabled -and ($_.AllowedMemberTypes -contains 'Application') })
		$roleIds = foreach ($perm in $Permissions) {
			($graphAppRoles | Where-Object { $_.Value -eq $perm } | Select-Object -ExpandProperty Id -First 1) ?? (Write-Error "Not a Microsoft Graph **application** permission: $perm")
		}

		$requiredResourceAccess = @(
			@{
				resourceAppId  = $graphAppId
				resourceAccess = ($roleIds | ForEach-Object { @{ id = $_; type = 'Role' } })
			}
		)

		$exoAppId = '00000002-0000-0ff1-ce00-000000000000'
		if ($EnableExchangeManageAsApp) {
			$exoSp = Get-MgServicePrincipal -Filter "appId eq '$exoAppId'" -Property 'Id,AppRoles'
			if (-not $exoSp) {
				Write-Error 'Exchange Online service principal not found in this tenant.'
			}
			$exoRole = $exoSp.AppRoles | Where-Object { $_.IsEnabled -and $_.Value -eq 'Exchange.ManageAsApp' } | Select-Object -First 1
			if (-not $exoRole) {
				Write-Error "Couldn't locate the Exchange.ManageAsApp app role on Exchange Online."
			}
			$requiredResourceAccess += @{
				resourceAppId  = $exoAppId
				resourceAccess = @(@{ id = $exoRole.Id; type = 'Role' })
			}
		}

		if ($PSCmdlet.ShouldProcess("App '$Name'", 'Create and configure')) {
			$app = New-MgApplication -DisplayName $Name -RequiredResourceAccess $requiredResourceAccess
			$sp = New-MgServicePrincipal -AppId $app.AppId -DisplayName $Name
			for ($i = 0; $i -lt 10 -and -not $sp.Id; $i++) {
				Start-Sleep -Seconds 2
				$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -Property Id, AppId, DisplayName
			}

			$out = [ordered]@{
				Name               = $Name
				TenantId           = $TenantId
				ClientId           = $app.AppId
				AppObjectId        = $app.Id
				ServicePrincipalId = $sp.Id
				Permissions        = $Permissions
				AuthType           = $AuthType
			}

			$pwd = $null
			$cert = $null
			if ($AuthType -eq 'ClientCredential') {
				$start = (Get-Date).AddMinutes(-5)
				$end = $start.AddMonths($SecretMonths)
				$pwd = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
					displayName   = 'DefaultSecret'
					startDateTime = $start
					endDateTime   = $end
				}
				$out['ClientSecret'] = $pwd.SecretText
				$out['SecretEndsOn'] = $pwd.EndDateTime
				$out['ConnectExample'] = "Connect-MgGraph -TenantId $TenantId -ClientId $($app.AppId) -ClientSecret (ConvertTo-SecureString '$($pwd.SecretText)' -AsPlainText -Force)"
			}
			else {
				$cert = New-SelfSignedCertificate -Subject "CN=$Name" `
					-KeyAlgorithm RSA -KeyLength 2048 -KeySpec KeyExchange `
					-Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' `
					-CertStoreLocation 'Cert:\\CurrentUser\\My' `
					-NotAfter (Get-Date).AddYears($CertYears)

				[byte[]]$keyBytes = $cert.RawData
				[byte[]]$thumbBytes = $cert.GetCertHash()
				$keyCred = @{
					type                = 'AsymmetricX509Cert'
					usage               = 'Verify'
					key                 = $keyBytes
					customKeyIdentifier = $thumbBytes
					displayName         = $cert.Subject
					startDateTime       = $cert.NotBefore
					endDateTime         = $cert.NotAfter
				}
				Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCred) | Out-Null
				$out['CertificateThumbprint'] = $cert.Thumbprint
				$out['CertificateStore'] = 'Cert:\\CurrentUser\\My'
				$out['CertValidTo'] = $cert.NotAfter
				$out['ConnectExample'] = "Connect-MgGraph -TenantId $TenantId -ClientId $($app.AppId) -CertificateThumbprint $($cert.Thumbprint)"
			}

			foreach ($rid in $roleIds) {
				New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $rid | Out-Null
			}

			if ($EnableExchangeManageAsApp) {
				New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $exoSp.Id -AppRoleId $exoRole.Id | Out-Null
				$existing = (Get-MgApplication -ApplicationId $app.Id -Property 'RequiredResourceAccess').RequiredResourceAccess
				if (-not ($existing | Where-Object { $_.resourceAppId -eq $exoAppId })) {
					$existing += @{
						resourceAppId  = $exoAppId
						resourceAccess = @(@{ id = $exoRole.Id; type = 'Role' })
					}
					Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $existing | Out-Null
				}
				foreach ($roleName in @('Exchange Administrator', 'Global Reader')) {
					$def = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$roleName'"
					if ($def) {
						New-MgRoleManagementDirectoryRoleAssignment -BodyParameter @{
							"@odata.type"    = '#microsoft.graph.unifiedRoleAssignment'
							roleDefinitionId = $def.Id
							principalId      = $sp.Id
							directoryScopeId = '/'
						} | Out-Null
					}
				}
				$out['ConnectEXOExample'] = "Connect-ExchangeOnline -CertificateThumbprint $($out.CertificateThumbprint) -AppId $($app.AppId) -Organization '<tenant>.onmicrosoft.com'"
			}

			$profilePath = $null
			if ($ProfileDirectory) {
				$profile = [ordered]@{
					profileName        = $ProfileName
					name               = $Name
					tenantId           = $TenantId
					clientId           = $app.AppId
					permissions        = $Permissions
					companyName        = $CompanyName
					defaultDomain      = $DefaultDomain
					projectDescription = $ProjectDescription
					authType           = $AuthType
					createdOnUtc       = (Get-Date).ToUniversalTime().ToString('o')
				}
				if ($AuthType -eq 'ClientCredential' -and $pwd -and $pwd.SecretText) {
					$secure = ConvertTo-SecureString $pwd.SecretText -AsPlainText -Force
					$profile['clientSecret'] = $secure | ConvertFrom-SecureString
					$profile['secretEncrypted'] = $true
					$profile['secretEndsOn'] = $pwd.EndDateTime
				}
				if ($cert) {
					$profile['certificateThumbprint'] = $cert.Thumbprint
					$profile['certificateStore'] = 'Cert:\\CurrentUser\\My'
					$profile['certValidTo'] = $cert.NotAfter
				}
				if (-not (Test-Path -LiteralPath $ProfileDirectory)) {
					New-Item -ItemType Directory -Path $ProfileDirectory -Force | Out-Null
				}
				$profilePath = Join-Path $ProfileDirectory ("$ProfileName.json")
				$profile | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -Path $profilePath
				$out['ProfilePath'] = $profilePath
			}

			if ($ConfigPath) {
				if (-not (Test-Path -LiteralPath $ConfigPath)) {
					$emptyConfig = [ordered]@{}
					$emptyConfig | ConvertTo-Json -Depth 2 | Set-Content -Encoding UTF8 -Path $ConfigPath
				}
				$rawConfig = Get-Content -LiteralPath $ConfigPath -Raw
				if ([string]::IsNullOrWhiteSpace($rawConfig)) {
					$configObj = [pscustomobject]@{}
				}
				else {
					$configObj = $rawConfig | ConvertFrom-Json -Depth 10
				}
				$tenantKey = $ConfigTenantKey
				if (-not $tenantKey) {
					$tenantKey = $ProfileName
				}
				$tenantConfig = [ordered]@{
					clientId           = $app.AppId
					tenantId           = $TenantId
					authType           = $AuthType
					companyName        = $CompanyName
					defaultDomain      = $DefaultDomain
					projectDescription = $ProjectDescription
					profilePath        = $profilePath
					permissions        = $Permissions
				}
				if ($profile['clientSecret']) {
					$tenantConfig['clientSecret'] = $profile['clientSecret']
					$tenantConfig['secretEncrypted'] = $true
					$tenantConfig['secretEndsOn'] = $profile['secretEndsOn']
				}
				if ($cert) {
					$tenantConfig['certificateThumbprint'] = $cert.Thumbprint
					$tenantConfig['certificateStore'] = 'Cert:\\CurrentUser\\My'
					$tenantConfig['certValidTo'] = $cert.NotAfter
				}
				if ($configObj -isnot [pscustomobject]) {
					$configObj = [pscustomobject]@{}
				}
				$configObj | Add-Member -NotePropertyName $tenantKey -NotePropertyValue ([pscustomobject]$tenantConfig) -Force
				$configObj | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -Path $ConfigPath
				$out['ConfigPath'] = $ConfigPath
				$out['ConfigTenantKey'] = $tenantKey
			}

			return [pscustomobject]$out
		}
	}
	finally {
		$ErrorActionPreference = $previousErrorAction
	}
}

function Set-ActiveMigrationProject {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$ProjectPath
	)

	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	$script:ActiveProjectContext = $context
	Set-MgProfileDirectory -Path $context.ProfilesPath -PersistToProcess
	Set-MgConfigPath -Path $context.ConfigPath -PersistToProcess
	$env:MIGRATION_PROJECT_NAME = $context.ProjectName
	Write-Host "Active migration project set to '$($context.ProjectName)'" -ForegroundColor Cyan
	return $context
}

function Clear-ActiveMigrationProject {
	[CmdletBinding()]
	param()

	$script:ActiveProjectContext = $null
	[Environment]::SetEnvironmentVariable('MG_PROFILE_DIRECTORY', $null, 'Process')
	[Environment]::SetEnvironmentVariable('MG_CONFIG_PATH', $null, 'Process')
	Remove-Item Env:MIGRATION_PROJECT_NAME -ErrorAction SilentlyContinue
	Write-Host "Active migration project cleared" -ForegroundColor Yellow
}

function Get-ActiveMigrationProject {
	[CmdletBinding()]
	param()

	if ($script:ActiveProjectContext) {
		return $script:ActiveProjectContext
	}
	if ($env:MG_CONFIG_PATH) {
		try {
			return Get-MigrationProjectContext -ProjectPath (Split-Path -Parent $env:MG_CONFIG_PATH)
		}
		catch {
		}
	}
	$null
}

function Get-MigrationProject {
	[CmdletBinding()]
	[OutputType([pscustomobject])]
	param(
		[string]$ProjectPath
	)

	$resolvedProject = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $resolvedProject
	$config = $context.Config
	[pscustomobject]@{
		ProjectName        = $context.ProjectName
		ProjectPath        = $context.ProjectPath
		ConfigPath         = $context.ConfigPath
		ProfilesPath       = $context.ProfilesPath
		ProjectDescription = $config.projectDescription
		CreatedOnUtc       = $config.createdOnUtc
		SourceTenant       = $config.sourceTenant
		TargetTenant       = $config.targetTenant
	}
}

function Update-ProjectTenantConfig {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$ConfigPath,
		[Parameter(Mandatory)][string]$TenantKey,
		[Parameter(Mandatory)][hashtable]$Properties
	)

	$raw = Get-Content -LiteralPath $ConfigPath -Raw
	if ([string]::IsNullOrWhiteSpace($raw)) {
		$configObj = [ordered]@{}
	}
	else {
		$configObj = $raw | ConvertFrom-Json -Depth 10
	}
	if ($configObj -isnot [pscustomobject]) {
		$configObj = [pscustomobject]@{}
	}
	$tenant = $configObj.$TenantKey
	if (-not $tenant) {
		$tenant = [pscustomobject]@{}
		$configObj | Add-Member -NotePropertyName $TenantKey -NotePropertyValue $tenant -Force
	}
	foreach ($key in $Properties.Keys) {
		$tenant | Add-Member -NotePropertyName $key -NotePropertyValue $Properties[$key] -Force
	}
	$configObj | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -Path $ConfigPath
}

function Import-ProjectTenantConfig {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$SourceConfigPath,
		[Parameter(Mandatory)][string]$DestinationConfigPath,
		[Parameter(Mandatory)][string]$DestinationProfilesPath,
		[string]$SourceTenantKey = 'targetTenant',
		[string]$DestinationTenantKey = 'targetTenant',
		[string]$ProfileName = 'Target'
	)

	if (-not (Test-Path -LiteralPath $SourceConfigPath)) {
		throw "Source config '$SourceConfigPath' not found."
	}
	$resolvedSource = (Resolve-Path -LiteralPath $SourceConfigPath).Path
	$raw = Get-Content -LiteralPath $resolvedSource -Raw
	if ([string]::IsNullOrWhiteSpace($raw)) {
		throw "Source config $resolvedSource is empty."
	}
	$sourceConfig = $raw | ConvertFrom-Json -Depth 10
	$tenant = $sourceConfig.$SourceTenantKey
	if (-not $tenant) {
		throw "Source config $resolvedSource does not contain tenant key '$SourceTenantKey'."
	}

	if (-not (Test-Path -LiteralPath $DestinationConfigPath)) {
		'{}' | Set-Content -Encoding UTF8 -Path $DestinationConfigPath
	}
	if (-not (Test-Path -LiteralPath $DestinationProfilesPath)) {
		New-Item -ItemType Directory -Path $DestinationProfilesPath -Force | Out-Null
	}

	$copiedProfilePath = $null
	if ($tenant.profilePath -and (Test-Path -LiteralPath $tenant.profilePath)) {
		$copiedProfilePath = Join-Path $DestinationProfilesPath ("$ProfileName.json")
		Copy-Item -LiteralPath $tenant.profilePath -Destination $copiedProfilePath -Force
		$copiedProfilePath = (Resolve-Path -LiteralPath $copiedProfilePath).Path
	}

	$properties = [ordered]@{}
	foreach ($prop in $tenant.PSObject.Properties) {
		$properties[$prop.Name] = $prop.Value
	}
	if ($copiedProfilePath) {
		$properties['profilePath'] = $copiedProfilePath
	}

	Update-ProjectTenantConfig -ConfigPath $DestinationConfigPath -TenantKey $DestinationTenantKey -Properties $properties

	[pscustomobject]@{
		SourceConfigPath = $resolvedSource
		ConfigPath       = (Resolve-Path -LiteralPath $DestinationConfigPath).Path
		TenantKey        = $DestinationTenantKey
		ProfilePath      = $copiedProfilePath
	}
}

function Get-MgProfilePath {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$ProfileName,
		[string]$ProfilePath
	)

	if ($ProfilePath) {
		if (Test-Path -LiteralPath $ProfilePath) {
			return (Resolve-Path -LiteralPath $ProfilePath).Path
		}
		return $ProfilePath
	}
	return (Join-Path $script:MgProfileDirectory ("$ProfileName.json"))
}

function Get-MgConfigTenant {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$TenantKey,
		[string]$ConfigPath
	)

	$resolvedPath = $ConfigPath
	if (-not $resolvedPath) {
		$resolvedPath = $script:MgConfigPath
	}
	$candidates = @()
	if ($resolvedPath) {
		$candidates += $resolvedPath
	}
	$cwdCandidate = Join-Path (Get-Location).Path 'Config.json'
	if (-not $candidates.Contains($cwdCandidate)) {
		$candidates += $cwdCandidate
	}
	if ($script:MgConfigPath -and -not $candidates.Contains($script:MgConfigPath)) {
		$candidates += $script:MgConfigPath
	}
	$resolvedPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
	if (-not $resolvedPath) {
		throw "Config file not found. Checked: $($candidates -join ', ')"
	}
	$resolvedPath = (Resolve-Path -LiteralPath $resolvedPath).Path
	$raw = Get-Content -LiteralPath $resolvedPath -Raw
	if ([string]::IsNullOrWhiteSpace($raw)) {
		throw "Config file $resolvedPath is empty."
	}
	$config = $raw | ConvertFrom-Json -Depth 10
	$tenant = $config.$TenantKey
	if (-not $tenant) {
		throw "Tenant key '$TenantKey' not found in config $resolvedPath."
	}
	[pscustomobject]@{
		Tenant     = $tenant
		ConfigPath = (Resolve-Path -LiteralPath $resolvedPath).Path
	}
}

function Connect-MgCachedTenant {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$ProfileName,
		[string]$ProfilePath,
		[switch]$PassThruProfile
	)

	$resolvedPath = Get-MgProfilePath -ProfileName $ProfileName -ProfilePath $ProfilePath
	if (-not (Test-Path -LiteralPath $resolvedPath)) {
		throw "Profile '$ProfileName' not found at $resolvedPath"
	}
	$profile = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
	if (-not $profile.clientId -or -not $profile.tenantId) {
		throw "Profile $resolvedPath is missing clientId or tenantId."
	}

	Write-Verbose "Connecting to tenant $($profile.tenantId) with profile $ProfileName"

	if ($profile.clientSecret) {
		$secureSecret = ConvertTo-SecureString $profile.clientSecret
		Connect-MgGraph -TenantId $profile.tenantId -ClientId $profile.clientId -ClientSecret $secureSecret -NoWelcome | Out-Null
	}
	elseif ($profile.certificateThumbprint) {
		Connect-MgGraph -TenantId $profile.tenantId -ClientId $profile.clientId -CertificateThumbprint $profile.certificateThumbprint -NoWelcome | Out-Null
	}
	else {
		throw "Profile $resolvedPath does not contain a clientSecret or certificateThumbprint to authenticate with."
	}

	if ($PassThruProfile) {
		return $profile
	}
	return Get-MgContext
}

function Connect-MgConfiguredTenant {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$TenantKey,
		[string]$ConfigPath,
		[switch]$PassThruProfile
	)

	$result = Get-MgConfigTenant -TenantKey $TenantKey -ConfigPath $ConfigPath
	$tenant = $result.Tenant
	$tenantId = $tenant.TenantId ?? $tenant.tenantId
	$clientId = $tenant.ClientId ?? $tenant.clientId
	if (-not $tenantId -or -not $clientId) {
		throw "Config entry '$TenantKey' is missing tenantId or clientId."
	}
	Write-Verbose "Connecting via config '$TenantKey' defined in $($result.ConfigPath)"

	if ($tenant.PSObject.Properties['clientSecretPlain'] -and $tenant.clientSecretPlain) {
		$secureSecret = ConvertTo-SecureString $tenant.clientSecretPlain -AsPlainText -Force
		Connect-MgGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $secureSecret -NoWelcome | Out-Null
	}
	elseif ($tenant.PSObject.Properties['clientSecret'] -and $tenant.clientSecret) {
		$secureSecret = ConvertTo-SecureString $tenant.clientSecret
		Connect-MgGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $secureSecret -NoWelcome | Out-Null
	}
	elseif ($tenant.PSObject.Properties['certificateThumbprint'] -and $tenant.certificateThumbprint) {
		Connect-MgGraph -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $tenant.certificateThumbprint -NoWelcome | Out-Null
	}
	else {
		throw "Config entry '$TenantKey' lacks a clientSecret or certificateThumbprint."
	}

	if ($PassThruProfile) {
		return $tenant
	}
	return Get-MgContext
}

function Connect-ExchangeConfiguredTenant {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$TenantKey,
		[string]$ConfigPath
	)

	$result = Get-MgConfigTenant -TenantKey $TenantKey -ConfigPath $ConfigPath
	$tenant = $result.Tenant
	$clientId = $tenant.ClientId ?? $tenant.clientId
	$thumbprint = $tenant.certificateThumbprint ?? $tenant.CertificateThumbprint
	$organization = $tenant.defaultDomain ?? $tenant.DefaultDomain

	if (-not $clientId) { throw "Config entry '$TenantKey' is missing clientId required for Exchange connection." }
	if (-not $thumbprint) { throw "Config entry '$TenantKey' is missing certificateThumbprint required for Exchange ManageAsApp." }
	if (-not $organization) { throw "Config entry '$TenantKey' is missing defaultDomain (e.g., contoso.onmicrosoft.com) required for Exchange connection." }

	Write-Verbose "Connecting to Exchange Online for tenant '$TenantKey' using appId $clientId and certificate $thumbprint"
	Connect-ExchangeOnline -CertificateThumbprint $thumbprint -AppId $clientId -Organization $organization -ShowBanner:$false | Out-Null
	return $organization
}

function Connect-Source {
	[CmdletBinding(DefaultParameterSetName = 'Config')]
	param(
		[Parameter(ParameterSetName = 'Profile')]
		[string]$ProfilePath,
		[Parameter(ParameterSetName = 'Config')]
		[string]$ConfigPath,
		[Parameter(ParameterSetName = 'Config')]
		[string]$TenantKey = 'sourceTenant',
		[switch]$PassThruProfile
	)

	if ($PSCmdlet.ParameterSetName -eq 'Profile') {
		$invokeParams = @{
			ProfileName     = 'Source'
			ProfilePath     = $ProfilePath
			PassThruProfile = $PassThruProfile
		}
		if ($PSBoundParameters.ContainsKey('Verbose')) {
			$invokeParams['Verbose'] = $true
		}
		return Connect-MgCachedTenant @invokeParams
	}

	if (-not $ConfigPath -and $script:ActiveProjectContext) {
		$ConfigPath = $script:ActiveProjectContext.ConfigPath
	}
	$invokeParams = @{
		TenantKey       = $TenantKey
		ConfigPath      = $ConfigPath
		PassThruProfile = $PassThruProfile
	}
	if ($PSBoundParameters.ContainsKey('Verbose')) {
		#$invokeParams['Verbose'] = $true
	}
	Connect-MgConfiguredTenant @invokeParams
}

function Connect-Target {
	[CmdletBinding(DefaultParameterSetName = 'Config')]
	param(
		[Parameter(ParameterSetName = 'Profile')]
		[string]$ProfilePath,
		[Parameter(ParameterSetName = 'Config')]
		[string]$ConfigPath,
		[Parameter(ParameterSetName = 'Config')]
		[string]$TenantKey = 'targetTenant',
		[switch]$PassThruProfile
	)

	if ($PSCmdlet.ParameterSetName -eq 'Profile') {
		$invokeParams = @{
			ProfileName     = 'Target'
			ProfilePath     = $ProfilePath
			PassThruProfile = $PassThruProfile
		}
		if ($PSBoundParameters.ContainsKey('Verbose')) {
			$invokeParams['Verbose'] = $true
		}
		return Connect-MgCachedTenant @invokeParams
	}

	if (-not $ConfigPath -and $script:ActiveProjectContext) {
		$ConfigPath = $script:ActiveProjectContext.ConfigPath
	}
	$invokeParams = @{
		TenantKey       = $TenantKey
		ConfigPath      = $ConfigPath
		PassThruProfile = $PassThruProfile
	}
	if ($PSBoundParameters.ContainsKey('Verbose')) {
		#$invokeParams['Verbose'] = $true
	}
	Connect-MgConfiguredTenant @invokeParams
}

function Connect-ProjectSourceExchange {
	[CmdletBinding()]
	param(
		[string]$ProjectPath
	)

	$ProjectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	Connect-ExchangeConfiguredTenant -TenantKey 'sourceTenant' -ConfigPath $context.ConfigPath
}

function Connect-ProjectTargetExchange {
	[CmdletBinding()]
	param(
		[string]$ProjectPath
	)

	$ProjectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	Connect-ExchangeConfiguredTenant -TenantKey 'targetTenant' -ConfigPath $context.ConfigPath
}

function New-MigrationProject {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$Name,
		[string]$Path = (Get-Location).Path,
		[string]$Description,
		[switch]$Force
	)

	if (-not (Test-Path -LiteralPath $Path)) {
		New-Item -ItemType Directory -Path $Path -Force | Out-Null
	}
	$resolvedBase = (Resolve-Path -LiteralPath $Path).Path
	$sanitized = ($Name -replace '[^0-9A-Za-z _-]', '').Trim()
	if (-not $sanitized) {
		$sanitized = "MigrationProject"
	}
	$folderName = ($sanitized -replace '\s+', '-').Trim('-')
	if (-not $folderName) {
		$folderName = "MigrationProject"
	}
	$projectRoot = Join-Path $resolvedBase $folderName
	if (Test-Path -LiteralPath $projectRoot) {
		if (-not $Force) {
			throw "Project folder '$projectRoot' already exists. Use -Force to reuse."
		}
	}
	else {
		New-Item -ItemType Directory -Path $projectRoot | Out-Null
	}
	$profilesDir = Join-Path $projectRoot 'profiles'
	if (-not (Test-Path -LiteralPath $profilesDir)) {
		New-Item -ItemType Directory -Path $profilesDir | Out-Null
	}
	$configPath = Join-Path $projectRoot 'Config.json'
	if ((-not (Test-Path -LiteralPath $configPath)) -or $Force) {
		$config = [ordered]@{
			projectName        = $Name
			projectDescription = $Description ?? $Name
			createdOnUtc       = (Get-Date).ToUniversalTime().ToString('o')
			sourceTenant       = [ordered]@{}
			targetTenant       = [ordered]@{}
		}
		$config | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -Path $configPath
	}
	Get-MigrationProjectContext -ProjectPath $projectRoot
}

function Export-MigrationProject {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[string]$DestinationPath,
		[System.Security.SecureString]$PfxPassword,
		[switch]$Overwrite
	)

	$context = Get-MigrationProjectContext -ProjectPath (Resolve-ProjectPath $ProjectPath)
	$exportRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("MigrationProjectExport-" + [guid]::NewGuid())
	$pfxDir = Join-Path $exportRoot 'pfx'
	$profilesDir = Join-Path $exportRoot 'profiles'
	$manifestPath = Join-Path $exportRoot 'manifest.json'
	try {
		New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
		Copy-Item -LiteralPath $context.ConfigPath -Destination (Join-Path $exportRoot 'Config.json') -Force
		if (Test-Path -LiteralPath $context.ProfilesPath) {
			New-Item -ItemType Directory -Path $profilesDir -Force | Out-Null
			$profileItems = Get-ChildItem -Path $context.ProfilesPath -Force -ErrorAction SilentlyContinue
			if ($profileItems) {
				Copy-Item -Path ($profileItems | Select-Object -ExpandProperty FullName) -Destination $profilesDir -Recurse -Force
			}
		}

		$thumbprints = @()
		foreach ($tenantKey in @('sourceTenant','targetTenant')) {
			$thumb = $context.Config.$tenantKey.certificateThumbprint ?? $context.Config.$tenantKey.CertificateThumbprint
			if ($thumb) { $thumbprints += $thumb }
		}
		$thumbprints = $thumbprints | Where-Object { $_ } | Select-Object -Unique
		if ($thumbprints.Count -gt 0) {
			if (-not $PfxPassword) {
				$PfxPassword = Read-Host -AsSecureString 'Enter a password to protect exported PFX files'
			}
			New-Item -ItemType Directory -Path $pfxDir -Force | Out-Null
			foreach ($thumb in $thumbprints) {
				$cert = Get-ChildItem -Path "Cert:\\CurrentUser\\My\\$thumb" -ErrorAction SilentlyContinue
				if (-not $cert) {
					throw "Certificate with thumbprint $thumb not found in Cert:\CurrentUser\My"
				}
				$pfxPath = Join-Path $pfxDir ("$thumb.pfx")
				Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $PfxPassword -Force | Out-Null
			}
		}

		$manifest = [ordered]@{
			projectName   = $context.ProjectName
			projectPath   = $context.ProjectPath
			createdOnUtc  = $context.Config.createdOnUtc
			exportedOnUtc = (Get-Date).ToUniversalTime().ToString('o')
			certificates  = $thumbprints
			config        = 'Config.json'
			profiles      = 'profiles'
			pfx           = if ($thumbprints.Count -gt 0) { 'pfx' } else { $null }
		}
		$manifest | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -Path $manifestPath

		$zipName = $context.ProjectName
		if (-not $zipName) { $zipName = Split-Path -Leaf $context.ProjectPath }
		$zipName = ($zipName -replace '[^0-9A-Za-z _-]', '').Trim()
		if (-not $zipName) { $zipName = 'MigrationProject' }
		$zipName = ($zipName -replace '\s+', '-').Trim('-')
		if (-not $DestinationPath) {
			$DestinationPath = Join-Path (Split-Path -Parent $context.ProjectPath) ("$zipName-export.zip")
		}
		if (Test-Path -LiteralPath $DestinationPath) {
			if (-not $Overwrite) {
				throw "Destination $DestinationPath exists. Use -Overwrite to replace."
			}
			Remove-Item -LiteralPath $DestinationPath -Force
		}
		Compress-Archive -Path (Join-Path $exportRoot '*') -DestinationPath $DestinationPath -Force
		return [pscustomobject]@{
			ProjectName    = $context.ProjectName
			ProjectPath    = $context.ProjectPath
			ArchivePath    = (Resolve-Path -LiteralPath $DestinationPath).Path
			CertificatePfx = $thumbprints
		}
	}
	finally {
		if (Test-Path -LiteralPath $exportRoot) {
			Remove-Item -LiteralPath $exportRoot -Recurse -Force -ErrorAction SilentlyContinue
		}
	}
}

function Import-MigrationProject {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$ZipPath,
		[string]$DestinationRoot = (Get-Location).Path,
		[System.Security.SecureString]$PfxPassword,
		[switch]$Force
	)

	if (-not (Test-Path -LiteralPath $ZipPath)) {
		throw "Archive '$ZipPath' not found."
	}
	$resolvedZip = (Resolve-Path -LiteralPath $ZipPath).Path
	$working = Join-Path ([System.IO.Path]::GetTempPath()) ("MigrationProjectImport-" + [guid]::NewGuid())
	try {
		New-Item -ItemType Directory -Path $working -Force | Out-Null
		Expand-Archive -LiteralPath $resolvedZip -DestinationPath $working -Force

		$configPath = Join-Path $working 'Config.json'
		if (-not (Test-Path -LiteralPath $configPath)) {
			$configCandidate = Get-ChildItem -Path $working -Filter 'Config.json' -Recurse -File | Select-Object -First 1
			if (-not $configCandidate) {
				throw "Config.json not found inside archive $resolvedZip."
			}
			$configPath = $configCandidate.FullName
		}
		$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json -Depth 10

		$projectName = $config.projectName
		if (-not $projectName) { $projectName = 'MigrationProject' }
		$sanitized = ($projectName -replace '[^0-9A-Za-z _-]', '').Trim()
		if (-not $sanitized) { $sanitized = 'MigrationProject' }
		$folderName = ($sanitized -replace '\s+', '-').Trim('-')
		if (-not (Test-Path -LiteralPath $DestinationRoot)) {
			New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null
		}
		$destRoot = (Resolve-Path -LiteralPath $DestinationRoot).Path
		$projectRoot = Join-Path $destRoot $folderName
		if (Test-Path -LiteralPath $projectRoot) {
			if (-not $Force) {
				throw "Destination project folder '$projectRoot' already exists. Use -Force to overwrite."
			}
			Remove-Item -LiteralPath $projectRoot -Recurse -Force
		}
		New-Item -ItemType Directory -Path $projectRoot -Force | Out-Null
		$destProfiles = Join-Path $projectRoot 'profiles'
		New-Item -ItemType Directory -Path $destProfiles -Force | Out-Null

		$sourceProfiles = Join-Path (Split-Path -Parent $configPath) 'profiles'
		if (Test-Path -LiteralPath $sourceProfiles) {
			Copy-Item -Path (Join-Path $sourceProfiles '*') -Destination $destProfiles -Recurse -Force
		}
		Copy-Item -LiteralPath $configPath -Destination (Join-Path $projectRoot 'Config.json') -Force
		$configPath = Join-Path $projectRoot 'Config.json'

		foreach ($tenantKey in @('sourceTenant','targetTenant')) {
			$tenant = $config.$tenantKey
			if (-not $tenant) { continue }
			$fileName = if ($tenant.profilePath) { Split-Path -Leaf $tenant.profilePath } else {
				if ($tenantKey -eq 'sourceTenant') { 'Source.json' } elseif ($tenantKey -eq 'targetTenant') { 'Target.json' } else { "$tenantKey.json" }
			}
			$newProfilePath = Join-Path $destProfiles $fileName
			$tenant.profilePath = $newProfilePath
		}
		$config | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -Path $configPath

		$pfxDir = Join-Path (Split-Path -Parent $configPath) 'pfx'
		if (-not (Test-Path -LiteralPath $pfxDir)) {
			$pfxDir = Join-Path $working 'pfx'
		}
		$importedThumbs = @()
		if (Test-Path -LiteralPath $pfxDir) {
			$pfxFiles = Get-ChildItem -LiteralPath $pfxDir -Filter '*.pfx' -File
			if ($pfxFiles.Count -gt 0) {
				if (-not $PfxPassword) {
					$PfxPassword = Read-Host -AsSecureString 'Enter password used to protect exported PFX files'
				}
				foreach ($pfx in $pfxFiles) {
					Import-PfxCertificate -FilePath $pfx.FullName -CertStoreLocation 'Cert:\CurrentUser\My' -Password $PfxPassword -Exportable | Out-Null
					$importedThumbs += ($pfx.BaseName)
				}
			}
		}

		return Get-MigrationProjectContext -ProjectPath $projectRoot
	}
	finally {
		if (Test-Path -LiteralPath $working) {
			Remove-Item -LiteralPath $working -Recurse -Force -ErrorAction SilentlyContinue
		}
	}
}

function Invoke-MigrateUser {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$Source,
		[Parameter(Mandatory)][string]$Target,
		[string]$ProjectPath,
		[string[]]$Licenses,
		[Parameter(Mandatory)][string]$Password,
		[switch]$BlockSignin
	)

	$projectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $projectPath

	$sourceProps = @(
		'displayName','givenName','surname','mailNickname','jobTitle','department',
		'officeLocation','businessPhones','mobilePhone','usageLocation'
	)
	Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	Connect-ProjectSourceTenant -ProjectPath $context.ProjectPath -Silent | Out-Null
	try {
		$srcUser = Get-MgUser -UserId $Source -Property $sourceProps -ErrorAction Stop
	}
	finally {
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	}

	Connect-ProjectTargetTenant -ProjectPath $context.ProjectPath -Silent | Out-Null
	try {
		try {
			$existing = Get-MgUser -UserId $Target -ErrorAction Stop
			if ($existing) { throw "User '$Target' already exists in target tenant." }
		}
		catch {
			$message = $_.Exception.Message
			if ($message -notmatch 'Request_ResourceNotFound' -and $message -notmatch 'ResourceNotFound' -and $message -notmatch '404') {
				throw
			}
		}

		$body = [ordered]@{
			userPrincipalName = $Target
			accountEnabled    = (-not $BlockSignin)
			passwordProfile   = @{
				password = $Password
				forceChangePasswordNextSignIn = $true
			}
		}
		if ($srcUser.DisplayName) { $body['displayName'] = $srcUser.DisplayName }
		if ($srcUser.MailNickname) { $body['mailNickname'] = $srcUser.MailNickname } else {
			$local = ($Target -split '@')[0]
			$body['mailNickname'] = $local
		}
		if ($srcUser.GivenName) { $body['givenName'] = $srcUser.GivenName }
		if ($srcUser.Surname) { $body['surname'] = $srcUser.Surname }
		if ($srcUser.JobTitle) { $body['jobTitle'] = $srcUser.JobTitle }
		if ($srcUser.Department) { $body['department'] = $srcUser.Department }
		if ($srcUser.OfficeLocation) { $body['officeLocation'] = $srcUser.OfficeLocation }
		if ($srcUser.MobilePhone) { $body['mobilePhone'] = $srcUser.MobilePhone }
		if ($srcUser.BusinessPhones) { $body['businessPhones'] = @($srcUser.BusinessPhones | Where-Object { $_ }) }
		if ($srcUser.UsageLocation) { $body['usageLocation'] = $srcUser.UsageLocation }
		elseif ($Licenses -and $Licenses.Count -gt 0) { $body['usageLocation'] = 'US' }

		$newUser = New-MgUser -BodyParameter $body -ErrorAction Stop

		if ($Licenses -and $Licenses.Count -gt 0) {
			$licenseAdd = @()
			foreach ($sku in $Licenses) {
				$licenseAdd += @{ skuId = $sku }
			}
			Set-MgUserLicense -UserId $newUser.Id -AddLicenses $licenseAdd -ErrorAction Stop | Out-Null
		}

		return [pscustomobject]@{
			SourceUPN = $Source
			TargetUPN = $Target
			UserId    = $newUser.Id
			Licenses  = $Licenses
			AccountEnabled = (-not $BlockSignin)
		}
	}
	finally {
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	}
}

function Invoke-DiscoverSourceAccounts {
    [CmdletBinding()]
    param(
        [string]$ProjectPath,
        [string]$OutputPath,
        [switch]$UseSqlite
    )

    $projectPath = Resolve-ProjectPath $ProjectPath
    $context = Get-MigrationProjectContext -ProjectPath $projectPath
    if (-not $OutputPath) {
        $OutputPath = Join-Path $context.ProjectPath 'DiscoveredAccounts.csv'
    }

    $dbPath = $null
    $taskId = [guid]::NewGuid().ToString()
    $statusUpdated = $false
    if ($UseSqlite) {
        $dbPath = Get-MigrationDatabase -ProjectPath $context.ProjectPath
        $nowIso = (Get-Date).ToUniversalTime().ToString('o')
        Invoke-MySQLiteQuery -Path $dbPath -Query (@"
INSERT OR REPLACE INTO Tasks (taskId, name, total, processed, status, message, startedOnUtc, updatedOnUtc)
VALUES ('{0}', 'DiscoverSourceAccounts', 0, 0, 'Running', '', '{1}', '{1}');
"@ -f $taskId, $nowIso) | Out-Null
    }

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-ProjectSourceTenant -ProjectPath $context.ProjectPath -Silent | Out-Null
        $users = $null
        try {
            $users = Get-MgUser -All -Property 'id','displayName','userPrincipalName','mail','proxyAddresses','assignedLicenses' -ErrorAction Stop |
                Select-Object 'id','displayName','userPrincipalName','mail','proxyAddresses','assignedLicenses'
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $total = @($users).Count
        if ($UseSqlite -and $dbPath) {
            $nowIso = (Get-Date).ToUniversalTime().ToString('o')
            Invoke-MySQLiteQuery -Path $dbPath -Query (@"
UPDATE Tasks SET total = {0}, updatedOnUtc = '{1}' WHERE taskId = '{2}';
"@ -f $total, $nowIso, $taskId) | Out-Null
        }

        $mailboxMap = @{}
        try {
            Connect-ProjectSourceExchange -ProjectPath $context.ProjectPath | Out-Null
            try {
                $recipients = Get-EXORecipient -ResultSize Unlimited -PropertySets All -ErrorAction Stop
                foreach ($recip in $recipients) {
                    if ($recip.ExternalDirectoryObjectId) {
                        $mailboxMap[$recip.ExternalDirectoryObjectId] = $recip.RecipientTypeDetails
                    }
                }
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch {
            Write-Warning "Exchange discovery failed: $($_.Exception.Message). Mailbox type will be blank."
        }

        $rows = foreach ($u in $users) {
            $type = $null
            if ($mailboxMap.ContainsKey($u.Id)) {
                $type = $mailboxMap[$u.Id]
            }
            elseif ($u.Mail -or ($u.ProxyAddresses -and $u.ProxyAddresses.Count -gt 0)) {
                $type = 'MailEnabledUser'
            }
            else {
                $type = 'NonMailEnabledUser'
            }

            $assignedSkus = @()
            if ($u.AssignedLicenses) {
                $assignedSkus = @($u.AssignedLicenses | ForEach-Object { (Get-MicrosoftOfficeProduct -SkuId $_.SkuId).ProductDisplayName })
            }
            $AssignedLicenses = ($assignedSkus -join ",")

            if ($UseSqlite -and $dbPath) {
                $nowLoop = (Get-Date).ToUniversalTime().ToString('o')
                $insertAccount = @"
INSERT OR REPLACE INTO AccountsSource (objectId, sourceUpn, displayName, mail, proxyAddresses, assignedLicenses, mailboxType, usageLocation, lastSeenUtc)
VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}');
"@ -f ($u.Id -replace "'","''"), ($u.UserPrincipalName -replace "'","''"), ($u.DisplayName -replace "'","''"), ($u.Mail -replace "'","''"), (($u.ProxyAddresses -join ',') -replace "'","''"), ($AssignedLicenses -replace "'","''"), ($type -replace "'","''"), (($u.UsageLocation) -replace "'","''"), $nowLoop
                Invoke-MySQLiteQuery -Path $dbPath -Query $insertAccount | Out-Null
                if ($mailboxMap.ContainsKey($u.Id)) {
                    $mbType = $mailboxMap[$u.Id]
                    $insertMailbox = @"
INSERT OR REPLACE INTO MailboxesSource (objectId, recipientType, mailboxType, lastSeenUtc)
VALUES ('{0}', '{1}', '{2}', '{3}');
"@ -f ($u.Id -replace "'","''"), ($mbType -replace "'","''"), ($mbType -replace "'","''"), $nowLoop
                    Invoke-MySQLiteQuery -Path $dbPath -Query $insertMailbox | Out-Null
                }
                Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET processed = processed + 1, updatedOnUtc = '{0}' WHERE taskId = '{1}';" -f (Get-Date).ToUniversalTime().ToString('o'), $taskId) | Out-Null
            }

            [pscustomobject]@{
                SourceId          = $u.Id
                Type              = $type
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                EmailAddress      = $u.Mail
                ProxyAddresses    = ($u.ProxyAddresses -join ',')
                AssignedLicences  = $AssignedLicenses
            }
        }

        if ($OutputPath) {
            $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        }
        if ($UseSqlite -and $dbPath) {
            $nowIso = (Get-Date).ToUniversalTime().ToString('o')
            Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET status = 'Completed', updatedOnUtc = '{0}', message = 'OK' WHERE taskId = '{1}';" -f $nowIso, $taskId) | Out-Null
            $statusUpdated = $true
            return [pscustomobject]@{ TaskId = $taskId; OutputPath = $OutputPath; Total = $total; Status = 'Completed' }
        }
        return $OutputPath
    }
    catch {
        if ($UseSqlite -and $dbPath -and -not $statusUpdated) {
            $errMsg = $_.Exception.Message -replace "'","''"
            Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET status = 'Failed', message = '{0}', updatedOnUtc = '{1}' WHERE taskId = '{2}';" -f $errMsg, (Get-Date).ToUniversalTime().ToString('o'), $taskId) | Out-Null
        }
        throw
    }
}
function Invoke-DiscoverTargetAccounts {
    [CmdletBinding()]
    param(
        [string]$ProjectPath,
        [string]$OutputPath,
        [switch]$UseSqlite
    )

    $projectPath = Resolve-ProjectPath $ProjectPath
    $context = Get-MigrationProjectContext -ProjectPath $projectPath
    if (-not $OutputPath) {
        $OutputPath = Join-Path $context.ProjectPath 'DiscoveredTargetAccounts.csv'
    }

    $dbPath = $null
    $taskId = [guid]::NewGuid().ToString()
    $statusUpdated = $false
    if ($UseSqlite) {
        $dbPath = Get-MigrationDatabase -ProjectPath $context.ProjectPath
        $nowIso = (Get-Date).ToUniversalTime().ToString('o')
        Invoke-MySQLiteQuery -Path $dbPath -Query (@"
INSERT OR REPLACE INTO Tasks (taskId, name, total, processed, status, message, startedOnUtc, updatedOnUtc)
VALUES ('{0}', 'DiscoverTargetAccounts', 0, 0, 'Running', '', '{1}', '{1}');
"@ -f $taskId, $nowIso) | Out-Null
    }

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-ProjectTargetTenant -ProjectPath $context.ProjectPath -Silent | Out-Null
        $users = $null
        try {
            $users = Get-MgUser -All -Property 'id','displayName','userPrincipalName','mail','proxyAddresses','assignedLicenses' -ErrorAction Stop |
                Select-Object 'id','displayName','userPrincipalName','mail','proxyAddresses','assignedLicenses'
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $total = @($users).Count
        if ($UseSqlite -and $dbPath) {
            $nowIso = (Get-Date).ToUniversalTime().ToString('o')
            Invoke-MySQLiteQuery -Path $dbPath -Query (@"
UPDATE Tasks SET total = {0}, updatedOnUtc = '{1}' WHERE taskId = '{2}';
"@ -f $total, $nowIso, $taskId) | Out-Null
        }

        $mailboxMap = @{}
        try {
            Connect-ProjectTargetExchange -ProjectPath $context.ProjectPath | Out-Null
            try {
                $recipients = Get-EXORecipient -ResultSize Unlimited -PropertySets All -ErrorAction Stop
                foreach ($recip in $recipients) {
                    if ($recip.ExternalDirectoryObjectId) {
                        $mailboxMap[$recip.ExternalDirectoryObjectId] = $recip.RecipientTypeDetails
                    }
                }
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch {
            Write-Warning "Exchange discovery (target) failed: $($_.Exception.Message). Mailbox type will be blank."
        }

        $rows = foreach ($u in $users) {
            $type = $null
            if ($mailboxMap.ContainsKey($u.Id)) {
                $type = $mailboxMap[$u.Id]
            }
            elseif ($u.Mail -or ($u.ProxyAddresses -and $u.ProxyAddresses.Count -gt 0)) {
                $type = 'MailEnabledUser'
            }
            else {
                $type = 'NonMailEnabledUser'
	        }

	        $assignedSkus = @()
	        if ($u.AssignedLicenses) {
	            $assignedSkus = @($u.AssignedLicenses | ForEach-Object { (Get-MicrosoftOfficeProduct -SkuId $_.SkuId).ProductDisplayName })
	        }
	        $AssignedLicenses = ($assignedSkus -join ",")

            if ($UseSqlite -and $dbPath) {
                $nowLoop = (Get-Date).ToUniversalTime().ToString('o')
                $insertAccount = @"
INSERT OR REPLACE INTO AccountsTarget (objectId, targetUpn, displayName, mail, proxyAddresses, assignedLicenses, mailboxType, usageLocation, lastSeenUtc)
VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}');
"@ -f ($u.Id -replace "'","''"), ($u.UserPrincipalName -replace "'","''"), ($u.DisplayName -replace "'","''"), ($u.Mail -replace "'","''"), (($u.ProxyAddresses -join ',') -replace "'","''"), ($AssignedLicenses -replace "'","''"), ($type -replace "'","''"), (($u.UsageLocation) -replace "'","''"), $nowLoop
                Invoke-MySQLiteQuery -Path $dbPath -Query $insertAccount | Out-Null
                if ($mailboxMap.ContainsKey($u.Id)) {
                    $mbType = $mailboxMap[$u.Id]
                    $insertMailbox = @"
INSERT OR REPLACE INTO MailboxesTarget (objectId, recipientType, mailboxType, lastSeenUtc)
VALUES ('{0}', '{1}', '{2}', '{3}');
"@ -f ($u.Id -replace "'","''"), ($mbType -replace "'","''"), ($mbType -replace "'","''"), $nowLoop
                    Invoke-MySQLiteQuery -Path $dbPath -Query $insertMailbox | Out-Null
                }
                Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET processed = processed + 1, updatedOnUtc = '{0}' WHERE taskId = '{1}';" -f (Get-Date).ToUniversalTime().ToString('o'), $taskId) | Out-Null
            }

            [pscustomobject]@{
                TargetId          = $u.Id
                Type              = $type
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                EmailAddress      = $u.Mail
                ProxyAddresses    = ($u.ProxyAddresses -join ',')
                AssignedLicences  = $AssignedLicenses
            }
        }

        if ($OutputPath) {
            $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        }
        if ($UseSqlite -and $dbPath) {
            $nowIso = (Get-Date).ToUniversalTime().ToString('o')
            Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET status = 'Completed', updatedOnUtc = '{0}', message = 'OK' WHERE taskId = '{1}';" -f $nowIso, $taskId) | Out-Null
            $statusUpdated = $true
            return [pscustomobject]@{ TaskId = $taskId; OutputPath = $OutputPath; Total = $total; Status = 'Completed' }
        }
        return $OutputPath
    }
    catch {
        if ($UseSqlite -and $dbPath -and -not $statusUpdated) {
            $errMsg = $_.Exception.Message -replace "'","''"
            Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET status = 'Failed', message = '{0}', updatedOnUtc = '{1}' WHERE taskId = '{2}';" -f $errMsg, (Get-Date).ToUniversalTime().ToString('o'), $taskId) | Out-Null
        }
        throw
    }
}

function Resolve-ProjectPath {
	param([string]$Path)

	if ($Path) { return $Path }
	if ($script:ActiveProjectContext) { return $script:ActiveProjectContext.ProjectPath }
	return (Get-Location).Path
}

function Assert-MySQLiteAvailable {
	[CmdletBinding()]
	param()

	if (-not (Get-Module -ListAvailable -Name 'mySQLite')) {
		throw "mySQLite module is required for SQLite-backed features. Install-Module mySQLite"
	}
	if (-not (Get-Module -Name 'mySQLite')) {
		Import-Module mySQLite -ErrorAction Stop
	}
}

function Assert-ThreadJobAvailable {
	[CmdletBinding()]
	param()

	if (-not (Get-Command -Name Start-ThreadJob -ErrorAction SilentlyContinue)) {
		throw "Start-ThreadJob is required for -AsJob helpers. Install/enable ThreadJob (PowerShell 7+) or use Start-Job instead."
	}
}

function Get-MigrationDatabase {
	[CmdletBinding()]
	param(
		[string]$ProjectPath
	)

	$resolvedProject = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $resolvedProject
	Assert-MySQLiteAvailable
	$dbPath = Join-Path $context.ProjectPath 'Discovery.sqlite'
	if (-not (Test-Path -LiteralPath $dbPath)) {
		New-MySQLiteDB -Path $dbPath -Force | Out-Null
	}
	Invoke-MySQLiteQuery -Path $dbPath -Query "PRAGMA journal_mode=WAL;" | Out-Null
	Invoke-MySQLiteQuery -Path $dbPath -Query "PRAGMA busy_timeout=5000;" | Out-Null
	$createStatements = @(
		"CREATE TABLE IF NOT EXISTS Metadata (schemaVersion INTEGER)",
		"CREATE TABLE IF NOT EXISTS AccountsSource (objectId TEXT PRIMARY KEY, sourceUpn TEXT, displayName TEXT, mail TEXT, proxyAddresses TEXT, assignedLicenses TEXT, mailboxType TEXT, usageLocation TEXT, lastSeenUtc TEXT)",
		"CREATE TABLE IF NOT EXISTS AccountsTarget (objectId TEXT PRIMARY KEY, targetUpn TEXT, displayName TEXT, mail TEXT, proxyAddresses TEXT, assignedLicenses TEXT, mailboxType TEXT, usageLocation TEXT, lastSeenUtc TEXT)",
		"CREATE TABLE IF NOT EXISTS MailboxesSource (objectId TEXT PRIMARY KEY, recipientType TEXT, mailboxSize TEXT, itemCount INTEGER, archiveSize TEXT, lastLogon TEXT, forwarding TEXT, holdFlags TEXT, mailboxType TEXT, lastSeenUtc TEXT)",
		"CREATE TABLE IF NOT EXISTS MailboxesTarget (objectId TEXT PRIMARY KEY, recipientType TEXT, mailboxSize TEXT, itemCount INTEGER, archiveSize TEXT, lastLogon TEXT, forwarding TEXT, holdFlags TEXT, mailboxType TEXT, lastSeenUtc TEXT)"
	)
	foreach ($stmt in $createStatements) {
		Invoke-MySQLiteQuery -Path $dbPath -Query $stmt | Out-Null
	}
	$ensureColumn = {
		param($table, $column)
		$colInfo = Invoke-MySQLiteQuery -Path $dbPath -Query "PRAGMA table_info($table);" | Where-Object { $_.name -eq $column }
		if (-not $colInfo) {
			Invoke-MySQLiteQuery -Path $dbPath -Query "ALTER TABLE $table ADD COLUMN $column TEXT;" | Out-Null
		}
	}
	$ensureColumn.Invoke('MailboxesSource', 'mailboxType')
	$ensureColumn.Invoke('MailboxesTarget', 'mailboxType')
	Invoke-MySQLiteQuery -Path $dbPath -Query "CREATE TABLE IF NOT EXISTS Tasks (taskId TEXT PRIMARY KEY, name TEXT, total INTEGER, processed INTEGER, status TEXT, message TEXT, startedOnUtc TEXT, updatedOnUtc TEXT);" | Out-Null
	return $dbPath
}

function Get-MigrationTask {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[string]$TaskId,
		[string]$Name
	)

	$dbPath = Get-MigrationDatabase -ProjectPath (Resolve-ProjectPath $ProjectPath)
	$filter = ''
	if ($TaskId) {
		$filter = "WHERE taskId = '$($TaskId -replace '''','''''')'"
	}
	elseif ($Name) {
		$filter = "WHERE name = '$($Name -replace '''','''''')'"
	}
	$query = "SELECT taskId, name, total, processed, status, message, startedOnUtc, updatedOnUtc FROM Tasks $filter ORDER BY updatedOnUtc DESC;"
	$rows = Invoke-MySQLiteQuery -Path $dbPath -Query $query
	foreach ($row in $rows) {
		$percent = $null
		if ($row.total -and $row.processed -ne $null -and $row.total -gt 0) {
			$percent = [math]::Round(100.0 * ($row.processed / $row.total), 1)
		}
		[pscustomobject]@{
			TaskId       = $row.taskId
			Name         = $row.name
			Total        = $row.total
			Processed    = $row.processed
			Percent      = $percent
			Status       = $row.status
			Message      = $row.message
			StartedOnUtc = $row.startedOnUtc
			UpdatedOnUtc = $row.updatedOnUtc
			DatabasePath = $dbPath
		}
	}
}

function Invoke-DiscoverSourceMailboxStatistics {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[switch]$UseSqlite
	)

	$projectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $projectPath
	$dbPath = $null
	$taskId = [guid]::NewGuid().ToString()
	$statusUpdated = $false
	if ($UseSqlite) {
		$dbPath = Get-MigrationDatabase -ProjectPath $context.ProjectPath
		$nowIso = (Get-Date).ToUniversalTime().ToString('o')
		Invoke-MySQLiteQuery -Path $dbPath -Query (@"
INSERT OR REPLACE INTO Tasks (taskId, name, total, processed, status, message, startedOnUtc, updatedOnUtc)
VALUES ('{0}', 'DiscoverSourceMailboxStatistics', 0, 0, 'Running', '', '{1}', '{1}');
"@ -f $taskId, $nowIso) | Out-Null
	}
	elseif (-not $UseSqlite) {
		throw "Invoke-DiscoverSourceMailboxStatistics requires -UseSqlite to store results in Discovery.sqlite."
	}

	Connect-ProjectSourceTenant -ProjectPath $context.ProjectPath -Silent | Out-Null
	Connect-ProjectSourceExchange -ProjectPath $context.ProjectPath | Out-Null
	try {
		$mailboxes = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
		$total = @($mailboxes).Count
		if ($UseSqlite -and $dbPath) {
			$nowIso = (Get-Date).ToUniversalTime().ToString('o')
			Invoke-MySQLiteQuery -Path $dbPath -Query (@"
UPDATE Tasks SET total = {0}, updatedOnUtc = '{1}' WHERE taskId = '{2}';
"@ -f $total, $nowIso, $taskId) | Out-Null
		}
		foreach ($mb in $mailboxes) {
			$stats = $null
			try {
				$stats = Get-EXOMailboxStatistics -Identity $mb.ExternalDirectoryObjectId -ErrorAction Stop
			}
			catch {
				Write-Warning "Failed to get mailbox statistics for $($mb.UserPrincipalName): $($_.Exception.Message)"
			}

			$size = $null
			$itemCount = $null
			$archiveSize = $null
			$lastLogon = $null
			if ($stats) {
				$size = $stats.TotalItemSize
				$itemCount = $stats.ItemCount
				$archiveSize = $stats.TotalDeletedItemSize
				$lastLogon = $stats.LastLogonTime
			}

			$recipientType = $mb.RecipientTypeDetails
			$forwarding = $mb.ForwardingSmtpAddress ?? $mb.ForwardingAddress
			$holdFlags = @()
			if ($mb.LitigationHoldEnabled) { $holdFlags += 'LitigationHold' }
			if ($mb.InPlaceHolds -and $mb.InPlaceHolds.Count -gt 0) { $holdFlags += 'InPlaceHold' }

			$nowLoop = (Get-Date).ToUniversalTime().ToString('o')
			$insertMailbox = @"
INSERT OR REPLACE INTO MailboxesSource (objectId, recipientType, mailboxType, mailboxSize, itemCount, archiveSize, lastLogon, forwarding, holdFlags, lastSeenUtc)
VALUES ('{0}', '{1}', '{2}', '{3}', {4}, '{5}', '{6}', '{7}', '{8}', '{9}');
"@ -f ($mb.ExternalDirectoryObjectId -replace "'","''"), ($recipientType -replace "'","''"), ($recipientType -replace "'","''"), (($size) -replace "'","''"), ($itemCount ?? 'NULL'), (($archiveSize) -replace "'","''"), (($lastLogon) -replace "'","''"), (($forwarding) -replace "'","''"), (($holdFlags -join '|') -replace "'","''"), $nowLoop
			Invoke-MySQLiteQuery -Path $dbPath -Query $insertMailbox | Out-Null
			if ($UseSqlite -and $dbPath) {
				Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET processed = processed + 1, updatedOnUtc = '{0}' WHERE taskId = '{1}';" -f (Get-Date).ToUniversalTime().ToString('o'), $taskId) | Out-Null
			}
		}
	}
	finally {
		Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
	}
	if ($UseSqlite -and $dbPath) {
		$nowIso = (Get-Date).ToUniversalTime().ToString('o')
		Invoke-MySQLiteQuery -Path $dbPath -Query ("UPDATE Tasks SET status = 'Completed', updatedOnUtc = '{0}', message = 'OK' WHERE taskId = '{1}';" -f $nowIso, $taskId) | Out-Null
		$statusUpdated = $true
		return [pscustomobject]@{ TaskId = $taskId; DatabasePath = $dbPath; Status = 'Completed' }
	}
	return $dbPath
}

function Start-AccountDiscoveryJob {
	[CmdletBinding()]
	param(
		[ValidateSet('Source','Target')][string]$Scope = 'Source',
		[string]$ProjectPath,
		[string]$OutputPath,
		[switch]$UseSqlite
	)

	Assert-ThreadJobAvailable
	$modulePath = $MyInvocation.MyCommand.Module.Path
	$proj = Resolve-ProjectPath $ProjectPath
	$jobName = "Discover-$Scope"

	return Start-ThreadJob -Name $jobName -ScriptBlock {
		param($modulePath,$proj,$output,$useSqlite,$scope)
		Import-Module $modulePath -Force
		if ($scope -eq 'Target') {
			Invoke-DiscoverTargetAccounts -ProjectPath $proj -OutputPath $output -UseSqlite:$useSqlite
		}
		else {
			Invoke-DiscoverSourceAccounts -ProjectPath $proj -OutputPath $output -UseSqlite:$useSqlite
		}
	} -ArgumentList @($modulePath,$proj,$OutputPath,$UseSqlite,$Scope)
}

function Start-CollectMailboxStatisticsJob {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[switch]$UseSqlite
	)

	Assert-ThreadJobAvailable
	$modulePath = $MyInvocation.MyCommand.Module.Path
	$proj = Resolve-ProjectPath $ProjectPath
	$jobName = "Discover-SourceMailboxStats"

	return Start-ThreadJob -Name $jobName -ScriptBlock {
		param($modulePath,$proj,$useSqlite)
		Import-Module $modulePath -Force
		Invoke-DiscoverSourceMailboxStatistics -ProjectPath $proj -UseSqlite:$useSqlite
	} -ArgumentList @($modulePath,$proj,$UseSqlite)
}

function Invoke-ProjectTenantProvisioning {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$TenantId,
		[Parameter(Mandatory)][string]$TenantKey,
		[Parameter(Mandatory)][string]$ProfileName,
		[string]$ProjectPath,
		[string]$AppDisplayName,
		[ValidateSet('Certificate', 'ClientCredential')][string]$AuthType = 'Certificate',
		[string[]]$Permissions = @('Directory.Read.All'),
		[switch]$EnableExchangeManageAsApp,
		[string]$Description
	)

	$ProjectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	$displayName = if ($AppDisplayName) { $AppDisplayName } else { "{0} {1}" -f $context.ProjectName, $ProfileName }
	Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	try {
		$parameters = @{
			Name             = $displayName
			AuthType         = $AuthType
			Permissions      = $Permissions
			TenantId         = $TenantId
			CompanyName      = $context.ProjectName
			ProjectDescription = $Description ?? $context.Config.projectDescription
			ProfileName      = $ProfileName
			ProfileDirectory = $context.ProfilesPath
			ConfigPath       = $context.ConfigPath
			ConfigTenantKey  = $TenantKey
			EnableExchangeManageAsApp = $EnableExchangeManageAsApp
		}
		$outcome = New-MgCustomApp @parameters
		try {
			$org = Get-MgOrganization -Property DisplayName,VerifiedDomains -ErrorAction Stop | Select-Object -First 1
			if ($org) {
				$initialDomain = ($org.VerifiedDomains | Where-Object { $_.IsInitial } | Select-Object -First 1 -ExpandProperty Name)
				if (-not $initialDomain) {
					$initialDomain = ($org.VerifiedDomains | Select-Object -First 1 -ExpandProperty Name)
				}
				$metadata = @{}
				if ($org.DisplayName) {
					$metadata['tenantDisplayName'] = $org.DisplayName
					$metadata['companyName'] = $org.DisplayName
				}
				if ($initialDomain) {
					$metadata['defaultDomain'] = $initialDomain
				}
				if ($metadata.Count -gt 0) {
					Update-ProjectTenantConfig -ConfigPath $context.ConfigPath -TenantKey $TenantKey -Properties $metadata
				}
			}
		}
		catch {
			Write-Verbose "Unable to collect organization metadata: $($_.Exception.Message)"
		}
		return $outcome
	}
	finally {
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
	}
}

function Add-ProjectTargetTenant {
	[CmdletBinding(DefaultParameterSetName = 'Provision')]
	param(
		[Parameter(Mandatory, ParameterSetName = 'Provision')][string]$TenantId,
		[Parameter(ParameterSetName = 'Provision')]
		[Parameter(ParameterSetName = 'Import')]
		[string]$ProjectPath,
		[Parameter(ParameterSetName = 'Provision')][string]$AppDisplayName,
		[Parameter(ParameterSetName = 'Provision')][ValidateSet('Certificate', 'ClientCredential')][string]$AuthType = 'Certificate',
		[Parameter(ParameterSetName = 'Provision')][string[]]$Permissions = @('Directory.Read.All','User.Read.All'),
		[Parameter(ParameterSetName = 'Provision')][switch]$EnableExchangeManageAsApp,

		[Parameter(Mandatory, ParameterSetName = 'Import')][string]$ImportProjectPath,
		[Parameter(ParameterSetName = 'Import')][string]$ImportConfigPath,
		[Parameter(ParameterSetName = 'Import')][string]$ImportTenantKey = 'targetTenant',
		[Parameter(ParameterSetName = 'Import')][string]$ProfileName = 'Target'
	)

	if ($PSCmdlet.ParameterSetName -eq 'Import') {
		$destinationProject = Resolve-ProjectPath $ProjectPath
		$destinationContext = Get-MigrationProjectContext -ProjectPath $destinationProject
		$sourceConfig = $ImportConfigPath
		if ($ImportProjectPath -and -not $sourceConfig) {
			$resolvedImportProject = (Resolve-Path -LiteralPath $ImportProjectPath).Path
			$sourceConfig = Join-Path $resolvedImportProject 'Config.json'
		}
		if (-not $sourceConfig) {
			throw "Specify -ImportProjectPath (project root) or -ImportConfigPath (Config.json) to import the target tenant."
		}
		return Import-ProjectTenantConfig -SourceConfigPath $sourceConfig -DestinationConfigPath $destinationContext.ConfigPath -DestinationProfilesPath $destinationContext.ProfilesPath -SourceTenantKey $ImportTenantKey -DestinationTenantKey 'targetTenant' -ProfileName $ProfileName
	}

	Invoke-ProjectTenantProvisioning -TenantId $TenantId -TenantKey 'targetTenant' -ProfileName 'Target' -ProjectPath (Resolve-ProjectPath $ProjectPath) -AppDisplayName $AppDisplayName -AuthType $AuthType -Permissions $Permissions -EnableExchangeManageAsApp:$EnableExchangeManageAsApp -Description 'Target tenant'
}

function Add-ProjectSourceTenant {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$TenantId,
		[string]$ProjectPath,
		[string]$AppDisplayName,
		[ValidateSet('Certificate', 'ClientCredential')][string]$AuthType = 'Certificate',
		[string[]]$Permissions = @('Directory.Read.All','User.Read.All'),
		[switch]$EnableExchangeManageAsApp
	)

	Invoke-ProjectTenantProvisioning -TenantId $TenantId -TenantKey 'sourceTenant' -ProfileName 'Source' -ProjectPath (Resolve-ProjectPath $ProjectPath) -AppDisplayName $AppDisplayName -AuthType $AuthType -Permissions $Permissions -EnableExchangeManageAsApp:$EnableExchangeManageAsApp -Description 'Source tenant'
}

function Connect-ProjectSourceTenant {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[switch]$PassThruProfile,
		[switch]$Silent
	)

	$ProjectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	$connectParams = @{
		ConfigPath      = $context.ConfigPath
		TenantKey       = 'sourceTenant'
		PassThruProfile = $PassThruProfile
	}
	if ($Silent) {
		$connectParams['Verbose'] = $false
	}
	else {
		$connectParams['Verbose'] = $VerbosePreference
	}
	Connect-Source @connectParams
}

function Connect-ProjectTargetTenant {
	[CmdletBinding()]
	param(
		[string]$ProjectPath,
		[switch]$PassThruProfile,
		[switch]$Silent
	)

	$ProjectPath = Resolve-ProjectPath $ProjectPath
	$context = Get-MigrationProjectContext -ProjectPath $ProjectPath
	$connectParams = @{
		ConfigPath      = $context.ConfigPath
		TenantKey       = 'targetTenant'
		PassThruProfile = $PassThruProfile
	}
	if ($Silent) {
		$connectParams['Verbose'] = $false
	}
	else {
		$connectParams['Verbose'] = $VerbosePreference
	}
	Connect-Target @connectParams
}

Set-Alias -Name cpst -Value Connect-ProjectSourceTenant
Set-Alias -Name cptt -Value Connect-ProjectTargetTenant
Set-Alias -Name cpse -Value Connect-ProjectSourceExchange
Set-Alias -Name cpte -Value Connect-ProjectTargetExchange

Export-ModuleMember -Function `
	Set-MgProfileDirectory, `
	Set-MgConfigPath, `
	Get-MigrationProjectContext, `
	Get-MigrationProject, `
	Export-MigrationProject, `
	Import-MigrationProject, `
	Invoke-MigrateUser, `
	Get-MigrationTask, `
	Invoke-DiscoverSourceAccounts, `
	Invoke-DiscoverTargetAccounts, `
	Invoke-DiscoverSourceMailboxStatistics, `
	Start-AccountDiscoveryJob, `
	Start-CollectMailboxStatisticsJob, `
	New-MgCustomApp, `
	Get-MicrosoftProductSheet, `
	Import-MicrosoftProductSheet, `
	Get-MSProduct, `
	Get-MicrosoftOfficeProduct, `
	Set-ActiveMigrationProject, `
	Clear-ActiveMigrationProject, `
	Get-ActiveMigrationProject, `
	Connect-Source, `
	Connect-Target, `
	Connect-ProjectSourceExchange, `
	Connect-ProjectTargetExchange, `
	New-MigrationProject, `
	Add-ProjectSourceTenant, `
	Add-ProjectTargetTenant, `
	Connect-ProjectSourceTenant, `
	Connect-ProjectTargetTenant `
	-Alias `
	cpst, `
	cptt, `
	cpse, `
	cpte
