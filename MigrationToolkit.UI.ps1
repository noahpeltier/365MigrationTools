<#
.SYNOPSIS
    WPF front-end for the MigrationToolkit module, styled after the provided mockup.
    Run with: pwsh.exe -File .\MigrationToolkit.UI.ps1
#>

[CmdletBinding()]
param()

# Load dependencies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'MigrationToolkit.psm1'
if (-not (Test-Path $modulePath)) {
    [System.Windows.MessageBox]::Show("Module not found at `n$modulePath", 'Migration Toolkit UI', 'OK', 'Error') | Out-Null
    return
}

Import-Module $modulePath -Force

$defaultProjectsRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Projects'

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Migration Toolkit" Height="840" Width="1300" ResizeMode="CanResize"
        Background="#f5f7fb" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style x:Key="NavButtonStyle" TargetType="Button">
            <Setter Property="Margin" Value="4,0,0,0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Background" Value="#e7eef7"/>
            <Setter Property="Foreground" Value="#2d3e50"/>
            <Setter Property="BorderBrush" Value="#cfd8e3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="PrimaryTabStyle" TargetType="ToggleButton">
            <Setter Property="Width" Value="80"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Background" Value="#1677be"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#0f5d94"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="LightTabStyle" TargetType="ToggleButton">
            <Setter Property="Width" Value="100"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Background" Value="#eef2f7"/>
            <Setter Property="Foreground" Value="#2d3e50"/>
            <Setter Property="BorderBrush" Value="#cfd8e3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="ActionButtonStyle" TargetType="Button">
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Background" Value="#eef2f7"/>
            <Setter Property="Foreground" Value="#2d3e50"/>
            <Setter Property="BorderBrush" Value="#cfd8e3"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="LinkLike" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#1677be"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- nav -->
            <RowDefinition Height="Auto"/>   <!-- filters -->
            <RowDefinition Height="Auto"/>   <!-- actions -->
            <RowDefinition Height="*"/>      <!-- accounts grid -->
            <RowDefinition Height="220"/>    <!-- log + tasks -->
        </Grid.RowDefinitions>

        <!-- Navigation -->
        <DockPanel Grid.Row="0" LastChildFill="True" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                <Button Content="DASHBOARD" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="ACCOUNTS" Style="{StaticResource NavButtonStyle}" Background="#ffffff" BorderBrush="#1677be"/>
                <Button Content="MAILBOXES" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="ONEDRIVE" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="DESKTOP AGENTS" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="REPORTS" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="TEMPLATES" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="EVENTS" Style="{StaticResource NavButtonStyle}"/>
                <Button Content="TASKS" Style="{StaticResource NavButtonStyle}"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" DockPanel.Dock="Right" VerticalAlignment="Center" >
                <TextBlock Text="Projects Root:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBox Name="RootPathBox" Width="220" Margin="0,0,6,0"/>
                <Button Name="BrowseRootButton" Content="Browse" Width="70" Margin="0,0,6,0"/>
                <Button Name="RefreshProjectsButton" Content="Refresh" Width="70" Margin="0,0,6,0"/>
                <Button Name="NewProjectButton" Content="New" Width="60" Margin="0,0,6,0"/>
                <ComboBox Name="ProjectsCombo" Width="220" Margin="0,0,6,0"/>
                <Button Name="SetActiveButton" Content="Set Active" Width="90" Margin="0,0,6,0"/>
                <Button Name="ProjectDetailsButton" Content="Details" Width="70" Margin="0,0,6,0"/>
                <Button Name="OpenProjectButton" Content="Open" Width="60" Margin="0,0,6,0"/>
                <CheckBox Name="UseSqliteCheckBox" Content="Use SQLite" IsChecked="True" VerticalAlignment="Center"/>
            </StackPanel>
        </DockPanel>

        <!-- Filters -->
        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="80"/>
            </Grid.ColumnDefinitions>
            <ToggleButton Name="ListViewToggle" Content="List View" Style="{StaticResource PrimaryTabStyle}" IsChecked="True" Grid.Column="0"/>
            <ToggleButton Name="AssessmentToggle" Content="Assessment" Style="{StaticResource LightTabStyle}" Grid.Column="1"/>
            <StackPanel Orientation="Horizontal" Grid.Column="2" Margin="10,0,0,0">
                <TextBlock Text="Account State:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="AccountStateFilter" Width="120">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Migrated"/>
                    <ComboBoxItem Content="Partial"/>
                    <ComboBoxItem Content="Pending"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Grid.Column="3" Margin="10,0,0,0">
                <TextBlock Text="Matching:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="MatchingFilter" Width="100">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Matched"/>
                    <ComboBoxItem Content="Unmatched"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Grid.Column="4" Margin="10,0,0,0">
                <TextBlock Text="Source Type:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="SourceTypeFilter" Width="120">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Mailbox-Enabled User"/>
                    <ComboBoxItem Content="Mail-Enabled User"/>
                    <ComboBoxItem Content="Shared Mailbox"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Grid.Column="5" Margin="10,0,0,0">
                <TextBlock Text="Target Type:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="TargetTypeFilter" Width="120">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Mailbox-Enabled User"/>
                    <ComboBoxItem Content="Mail-Enabled User"/>
                    <ComboBoxItem Content="Shared Mailbox"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Grid.Column="6" Margin="10,0,0,0">
                <TextBlock Text="Environment:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="EnvironmentFilter" Width="100">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Production"/>
                    <ComboBoxItem Content="Test"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Grid.Column="7" Margin="10,0,0,0">
                <TextBlock Text="ODM Licensed:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox Name="OdmFilter" Width="100">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="Yes"/>
                    <ComboBoxItem Content="No"/>
                </ComboBox>
            </StackPanel>
            <TextBox Name="AccountsSearchBox" Grid.Column="8" Height="30" VerticalContentAlignment="Center" Margin="10,0,6,0" />
            <Button Name="SearchButton" Content="Search" Grid.Column="9" Width="70" />
        </Grid>

        <!-- Actions -->
        <Grid Grid.Row="2" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button Name="DiscoveryButton" Content="DISCOVERY" Style="{StaticResource ActionButtonStyle}" Grid.Column="0"/>
            <Button Name="MatchButton" Content="MATCH" Style="{StaticResource ActionButtonStyle}" Grid.Column="1"/>
            <Button Name="MigrateButton" Content="MIGRATE" Style="{StaticResource ActionButtonStyle}" Grid.Column="2"/>
            <Button Name="CollectStatsButton" Content="COLLECT STATISTICS" Style="{StaticResource ActionButtonStyle}" Grid.Column="3"/>
            <Button Name="CollectionsButton" Content="COLLECTIONS" Style="{StaticResource ActionButtonStyle}" Grid.Column="4"/>
            <Button Name="MoreButton" Content="MORE" Style="{StaticResource ActionButtonStyle}" Grid.Column="5"/>
            <Button Name="EditColumnsButton" Content="EDIT COLUMNS" Style="{StaticResource ActionButtonStyle}" Grid.Column="6"/>
            <TextBlock Text="" Grid.Column="7"/>
            <StackPanel Orientation="Horizontal" Grid.Column="8" HorizontalAlignment="Right">
                <TextBlock Text="Active Project:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock Name="ActiveProjectLabel" Text="(none)" VerticalAlignment="Center" FontWeight="Bold"/>
            </StackPanel>
        </Grid>

        <!-- Accounts Grid -->
        <Border Grid.Row="3" Background="White" BorderBrush="#dbe2ea" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="10" LastChildFill="True">
                    <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                        <TextBlock Text="Accounts" FontSize="14" FontWeight="SemiBold" Margin="0,0,10,0"/>
                        <TextBlock Name="AccountCountLabel" Text="" VerticalAlignment="Center" />
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" DockPanel.Dock="Right">
                        <TextBlock Text="Scope:" VerticalAlignment="Center" Margin="0,0,6,0"/>
                        <ComboBox Name="ScopeCombo" Width="120" SelectedIndex="0">
                            <ComboBoxItem Content="Source Discovery" IsSelected="True"/>
                            <ComboBoxItem Content="Target Discovery"/>
                        </ComboBox>
                        <Button Name="RefreshAccountsButton" Content="Refresh" Margin="10,0,0,0" Width="80"/>
                    </StackPanel>
                </DockPanel>
                <DataGrid Name="AccountsGrid" Grid.Row="1" AutoGenerateColumns="False" HeadersVisibility="Column"
                          IsReadOnly="True" GridLinesVisibility="Horizontal" RowHeaderWidth="0" >
                    <DataGrid.Columns>
                        <DataGridCheckBoxColumn Header="" Binding="{Binding Selected}" Width="40"/>
                        <DataGridTextColumn Header="Sync Status" Binding="{Binding SyncStatus}" Width="80"/>
                        <DataGridTextColumn Header="Source Name" Binding="{Binding SourceName}" Width="200"/>
                        <DataGridTextColumn Header="Target Name" Binding="{Binding TargetName}" Width="200"/>
                        <DataGridTextColumn Header="Source Type" Binding="{Binding SourceType}" Width="160"/>
                        <DataGridTextColumn Header="Target Type" Binding="{Binding TargetType}" Width="160"/>
                        <DataGridTextColumn Header="ODM Licensed" Binding="{Binding OdmLicensed}" Width="120"/>
                        <DataGridTextColumn Header="Account State" Binding="{Binding AccountState}" Width="140"/>
                        <DataGridTextColumn Header="Source UPN" Binding="{Binding SourceUpn}" Width="*" />
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>

        <!-- Bottom: Log + Tasks -->
        <Grid Grid.Row="4" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <GroupBox Header="Log" Grid.Column="0" Margin="0,0,8,0">
                <TextBox Name="LogBox" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            </GroupBox>
            <GroupBox Header="Tasks" Grid.Column="1">
                <DockPanel>
                    <StackPanel Orientation="Horizontal" DockPanel.Dock="Top" Margin="0,0,0,6">
                        <Button Name="RefreshTasksButton" Content="Refresh Tasks" Width="110" Margin="0,0,6,0"/>
                        <Button Name="ReceiveJobButton" Content="Receive Job Output" Width="140"/>
                    </StackPanel>
                    <DataGrid Name="TasksGrid" AutoGenerateColumns="False" HeadersVisibility="Column" IsReadOnly="True"
                              GridLinesVisibility="Vertical" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Id" Binding="{Binding TaskId}" Width="60"/>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="200"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120"/>
                            <DataGridTextColumn Header="Processed" Binding="{Binding Processed}" Width="80"/>
                            <DataGridTextColumn Header="Total" Binding="{Binding Total}" Width="80"/>
                            <DataGridTextColumn Header="Percent" Binding="{Binding Percent}" Width="80"/>
                            <DataGridTextColumn Header="Updated (UTC)" Binding="{Binding UpdatedOnUtc}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </DockPanel>
            </GroupBox>
        </Grid>
    </Grid>

    <Window.ContextMenu>
        <ContextMenu/>
    </Window.ContextMenu>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Controls
$rootPathBox = $window.FindName('RootPathBox')
$projectsCombo = $window.FindName('ProjectsCombo')
$activeProjectLabel = $window.FindName('ActiveProjectLabel')
$projectDetailsButton = $window.FindName('ProjectDetailsButton')
$tasksGrid = $window.FindName('TasksGrid')
$logBox = $window.FindName('LogBox')
$useSqliteCheckBox = $window.FindName('UseSqliteCheckBox')
$accountsGrid = $window.FindName('AccountsGrid')
$accountCountLabel = $window.FindName('AccountCountLabel')
$scopeCombo = $window.FindName('ScopeCombo')
$accountsSearchBox = $window.FindName('AccountsSearchBox')

$script:AccountsCache = @()

function Write-UiLog {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message
    $logBox.AppendText("$line`r`n")
    $logBox.ScrollToEnd()
}

function Get-SelectedProjectPath {
    if (-not $projectsCombo.SelectedItem) {
        throw 'Select a project first.'
    }
    return [string]$projectsCombo.SelectedItem.Tag
}

function Load-Projects {
    $root = $rootPathBox.Text
    if (-not (Test-Path $root)) {
        Write-UiLog "Projects root not found: $root"
        return
    }

    $projectsCombo.Items.Clear()
    Get-ChildItem -Path $root -Directory | ForEach-Object {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $_.Name
        $item.Tag = $_.FullName
        [void]$projectsCombo.Items.Add($item)
    }

    if ($projectsCombo.Items.Count -gt 0) {
        $projectsCombo.SelectedIndex = 0
    }
    Write-UiLog "Loaded projects from $root"
}

function Show-ProjectDetails {
    param([string]$ProjectPath)
    $details = Get-MigrationProject -ProjectPath $ProjectPath
    $json = $details | ConvertTo-Json -Depth 6
    [System.Windows.MessageBox]::Show($json, 'Project Details') | Out-Null
}

function Refresh-Tasks {
    try {
        $path = Get-SelectedProjectPath
    } catch {
        Write-UiLog $_.Exception.Message
        return
    }

    try {
        $tasks = Get-MigrationTask -ProjectPath $path | Sort-Object UpdatedOnUtc -Descending
        $tasksGrid.ItemsSource = $tasks
        Write-UiLog "Loaded tasks for $path"
    } catch {
        Write-UiLog "Task refresh failed: $($_.Exception.Message)"
    }
}

function Prompt-ProjectName {
    $input = [Microsoft.VisualBasic.Interaction]::InputBox('Project Name', 'New Migration Project')
    if ([string]::IsNullOrWhiteSpace($input)) {
        return $null
    }
    return $input.Trim()
}

function Load-AccountsGrid {
    param(
        [string]$ProjectPath,
        [string]$Scope = 'Source'
    )

    $sourceFile = Join-Path $ProjectPath 'DiscoveredAccounts.csv'
    $targetFile = Join-Path $ProjectPath 'DiscoveredTargetAccounts.csv'
    $fileToUse = if ($Scope -eq 'Target' -and (Test-Path $targetFile)) { $targetFile }
                 elseif (Test-Path $sourceFile) { $sourceFile }
                 elseif (Test-Path $targetFile) { $targetFile }
                 else { $null }

    if (-not $fileToUse) {
        $accountsGrid.ItemsSource = @()
        $accountCountLabel.Text = 'No discovery CSV found'
        Write-UiLog "No discovery CSV found in $ProjectPath"
        return
    }

    try {
        $records = Import-Csv -Path $fileToUse
        $friendly = foreach ($r in $records) {
            $sourceName = $r.SourceName
            if (-not $sourceName) { $sourceName = $r.DisplayName }
            if (-not $sourceName) { $sourceName = $r.displayName }
            $sourceUpn = $r.SourceUpn
            if (-not $sourceUpn) { $sourceUpn = $r.UserPrincipalName }
            if (-not $sourceUpn) { $sourceUpn = $r.UPN }

            $odm = if ($r.AssignedLicenses -and $r.AssignedLicenses.Trim() -ne '') { 'Yes' } else { 'No' }
            $accountState = $r.AccountState
            if (-not $accountState) { $accountState = $r.Status }

            [pscustomobject]@{
                Selected    = $false
                SyncStatus  = $r.SyncStatus
                SourceName  = $sourceName
                TargetName  = $r.TargetName
                SourceType  = $r.SourceType ?? $r.MailboxType ?? $r.Type
                TargetType  = $r.TargetType
                OdmLicensed = $odm
                AccountState= $accountState
                SourceUpn   = $sourceUpn
            }
        }
        $script:AccountsCache = $friendly
        $accountsGrid.ItemsSource = $friendly
        $accountCountLabel.Text = ("{0} accounts" -f ($friendly.Count))
        Write-UiLog "Loaded $($friendly.Count) accounts from $fileToUse"
    } catch {
        Write-UiLog "Failed to load accounts: $($_.Exception.Message)"
    }
}

function Apply-SearchFilter {
    $term = $accountsSearchBox.Text
    if (-not $term -or $term.Trim().Length -eq 0) {
        $accountsGrid.ItemsSource = $script:AccountsCache
        $accountCountLabel.Text = ("{0} accounts" -f ($script:AccountsCache.Count))
        return
    }

    $term = $term.ToLowerInvariant()
    $filtered = $script:AccountsCache | Where-Object {
        $_.SourceName -and $_.SourceName.ToLower().Contains($term) -or
        $_.TargetName -and $_.TargetName.ToLower().Contains($term) -or
        $_.SourceUpn   -and $_.SourceUpn.ToLower().Contains($term)
    }
    $accountsGrid.ItemsSource = $filtered
    $accountCountLabel.Text = ("{0} account(s) (filtered)" -f ($filtered.Count))
}

function Start-Discovery {
    param(
        [ValidateSet('Source','Target')]
        [string]$Scope
    )
    $useSql = $useSqliteCheckBox.IsChecked -eq $true
    $path = Get-SelectedProjectPath
    $result = Start-AccountDiscoveryJob -Scope $Scope -ProjectPath $path -UseSqlite:$useSql
    if ($result.TaskId) {
        Write-UiLog "Started $Scope discovery task $($result.TaskId) with SQLite tracking."
    } else {
        Write-UiLog "Started $Scope discovery job."
    }
    Refresh-Tasks
}

function Start-MailboxStats {
    $useSql = $useSqliteCheckBox.IsChecked -eq $true
    $path = Get-SelectedProjectPath
    $result = Start-CollectMailboxStatisticsJob -ProjectPath $path -UseSqlite:$useSql
    if ($result.TaskId) {
        Write-UiLog "Started mailbox statistics task $($result.TaskId) with SQLite tracking."
    } else {
        Write-UiLog "Started mailbox statistics job."
    }
    Refresh-Tasks
}

# Event wiring
($window.FindName('RefreshProjectsButton')).Add_Click({
    Load-Projects
})

($window.FindName('BrowseRootButton')).Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $rootPathBox.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $rootPathBox.Text = $dialog.SelectedPath
        Load-Projects
    }
})

($window.FindName('NewProjectButton')).Add_Click({
    $name = Prompt-ProjectName
    if (-not $name) { return }
    try {
        $project = New-MigrationProject -Name $name -Path $rootPathBox.Text
        Write-UiLog "Created project $name at $($project.ProjectPath)"
        Load-Projects
    } catch {
        Write-UiLog "Failed to create project: $($_.Exception.Message)"
    }
})

$projectsCombo.Add_SelectionChanged({
    if (-not $projectsCombo.SelectedItem) { return }
    try {
        $path = Get-SelectedProjectPath
        $activeProjectLabel.Text = $path
        Write-UiLog "Selected project $path"
        $scopeLabel = ($scopeCombo.SelectedItem).Content
        $scope = if ($scopeLabel -like '*Target*') { 'Target' } else { 'Source' }
        Load-AccountsGrid -ProjectPath $path -Scope $scope
        Refresh-Tasks
    } catch {
        Write-UiLog $_.Exception.Message
    }
})

($window.FindName('SetActiveButton')).Add_Click({
    try {
        $path = Get-SelectedProjectPath
        Set-ActiveMigrationProject -ProjectPath $path
        $activeProjectLabel.Text = $path
        Write-UiLog "Active project set to $path"
    } catch {
        Write-UiLog $_.Exception.Message
    }
})

$projectDetailsButton.Add_Click({
    try {
        $path = Get-SelectedProjectPath
        Show-ProjectDetails -ProjectPath $path
    } catch {
        Write-UiLog $_.Exception.Message
    }
})

($window.FindName('OpenProjectButton')).Add_Click({
    try {
        $path = Get-SelectedProjectPath
        Start-Process explorer.exe $path
    } catch {
        Write-UiLog $_.Exception.Message
    }
})

($window.FindName('DiscoveryButton')).Add_Click({
    $menu = New-Object System.Windows.Controls.ContextMenu
    foreach ($scope in @('Source','Target')) {
        $item = New-Object System.Windows.Controls.MenuItem
        $item.Header = "Start $scope Discovery"
        $item.Add_Click({
            try { Start-Discovery -Scope $scope } catch { Write-UiLog $_.Exception.Message }
        })
        $menu.Items.Add($item) | Out-Null
    }
    $menu.IsOpen = $true
})

($window.FindName('CollectStatsButton')).Add_Click({
    try { Start-MailboxStats } catch { Write-UiLog $_.Exception.Message }
})

($window.FindName('RefreshAccountsButton')).Add_Click({
    try {
        $path = Get-SelectedProjectPath
        $scopeLabel = ($scopeCombo.SelectedItem).Content
        $scope = if ($scopeLabel -like '*Target*') { 'Target' } else { 'Source' }
        Load-AccountsGrid -ProjectPath $path -Scope $scope
    } catch {
        Write-UiLog $_.Exception.Message
    }
})

($window.FindName('RefreshTasksButton')).Add_Click({
    Refresh-Tasks
})

($window.FindName('ReceiveJobButton')).Add_Click({
    $jobId = [Microsoft.VisualBasic.Interaction]::InputBox('Job Id to receive (Get-Job)', 'Receive Job Output')
    if ([string]::IsNullOrWhiteSpace($jobId)) { return }
    try {
        $job = Get-Job -Id [int]$jobId -ErrorAction Stop
        $output = $job | Receive-Job -Keep
        Write-UiLog "Job $jobId output:`r`n$($output | Out-String)"
    } catch {
        Write-UiLog "Receive-Job failed: $($_.Exception.Message)"
    }
})

$accountsSearchBox.Add_TextChanged({ Apply-SearchFilter })
($window.FindName('SearchButton')).Add_Click({ Apply-SearchFilter })

$scopeCombo.Add_SelectionChanged({
    try {
        $path = Get-SelectedProjectPath
        $scopeLabel = ($scopeCombo.SelectedItem).Content
        $scope = if ($scopeLabel -like '*Target*') { 'Target' } else { 'Source' }
        Load-AccountsGrid -ProjectPath $path -Scope $scope
    } catch {
        # ignore when no project yet
    }
})

($window.FindName('MatchButton')).Add_Click({
    Write-UiLog "Match workflow not yet automated in UI. Use module cmdlets as needed."
})
($window.FindName('MigrateButton')).Add_Click({
    Write-UiLog "Migration actions can be scripted via Invoke-MigrateUser; UI hook is a placeholder."
})
($window.FindName('CollectionsButton')).Add_Click({
    Write-UiLog "Collections UI placeholder."
})
($window.FindName('MoreButton')).Add_Click({
    Write-UiLog "More menu placeholder."
})
($window.FindName('EditColumnsButton')).Add_Click({
    Write-UiLog "Column editing not implemented in this UI."
})

# Initial state
$rootPathBox.Text = $defaultProjectsRoot
Load-Projects

# Show UI
$null = $window.ShowDialog()
