if (-not (Get-Command Write-AppLog -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'utils.ps1')
}
if (-not (Get-Command Invoke-StatusCheck -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'installer.ps1')
}
if (-not (Get-Command Install-hermes-agent-windowsApp -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'app-manager.ps1')
}

function Start-hermes-agent-windowsGui {
    $projectRoot = Get-ProjectRoot
    $global:HermesAgentWindowsRoot = $projectRoot
    $global:HermesAgentWindowsLogPath = Get-LogFilePath -Kind 'app'

    if (-not (Test-Path (Join-Path $projectRoot 'logs'))) {
        New-ProjectFolder -RootPath $projectRoot | Out-Null
    }

    try {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml -ErrorAction Stop | Out-Null
    }
    catch {
        Show-ErrorMessage -Message 'WPF is not available on this system. hermes-agent-windows needs Windows desktop PowerShell support.' -Title 'hermes-agent-windows'
        return
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="hermes-agent-windows"
        Height="760"
        Width="1180"
        MinHeight="650"
        MinWidth="980"
        WindowStartupLocation="CenterScreen"
        Background="#14161A"
        Foreground="#F2F4F8"
        FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="CardBorder" Color="#2B3138" />
        <SolidColorBrush x:Key="CardFill" Color="#1B1F24" />
        <SolidColorBrush x:Key="Accent" Color="#4CC2FF" />
        <SolidColorBrush x:Key="Good" Color="#2ECC71" />
        <SolidColorBrush x:Key="Warn" Color="#F5A623" />
        <SolidColorBrush x:Key="Bad" Color="#E74C3C" />
        <SolidColorBrush x:Key="Muted" Color="#A0A7B4" />
        <Style TargetType="Button">
            <Setter Property="Margin" Value="3" />
            <Setter Property="Padding" Value="8,5" />
            <Setter Property="Background" Value="#263041" />
            <Setter Property="Foreground" Value="#F2F4F8" />
            <Setter Property="BorderBrush" Value="#3B4657" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="FontSize" Value="11" />
            <Setter Property="MinHeight" Value="28" />
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#101215" />
            <Setter Property="Foreground" Value="#F2F4F8" />
            <Setter Property="BorderBrush" Value="#2B3138" />
            <Setter Property="FontFamily" Value="Consolas" />
        </Style>
    </Window.Resources>
    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="150" />
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="hermes-agent-windows" FontSize="26" FontWeight="Bold" Foreground="#F8FAFC" />
            <TextBlock Text="Smart Windows Setup Tool for Hermes Agent" FontSize="13" Foreground="#A0A7B4" Margin="0,4,0,0" />
        </StackPanel>

        <Border Grid.Row="1" Background="#171B20" BorderBrush="#2B3138" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,10">
            <StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="StartSetupButton" Content="Start Hermes Agent Setup" Width="205" Height="36" Background="#2D7FF9" />
                    <Button x:Name="InstallAppButton" Content="Install hermes-agent-windows Shortcut" Width="180" Height="36" />
                    <Button x:Name="UninstallAppButton" Content="Uninstall Shortcut" Width="140" Height="36" Background="#4A2630" />
                    <TextBlock x:Name="TopStatusText" VerticalAlignment="Center" Margin="16,0,0,0" Foreground="#A0A7B4" Text="Ready." FontSize="14" />
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                    <TextBlock Text="Ollama API Key" VerticalAlignment="Center" Foreground="#A0A7B4" Margin="4,0,8,0" />
                    <PasswordBox x:Name="OllamaApiKeyBox" Width="260" Height="30" Background="#101215" Foreground="#F2F4F8" BorderBrush="#2B3138" />
                    <TextBlock Text="Model" VerticalAlignment="Center" Foreground="#A0A7B4" Margin="12,0,8,0" />
                    <ComboBox x:Name="OllamaModelBox" Width="230" Height="30" IsEditable="True" Text="kimi-k2.6:cloud" Background="#101215" Foreground="#111418" />
                    <Button x:Name="SaveOllamaCloudButton" Content="Save Cloud API" Width="120" Height="30" />
                    <Button x:Name="RefreshOllamaModelsButton" Content="Refresh Models" Width="120" Height="30" />
                    <Button x:Name="TestOllamaCloudButton" Content="Test Cloud API" Width="120" Height="30" />
                </StackPanel>
            </StackPanel>
        </Border>

        <ScrollViewer Grid.Row="2" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Auto" MaxHeight="270">
            <UniformGrid Columns="3" Rows="5" Margin="0,0,0,8">
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="hermes-agent-windows App" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="AppStatusValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="AppStatusDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Admin Check" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="AdminValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="AdminDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Windows Version" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="WindowsValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="WindowsDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="PowerShell Version" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="PowerShellValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="PowerShellDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="WSL Status" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="WslStatusValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="WslStatusDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="WSL Distro" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="WslDistroValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="WslDistroDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="WSL Account" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="WslAccountValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="WslAccountDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Ollama Status" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="OllamaStatusValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="OllamaStatusDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Ollama Version" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="OllamaVersionValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="OllamaVersionDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Hermes Agent Status" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="HermesStatusValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="HermesStatusDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Hermes Agent Version" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="HermesVersionValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="HermesVersionDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Hermes Gateway" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="GatewayValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="GatewayDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
                <Border Background="{StaticResource CardFill}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Margin="0,0,8,8" Padding="9">
                    <StackPanel>
                        <TextBlock Text="Updates" Foreground="{StaticResource Accent}" FontSize="12" FontWeight="SemiBold" />
                        <TextBlock x:Name="UpdatesValue" Text="Unknown" FontSize="16" FontWeight="Bold" Margin="0,4,0,0" />
                        <TextBlock x:Name="UpdatesDetail" Text="Waiting for status check." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,3,0,0" FontSize="11" MaxHeight="42" TextTrimming="CharacterEllipsis" />
                    </StackPanel>
                </Border>
            </UniformGrid>
        </ScrollViewer>

        <Border Grid.Row="3" Background="#111418" BorderBrush="#2B3138" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,10" MinHeight="150">
            <DockPanel>
                <TextBlock Text="Live Log" DockPanel.Dock="Top" FontSize="13" FontWeight="SemiBold" Foreground="#A0A7B4" Margin="0,0,0,6" />
                <TextBox x:Name="LogBox"
                         AcceptsReturn="True"
                         IsReadOnly="True"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         FontSize="12" />
            </DockPanel>
        </Border>

        <Border Grid.Row="4" Background="#171B20" BorderBrush="#2B3138" BorderThickness="1" CornerRadius="8" Padding="8">
            <DockPanel LastChildFill="True">
                <DockPanel DockPanel.Dock="Bottom" LastChildFill="False">
                    <TextBlock x:Name="BottomStatus" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="4,4,0,0" Foreground="#A0A7B4" FontSize="12" />
                    <Button x:Name="GitHubButton" Content="Built by jlaiii" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,4,4,0" Background="Transparent" BorderBrush="Transparent" Foreground="#607080" FontSize="11" FontFamily="Consolas" Padding="3,1" />
                </DockPanel>
                <WrapPanel VerticalAlignment="Top">
                <Button x:Name="CheckStatusButton" Content="Check Status" />
                <Button x:Name="InstallWslButton" Content="Install WSL" />
                <Button x:Name="RestartWslButton" Content="Restart WSL" />
                <Button x:Name="AdminAccountButton" Content="Create/Reset WSL Admin" />
                <Button x:Name="InstallOllamaButton" Content="Install Ollama in WSL" />
                <Button x:Name="StartOllamaButton" Content="Start Ollama" />
                <Button x:Name="InstallHermesButton" Content="Install Hermes in WSL" />
                <Button x:Name="UpdateHermesButton" Content="Update Hermes Agent" />
                <Button x:Name="HermesDoctorButton" Content="Hermes Doctor" />
                <Button x:Name="LaunchHermesCliButton" Content="Launch Hermes CLI" />
                <Button x:Name="EnableGatewayButton" Content="Enable Gateway" />
                <Button x:Name="OpenGatewayButton" Content="Open Dashboard" />
                <Button x:Name="StartHermesButton" Content="Start Hermes Agent" />
                <Button x:Name="StopHermesButton" Content="Stop Hermes Agent" />
                <Button x:Name="RestartHermesButton" Content="Restart Hermes Agent" />
                <Button x:Name="CleanHermesButton" Content="Clean Hermes WSL Files" />
                <Button x:Name="ReinstallWslButton" Content="Reinstall WSL" Background="#5A2A1A" />
                <Button x:Name="WipeWslButton" Content="Wipe WSL" Background="#6B1F1F" />
                <Button x:Name="OpenConfigButton" Content="Open Config Folder" />
                <Button x:Name="OpenLogsButton" Content="Open Logs Folder" />
                <Button x:Name="ExitButton" Content="Exit" />
                </WrapPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $controls = @{
        StartSetupButton    = $window.FindName('StartSetupButton')
        InstallAppButton    = $window.FindName('InstallAppButton')
        UninstallAppButton  = $window.FindName('UninstallAppButton')
        OllamaApiKeyBox     = $window.FindName('OllamaApiKeyBox')
        OllamaModelBox      = $window.FindName('OllamaModelBox')
        SaveOllamaCloudButton = $window.FindName('SaveOllamaCloudButton')
        RefreshOllamaModelsButton = $window.FindName('RefreshOllamaModelsButton')
        TestOllamaCloudButton = $window.FindName('TestOllamaCloudButton')
        TopStatusText       = $window.FindName('TopStatusText')
        AppStatusValue      = $window.FindName('AppStatusValue')
        AppStatusDetail     = $window.FindName('AppStatusDetail')
        AdminValue          = $window.FindName('AdminValue')
        AdminDetail         = $window.FindName('AdminDetail')
        WindowsValue        = $window.FindName('WindowsValue')
        WindowsDetail       = $window.FindName('WindowsDetail')
        PowerShellValue     = $window.FindName('PowerShellValue')
        PowerShellDetail    = $window.FindName('PowerShellDetail')
        WslStatusValue      = $window.FindName('WslStatusValue')
        WslStatusDetail     = $window.FindName('WslStatusDetail')
        WslDistroValue      = $window.FindName('WslDistroValue')
        WslDistroDetail     = $window.FindName('WslDistroDetail')
        WslAccountValue     = $window.FindName('WslAccountValue')
        WslAccountDetail    = $window.FindName('WslAccountDetail')
        OllamaStatusValue   = $window.FindName('OllamaStatusValue')
        OllamaStatusDetail  = $window.FindName('OllamaStatusDetail')
        OllamaVersionValue  = $window.FindName('OllamaVersionValue')
        OllamaVersionDetail = $window.FindName('OllamaVersionDetail')
        HermesStatusValue   = $window.FindName('HermesStatusValue')
        HermesStatusDetail  = $window.FindName('HermesStatusDetail')
        HermesVersionValue  = $window.FindName('HermesVersionValue')
        HermesVersionDetail = $window.FindName('HermesVersionDetail')
        GatewayValue        = $window.FindName('GatewayValue')
        GatewayDetail       = $window.FindName('GatewayDetail')
        UpdatesValue        = $window.FindName('UpdatesValue')
        UpdatesDetail       = $window.FindName('UpdatesDetail')
        LogBox              = $window.FindName('LogBox')
        CheckStatusButton   = $window.FindName('CheckStatusButton')
        InstallWslButton    = $window.FindName('InstallWslButton')
        RestartWslButton    = $window.FindName('RestartWslButton')
        AdminAccountButton  = $window.FindName('AdminAccountButton')
        InstallOllamaButton = $window.FindName('InstallOllamaButton')
        StartOllamaButton   = $window.FindName('StartOllamaButton')
        InstallHermesButton = $window.FindName('InstallHermesButton')
        UpdateHermesButton  = $window.FindName('UpdateHermesButton')
        HermesDoctorButton  = $window.FindName('HermesDoctorButton')
        LaunchHermesCliButton = $window.FindName('LaunchHermesCliButton')
        EnableGatewayButton = $window.FindName('EnableGatewayButton')
        OpenGatewayButton   = $window.FindName('OpenGatewayButton')
        StartHermesButton   = $window.FindName('StartHermesButton')
        StopHermesButton    = $window.FindName('StopHermesButton')
        RestartHermesButton = $window.FindName('RestartHermesButton')
        CleanHermesButton   = $window.FindName('CleanHermesButton')
        ReinstallWslButton  = $window.FindName('ReinstallWslButton')
        WipeWslButton       = $window.FindName('WipeWslButton')
        OpenConfigButton    = $window.FindName('OpenConfigButton')
        OpenLogsButton      = $window.FindName('OpenLogsButton')
        ExitButton          = $window.FindName('ExitButton')
        GitHubButton        = $window.FindName('GitHubButton')
        BottomStatus        = $window.FindName('BottomStatus')
    }

    $script:GuiState = @{
        ProjectRoot = $projectRoot
        LastLogLineCount = 0
        Jobs = @{}
        Window = $window
        Controls = $controls
    }

    function Add-GuiLogLine {
        param([string]$Text)
        if ([string]::IsNullOrWhiteSpace($Text)) {
            return
        }

        $line = ConvertTo-GuiSafeText $Text
        if (-not $line) {
            return
        }

        $controls.LogBox.AppendText($line + [Environment]::NewLine)
        $controls.LogBox.ScrollToEnd()
    }

    function Set-StatusVisual {
        param(
            [System.Windows.Controls.TextBlock]$ValueControl,
            [System.Windows.Controls.TextBlock]$DetailControl,
            [string]$Status,
            [string]$Message,
            [string]$Details
        )

        $ValueControl.Text = if ($Status) { $Status } else { 'Unknown' }
        $DetailControl.Text = if ($Details) { "$Message`n$Details" } else { $Message }

        switch ($Status) {
            'Installed' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::LightGreen }
            'Running' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::LightGreen }
            'Stopped' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::Khaki }
            'Missing' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::OrangeRed }
            'Needs Update' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::Gold }
            'NeedsReboot' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::LightSalmon }
            'Error' { $ValueControl.Foreground = [System.Windows.Media.Brushes]::Tomato }
            default { $ValueControl.Foreground = [System.Windows.Media.Brushes]::LightGray }
        }
    }

    function Update-FromSummary {
        param([pscustomobject]$Summary)
        if (-not $Summary) { return }

        Set-StatusVisual $controls.AppStatusValue $controls.AppStatusDetail $Summary.AppStatus.Status $Summary.AppStatus.Message $Summary.AppStatus.Details
        Set-StatusVisual $controls.AdminValue $controls.AdminDetail $Summary.AdminCheck.Status $Summary.AdminCheck.Message $Summary.AdminCheck.Details
        Set-StatusVisual $controls.WindowsValue $controls.WindowsDetail $Summary.WindowsVersion.Status $Summary.WindowsVersion.Message ($Summary.WindowsVersion.Message)
        Set-StatusVisual $controls.PowerShellValue $controls.PowerShellDetail $Summary.PowerShellVersion.Status $Summary.PowerShellVersion.Message ($Summary.PowerShellVersion.Message)
        Set-StatusVisual $controls.WslStatusValue $controls.WslStatusDetail $Summary.WslStatus.Status $Summary.WslStatus.Message $Summary.WslStatus.Details
        Set-StatusVisual $controls.WslDistroValue $controls.WslDistroDetail $Summary.WslDistro.Status $Summary.WslDistro.Message $Summary.WslDistro.Details
        Set-StatusVisual $controls.WslAccountValue $controls.WslAccountDetail $Summary.WslAccount.Status $Summary.WslAccount.Message $Summary.WslAccount.Details
        Set-StatusVisual $controls.OllamaStatusValue $controls.OllamaStatusDetail $Summary.OllamaStatus.Status $Summary.OllamaStatus.Message $Summary.OllamaStatus.Details
        Set-StatusVisual $controls.OllamaVersionValue $controls.OllamaVersionDetail $Summary.OllamaVersion.Status $Summary.OllamaVersion.Message $Summary.OllamaVersion.Details
        Set-StatusVisual $controls.HermesStatusValue $controls.HermesStatusDetail $Summary.HermesStatus.Status $Summary.HermesStatus.Message $Summary.HermesStatus.Details
        Set-StatusVisual $controls.HermesVersionValue $controls.HermesVersionDetail $Summary.HermesVersion.Status $Summary.HermesVersion.Message $Summary.HermesVersion.Details
        Set-StatusVisual $controls.GatewayValue $controls.GatewayDetail $Summary.GatewayStatus.Status $Summary.GatewayStatus.Message $Summary.GatewayStatus.Details
        Set-StatusVisual $controls.UpdatesValue $controls.UpdatesDetail $Summary.Updates.Status $Summary.Updates.Message $Summary.Updates.Details
        $controls.BottomStatus.Text = $Summary.Summary
        $controls.TopStatusText.Text = 'Status refreshed.'
    }

    function Refresh-AppLogTail {
        $logFile = Get-LogFilePath -Kind 'app'
        if (-not (Test-Path $logFile)) {
            return
        }

        try {
            $lines = @(Get-Content -Path $logFile -ErrorAction SilentlyContinue)
            if ($null -eq $lines) { return }
            if ($lines.Count -gt $script:GuiState.LastLogLineCount) {
                $newLines = $lines[$script:GuiState.LastLogLineCount..($lines.Count - 1)]
                foreach ($line in $newLines) {
                    Add-GuiLogLine $line
                }
                $script:GuiState.LastLogLineCount = $lines.Count
            }
        }
        catch {
        }
    }

    function Start-GuiJob {
        param(
            [string]$TaskName,
            [string]$FunctionName,
            [object[]]$FunctionArguments = @()
        )

        if ($script:GuiState.Jobs.ContainsKey($TaskName)) {
            Add-GuiLogLine "[$TaskName] Task already running."
            return
        }

        Write-AppLog -Message "Starting task: $TaskName" -Level 'INFO' | Out-Null
        $controls.TopStatusText.Text = "$TaskName in progress..."
        $controls.BottomStatus.Text = "$TaskName started."

        $job = Start-Job -Name $TaskName -ArgumentList $script:GuiState.ProjectRoot, $FunctionName, $FunctionArguments -ScriptBlock {
            param($root, $funcName, $funcArgs)
            $global:HermesAgentWindowsRoot = $root
            $global:HermesAgentWindowsLogPath = Join-Path $root 'logs\app.log'
            . (Join-Path $root 'src\utils.ps1')
            . (Join-Path $root 'src\checks.ps1')
            . (Join-Path $root 'src\wsl-manager.ps1')
            . (Join-Path $root 'src\ollama-manager.ps1')
            . (Join-Path $root 'src\hermes-manager.ps1')
            . (Join-Path $root 'src\app-manager.ps1')
            . (Join-Path $root 'src\installer.ps1')
            & $funcName @funcArgs
        }

        $script:GuiState.Jobs[$TaskName] = $job
    }

    function Start-StatusCheckJob {
        Start-GuiJob -TaskName 'CheckStatus' -FunctionName 'Invoke-StatusCheck'
    }

    function Start-InstallAppJob {
        Start-GuiJob -TaskName 'InstallApp' -FunctionName 'Install-hermes-agent-windowsApp'
    }

    function Start-UninstallAppJob {
        Start-GuiJob -TaskName 'UninstallApp' -FunctionName 'Uninstall-hermes-agent-windowsApp'
    }

    function Start-InstallWslJob {
        Start-GuiJob -TaskName 'InstallWsl' -FunctionName 'Install-Wsl'
    }

    function Start-RestartWslJob {
        Start-GuiJob -TaskName 'RestartWsl' -FunctionName 'Restart-Wsl'
    }

    function Start-AdminAccountJob {
        Start-GuiJob -TaskName 'AdminAccount' -FunctionName 'Ensure-WslAdminAccount'
    }

    function Start-InstallOllamaJob {
        Start-GuiJob -TaskName 'InstallOllama' -FunctionName 'Install-Ollama'
    }

    function Start-OllamaJob {
        Start-GuiJob -TaskName 'StartOllama' -FunctionName 'Start-Ollama'
    }

    function Start-SaveOllamaCloudJob {
        $apiKey = $controls.OllamaApiKeyBox.Password
        $model = $controls.OllamaModelBox.Text
        Start-GuiJob -TaskName 'SaveOllamaCloud' -FunctionName 'Set-OllamaCloudConfig' -FunctionArguments @($apiKey, $model)
    }

    function Start-RefreshOllamaModelsJob {
        Start-GuiJob -TaskName 'RefreshOllamaModels' -FunctionName 'Get-OllamaCloudModels'
    }

    function Start-TestOllamaCloudJob {
        $model = $controls.OllamaModelBox.Text
        Start-GuiJob -TaskName 'TestOllamaCloud' -FunctionName 'Test-OllamaCloudApi' -FunctionArguments @($model)
    }

    function Start-InstallHermesJob {
        Start-GuiJob -TaskName 'InstallHermes' -FunctionName 'Install-HermesAgent'
    }

    function Start-UpdateHermesJob {
        Start-GuiJob -TaskName 'UpdateHermes' -FunctionName 'Update-HermesAgent'
    }

    function Start-HermesDoctorJob {
        Start-GuiJob -TaskName 'HermesDoctor' -FunctionName 'Invoke-HermesDoctor'
    }

    function Start-EnableGatewayJob {
        Start-GuiJob -TaskName 'EnableGateway' -FunctionName 'Enable-HermesGateway'
    }

    function Start-CleanHermesJob {
        Start-GuiJob -TaskName 'CleanHermes' -FunctionName 'Clear-HermesWslFiles'
    }

    function Start-ReinstallWslJob {
        Start-GuiJob -TaskName 'ReinstallWsl' -FunctionName 'Reinstall-DefaultWslDistro'
    }

    function Start-WipeWslJob {
        Start-GuiJob -TaskName 'WipeWsl' -FunctionName 'Unregister-DefaultWslDistro'
    }

    function Start-FullSetupJob {
        Start-GuiJob -TaskName 'FullSetup' -FunctionName 'Start-FullSetup'
    }

    function Start-HermesControlJob {
        param(
            [string]$TaskName,
            [string]$FunctionName
        )
        Start-GuiJob -TaskName $TaskName -FunctionName $FunctionName
    }

    function Start-GuidedSetup {
        Add-GuiLogLine 'Starting full automated setup.'
        Start-FullSetupJob
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(700)
    $timer.Add_Tick({
        Refresh-AppLogTail

        foreach ($entry in @($script:GuiState.Jobs.GetEnumerator())) {
            $taskName = $entry.Key
            $job = $entry.Value
            if ($job.State -in 'Completed', 'Failed', 'Stopped') {
                try {
                    $output = Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue
                    $outputItems = @($output)
                    if ($outputItems.Count -gt 0) {
                        foreach ($item in $outputItems) {
                            if ($item -is [string]) {
                                Add-GuiLogLine $item
                            }
                        }
                    }

                    if ($outputItems.Count -gt 0 -and $outputItems[0] -is [pscustomobject]) {
                        $result = $outputItems | Select-Object -First 1
                        if ($taskName -eq 'CheckStatus') {
                            Update-FromSummary $result
                        }
                        elseif ($taskName -eq 'RefreshOllamaModels') {
                            $controls.TopStatusText.Text = "$taskName finished with $($result.Status)."
                            Add-GuiLogLine "[$taskName] $($result.Message)"
                            if ($result.Models) {
                                $controls.OllamaModelBox.Items.Clear()
                                foreach ($modelName in @($result.Models)) {
                                    [void]$controls.OllamaModelBox.Items.Add($modelName)
                                }
                                if (-not $controls.OllamaModelBox.Text) {
                                    $controls.OllamaModelBox.Text = 'kimi-k2.6:cloud'
                                }
                            }
                        }
                        elseif ($taskName -eq 'SaveOllamaCloud' -or $taskName -eq 'TestOllamaCloud' -or $taskName -eq 'InstallApp' -or $taskName -eq 'UninstallApp' -or $taskName -eq 'InstallWsl' -or $taskName -eq 'RestartWsl' -or $taskName -eq 'AdminAccount' -or $taskName -eq 'InstallOllama' -or $taskName -eq 'StartOllama' -or $taskName -eq 'InstallHermes' -or $taskName -eq 'UpdateHermes' -or $taskName -eq 'HermesDoctor' -or $taskName -eq 'EnableGateway' -or $taskName -eq 'CleanHermes' -or $taskName -eq 'ReinstallWsl' -or $taskName -eq 'WipeWsl' -or $taskName -eq 'FullSetup') {
                            $controls.TopStatusText.Text = "$taskName finished with $($result.Status)."
                            Add-GuiLogLine "[$taskName] $($result.Message)"
                            Start-StatusCheckJob
                        }
                        elseif ($taskName -eq 'StartHermes' -or $taskName -eq 'StopHermes' -or $taskName -eq 'RestartHermes') {
                            $controls.TopStatusText.Text = "$taskName finished with $($result.Status)."
                            Start-StatusCheckJob
                        }
                    }
                    elseif ($taskName -eq 'CheckStatus' -and $outputItems.Count -gt 0) {
                        Update-FromSummary $outputItems[0]
                    }
                }
                finally {
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                    $script:GuiState.Jobs.Remove($taskName) | Out-Null
                }
            }
        }
    })
    $timer.Start()

    $controls.StartSetupButton.Add_Click({
        Start-GuidedSetup
    })
    $controls.InstallAppButton.Add_Click({
        Start-InstallAppJob
    })
    $controls.UninstallAppButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Remove the hermes-agent-windows Desktop and Start Menu shortcuts? This will not delete the project folder or WSL data.', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-UninstallAppJob }
    })
    $controls.CheckStatusButton.Add_Click({
        Start-StatusCheckJob
    })
    $controls.InstallWslButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Install WSL now? A reboot may be required.', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-InstallWslJob }
    })
    $controls.RestartWslButton.Add_Click({
        Start-RestartWslJob
    })
    $controls.AdminAccountButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Create or reset the WSL admin account to username admin and password admin? This is convenient for setup but should be changed later on shared machines.', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-AdminAccountJob }
    })
    $controls.InstallOllamaButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Install Ollama inside WSL now?', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-InstallOllamaJob }
    })
    $controls.StartOllamaButton.Add_Click({
        Start-OllamaJob
    })
    $controls.SaveOllamaCloudButton.Add_Click({
        Start-SaveOllamaCloudJob
    })
    $controls.RefreshOllamaModelsButton.Add_Click({
        Start-RefreshOllamaModelsJob
    })
    $controls.TestOllamaCloudButton.Add_Click({
        Start-TestOllamaCloudJob
    })
    $controls.InstallHermesButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Install Hermes Agent inside WSL now?', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-InstallHermesJob }
    })
    $controls.UpdateHermesButton.Add_Click({
        Start-UpdateHermesJob
    })
    $controls.HermesDoctorButton.Add_Click({
        Start-HermesDoctorJob
    })
    $controls.LaunchHermesCliButton.Add_Click({
        $result = Open-HermesCli
        Add-GuiLogLine "Launch Hermes CLI: $($result.Message)"
    })
    $controls.EnableGatewayButton.Add_Click({
        Start-EnableGatewayJob
    })
    $controls.OpenGatewayButton.Add_Click({
        $result = Open-HermesGateway
        Add-GuiLogLine "Open Dashboard: $($result.Message)"
    })
    $controls.StartHermesButton.Add_Click({
        Start-HermesControlJob -TaskName 'StartHermes' -FunctionName 'Start-HermesAgent'
    })
    $controls.StopHermesButton.Add_Click({
        Start-HermesControlJob -TaskName 'StopHermes' -FunctionName 'Stop-HermesAgent'
    })
    $controls.RestartHermesButton.Add_Click({
        Start-HermesControlJob -TaskName 'RestartHermes' -FunctionName 'Restart-HermesAgent'
    })
    $controls.CleanHermesButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('Remove Hermes user files inside WSL? This does not wipe the whole WSL distro.', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-CleanHermesJob }
    })
    $controls.ReinstallWslButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('This will unregister the default WSL distro and reinstall WSL. It can delete Linux files in that distro. Continue only if you have backups.', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Error)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-ReinstallWslJob }
    })
    $controls.WipeWslButton.Add_Click({
        $choice = [System.Windows.MessageBox]::Show('DANGER: This unregisters the default WSL distro and deletes its Linux filesystem. This cannot be undone. Continue?', 'hermes-agent-windows', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Error)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Start-WipeWslJob }
    })
    $controls.OpenConfigButton.Add_Click({
        $result = Open-HermesConfigFolder
        Add-GuiLogLine "Open Config Folder: $($result.Message)"
    })
    $controls.OpenLogsButton.Add_Click({
        $result = Open-FolderSafe -Path (Join-Path $script:GuiState.ProjectRoot 'logs')
        Add-GuiLogLine "Open Logs Folder: $($result.Message)"
    })
    $controls.ExitButton.Add_Click({
        $window.Close()
    })

    $controls.GitHubButton.Add_Click({
        try {
            Start-Process 'https://github.com/jlaiii/hermes-agent-windows' | Out-Null
        }
        catch {
            Add-GuiLogLine "Could not open browser: $($_.Exception.Message)"
        }
    })

    Add-GuiLogLine 'hermes-agent-windows GUI started.'
    Start-StatusCheckJob
    $window.Add_Closed({
        $timer.Stop()
        foreach ($entry in @($script:GuiState.Jobs.GetEnumerator())) {
            try {
                Stop-Job -Job $entry.Value -ErrorAction SilentlyContinue | Out-Null
                Remove-Job -Job $entry.Value -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }
    })
    $window.ShowDialog() | Out-Null
}

