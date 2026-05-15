#requires -Version 7.4
[CmdletBinding()]
param()

# GREX365 GUI launcher — WPF + PowerShell, no compilation.
# - XAML inline.
# - Background work in runspaces from a runspace pool.
# - IPC via synchronized hashtable (queues drained by DispatcherTimer).
# - UI mutated only via Dispatcher.Invoke.
#
# Architecture:
#   $SyncHash.LogQueue         <- runspaces enqueue log events
#   $SyncHash.StatusQueue      <- runspaces enqueue connection state updates
#   $SyncHash.OpStatusQueue    <- runspaces enqueue per-operation progress messages
#   DispatcherTimer every 150ms drains queues on UI thread.

if ($PSVersionTable.PSVersion.Major -lt 7) { Write-Host 'Requiere PowerShell 7+.' -ForegroundColor Red; exit 1 }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
try { Add-Type -AssemblyName System.Windows.Forms } catch {}
try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}

# Diagnostic log path: every GUI session appends here. Use to debug dispatcher errors
# and other context-loss issues.
$global:GREX365_GuiDebugLog = Join-Path ([System.IO.Path]::GetTempPath()) 'grex365-gui.debug.log'
function global:Write-GuiDebug {
    param([string]$Message,[string]$Source = 'GUI')
    try {
        $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Source, $Message
        Add-Content -LiteralPath $global:GREX365_GuiDebugLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}
Write-GuiDebug "==== GUI bootstrap started · PS $($PSVersionTable.PSVersion) ===="

# --- Module bootstrap ---

# Resolve script root robustly. PSScriptRoot is reliable when dot-sourced from a file;
# fallback to $MyInvocation; final fallback to a known test base path injected via env.
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              elseif ($env:GREX365_GUI_DIR) { $env:GREX365_GUI_DIR }
              else { throw 'No se pudo resolver la ruta del script GUI.' }

$RepoRoot     = Split-Path -Parent $ScriptRoot
$LauncherRoot = Split-Path -Parent $RepoRoot
$global:GREX365_BasePath = $RepoRoot
$ModulesPath  = Join-Path $RepoRoot 'Modules'
$ScriptsPath  = Join-Path $RepoRoot 'Scripts'

$modulesInOrder = @(
    'Logging.ps1','Console.ps1','Validation.ps1','Csv.ps1','Preferences.ps1',
    'Retry.ps1','Audit.ps1','Report.ps1','Roles.ps1','Templates.ps1','Jobs.ps1',
    'Connection.ps1','GroupResolver.ps1','CertWizard.ps1'
)
foreach ($m in $modulesInOrder) {
    $p = Join-Path $ModulesPath $m
    if (Test-Path $p) { . $p }
}

# --- Synchronized IPC state ---

$SyncHash = [hashtable]::Synchronized(@{
    LogQueue        = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    StatusQueue     = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    OpStatusQueue   = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    PanelQueue      = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    JobDoneQueue    = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    SearchResultQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    BusyCount       = 0
})

# Runspace pool sized for parallel ops without flooding Graph.
# MTA apartment: STA serializes async continuations (MSAL/Graph SDK use them
# under the hood for token acquisition). Cert-based Connect-MgGraph and
# Connect-ExchangeOnline observed to deadlock in STA pools — switching to MTA
# fixes the "Connect hangs forever" symptom.
# InitialSessionState pre-sets preferences so background jobs never block on
# confirmation/progress/warning prompts (no host UI to answer them).
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$iss.ExecutionPolicy = 'Bypass'
$iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ConfirmPreference','None',''))
$iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ProgressPreference','SilentlyContinue',''))
$iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('WarningPreference','SilentlyContinue',''))
$iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ErrorActionPreference','Continue',''))
$iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('GREX365_BasePath',$RepoRoot,''))

$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 4, $iss, $Host)
$RunspacePool.ApartmentState = 'MTA'
$RunspacePool.ThreadOptions  = 'ReuseThread'
$RunspacePool.Open()

$JobsList = New-Object System.Collections.Generic.List[object]

# --- XAML ---

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GREX365 · Microsoft 365 toolkit"
        Width="1280" Height="800" MinWidth="1000" MinHeight="600"
        WindowStartupLocation="CenterScreen"
        Background="#0F1116" Foreground="#E6E6E6" FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <SolidColorBrush x:Key="bg.dark"   Color="#0F1116"/>
    <SolidColorBrush x:Key="bg.panel"  Color="#161A22"/>
    <SolidColorBrush x:Key="bg.card"   Color="#1C2230"/>
    <SolidColorBrush x:Key="fg.dim"    Color="#8B95A6"/>
    <SolidColorBrush x:Key="fg.text"   Color="#E6E6E6"/>
    <SolidColorBrush x:Key="accent"    Color="#3B82F6"/>
    <SolidColorBrush x:Key="accent.dim" Color="#1F3055"/>
    <SolidColorBrush x:Key="ok"        Color="#22C55E"/>
    <SolidColorBrush x:Key="warn"      Color="#F59E0B"/>
    <SolidColorBrush x:Key="err"       Color="#EF4444"/>
    <SolidColorBrush x:Key="divider"   Color="#262B36"/>

    <Style TargetType="Button">
      <Setter Property="Background"  Value="{StaticResource accent}"/>
      <Setter Property="Foreground"  Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"     Value="14,8"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border CornerRadius="6" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#2563EB"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Background" Value="#384156"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="btn.ghost" TargetType="Button">
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Foreground"  Value="{StaticResource fg.text}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{StaticResource divider}"/>
      <Setter Property="Padding"     Value="12,7"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border CornerRadius="6" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#1F2533"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="sideItem" TargetType="ListBoxItem">
      <Setter Property="Padding" Value="14,10"/>
      <Setter Property="Foreground" Value="{StaticResource fg.text}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="bd" CornerRadius="6" Background="Transparent" Padding="{TemplateBinding Padding}" Margin="6,2">
              <ContentPresenter />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource accent.dim}"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1A1F2B"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#0E1320"/>
      <Setter Property="Foreground" Value="{StaticResource fg.text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource divider}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="CaretBrush" Value="{StaticResource fg.text}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border CornerRadius="4" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
              <ScrollViewer x:Name="PART_ContentHost" Padding="{TemplateBinding Padding}"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource fg.text}"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="56"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="200"/>
      <RowDefinition Height="32"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="{StaticResource bg.panel}" BorderBrush="{StaticResource divider}" BorderThickness="0,0,0,1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="260"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0">
          <Border Width="22" Height="22" CornerRadius="5" Background="{StaticResource accent}"/>
          <TextBlock Text="GREX365" Margin="10,0,0,0" FontWeight="Bold" FontSize="16" VerticalAlignment="Center"/>
          <TextBlock Text="·  M365 toolkit" Margin="8,2,0,0" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="HeaderSubtitle" Text="Welcome" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
          <Border x:Name="ChipRole" Background="{StaticResource accent.dim}" CornerRadius="4" Padding="8,2" Margin="16,0,4,0">
            <TextBlock x:Name="ChipRoleText" Text="role" Foreground="{StaticResource fg.text}" FontSize="11" FontWeight="SemiBold"/>
          </Border>
          <Border x:Name="ChipUiMode" Background="#2A3140" CornerRadius="4" Padding="8,2" Margin="4,0,4,0">
            <TextBlock x:Name="ChipUiModeText" Text="mode" Foreground="{StaticResource fg.text}" FontSize="11" FontWeight="SemiBold"/>
          </Border>
          <Border x:Name="ChipJobs" Background="#2A3140" CornerRadius="4" Padding="8,2" Margin="4,0,4,0" Visibility="Collapsed">
            <TextBlock x:Name="ChipJobsText" Text="0 jobs" Foreground="{StaticResource fg.text}" FontSize="11" FontWeight="SemiBold"/>
          </Border>
        </StackPanel>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0">
          <Button x:Name="BtnConnect"    Content="Connect"    Width="100" Margin="0,0,8,0"/>
          <Button x:Name="BtnDisconnect" Content="Disconnect" Width="100" Style="{StaticResource btn.ghost}"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Body: sidebar + content -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="260"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Sidebar -->
      <Border Grid.Column="0" Background="{StaticResource bg.panel}" BorderBrush="{StaticResource divider}" BorderThickness="0,0,1,0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Margin="20,16,20,8">
            <TextBlock Text="OPERACIONES" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold"/>
          </StackPanel>
          <ListBox x:Name="SideNav" Grid.Row="1" Background="Transparent" BorderThickness="0"
                   ItemContainerStyle="{StaticResource sideItem}">
            <ListBoxItem Tag="welcome"      IsSelected="True">Inicio</ListBoxItem>
            <ListBoxItem Tag="health">Salud del tenant</ListBoxItem>
            <ListBoxItem Tag="audit">Auditoría de identidad</ListBoxItem>
            <ListBoxItem Tag="groups">Grupos · workflow / miembros</ListBoxItem>
            <ListBoxItem Tag="export">Exportar miembros</ListBoxItem>
            <ListBoxItem Tag="permissions">Permisos de buzón (bulk)</ListBoxItem>
            <ListBoxItem Tag="convert">Convertir buzón shared→user</ListBoxItem>
            <ListBoxItem Tag="offboarding">Offboarding wizard</ListBoxItem>
            <ListBoxItem Tag="selftest">Self-test (testeo*)</ListBoxItem>
            <ListBoxItem Tag="jobs">Jobs en background</ListBoxItem>
            <ListBoxItem Tag="cert">Certificado</ListBoxItem>
            <ListBoxItem Tag="prefs">Preferencias</ListBoxItem>
          </ListBox>
        </Grid>
      </Border>

      <!-- Content -->
      <Grid Grid.Column="1">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <Grid Margin="32,24,32,24">
            <!-- Welcome panel -->
            <StackPanel x:Name="PanelWelcome">
              <TextBlock Text="Bienvenido" FontSize="22" FontWeight="Bold" Margin="0,0,0,4"/>
              <TextBlock Text="Selecciona una operación en la barra lateral. La sesión, el rol y el modo se configuran en Preferencias." Foreground="{StaticResource fg.dim}"/>
              <Grid Margin="0,24,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="16"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="16"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" Background="{StaticResource bg.card}" CornerRadius="8" Padding="16">
                  <StackPanel>
                    <TextBlock Text="Tenant" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold"/>
                    <TextBlock x:Name="CardTenant" Text="—" FontSize="18" FontWeight="Bold" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>
                <Border Grid.Column="2" Background="{StaticResource bg.card}" CornerRadius="8" Padding="16">
                  <StackPanel>
                    <TextBlock Text="Cuenta" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold"/>
                    <TextBlock x:Name="CardAccount" Text="—" FontSize="18" FontWeight="Bold" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>
                <Border Grid.Column="4" Background="{StaticResource bg.card}" CornerRadius="8" Padding="16">
                  <StackPanel>
                    <TextBlock Text="Conexión" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold"/>
                    <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                      <TextBlock x:Name="CardGraph" Text="Graph: —" Margin="0,0,12,0" FontWeight="SemiBold"/>
                      <TextBlock x:Name="CardExo"   Text="EXO: —" FontWeight="SemiBold"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
              </Grid>
              <TextBlock Text="Acciones rápidas" FontSize="14" FontWeight="SemiBold" Margin="0,28,0,8"/>
              <WrapPanel>
                <Button x:Name="QuickHealth"        Content="Salud del tenant"        Margin="0,0,8,8"/>
                <Button x:Name="QuickAudit"         Content="Identity audit"          Margin="0,0,8,8" Style="{StaticResource btn.ghost}"/>
                <Button x:Name="QuickSelfTest"      Content="Self-test (testeo*)"     Margin="0,0,8,8" Style="{StaticResource btn.ghost}"/>
                <Button x:Name="QuickToolkitCheck"  Content="Health check del toolkit" Margin="0,0,8,8" Style="{StaticResource btn.ghost}"/>
                <Button x:Name="QuickOpenLogs"      Content="Carpeta de logs"         Margin="0,0,8,8" Style="{StaticResource btn.ghost}"/>
                <Button x:Name="QuickOpenReports"   Content="Carpeta de informes"     Margin="0,0,8,8" Style="{StaticResource btn.ghost}"/>
              </WrapPanel>
            </StackPanel>

            <!-- Health panel -->
            <StackPanel x:Name="PanelHealth" Visibility="Collapsed">
              <TextBlock Text="Salud del tenant" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Snapshot read-only: licencias, cuotas, MFA, roles privilegiados, app secrets, usuarios stale, salud de servicios, avisos." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Button x:Name="BtnRunHealth" Content="Ejecutar análisis" Width="200" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="HealthStatus" Text="Listo." Foreground="{StaticResource fg.dim}"/>
            </StackPanel>

            <!-- Audit panel -->
            <StackPanel x:Name="PanelAudit" Visibility="Collapsed">
              <TextBlock Text="Identity audit" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Detecta stale members (>180d), stale guests (>90d), deshabilitados con licencia, grupos sin owner, M365/DL vacíos." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Button x:Name="BtnRunAudit" Content="Ejecutar auditoría" Width="200" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="AuditStatus" Text="Listo." Foreground="{StaticResource fg.dim}"/>
            </StackPanel>

            <!-- Groups · 3 modos -->
            <StackPanel x:Name="PanelGroups" Visibility="Collapsed">
              <TextBlock Text="Grupos · operaciones" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Elige qué quieres hacer. Cada modo tiene su propio formulario; no es necesario CSV con columna Action." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>

              <Border Background="{StaticResource bg.card}" CornerRadius="8" Padding="16" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Text="MODO" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold" Margin="0,0,0,8"/>
                  <RadioButton x:Name="GroupsModeAdd"    GroupName="GroupsMode" IsChecked="True" Margin="0,2,0,2" Foreground="{StaticResource fg.text}">
                    <TextBlock TextWrapping="Wrap"><Run FontWeight="SemiBold">1 · Añadir miembros a un grupo existente</Run> · sin crear nada · selector de grupo + CSV mínimo <Run FontStyle="Italic">Email;Id</Run></TextBlock>
                  </RadioButton>
                  <RadioButton x:Name="GroupsModeCreate" GroupName="GroupsMode" Margin="0,8,0,2" Foreground="{StaticResource fg.text}">
                    <TextBlock TextWrapping="Wrap"><Run FontWeight="SemiBold">2 · Crear grupo nuevo</Run> · DL · M365 · Mail-enabled Security · sólo crea, no añade miembros</TextBlock>
                  </RadioButton>
                  <RadioButton x:Name="GroupsModeCreateAdd" GroupName="GroupsMode" Margin="0,8,0,2" Foreground="{StaticResource fg.text}">
                    <TextBlock TextWrapping="Wrap"><Run FontWeight="SemiBold">3 · Crear grupo y añadir miembros</Run> · combinado · creación + carga inicial de miembros desde lista o CSV</TextBlock>
                  </RadioButton>
                </StackPanel>
              </Border>

              <!-- Sub-panel: añadir miembros (legacy) -->
              <StackPanel x:Name="GroupsAddBox" Margin="0,16,0,0">
                <TextBlock Text="1) Busca el grupo destino · 2) Elige uno · 3) Indica CSV de miembros · 4) Añadir." Foreground="{StaticResource fg.dim}" Margin="0,0,0,8" TextWrapping="Wrap"/>

                <TextBlock Text="Búsqueda grupo destino (correo / nombre / alias)" Foreground="{StaticResource fg.dim}" Margin="0,6,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox Grid.Column="0" x:Name="LegacyGroupSearch"/>
                  <Button  Grid.Column="1" x:Name="BtnLegacySearch" Content="Buscar" Width="120" Margin="8,0,0,0"/>
                </Grid>

                <TextBlock Text="Candidatos" Foreground="{StaticResource fg.dim}" Margin="0,12,0,4"/>
                <ListBox x:Name="LegacyCandidates" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                         BorderBrush="{StaticResource divider}" BorderThickness="1" Height="100"
                         FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>

                <TextBlock Text="CSV de miembros (Email;Id)" Foreground="{StaticResource fg.dim}" Margin="0,12,0,4"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="LegacyMembersCsv" Grid.Column="0"/>
                  <Button x:Name="LegacyBrowse" Grid.Column="1" Content="Examinar..." Width="120" Margin="8,0,0,0" Style="{StaticResource btn.ghost}"/>
                </Grid>
                <Button x:Name="BtnRunLegacyMembers" Content="Añadir miembros" Width="220" HorizontalAlignment="Left" Margin="0,14,0,0"/>
              </StackPanel>

              <!-- Sub-panel: crear grupo solo -->
              <StackPanel x:Name="GroupsCreateBox" Margin="0,16,0,0" Visibility="Collapsed">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Tipo"           Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <ComboBox  Grid.Row="0" Grid.Column="1" x:Name="CreateType"   Background="#0E1320" Foreground="White" Margin="0,0,0,8">
                    <ComboBoxItem Tag="DL" IsSelected="True">Distribution List (DL)</ComboBoxItem>
                    <ComboBoxItem Tag="M365">Microsoft 365 Group</ComboBoxItem>
                    <ComboBoxItem Tag="MailSecurity">Mail-enabled Security Group</ComboBoxItem>
                  </ComboBox>
                  <TextBlock Grid.Row="1" Grid.Column="0" Text="Correo primario" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="1" Grid.Column="1" x:Name="CreateEmail" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Nombre visible" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="2" Grid.Column="1" x:Name="CreateDisplay" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="3" Grid.Column="0" Text="Alias (opc.)"   Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="3" Grid.Column="1" x:Name="CreateAlias" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="4" Grid.Column="0" Text="Owners (M365)"  Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="4" Grid.Column="1" x:Name="CreateOwners" Margin="0,0,0,8"/>
                  <CheckBox  Grid.Row="5" Grid.Column="1" x:Name="CreateHidden" Content="Oculto del GAL" Margin="0,4,0,0"/>
                </Grid>
                <Button x:Name="BtnRunCreateGroup" Content="Crear grupo" Width="220" HorizontalAlignment="Left" Margin="0,14,0,0"/>
              </StackPanel>

              <!-- Sub-panel: crear + añadir -->
              <StackPanel x:Name="GroupsCreateAddBox" Margin="0,16,0,0" Visibility="Collapsed">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Tipo"           Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <ComboBox  Grid.Row="0" Grid.Column="1" x:Name="CreateAddType" Background="#0E1320" Foreground="White" Margin="0,0,0,8">
                    <ComboBoxItem Tag="DL" IsSelected="True">Distribution List (DL)</ComboBoxItem>
                    <ComboBoxItem Tag="M365">Microsoft 365 Group</ComboBoxItem>
                    <ComboBoxItem Tag="MailSecurity">Mail-enabled Security Group</ComboBoxItem>
                  </ComboBox>
                  <TextBlock Grid.Row="1" Grid.Column="0" Text="Correo primario" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="1" Grid.Column="1" x:Name="CreateAddEmail" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Nombre visible" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="2" Grid.Column="1" x:Name="CreateAddDisplay" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="3" Grid.Column="0" Text="Alias (opc.)"   Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="3" Grid.Column="1" x:Name="CreateAddAlias" Margin="0,0,0,8"/>
                  <TextBlock Grid.Row="4" Grid.Column="0" Text="Owners (M365)"  Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                  <TextBox   Grid.Row="4" Grid.Column="1" x:Name="CreateAddOwners" Margin="0,0,0,8"/>
                  <CheckBox  Grid.Row="5" Grid.Column="1" x:Name="CreateAddHidden" Content="Oculto del GAL" Margin="0,4,0,0"/>
                </Grid>
                <TextBlock Text="Miembros · uno por línea, o separados por coma" Foreground="{StaticResource fg.dim}" Margin="0,12,0,4"/>
                <TextBox x:Name="CreateAddMembers" AcceptsReturn="True" TextWrapping="Wrap" Height="100" VerticalScrollBarVisibility="Auto"/>
                <Button x:Name="BtnRunCreateAddGroup" Content="Crear + añadir miembros" Width="260" HorizontalAlignment="Left" Margin="0,14,0,0"/>
              </StackPanel>

              <TextBlock x:Name="GroupsStatus" Text="Modo: Añadir miembros." Foreground="{StaticResource fg.dim}" Margin="0,14,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <!-- Export members -->
            <StackPanel x:Name="PanelExport" Visibility="Collapsed">
              <TextBlock Text="Exportar miembros de grupo/DL" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="1) Busca el grupo · 2) Elige candidato · 3) Selecciona carpeta destino · 4) Exporta CSV (Email;Id)." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>

              <TextBlock Text="Búsqueda (correo / nombre / alias)" Foreground="{StaticResource fg.dim}" Margin="0,16,0,4"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="ExportSearch"/>
                <Button  Grid.Column="1" x:Name="BtnExportSearch" Content="Buscar" Width="120" Margin="8,0,0,0"/>
              </Grid>

              <TextBlock Text="Candidatos (elige uno)" Foreground="{StaticResource fg.dim}" Margin="0,14,0,4"/>
              <ListBox x:Name="ExportCandidates" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                       BorderBrush="{StaticResource divider}" BorderThickness="1" Height="120"
                       FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>

              <TextBlock Text="Carpeta destino" Foreground="{StaticResource fg.dim}" Margin="0,14,0,4"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="ExportFolder" Text="C:\Temp"/>
                <Button  Grid.Column="1" x:Name="ExportBrowse" Content="Examinar..." Width="120" Margin="8,0,0,0" Style="{StaticResource btn.ghost}"/>
              </Grid>

              <Button x:Name="BtnRunExport" Content="Exportar CSV" Width="220" HorizontalAlignment="Left" Margin="0,16,0,0"/>
              <TextBlock x:Name="ExportStatus" Text="Listo. Conecta primero si Graph/EXO están desconectados." Foreground="{StaticResource fg.dim}" Margin="0,12,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <!-- Convert mailbox -->
            <StackPanel x:Name="PanelConvert" Visibility="Collapsed">
              <TextBlock Text="Convertir buzón SharedMailbox → UserMailbox" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Convierte el tipo del buzón y espera la sincronización de Exchange Online (30–90s). Requiere licencia de Teams asignada después para que el usuario aparezca." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Grid Margin="0,16,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="160"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Email / UPN" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
                <TextBox   Grid.Column="1" x:Name="ConvertUpn" Margin="0,0,0,4"/>
              </Grid>
              <CheckBox x:Name="ConvertForce" Content="Saltar confirmación 'No empieza por testeo'" Margin="160,8,0,0"/>
              <Button x:Name="BtnRunConvert" Content="Convertir buzón" Width="220" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="ConvertStatus" Text="Listo." Foreground="{StaticResource fg.dim}"/>
            </StackPanel>

            <!-- Self-test -->
            <StackPanel x:Name="PanelSelfTest" Visibility="Collapsed">
              <TextBlock Text="Self-test sobre objetos testeo*" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Crea una DL temporal, añade miembros, prueba permisos, oculta GAL y limpia. Nunca toca objetos existentes." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Grid Margin="0,16,0,0">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="160"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Target UPN" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                <TextBox   Grid.Row="0" Grid.Column="1" x:Name="SelfTestTarget" Text="testeo224@es.andersen.com" Margin="0,0,0,8"/>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Delegate UPN" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,8"/>
                <TextBox   Grid.Row="1" Grid.Column="1" x:Name="SelfTestDelegate" Text="testeo6@es.andersen.com" Margin="0,0,0,8"/>
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Seed nombre DL" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
                <TextBox   Grid.Row="2" Grid.Column="1" x:Name="SelfTestSeed" Text="testeo-selftest"/>
              </Grid>
              <CheckBox x:Name="SelfTestSkipCleanup" Content="Saltar limpieza (deja la DL para inspección)" Margin="160,12,0,0"/>
              <Button x:Name="BtnRunSelfTest" Content="Ejecutar self-test" Width="220" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="SelfTestStatus" Text="Listo." Foreground="{StaticResource fg.dim}"/>
            </StackPanel>

            <!-- Jobs viewer -->
            <StackPanel x:Name="PanelJobs" Visibility="Collapsed">
              <TextBlock Text="Jobs en background" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Jobs ejecutados desde la GUI durante esta sesión. Se autocull tras completarse o fallar." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <StackPanel Orientation="Horizontal" Margin="0,16,0,8">
                <Button x:Name="BtnJobsRefresh" Content="Refrescar" Width="140" Style="{StaticResource btn.ghost}" Margin="0,0,8,0"/>
                <Button x:Name="BtnJobsClear"   Content="Limpiar terminados" Width="180" Style="{StaticResource btn.ghost}"/>
              </StackPanel>
              <DataGrid x:Name="JobsGrid" AutoGenerateColumns="False" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                        BorderThickness="0" RowBackground="#1C2230" AlternatingRowBackground="#161A22"
                        HeadersVisibility="Column" GridLinesVisibility="None" IsReadOnly="True"
                        MinHeight="200">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Job"      Binding="{Binding Name}"    Width="*"/>
                  <DataGridTextColumn Header="Estado"   Binding="{Binding State}"   Width="120"/>
                  <DataGridTextColumn Header="Iniciado" Binding="{Binding Started}" Width="180"/>
                  <DataGridTextColumn Header="Edad"     Binding="{Binding Age}"     Width="100"/>
                </DataGrid.Columns>
              </DataGrid>
              <TextBlock x:Name="JobsStatus" Text="" Foreground="{StaticResource fg.dim}" Margin="0,12,0,0"/>
            </StackPanel>

            <!-- Certificate panel -->
            <StackPanel x:Name="PanelCert" Visibility="Collapsed">
              <TextBlock Text="Certificado · ExO + Graph (app-only)" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Estado actual del certificado y App Registration. El asistente debe ejecutarse desde consola (29 pasos interactivos)." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Border Background="{StaticResource bg.card}" CornerRadius="8" Padding="16" Margin="0,16,0,0">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Estado"       Foreground="{StaticResource fg.dim}" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="0" Grid.Column="1" x:Name="CertState"  Text="—" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="1" Grid.Column="0" Text="AppId"        Foreground="{StaticResource fg.dim}" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="1" Grid.Column="1" x:Name="CertAppId"  Text="—" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Thumbprint"   Foreground="{StaticResource fg.dim}" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="2" Grid.Column="1" x:Name="CertThumb"  Text="—" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="3" Grid.Column="0" Text="Tenant"       Foreground="{StaticResource fg.dim}" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="3" Grid.Column="1" x:Name="CertTenant" Text="—" Margin="0,0,0,6"/>
                  <TextBlock Grid.Row="4" Grid.Column="0" Text="Organización" Foreground="{StaticResource fg.dim}"/>
                  <TextBlock Grid.Row="4" Grid.Column="1" x:Name="CertOrg"    Text="—"/>
                </Grid>
              </Border>
              <StackPanel Orientation="Horizontal" Margin="0,16,0,0">
                <Button x:Name="BtnCertRefresh" Content="Refrescar" Width="140" Style="{StaticResource btn.ghost}" Margin="0,0,8,0"/>
                <Button x:Name="BtnCertOpenConsole" Content="Abrir asistente (consola)" Width="220" Margin="0,0,8,0"/>
                <Button x:Name="BtnCertOpenFolder" Content="Carpeta config" Width="160" Style="{StaticResource btn.ghost}"/>
              </StackPanel>
              <TextBlock x:Name="CertStatus" Text="" Foreground="{StaticResource fg.dim}" Margin="0,12,0,0"/>
            </StackPanel>


            <!-- Permissions -->
            <StackPanel x:Name="PanelPermissions" Visibility="Collapsed">
              <TextBlock Text="Permisos de buzón (bulk)" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="CSV: Action;Permission;Mailbox;Principal — FullAccess / SendAs / SendOnBehalf · add / remove." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
              <Grid Margin="0,16,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="PermsCsvPath" Grid.Column="0" Text=""/>
                <Button x:Name="PermsBrowse" Grid.Column="1" Content="Examinar..." Width="120" Margin="8,0,0,0" Style="{StaticResource btn.ghost}"/>
              </Grid>
              <Button x:Name="BtnRunPerms" Content="Aplicar permisos" Width="200" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="PermsStatus" Text="Listo." Foreground="{StaticResource fg.dim}"/>
            </StackPanel>

            <!-- Offboarding -->
            <StackPanel x:Name="PanelOffboarding" Visibility="Collapsed">
              <TextBlock Text="Offboarding wizard" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="14 pasos: bloquear sign-in · revocar sesiones · quitar licencias · convertir buzón a shared · auto-reply · forward · FullAccess · SendAs · ocultar GAL · quitar DLs/M365 · MFA · handover · informe HTML." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>

              <!-- Usuario saliente -->
              <TextBlock Text="Correo del usuario saliente" Foreground="{StaticResource fg.dim}" Margin="0,16,0,4" FontWeight="SemiBold"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="OffUserSearch"/>
                <Button  Grid.Column="1" x:Name="BtnOffUserSearch" Content="Buscar" Width="110" Margin="8,0,0,0"/>
              </Grid>
              <ListBox x:Name="OffUserCandidates" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                       BorderBrush="{StaticResource divider}" BorderThickness="1" Height="90" Margin="0,4,0,0"
                       FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>

              <!-- Delegados (multi) -->
              <TextBlock Text="Delegados que heredan el buzón (uno o varios)" Foreground="{StaticResource fg.dim}" Margin="0,14,0,4" FontWeight="SemiBold"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="OffDelegateSearch"/>
                <Button  Grid.Column="1" x:Name="BtnOffDelegateSearch" Content="Buscar" Width="110" Margin="8,0,0,0"/>
              </Grid>
              <ListBox x:Name="OffDelegateCandidates" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                       BorderBrush="{StaticResource divider}" BorderThickness="1" Height="90" Margin="0,4,0,0"
                       FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,4">
                <Button x:Name="BtnAddDelegate"    Content="Añadir →"      Width="120" Style="{StaticResource btn.ghost}" Margin="0,0,8,0"/>
                <Button x:Name="BtnRemoveDelegate" Content="Quitar selección" Width="160" Style="{StaticResource btn.ghost}"/>
              </StackPanel>
              <ListBox x:Name="OffDelegateChosen" Background="{StaticResource bg.card}" Foreground="LightGreen"
                       BorderBrush="{StaticResource divider}" BorderThickness="1" Height="70" Margin="0,0,0,0"
                       FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>

              <!-- Manager opcional -->
              <TextBlock Text="Manager (opcional, recibe la nota de handover)" Foreground="{StaticResource fg.dim}" Margin="0,14,0,4" FontWeight="SemiBold"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="OffManagerSearch"/>
                <Button  Grid.Column="1" x:Name="BtnOffManagerSearch" Content="Buscar" Width="110" Margin="8,0,0,0"/>
              </Grid>
              <ListBox x:Name="OffManagerCandidates" Background="{StaticResource bg.card}" Foreground="{StaticResource fg.text}"
                       BorderBrush="{StaticResource divider}" BorderThickness="1" Height="60" Margin="0,4,0,0"
                       FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"/>

              <!-- Idioma + tipo -->
              <Grid Margin="0,14,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="180"/>
                  <ColumnDefinition Width="200"/>
                  <ColumnDefinition Width="180"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Idioma plantilla" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
                <ComboBox  Grid.Column="1" x:Name="OffLang" Background="#0E1320" Foreground="White">
                  <ComboBoxItem Tag="es" IsSelected="True">Español</ComboBoxItem>
                  <ComboBoxItem Tag="en">English</ComboBoxItem>
                  <ComboBoxItem Tag="pt">Português</ComboBoxItem>
                </ComboBox>
                <TextBlock Grid.Column="2" Text="Tipo offboarding" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="16,0,0,0"/>
                <ComboBox  Grid.Column="3" x:Name="OffKind" Background="#0E1320" Foreground="White">
                  <ComboBoxItem Tag="permanent" IsSelected="True">Permanente (salida definitiva)</ComboBoxItem>
                  <ComboBoxItem Tag="temporary">Temporal (baja médica, sabático)</ComboBoxItem>
                </ComboBox>
              </Grid>

              <!-- Mensaje personalizado -->
              <TextBlock Text="Mensaje del auto-reply (opcional · sobrescribe la plantilla)" Foreground="{StaticResource fg.dim}" Margin="0,16,0,4" FontWeight="SemiBold"/>
              <TextBox x:Name="OffCustomBody" AcceptsReturn="True" TextWrapping="Wrap" Height="100"
                       VerticalScrollBarVisibility="Auto"
                       Text=""/>
              <TextBlock Text="Placeholders disponibles: {user}, {delegates}, {manager}, {date}. Si dejas vacío se usa la plantilla del idioma elegido." Foreground="{StaticResource fg.dim}" FontSize="11" Margin="0,4,0,0"/>

              <CheckBox x:Name="OffDryRun" Content="Forzar dry-run (recomendado)" IsChecked="True" Margin="0,12,0,0"/>

              <Button x:Name="BtnRunOff" Content="Ejecutar offboarding" Width="240" HorizontalAlignment="Left" Margin="0,16,0,16"/>
              <TextBlock x:Name="OffStatus" Text="Listo. Busca y selecciona el usuario, añade al menos un delegado, opcionalmente manager, y pulsa Ejecutar." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>
            </StackPanel>

            <!-- Preferences -->
            <StackPanel x:Name="PanelPrefs" Visibility="Collapsed">
              <TextBlock Text="Preferencias" FontSize="22" FontWeight="Bold"/>
              <TextBlock Text="Cambia método de conexión, rol y UI mode. Persiste en config/user_preferences.json." Foreground="{StaticResource fg.dim}" TextWrapping="Wrap"/>

              <Grid Margin="0,20,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="200"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Método activo" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,10"/>
                <ComboBox  Grid.Row="0" Grid.Column="1" x:Name="PrefMethod" Margin="0,0,0,10" Background="#0E1320" Foreground="White">
                  <ComboBoxItem Tag="cert">Certificate (app-only)</ComboBoxItem>
                  <ComboBoxItem Tag="traditional">Traditional (device code)</ComboBoxItem>
                </ComboBox>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Rol" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,10"/>
                <ComboBox  Grid.Row="1" Grid.Column="1" x:Name="PrefRole" Margin="0,0,0,10" Background="#0E1320" Foreground="White">
                  <ComboBoxItem Tag="viewer">viewer (solo lectura)</ComboBoxItem>
                  <ComboBoxItem Tag="operator">operator</ComboBoxItem>
                  <ComboBoxItem Tag="admin">admin (offboarding, bulk destructivo)</ComboBoxItem>
                </ComboBox>
                <TextBlock Grid.Row="2" Grid.Column="0" Text="UI Mode" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center" Margin="0,0,0,10"/>
                <ComboBox  Grid.Row="2" Grid.Column="1" x:Name="PrefUiMode" Margin="0,0,0,10" Background="#0E1320" Foreground="White">
                  <ComboBoxItem Tag="support">support (dry-run forzado)</ComboBoxItem>
                  <ComboBoxItem Tag="advanced">advanced</ComboBoxItem>
                </ComboBox>
                <Button Grid.Row="3" Grid.Column="1" x:Name="BtnSavePrefs" Content="Guardar preferencias" Width="200" HorizontalAlignment="Left"/>
              </Grid>
              <TextBlock x:Name="PrefStatus" Text="" Foreground="{StaticResource ok}" Margin="0,12,0,0"/>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </Grid>
    </Grid>

    <!-- Log -->
    <Border Grid.Row="2" Background="{StaticResource bg.panel}" BorderBrush="{StaticResource divider}" BorderThickness="0,1,0,1">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0" Margin="20,8,20,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Text="LOG EN VIVO" Foreground="{StaticResource fg.dim}" FontSize="11" FontWeight="Bold" VerticalAlignment="Center"/>
          <TextBlock x:Name="OpStatusText" Grid.Column="1" Text="" Foreground="{StaticResource fg.dim}" Margin="16,0,0,0" VerticalAlignment="Center"/>
          <Button   Grid.Column="2" x:Name="BtnClearLog" Content="Limpiar"  Width="100" Style="{StaticResource btn.ghost}"/>
        </Grid>
        <ListBox x:Name="LogList" Grid.Row="1" Background="#0B0E14" Foreground="White" BorderThickness="0" Margin="0"
                 FontFamily="Cascadia Mono, Consolas, monospace" FontSize="12"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
      </Grid>
    </Border>

    <!-- Status bar -->
    <Border Grid.Row="3" Background="#080A10" BorderBrush="{StaticResource divider}" BorderThickness="0,1,0,0">
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="DotGraph" Width="8" Height="8" Fill="#666" Margin="0,0,6,0"/>
          <TextBlock x:Name="SbGraph" Text="Graph: —" Margin="0,0,16,0" Foreground="{StaticResource fg.dim}"/>
          <Ellipse x:Name="DotExo" Width="8" Height="8" Fill="#666" Margin="0,0,6,0"/>
          <TextBlock x:Name="SbExo"   Text="EXO: —"   Margin="0,0,16,0" Foreground="{StaticResource fg.dim}"/>
          <TextBlock x:Name="SbTenant" Text="—"        Margin="0,0,16,0" Foreground="{StaticResource fg.dim}"/>
          <TextBlock x:Name="SbAccount" Text=""       Foreground="{StaticResource fg.dim}"/>
        </StackPanel>
        <TextBlock x:Name="SbClock" Grid.Column="1" Text="" Foreground="{StaticResource fg.dim}" VerticalAlignment="Center"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

# --- Load XAML ---

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Find named elements
function Find { param([string]$Name); return $Window.FindName($Name) }

$HeaderSubtitle = Find 'HeaderSubtitle'
$ChipRoleText   = Find 'ChipRoleText'
$ChipUiModeText = Find 'ChipUiModeText'
$ChipJobs       = Find 'ChipJobs'
$ChipJobsText   = Find 'ChipJobsText'
$BtnConnect     = Find 'BtnConnect'
$BtnDisconnect  = Find 'BtnDisconnect'
$SideNav        = Find 'SideNav'
$CardTenant     = Find 'CardTenant'
$CardAccount    = Find 'CardAccount'
$CardGraph      = Find 'CardGraph'
$CardExo        = Find 'CardExo'
$QuickHealth    = Find 'QuickHealth'
$QuickAudit     = Find 'QuickAudit'
$QuickSelfTest  = Find 'QuickSelfTest'
$QuickToolkitCheck = Find 'QuickToolkitCheck'
$QuickOpenLogs  = Find 'QuickOpenLogs'
$QuickOpenReports = Find 'QuickOpenReports'

$PanelWelcome     = Find 'PanelWelcome'
$PanelHealth      = Find 'PanelHealth'
$PanelAudit       = Find 'PanelAudit'
$PanelGroups      = Find 'PanelGroups'
$PanelExport      = Find 'PanelExport'
$PanelPermissions = Find 'PanelPermissions'
$PanelConvert     = Find 'PanelConvert'
$PanelOffboarding = Find 'PanelOffboarding'
$PanelSelfTest    = Find 'PanelSelfTest'
$PanelJobs        = Find 'PanelJobs'
$PanelCert        = Find 'PanelCert'
$PanelPrefs       = Find 'PanelPrefs'

$BtnRunHealth   = Find 'BtnRunHealth'
$HealthStatus   = Find 'HealthStatus'
$BtnRunAudit    = Find 'BtnRunAudit'
$AuditStatus    = Find 'AuditStatus'

$GroupsModeAdd       = Find 'GroupsModeAdd'
$GroupsModeCreate    = Find 'GroupsModeCreate'
$GroupsModeCreateAdd = Find 'GroupsModeCreateAdd'
$GroupsAddBox        = Find 'GroupsAddBox'
$GroupsCreateBox     = Find 'GroupsCreateBox'
$GroupsCreateAddBox  = Find 'GroupsCreateAddBox'
$LegacyGroupSearch   = Find 'LegacyGroupSearch'
$BtnLegacySearch     = Find 'BtnLegacySearch'
$LegacyCandidates    = Find 'LegacyCandidates'
$LegacyMembersCsv    = Find 'LegacyMembersCsv'
$LegacyBrowse        = Find 'LegacyBrowse'
$BtnRunLegacyMembers = Find 'BtnRunLegacyMembers'
$CreateType          = Find 'CreateType'
$CreateEmail         = Find 'CreateEmail'
$CreateDisplay       = Find 'CreateDisplay'
$CreateAlias         = Find 'CreateAlias'
$CreateOwners        = Find 'CreateOwners'
$CreateHidden        = Find 'CreateHidden'
$BtnRunCreateGroup   = Find 'BtnRunCreateGroup'
$CreateAddType       = Find 'CreateAddType'
$CreateAddEmail      = Find 'CreateAddEmail'
$CreateAddDisplay    = Find 'CreateAddDisplay'
$CreateAddAlias      = Find 'CreateAddAlias'
$CreateAddOwners     = Find 'CreateAddOwners'
$CreateAddHidden     = Find 'CreateAddHidden'
$CreateAddMembers    = Find 'CreateAddMembers'
$BtnRunCreateAddGroup = Find 'BtnRunCreateAddGroup'
$GroupsStatus        = Find 'GroupsStatus'

$ExportSearch      = Find 'ExportSearch'
$BtnExportSearch   = Find 'BtnExportSearch'
$ExportCandidates  = Find 'ExportCandidates'
$ExportFolder      = Find 'ExportFolder'
$ExportBrowse      = Find 'ExportBrowse'
$BtnRunExport      = Find 'BtnRunExport'
$ExportStatus      = Find 'ExportStatus'

$PermsCsvPath   = Find 'PermsCsvPath'
$PermsBrowse    = Find 'PermsBrowse'
$BtnRunPerms    = Find 'BtnRunPerms'
$PermsStatus    = Find 'PermsStatus'

$ConvertUpn    = Find 'ConvertUpn'
$ConvertForce  = Find 'ConvertForce'
$BtnRunConvert = Find 'BtnRunConvert'
$ConvertStatus = Find 'ConvertStatus'

$SelfTestTarget      = Find 'SelfTestTarget'
$SelfTestDelegate    = Find 'SelfTestDelegate'
$SelfTestSeed        = Find 'SelfTestSeed'
$SelfTestSkipCleanup = Find 'SelfTestSkipCleanup'
$BtnRunSelfTest      = Find 'BtnRunSelfTest'
$SelfTestStatus      = Find 'SelfTestStatus'

$JobsGrid       = Find 'JobsGrid'
$BtnJobsRefresh = Find 'BtnJobsRefresh'
$BtnJobsClear   = Find 'BtnJobsClear'
$JobsStatus     = Find 'JobsStatus'

$CertState     = Find 'CertState'
$CertAppId     = Find 'CertAppId'
$CertThumb     = Find 'CertThumb'
$CertTenant    = Find 'CertTenant'
$CertOrg       = Find 'CertOrg'
$BtnCertRefresh    = Find 'BtnCertRefresh'
$BtnCertOpenConsole = Find 'BtnCertOpenConsole'
$BtnCertOpenFolder = Find 'BtnCertOpenFolder'
$CertStatus    = Find 'CertStatus'


$OffUserSearch        = Find 'OffUserSearch'
$BtnOffUserSearch     = Find 'BtnOffUserSearch'
$OffUserCandidates    = Find 'OffUserCandidates'
$OffDelegateSearch    = Find 'OffDelegateSearch'
$BtnOffDelegateSearch = Find 'BtnOffDelegateSearch'
$OffDelegateCandidates = Find 'OffDelegateCandidates'
$BtnAddDelegate       = Find 'BtnAddDelegate'
$BtnRemoveDelegate    = Find 'BtnRemoveDelegate'
$OffDelegateChosen    = Find 'OffDelegateChosen'
$OffManagerSearch     = Find 'OffManagerSearch'
$BtnOffManagerSearch  = Find 'BtnOffManagerSearch'
$OffManagerCandidates = Find 'OffManagerCandidates'
$OffLang              = Find 'OffLang'
$OffKind              = Find 'OffKind'
$OffCustomBody        = Find 'OffCustomBody'
$OffDryRun            = Find 'OffDryRun'
$BtnRunOff            = Find 'BtnRunOff'
$OffStatus            = Find 'OffStatus'

$PrefMethod     = Find 'PrefMethod'
$PrefRole       = Find 'PrefRole'
$PrefUiMode     = Find 'PrefUiMode'
$BtnSavePrefs   = Find 'BtnSavePrefs'
$PrefStatus     = Find 'PrefStatus'

$LogList        = Find 'LogList'
$OpStatusText   = Find 'OpStatusText'
$BtnClearLog    = Find 'BtnClearLog'

$SbGraph    = Find 'SbGraph'
$SbExo      = Find 'SbExo'
$SbTenant   = Find 'SbTenant'
$SbAccount  = Find 'SbAccount'
$SbClock    = Find 'SbClock'
$DotGraph   = Find 'DotGraph'
$DotExo     = Find 'DotExo'

# --- Helpers ---

$panels = @{
    welcome      = $PanelWelcome
    health       = $PanelHealth
    audit        = $PanelAudit
    groups       = $PanelGroups
    export       = $PanelExport
    permissions  = $PanelPermissions
    convert      = $PanelConvert
    offboarding  = $PanelOffboarding
    selftest     = $PanelSelfTest
    jobs         = $PanelJobs
    cert         = $PanelCert
    prefs        = $PanelPrefs
}

function Show-Panel { param([string]$Tag)
    foreach ($k in $panels.Keys) {
        $panels[$k].Visibility = if ($k -eq $Tag) { 'Visible' } else { 'Collapsed' }
    }
    $HeaderSubtitle.Text = switch ($Tag) {
        'welcome'      { 'Bienvenido' }
        'health'       { 'Salud del tenant' }
        'audit'        { 'Identity audit' }
        'groups'       { 'Grupos · workflow / miembros' }
        'export'       { 'Exportar miembros' }
        'permissions'  { 'Permisos de buzón' }
        'convert'      { 'Convertir buzón shared → user' }
        'offboarding'  { 'Offboarding wizard' }
        'selftest'     { 'Self-test sobre testeo*' }
        'jobs'         { 'Jobs en background' }
        'cert'         { 'Certificado' }
        'prefs'        { 'Preferencias' }
        default        { '' }
    }
    if ($Tag -eq 'jobs') { Refresh-JobsGrid }
    if ($Tag -eq 'cert') { Refresh-CertPanel }
}

function global:Append-Log {
    param([string]$Message,[string]$Level = 'INFO',[string]$Source = '')
    try {
        $target = $LogList
        if (-not $target -and $global:GREX365_GUI) { $target = $global:GREX365_GUI.LogList }
        if (-not $target) { Write-GuiDebug 'Append-Log: no LogList'; return }
        $time = (Get-Date).ToString('HH:mm:ss')
        $tag = switch ($Level) { 'OK'{' OK '} 'WARN'{'WARN'} 'ERROR'{'FAIL'} default{'INFO'} }
        $brush = switch ($Level) { 'OK'{'LightGreen'} 'WARN'{'Gold'} 'ERROR'{'Salmon'} default{'Gainsboro'} }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "$time  $tag  $(if ($Source) {"[$Source] "})$Message"
        $tb.Foreground = [System.Windows.Media.Brushes]::$brush
        $tb.Margin = '8,1,8,1'
        $tb.FontFamily = 'Cascadia Mono, Consolas, monospace'
        $target.Items.Add($tb) | Out-Null
        if ($target.Items.Count -gt 1000) { $target.Items.RemoveAt(0) }
        $target.ScrollIntoView($tb)
    } catch {
        Write-GuiDebug ("Append-Log error: " + $_.Exception.Message)
    }
}

function global:Set-DotColor { param($Dot,[bool]$Connected)
    if (-not $Dot) { return }
    $Dot.Fill = if ($Connected) { [System.Windows.Media.Brushes]::LimeGreen } else { [System.Windows.Media.Brushes]::Gray }
}

function global:Refresh-StatusBar {
    try { Reset-SessionStateCache } catch {}
    $state = $null
    try { $state = Get-SessionState -Force } catch {}
    if (-not $state) {
        $SbGraph.Text = 'Graph: —'; $SbExo.Text = 'EXO: —'
        $SbTenant.Text = '—'; $SbAccount.Text = ''
        Set-DotColor -Dot $DotGraph -Connected $false
        Set-DotColor -Dot $DotExo   -Connected $false
        $CardTenant.Text = '—'; $CardAccount.Text = '—'
        $CardGraph.Text  = 'Graph: —'; $CardExo.Text = 'EXO: —'
        return
    }
    $SbGraph.Text   = if ($state.GraphConnected) { 'Graph: connected' } else { 'Graph: —' }
    $SbExo.Text     = if ($state.ExoConnected)   { 'EXO: connected'   } else { 'EXO: —'   }
    $SbTenant.Text  = if ($state.TenantDomain) { 'Tenant: ' + $state.TenantDomain } else { '—' }
    $SbAccount.Text = if ($state.Account) { '· ' + $state.Account } else { '' }
    Set-DotColor -Dot $DotGraph -Connected $state.GraphConnected
    Set-DotColor -Dot $DotExo   -Connected $state.ExoConnected
    $CardTenant.Text = if ($state.TenantDomain) { $state.TenantDomain } else { '—' }
    $CardAccount.Text = if ($state.Account) { $state.Account } else { '—' }
    $CardGraph.Text   = if ($state.GraphConnected) { 'Graph: connected' } else { 'Graph: —' }
    $CardExo.Text     = if ($state.ExoConnected)   { 'EXO: connected'   } else { 'EXO: —'   }
}

# --- Runspace runner ---

function Start-RunspaceJob {
    param(
        [string]$Name,
        [scriptblock]$Script,
        [hashtable]$JobArgs = @{},
        [string[]]$Inputs = @(),
        [hashtable]$Defaults = @{}
    )

    # Auto-inject inputs + defaults so child scripts run unattended.
    # Inputs is an ordered list consumed once each. Defaults map prompt-regex → answer
    # for prompts not in the queue. Unknown prompts return '' (or 'S' for S/N).
    if (-not $JobArgs.ContainsKey('inputs'))   { $JobArgs.inputs   = $Inputs }
    if (-not $JobArgs.ContainsKey('defaults')) { $JobArgs.defaults = $Defaults }

    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $RunspacePool
    [void]$ps.AddScript({
        param($Sync, $BasePath, $JobName, $Inner, $InnerArgs)
        $global:GREX365_BasePath = $BasePath
        $modOrder = @('Logging.ps1','Console.ps1','Validation.ps1','Csv.ps1','Preferences.ps1','Retry.ps1','Audit.ps1','Report.ps1','Roles.ps1','Templates.ps1','Jobs.ps1','Connection.ps1','GroupResolver.ps1')
        foreach ($m in $modOrder) {
            $p = Join-Path $BasePath ("Modules\$m")
            if (Test-Path $p) { . $p }
        }

        # Route Write-Log to the UI queue.
        function Write-Log {
            param([string]$Message,[ValidateSet('INFO','OK','WARN','ERROR','DEBUG')][string]$Level='INFO',[string]$Source='')
            $Sync.LogQueue.Enqueue(@{ msg=$Message; lvl=$Level; src=$Source; ts=(Get-Date).ToString('o') })
            if ($global:GREX365_LogSession) {
                $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $line = "[$stamp] [$Level]"
                if ($Source) { $line += " [$Source]" }
                $line += " $Message"
                [void]$global:GREX365_LogSession.Buffer.Add($line)
                if ($Level -eq 'ERROR') { $global:GREX365_LogSession.HasErrors = $true }
                if ($Level -eq 'OK')    { $global:GREX365_LogSession.HasSuccess = $true }
            }
        }
        function Update-OpStatus { param([string]$Text); $Sync.OpStatusQueue.Enqueue($Text) }

        # Ordered Read-Input mock. State held in $global: so child scripts (loaded via
        # `& $path` with Set-StrictMode Latest) can see it through the function call;
        # $script: would bind to the child script's own scope and miss the queue.
        # Globals are reset at every job start, and each runspace pool worker has its
        # own global scope, so cross-job contamination is bounded.
        $global:__inputQueue = New-Object 'System.Collections.Generic.Queue[string]'
        if ($InnerArgs.ContainsKey('inputs') -and $InnerArgs.inputs) {
            foreach ($v in @($InnerArgs.inputs)) { $global:__inputQueue.Enqueue([string]$v) }
        }
        $global:__inputDefaults = if ($InnerArgs.ContainsKey('defaults') -and $InnerArgs.defaults) { $InnerArgs.defaults } else { @{} }

        function global:Read-Input {
            param([string]$Prompt,[string]$Default='')
            if ($global:__inputQueue.Count -gt 0) {
                return $global:__inputQueue.Dequeue()
            }
            foreach ($key in $global:__inputDefaults.Keys) {
                if ($Prompt -match $key) { return [string]$global:__inputDefaults[$key] }
            }
            if ($Default) { return $Default }
            if ($Prompt -match '\(S/N\)|\(Y/N\)') { return 'S' }
            return ''
        }
        function global:Confirm-DestructiveAction { return $true }
        # Auto-pick highest-score candidate so scripts that fall through to
        # Show-GroupSelectionMenu don't deadlock on [Console]::ReadKey from GUI.
        function global:Show-GroupSelectionMenu {
            param([System.Collections.IList]$Options,[string]$SearchText)
            if (-not $Options -or $Options.Count -eq 0) { return $null }
            return $Options[0]
        }

        $Sync.LogQueue.Enqueue(@{ msg = "Job '$JobName' arrancado"; lvl='INFO'; src='GUI'; ts=(Get-Date).ToString('o') })
        $result = 'OK'
        try {
            & $Inner $Sync $InnerArgs
            $Sync.LogQueue.Enqueue(@{ msg = "Job '$JobName' completado"; lvl='OK'; src='GUI'; ts=(Get-Date).ToString('o') })
        } catch {
            $result = 'ERROR'
            $msg = $_.Exception.Message
            $stack = $_.ScriptStackTrace
            $Sync.LogQueue.Enqueue(@{ msg = "Job '$JobName' fallido: $msg"; lvl='ERROR'; src='GUI'; ts=(Get-Date).ToString('o') })
            if ($stack) {
                foreach ($line in ($stack -split "`r?`n" | Select-Object -First 6)) {
                    if ($line -and $line.Trim()) {
                        $Sync.LogQueue.Enqueue(@{ msg = '  ' + $line.Trim(); lvl='ERROR'; src='GUI'; ts=(Get-Date).ToString('o') })
                    }
                }
            }
        } finally {
            $Sync.StatusQueue.Enqueue(@{ refresh = $true })
            $Sync.OpStatusQueue.Enqueue('')
            $Sync.JobDoneQueue.Enqueue(@{ name = $JobName; result = $result })
        }
    }) | Out-Null
    [void]$ps.AddArgument($SyncHash)
    [void]$ps.AddArgument($global:GREX365_BasePath)
    [void]$ps.AddArgument($Name)
    [void]$ps.AddArgument($Script)
    [void]$ps.AddArgument($JobArgs)
    $async = $ps.BeginInvoke()

    [System.Threading.Interlocked]::Increment([ref]$SyncHash.BusyCount) | Out-Null
    $JobsList.Add([PSCustomObject]@{ PowerShell = $ps; Handle = $async; Name = $Name; Started = Get-Date })
    $SyncHash.OpStatusQueue.Enqueue("$Name · en ejecución...")
}

# --- Dispatcher timer drains queues ---
#
# Defensive design: every tick is wrapped in try/catch. Variables are pulled from
# $global: so the scriptblock works even if its captured closure context goes
# stale (observed when worker runspaces signal completion and the WPF dispatcher
# fires on a thread where the script's local scope is no longer in TLS).
# All errors land in $global:GREX365_GuiDebugLog so we can diagnose later.

# Surface state to globals for the dispatcher closure (and for diagnostics).
$global:GREX365_GUI = @{
    SyncHash       = $SyncHash
    JobsList       = $JobsList
    LogList        = $LogList
    OpStatusText   = $OpStatusText
    PanelJobs      = $PanelJobs
    ExportCandidates = $ExportCandidates
    ExportStatus     = $ExportStatus
    LegacyCandidates = $LegacyCandidates
    GroupsStatus     = $GroupsStatus
    OffUserCandidates = $OffUserCandidates
    OffDelegateCandidates = $OffDelegateCandidates
    OffManagerCandidates  = $OffManagerCandidates
    OffStatus        = $OffStatus
    ChipJobs         = $ChipJobs
    ChipJobsText     = $ChipJobsText
    BtnConnect       = $BtnConnect
    SbClock          = $SbClock
}

# Global helpers callable from anywhere — survive scope changes that the WPF
# dispatcher can introduce when worker runspaces complete on background threads.

# Drain helper isolates the `[ref]$item` inside its own function scope where
# the variable is unambiguously defined. The WPF-event-context scope confusion
# that breaks `[ref]$msg` directly inline cannot reach here.
function global:Get-AllQueueItems {
    param($Queue)
    $bag = New-Object System.Collections.Generic.List[object]
    if (-not $Queue) { return ,$bag.ToArray() }
    $item = $null
    while ($Queue.TryDequeue([ref]$item)) {
        if ($null -ne $item) { [void]$bag.Add($item) }
        $item = $null
    }
    return ,$bag.ToArray()
}

function global:Invoke-GuiTick {
    # The whole tick body lives here as a regular global function so any scope
    # issues with WPF event scriptblocks never reach it.

    # Make sure stdlib is available. Some WPF event invocation paths land on
    # threads where the default runspace state hasn't auto-imported Utility.
    if (-not (Get-Command Get-Date -ErrorAction SilentlyContinue)) {
        try { Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop } catch {}
    }

    $g = $global:GREX365_GUI
    if (-not $g) { Write-GuiDebug 'tick: no GUI global state yet'; return }
    $sh = $g.SyncHash
    if (-not $sh) { Write-GuiDebug 'tick: no SyncHash'; return }

    # Drain log queue.
    try {
        foreach ($msg in (Get-AllQueueItems $sh.LogQueue)) {
            Append-Log -Message $msg.msg -Level $msg.lvl -Source $msg.src
        }
    } catch { Write-GuiDebug ("tick.log: " + $_.Exception.Message) }

    # Status refresh queue.
    try {
        foreach ($s in (Get-AllQueueItems $sh.StatusQueue)) {
            if ($s.refresh) { try { Refresh-StatusBar } catch { Write-GuiDebug ("Refresh-StatusBar: " + $_.Exception.Message) } }
        }
    } catch { Write-GuiDebug ("tick.status: " + $_.Exception.Message) }

    # Op status text.
    try {
        foreach ($opMsg in (Get-AllQueueItems $sh.OpStatusQueue)) {
            if ($g.OpStatusText) { $g.OpStatusText.Text = [string]$opMsg }
        }
    } catch { Write-GuiDebug ("tick.op: " + $_.Exception.Message) }

    # Job-done queue.
    try {
        foreach ($jd in (Get-AllQueueItems $sh.JobDoneQueue)) {
            if ($g.PanelJobs -and [string]$g.PanelJobs.Visibility -eq 'Visible') {
                try { Refresh-JobsGrid } catch { Write-GuiDebug ("Refresh-JobsGrid: " + $_.Exception.Message) }
            }
        }
    } catch { Write-GuiDebug ("tick.jobdone: " + $_.Exception.Message) }

    # Search results.
    try {
        foreach ($sr in (Get-AllQueueItems $sh.SearchResultQueue)) {
            if (-not $sr) { continue }
            $listbox  = $null
            $statusTb = $null
            $kind     = if ($sr.kind) { [string]$sr.kind } else { 'group' }
            switch ([string]$sr.target) {
                'export'       { $listbox = $g.ExportCandidates;       $statusTb = $g.ExportStatus }
                'legacy'       { $listbox = $g.LegacyCandidates;       $statusTb = $g.GroupsStatus }
                'off-user'     { $listbox = $g.OffUserCandidates;      $statusTb = $g.OffStatus }
                'off-delegate' { $listbox = $g.OffDelegateCandidates;  $statusTb = $g.OffStatus }
                'off-manager'  { $listbox = $g.OffManagerCandidates;   $statusTb = $g.OffStatus }
            }
            if (-not $listbox) { continue }
            $listbox.Items.Clear()
            $listbox.Tag = $null
            if ([string]$sr.status -eq 'error') {
                if ($statusTb) { $statusTb.Text = [string]$sr.message }
                continue
            }
            $results = @($sr.results)
            if ($results.Count -eq 0) {
                if ($statusTb) { $statusTb.Text = "Sin coincidencias para '$($sr.needle)'." }
                continue
            }
            foreach ($cand in $results) {
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = if ($kind -eq 'user') { Format-UserCandidateLine -Candidate $cand } else { Format-GroupCandidateLine -Candidate $cand }
                $item.Tag = $cand
                $listbox.Items.Add($item) | Out-Null
            }
            $listbox.SelectedIndex = 0
            if ($statusTb) { $statusTb.Text = ("Coincidencias: " + $results.Count + ". Elige una y continúa.") }
        }
    } catch { Write-GuiDebug ("tick.search: " + $_.Exception.Message) }

    # Cull completed runspaces (defensive against null entries).
    try {
        if ($g.JobsList) {
            $done = @($g.JobsList | Where-Object { $_ -and $_.Handle -and $_.Handle.IsCompleted })
            foreach ($d in $done) {
                try { $d.PowerShell.EndInvoke($d.Handle) | Out-Null }
                catch {
                    if ($sh.LogQueue) {
                        $sh.LogQueue.Enqueue(@{ msg = "EndInvoke error en '$($d.Name)': $($_.Exception.Message)"; lvl='ERROR'; src='GUI' })
                    }
                }
                try { $d.PowerShell.Dispose() } catch {}
                try { [void]$g.JobsList.Remove($d) } catch {}
                try { [System.Threading.Interlocked]::Decrement([ref]$sh.BusyCount) | Out-Null } catch {}
            }
        }
    } catch { Write-GuiDebug ("tick.cull: " + $_.Exception.Message) }

    # Update busy chip + connect button state.
    try {
        $busy = 0
        try { $busy = [int]$sh.BusyCount } catch {}
        if ($g.ChipJobs -and $g.ChipJobsText) {
            if ($busy -gt 0) {
                $g.ChipJobs.Visibility = 'Visible'
                $g.ChipJobsText.Text = ('● ' + $busy + ' job' + $(if ($busy -gt 1) { 's' } else { '' }))
            } else {
                $g.ChipJobs.Visibility = 'Collapsed'
            }
        }
        if ($g.BtnConnect) { $g.BtnConnect.IsEnabled = ($busy -le 0) }
    } catch { Write-GuiDebug ("tick.chip: " + $_.Exception.Message) }

    # Clock.
    try {
        if ($g.SbClock) { $g.SbClock.Text = (Get-Date).ToString('HH:mm:ss') }
    } catch { Write-GuiDebug ("tick.clock: " + $_.Exception.Message) }
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
# Register-ObjectEvent routes the event through PSEventManager which preserves
# the runspace context. Add_Tick scriptblocks can sometimes execute in a context
# where global functions/cmdlets are not visible — observed empirically in this
# tool — so we avoid that path entirely.
$global:GREX365_TimerSubscriber = Register-ObjectEvent -InputObject $timer -EventName Tick -SourceIdentifier 'GREX365.GuiTick' -Action {
    try { Invoke-GuiTick }
    catch {
        try { Write-GuiDebug ("tick.toplevel: " + $_.Exception.Message) } catch {}
    }
}
$timer.Start()

# --- Wire navigation ---

$SideNav.Add_SelectionChanged({
    if (-not $SideNav.SelectedItem) { return }
    $tag = [string]$SideNav.SelectedItem.Tag
    Show-Panel -Tag $tag
})

# --- Wire actions ---

$BtnConnect.Add_Click({
    $prefs = Get-UserPreferences
    if (-not $prefs.ConnectionMethod) {
        [System.Windows.MessageBox]::Show(
            "Define el método de conexión primero.`nSidebar → Preferencias → 'Método activo'.",
            'GREX365 · Falta método', 'OK', 'Warning') | Out-Null
        return
    }

    # Pre-flight: cert mode needs config file, traditional needs admin UPN.
    if ($prefs.ConnectionMethod -eq 'cert') {
        if (-not (Test-CertConfigExists)) {
            [System.Windows.MessageBox]::Show(
                "Método actual: certificado. No hay configuración válida.`nVe a Certificado en la sidebar y lanza el asistente desde consola, o cambia el método a 'traditional' en Preferencias.",
                'GREX365 · Cert no configurado', 'OK', 'Warning') | Out-Null
            return
        }
    }
    $adminUpn = $null
    if ($prefs.ConnectionMethod -eq 'traditional') {
        $adminUpn = if ($prefs.PSObject.Properties.Name -contains 'TraditionalAdminUpn') { [string]$prefs.TraditionalAdminUpn } else { '' }
        if (-not $adminUpn) {
            $dlgInput = [Microsoft.VisualBasic.Interaction]::InputBox(
                'UPN del admin para device code (se guarda en preferencias):',
                'GREX365 · UPN admin', '')
            if ([string]::IsNullOrWhiteSpace($dlgInput)) { return }
            Set-PreferenceValue -Key 'TraditionalAdminUpn' -Value (Normalize-Input -Value $dlgInput)
            $adminUpn = (Normalize-Input -Value $dlgInput)
        }
    }

    Append-Log -Message ("Connect iniciado · método=" + $prefs.ConnectionMethod) -Level 'INFO' -Source 'Connect'

    Start-RunspaceJob -Name 'Connect' -JobArgs @{ method = [string]$prefs.ConnectionMethod; upn = $adminUpn } -Script {
        param($Sync,$JobArgs)
        try {
            $Sync.OpStatusQueue.Enqueue('Conectando a Microsoft Graph y Exchange Online...')
            Connect-RequiredServices -MgGraph -ExchangeOnline -GraphScopes @('User.Read.All','Group.Read.All','Directory.Read.All','GroupMember.ReadWrite.All','User.ReadBasic.All','Group.ReadWrite.All')
            $state = Get-SessionState -Force
            $Sync.LogQueue.Enqueue(@{
                msg = ('Conectado · tenant=' + ($state.TenantDomain ? $state.TenantDomain : '?') + ' · account=' + ($state.Account ? $state.Account : '?'))
                lvl = 'OK'; src = 'Connect'
            })
        } catch {
            $Sync.LogQueue.Enqueue(@{
                msg = ('Connect FAIL: ' + $_.Exception.Message)
                lvl = 'ERROR'; src = 'Connect'
            })
            if ($_.Exception.Message -match 'cert|certificate') {
                $Sync.LogQueue.Enqueue(@{ msg = 'Sugerencia: comprueba que el thumbprint en exo-app-params.json sigue siendo válido.'; lvl='WARN'; src='Connect' })
            } elseif ($_.Exception.Message -match 'device|browser|interactive') {
                $Sync.LogQueue.Enqueue(@{ msg = 'Sugerencia: device code abre el navegador. Mira el output original de PowerShell o switchea a cert.'; lvl='WARN'; src='Connect' })
            }
        }
    }
})

$BtnDisconnect.Add_Click({
    Start-RunspaceJob -Name 'Disconnect' -JobArgs @{} -Script {
        param($Sync,$JobArgs)
        Disconnect-AllServices
    }
})

$BtnRunHealth.Add_Click({
    if (-not (Test-RequiredConnections)) { return }
    $HealthStatus.Text = 'Ejecutando...'
    Start-RunspaceJob -Name 'Tenant Health' -JobArgs @{ script = (Join-Path $ScriptsPath 'Show-TenantHealth.ps1') } -Script {
        param($Sync,$JobArgs)
        & $JobArgs.script
    }
})

$BtnRunAudit.Add_Click({
    $graph = $false; try { $graph = Test-GraphConnected } catch {}
    if (-not $graph) {
        [System.Windows.MessageBox]::Show('Identity audit necesita Microsoft Graph. Pulsa Connect arriba.', 'GREX365', 'OK', 'Warning') | Out-Null
        return
    }
    $AuditStatus.Text = 'Ejecutando...'
    Start-RunspaceJob -Name 'Identity Audit' -JobArgs @{ script = (Join-Path $ScriptsPath 'Invoke-IdentityAudit.ps1') } -Script {
        param($Sync,$JobArgs)
        & $JobArgs.script
    }
})

function Pick-CsvFile {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv|All files (*.*)|*.*'
    $dlg.Multiselect = $false
    if ($dlg.ShowDialog() -eq $true) { return $dlg.FileName }
    return $null
}

$PermsBrowse.Add_Click({  $p = Pick-CsvFile; if ($p) { $PermsCsvPath.Text  = $p } })

$BtnRunPerms.Add_Click({
    $path = $PermsCsvPath.Text
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        $PermsStatus.Text = 'CSV inválido.'; return
    }
    $exo = $false; try { $exo = Test-ExchangeOnlineConnected } catch {}
    if (-not $exo) {
        [System.Windows.MessageBox]::Show('Permisos de buzón necesita Exchange Online conectado.', 'GREX365', 'OK', 'Warning') | Out-Null
        return
    }
    $PermsStatus.Text = 'Ejecutando permisos...'
    Start-RunspaceJob -Name 'Mailbox Perms' `
        -Inputs @($path) `
        -Defaults @{ 'CSV|Ruta' = $path } `
        -JobArgs @{ script=(Join-Path $ScriptsPath 'Set-SharedMailboxPermissions.ps1'); path=$path } -Script {
            param($Sync,$JobArgs)
            & $JobArgs.script
        }
})

# --- Offboarding handlers ---

$BtnOffUserSearch.Add_Click({
    $term = $OffUserSearch.Text.Trim()
    if (-not $term) { $OffStatus.Text = 'Búsqueda usuario vacía.'; return }
    if (-not (Test-RequiredConnections)) { return }
    $OffStatus.Text = "Buscando usuario '$term'..."
    Invoke-UserSearchJob -SearchText $term -TargetTag 'off-user'
})

$BtnOffDelegateSearch.Add_Click({
    $term = $OffDelegateSearch.Text.Trim()
    if (-not $term) { $OffStatus.Text = 'Búsqueda delegado vacía.'; return }
    if (-not (Test-RequiredConnections)) { return }
    $OffStatus.Text = "Buscando delegado '$term'..."
    Invoke-UserSearchJob -SearchText $term -TargetTag 'off-delegate'
})

$BtnOffManagerSearch.Add_Click({
    $term = $OffManagerSearch.Text.Trim()
    if (-not $term) { $OffStatus.Text = 'Búsqueda manager vacía.'; return }
    if (-not (Test-RequiredConnections)) { return }
    $OffStatus.Text = "Buscando manager '$term'..."
    Invoke-UserSearchJob -SearchText $term -TargetTag 'off-manager'
})

$BtnAddDelegate.Add_Click({
    $cand = $OffDelegateCandidates.SelectedItem
    if (-not $cand -or -not $cand.Tag) { $OffStatus.Text = 'Selecciona un candidato en la lista de delegados.'; return }
    $u = $cand.Tag
    $mail = if ($u.Mail) { $u.Mail } else { $u.UserPrincipalName }
    if (-not $mail) { $OffStatus.Text = 'Candidato sin email válido.'; return }
    foreach ($it in $OffDelegateChosen.Items) {
        if ([string]$it.Tag -eq $mail) { $OffStatus.Text = "Ya añadido: $mail"; return }
    }
    $item = New-Object System.Windows.Controls.ListBoxItem
    $item.Content = ('+ ' + $u.DisplayName + '  <' + $mail + '>')
    $item.Tag = $mail
    $OffDelegateChosen.Items.Add($item) | Out-Null
    $OffStatus.Text = "Delegados seleccionados: " + $OffDelegateChosen.Items.Count
})

$BtnRemoveDelegate.Add_Click({
    $sel = @($OffDelegateChosen.SelectedItems)
    foreach ($it in $sel) { [void]$OffDelegateChosen.Items.Remove($it) }
    $OffStatus.Text = "Delegados seleccionados: " + $OffDelegateChosen.Items.Count
})

$BtnRunOff.Add_Click({
    # Required: target user + at least 1 delegate
    $userCand = $OffUserCandidates.SelectedItem
    if (-not $userCand -or -not $userCand.Tag) { $OffStatus.Text = 'Falta seleccionar usuario saliente (Buscar + elegir).'; return }
    if ($OffDelegateChosen.Items.Count -eq 0) { $OffStatus.Text = 'Añade al menos 1 delegado.'; return }
    if (-not (Test-RequiredConnections)) { return }

    $userEmail = if ($userCand.Tag.Mail) { [string]$userCand.Tag.Mail } else { [string]$userCand.Tag.UserPrincipalName }
    $delegates = @()
    foreach ($it in $OffDelegateChosen.Items) { $delegates += [string]$it.Tag }
    $managerEmail = ''
    $mgCand = $OffManagerCandidates.SelectedItem
    if ($mgCand -and $mgCand.Tag) {
        $managerEmail = if ($mgCand.Tag.Mail) { [string]$mgCand.Tag.Mail } else { [string]$mgCand.Tag.UserPrincipalName }
    }

    $lang = if ($OffLang.SelectedItem) { [string]$OffLang.SelectedItem.Tag } else { 'es' }
    $kind = if ($OffKind.SelectedItem) { [string]$OffKind.SelectedItem.Tag } else { 'permanent' }
    $tpl  = "offboarding-$kind-$lang"

    $custom = $OffCustomBody.Text
    $dry    = [bool]$OffDryRun.IsChecked

    if (-not $userEmail.ToLowerInvariant().StartsWith('testeo') -and -not $dry) {
        $r = [System.Windows.MessageBox]::Show(
            "Usuario '$userEmail' no empieza por 'testeo' y dry-run está desactivado. ¿Continuar?",
            'Offboarding · Confirmación', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $OffStatus.Text = 'Cancelado.'; return }
    }

    $OffStatus.Text = ('Ejecutando offboarding · usuario=' + $userEmail + ' · delegados=' + ($delegates -join ',') + ($managerEmail ? (' · manager=' + $managerEmail) : ''))

    Start-RunspaceJob -Name ("Offboarding · $userEmail") `
        -JobArgs @{
            script    = (Join-Path $ScriptsPath 'Invoke-OffboardingWizard.ps1')
            email     = $userEmail
            delegates = $delegates
            manager   = $managerEmail
            template  = $tpl
            custom    = $custom
            dry       = $dry
        } `
        -Script {
            param($Sync,$JobArgs)
            # Pre-load context so the script's Read-Input prompts find the answers.
            $delegatesCsv = ($JobArgs.delegates -join ',')
            $script:__offEmail     = $JobArgs.email
            $script:__offDelegates = $delegatesCsv
            $script:__offManager   = if ($JobArgs.manager) { $JobArgs.manager } else { '' }
            $script:__offCustom    = $JobArgs.custom
            $script:__offTemplate  = $JobArgs.template

            # New ordered-input flow: script reads email, delegate(s), manager, then proceeds.
            $global:__inputQueue.Enqueue($script:__offEmail)
            $global:__inputQueue.Enqueue($script:__offDelegates)
            $global:__inputQueue.Enqueue($script:__offManager)

            # Force-pick the requested template, ignore any console-based interactive picker.
            function global:Select-TemplateInteractive { param([string]$Category)
                try { return Get-TemplateByName -Name $JobArgs.template } catch { return $null }
            }
            function global:Confirm-DestructiveAction { return $true }

            if ($JobArgs.dry) { Set-CurrentUIMode -Mode 'support' } else { Set-CurrentUIMode -Mode 'advanced' }

            # Inject custom body via env var; the script picks it up if defined.
            $env:GREX365_OFF_CUSTOMBODY = [string]$JobArgs.custom

            & $JobArgs.script
            Remove-Item Env:\GREX365_OFF_CUSTOMBODY -ErrorAction SilentlyContinue
        }
})

$QuickHealth.Add_Click({ $SideNav.SelectedIndex = 1; $BtnRunHealth.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) })
$QuickAudit.Add_Click({  $SideNav.SelectedIndex = 2; $BtnRunAudit.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) })

$QuickOpenLogs.Add_Click({
    $p = Join-Path $global:GREX365_BasePath 'logs'
    if (Test-Path $p) { Start-Process $p }
})
$QuickOpenReports.Add_Click({
    $p = Join-Path $global:GREX365_BasePath 'logs\reports'
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    Start-Process $p
})

$BtnClearLog.Add_Click({ $LogList.Items.Clear() })

# Preferences panel: load + save

function Load-PrefsToUI {
    $prefs = Get-UserPreferences
    foreach ($it in $PrefMethod.Items)   { if ([string]$it.Tag -eq [string]$prefs.ConnectionMethod) { $PrefMethod.SelectedItem = $it } }
    $role = Get-CurrentRole
    foreach ($it in $PrefRole.Items)     { if ([string]$it.Tag -eq $role) { $PrefRole.SelectedItem = $it } }
    $uimode = Get-CurrentUIMode
    foreach ($it in $PrefUiMode.Items)   { if ([string]$it.Tag -eq $uimode) { $PrefUiMode.SelectedItem = $it } }
}

$BtnSavePrefs.Add_Click({
    try {
        if ($PrefMethod.SelectedItem) { Set-PreferenceValue -Key 'ConnectionMethod' -Value ([string]$PrefMethod.SelectedItem.Tag) }
        if ($PrefRole.SelectedItem)   { Set-CurrentRole   -Role ([string]$PrefRole.SelectedItem.Tag) }
        if ($PrefUiMode.SelectedItem) { Set-CurrentUIMode -Mode ([string]$PrefUiMode.SelectedItem.Tag) }
        $PrefStatus.Text = 'Guardado.'
        Refresh-Chips
    } catch {
        $PrefStatus.Text = 'Error: ' + $_.Exception.Message
    }
})

# --- New helpers: chips, jobs grid, cert panel ---

function global:Refresh-Chips {
    try { $r = Get-CurrentRole } catch { $r = '—' }
    try { $m = Get-CurrentUIMode } catch { $m = '—' }
    $ChipRoleText.Text   = 'role · ' + $r
    $ChipUiModeText.Text = 'mode · ' + $m
}

function global:Refresh-JobsGrid {
    try {
        $rows = New-Object System.Collections.Generic.List[object]
        $live = @($JobsList | ForEach-Object {
            $age = ((Get-Date) - $_.Started)
            $ageStr = if ($age.TotalSeconds -lt 60) { ('{0:N0}s' -f $age.TotalSeconds) }
                      elseif ($age.TotalMinutes -lt 60) { ('{0:N0}m {1:N0}s' -f [math]::Floor($age.TotalMinutes), ($age.Seconds)) }
                      else { ('{0:N1}h' -f $age.TotalHours) }
            [PSCustomObject]@{
                Name    = $_.Name
                State   = if ($_.Handle.IsCompleted) { 'Done' } else { 'Running' }
                Started = $_.Started.ToString('HH:mm:ss')
                Age     = $ageStr
            }
        })
        foreach ($r in $live) { $rows.Add($r) }
        $JobsGrid.ItemsSource = $rows
        $JobsStatus.Text = ('Jobs en ejecución: ' + @($rows | Where-Object State -eq 'Running').Count + ' · total visibles: ' + $rows.Count)
    } catch {
        $JobsStatus.Text = 'Error refrescando: ' + $_.Exception.Message
    }
}

function global:Refresh-CertPanel {
    try {
        if (Test-CertConfigExists) {
            $cfg = Get-CertConfig
            $CertState.Text  = 'Configurado'
            $CertState.Foreground = [System.Windows.Media.Brushes]::LightGreen
            $CertAppId.Text  = [string]$cfg.AppId
            $CertThumb.Text  = [string]$cfg.CertThumbprint
            $CertTenant.Text = [string]$cfg.TenantId
            $CertOrg.Text    = [string]$cfg.Organization
        } else {
            $CertState.Text = 'No configurado'
            $CertState.Foreground = [System.Windows.Media.Brushes]::Salmon
            $CertAppId.Text = '—'; $CertThumb.Text = '—'; $CertTenant.Text = '—'; $CertOrg.Text = '—'
        }
        $CertStatus.Text = ''
    } catch {
        $CertStatus.Text = 'Error: ' + $_.Exception.Message
    }
}

# --- Common: connection pre-flight, folder picker, group search ---

function global:Test-RequiredConnections {
    param([string]$StatusLabel = 'Conecta primero (botón Connect arriba a la derecha).')

    $graph = $false; $exo = $false
    try { $graph = Test-GraphConnected } catch {}
    try { $exo   = Test-ExchangeOnlineConnected } catch {}

    if (-not $graph -or -not $exo) {
        $missing = @()
        if (-not $graph) { $missing += 'Microsoft Graph' }
        if (-not $exo)   { $missing += 'Exchange Online' }
        [System.Windows.MessageBox]::Show(
            'No conectado a: ' + ($missing -join ', ') + ".`n`n" + $StatusLabel,
            'GREX365 · Conexión requerida', 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

function Pick-Folder {
    param([string]$InitialPath = '')

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($InitialPath -and (Test-Path -LiteralPath $InitialPath)) {
            $dlg.SelectedPath = $InitialPath
        }
        $dlg.ShowNewFolderButton = $true
        $result = $dlg.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
        return $null
    } catch {
        # Last-resort fallback: keep textbox value, just notify
        [System.Windows.MessageBox]::Show(
            "No se pudo abrir el selector de carpeta. Escribe la ruta manualmente.`n" + $_.Exception.Message,
            'GREX365', 'OK', 'Information') | Out-Null
        return $null
    }
}

function Invoke-GroupSearchJob {
    param([string]$SearchText,[string]$TargetTag)

    Start-RunspaceJob -Name ("GroupSearch · $SearchText") `
        -JobArgs @{ search = $SearchText; targetTag = $TargetTag } `
        -Script {
            param($Sync,$JobArgs)
            $needle = Normalize-Input -Value $JobArgs.search
            if ([string]::IsNullOrWhiteSpace($needle)) {
                $Sync.SearchResultQueue.Enqueue(@{ target=$JobArgs.targetTag; status='error'; message='Búsqueda vacía.' })
                return
            }
            $normalized = Normalize-SearchText -Value $needle
            $exactMail  = $normalized -like '*@*'
            if (-not $exactMail -and $normalized.Length -lt 3) {
                $Sync.SearchResultQueue.Enqueue(@{ target=$JobArgs.targetTag; status='error'; message='Indica al menos 3 caracteres o correo completo.' })
                return
            }

            $all = New-Object System.Collections.Generic.List[object]
            if (Get-Command Get-EXORecipient -ErrorAction SilentlyContinue) {
                try {
                    foreach ($it in @(Get-ExchangeGroupCandidates -SearchText $normalized -SearchWasExactMail:$exactMail -MaxResults 20)) { $all.Add($it) }
                } catch {
                    $Sync.LogQueue.Enqueue(@{ msg = "Exchange search error: $($_.Exception.Message)"; lvl='WARN'; src='GroupSearch' })
                }
            }
            if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
                try {
                    foreach ($it in @(Get-GraphGroupCandidates -SearchText $normalized -SearchWasExactMail:$exactMail -MaxResults 20)) { $all.Add($it) }
                } catch {
                    $Sync.LogQueue.Enqueue(@{ msg = "Graph search error: $($_.Exception.Message)"; lvl='WARN'; src='GroupSearch' })
                }
            }
            $merged = @(Merge-GroupCandidates -Candidates $all)
            $payload = @()
            foreach ($m in $merged) {
                $payload += [PSCustomObject]@{
                    DisplayName        = [string]$m.DisplayName
                    PrimarySmtpAddress = [string]$m.PrimarySmtpAddress
                    GroupType          = [string]$m.GroupType
                    Identity           = [string]$m.Identity
                    GroupId            = [string]$m.GroupId
                    Source             = [string]$m.Source
                    Score              = [int]$m.Score
                }
            }
            $Sync.SearchResultQueue.Enqueue(@{ target=$JobArgs.targetTag; status='ok'; results=$payload; needle=$needle })
        }
}

function global:Format-GroupCandidateLine {
    param([PSCustomObject]$Candidate)
    $badge = switch ($Candidate.GroupType) {
        'Microsoft365Group'        { 'M365' }
        'DistributionList'         { ' DL ' }
        'MailEnabledSecurityGroup' { 'MSEC' }
        'SecurityGroup'            { ' SEC' }
        default                    { '  ? ' }
    }
    $mail = if ($Candidate.PrimarySmtpAddress) { $Candidate.PrimarySmtpAddress } else { '(sin mail)' }
    return ('[{0}]  {1,-40}  {2}' -f $badge, $Candidate.DisplayName, $mail)
}

function global:Format-UserCandidateLine {
    param([PSCustomObject]$Candidate)
    $name = if ($Candidate.DisplayName) { $Candidate.DisplayName } else { '(sin nombre)' }
    $upn  = if ($Candidate.UserPrincipalName) { $Candidate.UserPrincipalName } else { '' }
    return ('{0,-32}  {1}' -f $name, $upn)
}

function Invoke-UserSearchJob {
    param([string]$SearchText,[string]$TargetTag)

    Start-RunspaceJob -Name ("UserSearch · $SearchText") `
        -JobArgs @{ search = $SearchText; targetTag = $TargetTag } `
        -Script {
            param($Sync,$JobArgs)
            $needle = Normalize-Input -Value $JobArgs.search
            if ([string]::IsNullOrWhiteSpace($needle)) {
                $Sync.SearchResultQueue.Enqueue(@{ kind='user'; target=$JobArgs.targetTag; status='error'; message='Búsqueda vacía.' })
                return
            }
            if ($needle.Length -lt 2) {
                $Sync.SearchResultQueue.Enqueue(@{ kind='user'; target=$JobArgs.targetTag; status='error'; message='Mínimo 2 caracteres.' })
                return
            }

            $results = New-Object System.Collections.Generic.List[object]
            $seen = @{}

            $safe = $needle.Replace("'", "''")
            $isMail = $needle -like '*@*'

            # Graph users (mail + UPN + displayName starts/contains)
            if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                $tries = @()
                if ($isMail) {
                    $tries += "mail eq '$safe' or userPrincipalName eq '$safe'"
                } else {
                    $tries += "startswith(displayName,'$safe') or startswith(mail,'$safe') or startswith(userPrincipalName,'$safe')"
                }
                foreach ($filter in $tries) {
                    try {
                        $users = Get-MgUser -Top 25 -Filter $filter -ConsistencyLevel eventual -Property Id,DisplayName,Mail,UserPrincipalName,AccountEnabled -ErrorAction Stop
                        foreach ($u in @($users)) {
                            $idKey = ([string]$u.Id).ToLowerInvariant()
                            if (-not $idKey -or $seen.ContainsKey($idKey)) { continue }
                            $seen[$idKey] = $true
                            $results.Add([PSCustomObject]@{
                                Id                = [string]$u.Id
                                DisplayName       = [string]$u.DisplayName
                                Mail              = [string]$u.Mail
                                UserPrincipalName = [string]$u.UserPrincipalName
                                Enabled           = [bool]$u.AccountEnabled
                                Source            = 'Graph'
                            })
                        }
                    } catch {
                        $Sync.LogQueue.Enqueue(@{ msg = "User search Graph error: $($_.Exception.Message)"; lvl='WARN'; src='UserSearch' })
                    }
                }
            }

            $payload = @($results | Sort-Object @{Expression={ if ($_.Enabled) { 0 } else { 1 } }}, DisplayName | Select-Object -First 25)
            $Sync.SearchResultQueue.Enqueue(@{ kind='user'; target=$JobArgs.targetTag; status='ok'; results=$payload; needle=$needle })
        }
}

# --- Groups: mode switch ---

$GroupsModeAdd.Add_Checked({
    $GroupsAddBox.Visibility       = 'Visible'
    $GroupsCreateBox.Visibility    = 'Collapsed'
    $GroupsCreateAddBox.Visibility = 'Collapsed'
    $GroupsStatus.Text = 'Modo: Añadir miembros a grupo existente.'
})
$GroupsModeCreate.Add_Checked({
    $GroupsAddBox.Visibility       = 'Collapsed'
    $GroupsCreateBox.Visibility    = 'Visible'
    $GroupsCreateAddBox.Visibility = 'Collapsed'
    $GroupsStatus.Text = 'Modo: Crear grupo nuevo (sólo creación).'
})
$GroupsModeCreateAdd.Add_Checked({
    $GroupsAddBox.Visibility       = 'Collapsed'
    $GroupsCreateBox.Visibility    = 'Collapsed'
    $GroupsCreateAddBox.Visibility = 'Visible'
    $GroupsStatus.Text = 'Modo: Crear grupo + añadir miembros.'
})

# --- Mode 2 + 3 helper: create a group via EXO cmdlets (form-based, no CSV) ---

function Invoke-CreateGroupJob {
    param(
        [Parameter(Mandatory)][string]$Type,         # DL | M365 | MailSecurity
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)][string]$DisplayName,
        [string]$Alias = '',
        [string[]]$Owners = @(),
        [string[]]$Members = @(),
        [bool]$Hidden = $false
    )

    Start-RunspaceJob -Name ("Group create · $Email") `
        -JobArgs @{
            type    = $Type
            email   = $Email
            display = $DisplayName
            alias   = ($Alias ? $Alias : ($Email -replace '@.*$',''))
            owners  = $Owners
            members = $Members
            hidden  = $Hidden
        } `
        -Script {
            param($Sync,$JobArgs)
            $created = $false
            try {
                switch ($JobArgs.type) {
                    'DL' {
                        Invoke-WithRetry -OperationName 'New-DistributionGroup' -ScriptBlock {
                            New-DistributionGroup -Name $JobArgs.display -DisplayName $JobArgs.display -Alias $JobArgs.alias -PrimarySmtpAddress $JobArgs.email -Type Distribution -ErrorAction Stop | Out-Null
                        }
                    }
                    'MailSecurity' {
                        Invoke-WithRetry -OperationName 'New-DistributionGroup Security' -ScriptBlock {
                            New-DistributionGroup -Name $JobArgs.display -DisplayName $JobArgs.display -Alias $JobArgs.alias -PrimarySmtpAddress $JobArgs.email -Type Security -ErrorAction Stop | Out-Null
                        }
                    }
                    'M365' {
                        Invoke-WithRetry -OperationName 'New-UnifiedGroup' -ScriptBlock {
                            New-UnifiedGroup -DisplayName $JobArgs.display -Alias $JobArgs.alias -PrimarySmtpAddress $JobArgs.email -AccessType Private -ErrorAction Stop | Out-Null
                        }
                    }
                }
                $created = $true
                Write-Log ('Grupo creado: ' + $JobArgs.email + ' (' + $JobArgs.type + ')') -Level OK -Source 'GroupsCreate'
            } catch {
                Write-Log ('Error creando grupo: ' + $_.Exception.Message) -Level ERROR -Source 'GroupsCreate'
                return
            }

            # Owners (M365 only)
            if ($created -and $JobArgs.type -eq 'M365' -and $JobArgs.owners.Count -gt 0) {
                $okO = 0; foreach ($own in $JobArgs.owners) {
                    try {
                        Invoke-WithRetry -OperationName 'Add-UnifiedGroupLinks Owners' -ScriptBlock {
                            Add-UnifiedGroupLinks -Identity $JobArgs.email -LinkType Owners -Links $own -ErrorAction Stop
                        }
                        $okO++
                    } catch { Write-Log ('Owner falla: ' + $own + ' — ' + $_.Exception.Message) -Level WARN -Source 'GroupsCreate' }
                }
                Write-Log ("Owners aplicados: $okO / $($JobArgs.owners.Count)") -Level OK -Source 'GroupsCreate'
            }

            # Members
            if ($created -and $JobArgs.members.Count -gt 0) {
                $okM = 0; $skipM = 0
                foreach ($m in $JobArgs.members) {
                    try {
                        if ($JobArgs.type -eq 'M365') {
                            Invoke-WithRetry -OperationName 'Add-UnifiedGroupLinks Members' -ScriptBlock {
                                Add-UnifiedGroupLinks -Identity $JobArgs.email -LinkType Members -Links $m -ErrorAction Stop
                            }
                        } else {
                            Invoke-WithRetry -OperationName 'Add-DistributionGroupMember' -ScriptBlock {
                                Add-DistributionGroupMember -Identity $JobArgs.email -Member $m -ErrorAction Stop
                            }
                        }
                        $okM++
                    } catch {
                        if ($_.Exception.Message -match 'already a member|exists|duplicate') { $skipM++ }
                        else { Write-Log ('Miembro falla: ' + $m + ' — ' + $_.Exception.Message) -Level WARN -Source 'GroupsCreate' }
                    }
                }
                Write-Log ("Miembros añadidos: $okM · ya existían: $skipM · total intentos: $($JobArgs.members.Count)") -Level OK -Source 'GroupsCreate'
            }

            # Hidden GAL
            if ($created -and $JobArgs.hidden) {
                try {
                    if ($JobArgs.type -eq 'M365') {
                        Invoke-WithRetry -OperationName 'Set-UnifiedGroup hide' -ScriptBlock {
                            Set-UnifiedGroup -Identity $JobArgs.email -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                        }
                    } else {
                        Invoke-WithRetry -OperationName 'Set-DistributionGroup hide' -ScriptBlock {
                            Set-DistributionGroup -Identity $JobArgs.email -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                        }
                    }
                    Write-Log 'HiddenFromGAL=true aplicado.' -Level OK -Source 'GroupsCreate'
                } catch { Write-Log ('No se pudo ocultar de GAL: ' + $_.Exception.Message) -Level WARN -Source 'GroupsCreate' }
            }
        }
}

function Split-EmailList {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @($Value -split '[,;\s\r\n]+' | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() })
}

$BtnRunCreateGroup.Add_Click({
    if (-not (Test-RequiredConnections)) { return }
    $type    = if ($CreateType.SelectedItem) { [string]$CreateType.SelectedItem.Tag } else { 'DL' }
    $email   = $CreateEmail.Text.Trim()
    $display = $CreateDisplay.Text.Trim()
    $alias   = $CreateAlias.Text.Trim()
    $owners  = Split-EmailList -Value $CreateOwners.Text
    $hidden  = [bool]$CreateHidden.IsChecked
    if (-not $email -or -not (Test-Email -Value $email)) { $GroupsStatus.Text = 'Correo del grupo inválido.'; return }
    if (-not $display) { $GroupsStatus.Text = 'Nombre visible obligatorio.'; return }
    $GroupsStatus.Text = "Creando $type · $email..."
    Invoke-CreateGroupJob -Type $type -Email $email -DisplayName $display -Alias $alias -Owners $owners -Members @() -Hidden $hidden
})

$BtnRunCreateAddGroup.Add_Click({
    if (-not (Test-RequiredConnections)) { return }
    $type    = if ($CreateAddType.SelectedItem) { [string]$CreateAddType.SelectedItem.Tag } else { 'DL' }
    $email   = $CreateAddEmail.Text.Trim()
    $display = $CreateAddDisplay.Text.Trim()
    $alias   = $CreateAddAlias.Text.Trim()
    $owners  = Split-EmailList -Value $CreateAddOwners.Text
    $members = Split-EmailList -Value $CreateAddMembers.Text
    $hidden  = [bool]$CreateAddHidden.IsChecked
    if (-not $email -or -not (Test-Email -Value $email)) { $GroupsStatus.Text = 'Correo del grupo inválido.'; return }
    if (-not $display) { $GroupsStatus.Text = 'Nombre visible obligatorio.'; return }
    if ($members.Count -eq 0) { $GroupsStatus.Text = 'Añade al menos un miembro o usa modo 2 (crear solo).'; return }
    $GroupsStatus.Text = ("Creando $type · $email · " + $members.Count + ' miembros...')
    Invoke-CreateGroupJob -Type $type -Email $email -DisplayName $display -Alias $alias -Owners $owners -Members $members -Hidden $hidden
})

$LegacyBrowse.Add_Click({ $p = Pick-CsvFile; if ($p) { $LegacyMembersCsv.Text = $p } })

$BtnLegacySearch.Add_Click({
    $term = $LegacyGroupSearch.Text.Trim()
    if (-not $term) { $GroupsStatus.Text = 'Búsqueda vacía.'; return }
    if (-not (Test-RequiredConnections)) { return }
    $LegacyCandidates.Items.Clear()
    $LegacyCandidates.Tag = $null
    $GroupsStatus.Text = "Buscando '$term'..."
    Invoke-GroupSearchJob -SearchText $term -TargetTag 'legacy'
})

$BtnRunLegacyMembers.Add_Click({
    $csv = $LegacyMembersCsv.Text.Trim()
    $cand = $LegacyCandidates.SelectedItem
    if (-not $cand -or -not $cand.Tag) { $GroupsStatus.Text = 'Selecciona un candidato de la lista primero.'; return }
    if (-not $csv -or -not (Test-Path -LiteralPath $csv)) { $GroupsStatus.Text = 'CSV inválido.'; return }
    if (-not (Test-RequiredConnections)) { return }

    $resolved = $cand.Tag  # PSCustomObject del candidato
    $smtp = [string]$resolved.PrimarySmtpAddress
    if (-not $smtp.ToLowerInvariant().StartsWith('testeo')) {
        $role = try { Get-CurrentRole } catch { 'operator' }
        if ($role -ne 'admin') {
            $r = [System.Windows.MessageBox]::Show("Grupo '$smtp' no empieza por 'testeo' y tu rol no es admin. ¿Continuar?", 'Confirmación', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { $GroupsStatus.Text = 'Cancelado.'; return }
        }
    }

    $GroupsStatus.Text = "Añadiendo miembros a $smtp..."
    Start-RunspaceJob -Name "Add Members → $smtp" `
        -Inputs @($csv) `
        -Defaults @{ 'CSV|Ruta' = $csv } `
        -JobArgs @{ script = (Join-Path $ScriptsPath 'Add-GroupMembers.ps1'); resolved = $resolved } `
        -Script {
            param($Sync,$JobArgs)
            # Pre-inject resolved group so the script's Resolve-GroupBySearch returns it.
            $r = $JobArgs.resolved
            $opType = switch ($r.GroupType) {
                'Microsoft365Group'        { 'UnifiedGroup' }
                'DistributionList'         { 'DistributionGroup' }
                'MailEnabledSecurityGroup' { 'DistributionGroup' }
                default                    { 'DistributionGroup' }
            }
            function global:Resolve-GroupBySearch { param([string]$Prompt)
                return [PSCustomObject]@{
                    GroupType          = $opType
                    DisplayType        = $r.GroupType
                    Identity           = [string]$r.Identity
                    Id                 = [string]$r.GroupId
                    DisplayName        = [string]$r.DisplayName
                    PrimarySmtpAddress = [string]$r.PrimarySmtpAddress
                }
            }
            & $JobArgs.script
        }
})

# --- Export ---

$ExportBrowse.Add_Click({
    $p = Pick-Folder -InitialPath $ExportFolder.Text
    if ($p) { $ExportFolder.Text = $p }
})

$BtnExportSearch.Add_Click({
    $term = $ExportSearch.Text.Trim()
    if (-not $term) { $ExportStatus.Text = 'Búsqueda vacía.'; return }
    if (-not (Test-RequiredConnections)) { return }
    $ExportCandidates.Items.Clear()
    $ExportCandidates.Tag = $null
    $ExportStatus.Text = "Buscando '$term'..."
    Invoke-GroupSearchJob -SearchText $term -TargetTag 'export'
})

$BtnRunExport.Add_Click({
    $cand = $ExportCandidates.SelectedItem
    if (-not $cand -or -not $cand.Tag) { $ExportStatus.Text = 'Selecciona un candidato primero (botón Buscar y luego elige uno).'; return }
    $folder = $ExportFolder.Text.Trim()
    if (-not $folder) { $ExportStatus.Text = 'Carpeta destino obligatoria.'; return }
    if (-not (Test-Path -LiteralPath $folder)) {
        try { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        catch { $ExportStatus.Text = 'No se pudo crear carpeta: ' + $_.Exception.Message; return }
    }
    if (-not (Test-RequiredConnections)) { return }

    $resolved = $cand.Tag
    $ExportStatus.Text = ("Exportando miembros de " + $resolved.PrimarySmtpAddress + "...")
    Start-RunspaceJob -Name ('Export · ' + $resolved.PrimarySmtpAddress) `
        -Inputs @('S', $folder) `
        -Defaults @{ '\(S/N\)' = 'S'; 'Carpeta' = $folder } `
        -JobArgs @{ script = (Join-Path $ScriptsPath 'Export-GroupMembers.ps1'); resolved = $resolved; folder = $folder } `
        -Script {
            param($Sync,$JobArgs)
            $r = $JobArgs.resolved
            # Override Resolve-GroupByMail so the export script skips interactive selection.
            function global:Resolve-GroupByMail { param([string]$GroupMail)
                return [PSCustomObject]@{
                    Found              = $true
                    Cancelled          = $false
                    GroupType          = [string]$r.GroupType
                    Alias              = $null
                    DisplayName        = [string]$r.DisplayName
                    PrimarySmtpAddress = [string]$r.PrimarySmtpAddress
                    Identity           = [string]$r.Identity
                    GroupId            = [string]$r.GroupId
                    RawObject          = $null
                    MatchCount         = 1
                    Source             = [string]$r.Source
                }
            }
            & $JobArgs.script
        }
})

# --- Convert mailbox ---

$BtnRunConvert.Add_Click({
    $upn = $ConvertUpn.Text.Trim()
    if (-not $upn) { $ConvertStatus.Text = 'UPN obligatorio.'; return }
    $exo = $false; try { $exo = Test-ExchangeOnlineConnected } catch {}
    if (-not $exo) {
        [System.Windows.MessageBox]::Show('Convertir buzón necesita Exchange Online conectado.', 'GREX365', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $upn.ToLowerInvariant().StartsWith('testeo') -and -not $ConvertForce.IsChecked) {
        $r = [System.Windows.MessageBox]::Show("El UPN '$upn' no empieza por 'testeo'. ¿Continuar?", 'Confirmación', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $ConvertStatus.Text = 'Cancelado.'; return }
    }
    $ConvertStatus.Text = 'Convirtiendo...'
    Start-RunspaceJob -Name "Convert mailbox $upn" `
        -Inputs @($upn, 'S', 'N') `
        -Defaults @{ 'Email|UPN' = $upn; 'Convertir' = 'S'; 'Comprobar otro' = 'N' } `
        -JobArgs @{ script = (Join-Path $ScriptsPath 'Convert-SharedToUserMailbox.ps1') } `
        -Script {
            param($Sync,$JobArgs)
            & $JobArgs.script
        }
})

# --- Self-test ---

$BtnRunSelfTest.Add_Click({
    $target   = $SelfTestTarget.Text.Trim()
    $delegate = $SelfTestDelegate.Text.Trim()
    $seed     = $SelfTestSeed.Text.Trim()
    $skip     = [bool]$SelfTestSkipCleanup.IsChecked

    if (-not (Test-RequiredConnections)) { return }

    foreach ($n in @($target, $delegate)) {
        if (-not $n.ToLowerInvariant().StartsWith('testeo')) {
            $r = [System.Windows.MessageBox]::Show("'$n' no empieza por 'testeo'. El self-test sólo debe correrse contra cuentas testeo*. ¿Continuar?", 'Confirmación', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { $SelfTestStatus.Text = 'Cancelado.'; return }
        }
    }

    $SelfTestStatus.Text = 'Ejecutando self-test...'
    Start-RunspaceJob -Name 'Self-test' `
        -JobArgs @{ script = (Join-Path $ScriptsPath 'Invoke-SelfTest.ps1'); target=$target; delegate=$delegate; seed=$seed; skip=$skip } `
        -Script {
            param($Sync,$JobArgs)
            $params = @{ TargetUpn = $JobArgs.target; DelegateUpn = $JobArgs.delegate; GroupNameSeed = $JobArgs.seed }
            if ($JobArgs.skip) { $params.SkipCleanup = $true }
            & $JobArgs.script @params
        }
})

# --- Jobs panel ---

$BtnJobsRefresh.Add_Click({ Refresh-JobsGrid })
$BtnJobsClear.Add_Click({
    try {
        Remove-FinishedJobs
        Refresh-JobsGrid
        $JobsStatus.Text = 'Jobs persistentes terminados eliminados (cola en disco).'
    } catch {
        $JobsStatus.Text = 'Error: ' + $_.Exception.Message
    }
})

# --- Cert panel ---

$BtnCertRefresh.Add_Click({ Refresh-CertPanel })
$BtnCertOpenConsole.Add_Click({
    $cmd = ('-NoExit -ExecutionPolicy Bypass -File "' + (Join-Path $LauncherRoot 'Main.ps1') + '"')
    try {
        Start-Process pwsh -ArgumentList $cmd
        $CertStatus.Text = 'Consola lanzada. Selecciona "Asistente de certificado" en el menú.'
    } catch {
        $CertStatus.Text = 'No se pudo abrir pwsh: ' + $_.Exception.Message
    }
})
$BtnCertOpenFolder.Add_Click({
    $p = Join-Path $RepoRoot 'config'
    if (Test-Path $p) { Start-Process $p } else { $CertStatus.Text = 'No existe la carpeta config.' }
})

# Welcome quick actions for new panels
$QuickSelfTest.Add_Click({ $SideNav.SelectedIndex = 8 })  # index of "selftest" in sidenav
$QuickToolkitCheck.Add_Click({
    $testAll = Join-Path $ScriptsPath 'Test-AllScripts.ps1'
    if (-not (Test-Path -LiteralPath $testAll)) {
        Append-Log -Message ('Test-AllScripts.ps1 no encontrado en ' + $testAll) -Level 'ERROR' -Source 'HealthCheck'
        return
    }
    Start-RunspaceJob -Name 'Toolkit health check' `
        -JobArgs @{ script = $testAll } `
        -Script {
            param($Sync,$JobArgs)
            & $JobArgs.script
        }
})

# --- Keyboard shortcuts ---
# Ctrl+L  clear log
# Ctrl+R  refresh status bar
# F5      refresh status bar (alias)
# Ctrl+1..Ctrl+9  jump to sidebar item
$Window.Add_PreviewKeyDown({
    $k = $_.Key
    $ctrl = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control
    if ($ctrl -and $k -eq 'L') { $LogList.Items.Clear(); $_.Handled = $true; return }
    if (($ctrl -and $k -eq 'R') -or $k -eq 'F5') { Refresh-StatusBar; Refresh-Chips; $_.Handled = $true; return }
    if ($ctrl) {
        $idxMap = @{ 'D1'=0; 'D2'=1; 'D3'=2; 'D4'=3; 'D5'=4; 'D6'=5; 'D7'=6; 'D8'=7; 'D9'=8 }
        if ($idxMap.ContainsKey([string]$k)) {
            $i = $idxMap[[string]$k]
            if ($i -lt $SideNav.Items.Count) { $SideNav.SelectedIndex = $i; $_.Handled = $true }
        }
    }
})

# Initial state
Load-PrefsToUI
Refresh-StatusBar
Refresh-Chips
Refresh-CertPanel
Append-Log -Message 'GUI iniciada. Carpeta Seguimiento Claude lista.' -Level 'OK' -Source 'GUI'
Append-Log -Message 'Atajos: Ctrl+L limpia log · Ctrl+R / F5 refresca · Ctrl+1..9 navega.' -Level 'INFO' -Source 'GUI'

$Window.Add_Closed({
    try { $timer.Stop() } catch {}
    try { Unregister-Event -SourceIdentifier 'GREX365.GuiTick' -ErrorAction SilentlyContinue } catch {}
    try {
        $snapshot = $JobsList.ToArray()
        foreach ($j in $snapshot) {
            try { $j.PowerShell.Stop() } catch {}
            try { $j.PowerShell.Dispose() } catch {}
        }
    } catch {}
    try { $RunspacePool.Close(); $RunspacePool.Dispose() } catch {}
})

$null = $Window.ShowDialog()
