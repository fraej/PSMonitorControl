# PS Monitor Control — PowerShell + WPF GUI
$SystemLanguage = (Get-UICulture).TwoLetterISOLanguageName

$Messages = @{
    'en' = @{
        'BrightnessLabel'     = 'Brightness:'
        'ContrastLabel'       = 'Contrast:'
        'VolumeLabel'         = 'Speaker volume (if available):'
        'SelectMonitor'       = 'Select Monitor:'
        'ResetWarningTitle'   = 'Confirmation'
        'ResetWarningMessage' = 'Are you sure you want to reset the monitor to factory values?'
        'ResetButton'         = 'Reset factory values'
        'InternalDisplay'     = 'Internal Display'
        'NoMonitorsFound'     = 'No compatible monitors were found. Please ensure your external monitor supports DDC/CI or that your laptop panel supports WMI brightness.'
    }
    'es' = @{
        'BrightnessLabel'     = 'Brillo:'
        'ContrastLabel'       = 'Contraste:'
        'VolumeLabel'         = 'Volumen altavoces (si tiene):'
        'SelectMonitor'       = 'Seleccionar Monitor:'
        'ResetWarningTitle'   = 'Confirmación'
        'ResetWarningMessage' = '¿Estás seguro de que quieres restablecer el monitor a los valores de fábrica?'
        'ResetButton'         = 'Restablecer valores de fábrica'
        'InternalDisplay'     = 'Pantalla interna'
        'NoMonitorsFound'     = 'No se encontraron monitores compatibles. Asegúrese de que su monitor externo sea compatible con DDC/CI o que el panel de su portátil soporte el brillo WMI.'
    }
    'fr' = @{
        'BrightnessLabel'     = 'Luminosité :'
        'ContrastLabel'       = 'Contraste :'
        'VolumeLabel'         = 'Volume des haut-parleurs (si disponible) :'
        'SelectMonitor'       = 'Sélectionner le moniteur :'
        'ResetWarningTitle'   = 'Confirmation'
        'ResetWarningMessage' = "Êtes-vous sûr de vouloir rétablir les valeurs d`'usine du moniteur ?"
        'ResetButton'         = "Réinitialiser les valeurs d`'usine"
        'InternalDisplay'     = 'Écran interne'
        'NoMonitorsFound'     = "Aucun moniteur compatible n`'a été trouvé. Veuillez vous assurer que votre moniteur externe prend en charge DDC/CI ou que la luminosité WMI est disponible."
    }
    'de' = @{
        'BrightnessLabel'     = 'Helligkeit:'
        'ContrastLabel'       = 'Kontrast:'
        'VolumeLabel'         = 'Lautsprecherlautstärke (falls verfügbar):'
        'SelectMonitor'       = 'Monitor auswählen:'
        'ResetWarningTitle'   = 'Bestätigung'
        'ResetWarningMessage' = 'Sind Sie sicher, dass Sie den Monitor auf die Werkseinstellungen zurücksetzen möchten?'
        'ResetButton'         = 'Werkseinstellungen zurücksetzen'
        'InternalDisplay'     = 'Internes Display'
        'NoMonitorsFound'     = 'Es wurden keine kompatiblen Monitore gefunden. Stellen Sie sicher, dass Ihr externer Monitor DDC/CI unterstützt oder dass Ihr Laptop-Display WMI-Helligkeit unterstützt.'
    }
    'ca' = @{
        'BrightnessLabel'     = 'Brillantor:'
        'ContrastLabel'       = 'Contrast:'
        'VolumeLabel'         = 'Volum dels altaveus (si en té):'
        'SelectMonitor'       = 'Seleccionar monitor:'
        'ResetWarningTitle'   = 'Confirmació'
        'ResetWarningMessage' = 'Esteu segur que voleu restablir el monitor als valors de fàbrica?'
        'ResetButton'         = 'Restablir els valors de fàbrica'
        'InternalDisplay'     = 'Pantalla interna'
        'NoMonitorsFound'     = "No s`'han trobat monitors compatibles. Assegureu-vos que el vostre monitor extern és compatible amb DDC/CI o que la pantalla del portàtil suporta la lluminositat WMI."
    }
}

function Get-LocalizedString {
    param($Key)
    if ($Messages.ContainsKey($SystemLanguage) -and $Messages[$SystemLanguage].ContainsKey($Key)) {
        return $Messages[$SystemLanguage][$Key]
    }
    return $Messages['en'][$Key]  # Fallback to English
}

# C# interop for DDC/CI monitor control
$monitorAPICode = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class MonitorControl 
{
    // --- Monitor enumeration via EnumDisplayMonitors ---
    public delegate bool MonitorEnumDelegate(IntPtr hMonitor, IntPtr hdcMonitor,
        ref RECT lprcMonitor, IntPtr dwData);

    [DllImport("user32.dll")]
    public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip,
        MonitorEnumDelegate lpfnEnum, IntPtr dwData);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int left, top, right, bottom;
    }

    private static List<IntPtr> _hMonitors = new List<IntPtr>();

    private static bool MonitorEnumCallback(IntPtr hMonitor, IntPtr hdcMonitor,
        ref RECT lprcMonitor, IntPtr dwData)
    {
        _hMonitors.Add(hMonitor);
        return true;
    }

    public static IntPtr[] GetAllMonitorHandles()
    {
        _hMonitors.Clear();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero,
            new MonitorEnumDelegate(MonitorEnumCallback), IntPtr.Zero);
        return _hMonitors.ToArray();
    }

    // --- Physical monitor APIs (dxva2.dll) ---
    [DllImport("dxva2.dll", SetLastError = true)]
    private static extern bool SetVCPFeature(IntPtr hMonitor, byte bVCPCode, uint dwNewValue);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetVCPFeatureAndVCPFeatureReply(IntPtr hMonitor, byte vcpCode, IntPtr pvct, ref uint pdwCurrentValue, ref uint pdwMaximum);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetNumberOfPhysicalMonitorsFromHMONITOR(IntPtr hMonitor, ref uint pdwNumberOfPhysicalMonitors);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetPhysicalMonitorsFromHMONITOR(IntPtr hMonitor, uint dwPhysicalMonitorArraySize, [Out] PHYSICAL_MONITOR[] pPhysicalMonitorArray);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool DestroyPhysicalMonitor(IntPtr hMonitor);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PHYSICAL_MONITOR {
        public IntPtr hPhysicalMonitor;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szPhysicalMonitorDescription;
    }

    public static bool SetMonitorBrightness(IntPtr handle, uint brightness) 
    {
        return SetVCPFeature(handle, 0x10, brightness);
    }

    public static bool SetMonitorContrast(IntPtr handle, uint contrast) 
    {
        return SetVCPFeature(handle, 0x12, contrast);
    }

    public static bool SetMonitorVolume(IntPtr handle, uint volume)
    {
        return SetVCPFeature(handle, 0x62, volume);
    }

    public static bool ResetMonitorFactory(IntPtr handle)
    {
        return SetVCPFeature(handle, 0x04, 1);
    }
}
"@

# Load C# type only if not already loaded in this session
if (-not ([System.Management.Automation.PSTypeName]'MonitorControl').Type) {
    try {
        Add-Type -TypeDefinition $monitorAPICode -Language CSharp
        Write-Host "Successfully loaded MonitorControl class"
    }
    catch {
        Write-Error "Failed to load MonitorControl class: $_"
        exit 1
    }
}

# Separate class for GetMonitorInfo (loaded independently so updates don't require a session restart)
$monitorInfoCode = @"
using System;
using System.Runtime.InteropServices;

public class MonitorInfoHelper
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int left, top, right, bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct MONITORINFOEX {
        public uint cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    public static string GetMonitorDeviceName(IntPtr hMonitor)
    {
        MONITORINFOEX info = new MONITORINFOEX();
        info.cbSize = (uint)Marshal.SizeOf(typeof(MONITORINFOEX));
        if (GetMonitorInfo(hMonitor, ref info))
        {
            return info.szDevice;
        }
        return null;
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'MonitorInfoHelper').Type) {
    try {
        Add-Type -TypeDefinition $monitorInfoCode -Language CSharp
        Write-Host "Successfully loaded MonitorInfoHelper class"
    }
    catch {
        Write-Error "Failed to load MonitorInfoHelper class: $_"
        exit 1
    }
}

# Separate class for DPI-aware monitor bounds (per-monitor DPI awareness)
$dpiAwareBoundsCode = @"
using System;
using System.Runtime.InteropServices;

public class DpiAwareBoundsHelper
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct MONITORINFOEX {
        public uint cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    // Per-thread DPI awareness (Windows 10 1607+)
    [DllImport("user32.dll")]
    public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

    public static int[] GetMonitorBounds(IntPtr hMonitor)
    {
        // Temporarily switch to per-monitor DPI awareness for accurate coordinates
        IntPtr prevContext = IntPtr.Zero;
        try { prevContext = SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2); } catch {}

        try
        {
            MONITORINFOEX info = new MONITORINFOEX();
            info.cbSize = (uint)Marshal.SizeOf(typeof(MONITORINFOEX));
            if (GetMonitorInfo(hMonitor, ref info))
            {
                return new int[] {
                    info.rcMonitor.left,
                    info.rcMonitor.top,
                    info.rcMonitor.right - info.rcMonitor.left,
                    info.rcMonitor.bottom - info.rcMonitor.top
                };
            }
        }
        finally
        {
            // Restore previous DPI context
            if (prevContext != IntPtr.Zero)
                try { SetThreadDpiAwarenessContext(prevContext); } catch {}
        }
        return new int[] { 0, 0, 1920, 1080 };
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'DpiAwareBoundsHelper').Type) {
    try {
        Add-Type -TypeDefinition $dpiAwareBoundsCode -Language CSharp
        Write-Host "Successfully loaded DpiAwareBoundsHelper class"
    }
    catch {
        Write-Error "Failed to load DpiAwareBoundsHelper class: $_"
        exit 1
    }
}

# Load required assemblies
Add-Type -AssemblyName PresentationFramework

# XAML UI definition
$xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PS Monitor Control" SizeToContent="Height" Width="580" WindowStyle="ToolWindow" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <StackPanel Margin="10">
        <!-- Top: Monitor selector -->
        <StackPanel Orientation="Horizontal" Margin="0,0,0,10" HorizontalAlignment="Center">
            <TextBlock x:Name="MonitorSelection" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <ComboBox x:Name="MonitorList" Width="200"/>
        </StackPanel>

        <!-- Bottom: Diagram (left) + Controls (right) -->
        <Grid x:Name="Stacked">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Left: Monitor layout diagram -->
            <Border Grid.Column="0" Background="#1E1E1E" CornerRadius="8" Margin="0,0,10,0">
                <Canvas x:Name="MonitorCanvas" ClipToBounds="True"/>
            </Border>

            <!-- Right: Controls -->
            <GroupBox Grid.Column="1" x:Name="MonitorGroup"
                     Header="{Binding SelectedItem, ElementName=MonitorList}">
                <StackPanel Margin="5">
                    <TextBlock x:Name="BrightnessLabel" Margin="0,5,0,5"/>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                        <Slider x:Name="BrightnessSlider" Width="220" Minimum="0" Maximum="100" 
                                IsEnabled="False" TickFrequency="10" TickPlacement="BottomRight" SmallChange="1" LargeChange="10"/>
                        <TextBlock Width="20" TextAlignment="Right" Margin="10,0,0,0"
                                 Text="{Binding Value, ElementName=BrightnessSlider, StringFormat={}{0:0}}"/>
                    </StackPanel>
                    
                    <TextBlock x:Name="ContrastLabel" Margin="0,0,0,5"/>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                        <Slider x:Name="ContrastSlider" Width="220" Minimum="0" Maximum="100" 
                                IsEnabled="False" TickFrequency="10" TickPlacement="BottomRight" SmallChange="1" LargeChange="10"/>
                        <TextBlock Width="20" TextAlignment="Right" Margin="10,0,0,0"
                                 Text="{Binding Value, ElementName=ContrastSlider, StringFormat={}{0:0}}"/>
                    </StackPanel>
                    
                    <TextBlock x:Name="VolumeLabel" Margin="0,0,0,5"/>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                        <Slider x:Name="VolumeSlider" Width="220" Minimum="0" Maximum="100" 
                                IsEnabled="False" TickFrequency="10" TickPlacement="BottomRight" SmallChange="1" LargeChange="10"/>
                        <TextBlock Width="20" TextAlignment="Right" Margin="10,0,0,0"
                                 Text="{Binding Value, ElementName=VolumeSlider, StringFormat={}{0:0}}"/>
                    </StackPanel>
                    
                    <Button x:Name="ResetButton"
                            Margin="0,10,0,0" Width="Auto" HorizontalAlignment="Center" Padding="5,0,5,0">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="5"/>
                            </Style>
                        </Button.Resources>
                        <Button.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                                <GradientStop Color="#FFEEEEEE" Offset="0"/>
                                <GradientStop Color="#FFDDDDDD" Offset="1"/>
                            </LinearGradientBrush>
                        </Button.Background>
                    </Button>
                </StackPanel>
            </GroupBox>
        </Grid>
    </StackPanel>
</Window>
"@

# Load XAML with proper error handling
try {
    $reader = [System.XML.XMLReader]::Create([System.IO.StringReader]::new($xamlContent))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

# Get controls
$monitorList = $window.FindName("MonitorList")
$monitorSelection = $window.FindName("MonitorSelection")
$monitorSelection.Text = Get-LocalizedString 'SelectMonitor'

$brightnessLabel = $window.FindName("BrightnessLabel")
$brightnessLabel.Text = Get-LocalizedString 'BrightnessLabel'
$brightnessSlider = $window.FindName("BrightnessSlider")

$contrastLabel = $window.FindName("ContrastLabel")
$contrastLabel.Text = Get-LocalizedString 'ContrastLabel'
$contrastSlider = $window.FindName("ContrastSlider")

$volumeLabel = $window.FindName("VolumeLabel")
$volumeLabel.Text = Get-LocalizedString 'VolumeLabel'
$volumeSlider = $window.FindName("VolumeSlider")

$resetButton = $window.FindName("ResetButton")
$resetButton.Content = Get-LocalizedString 'ResetButton'

$monitorGroup = $window.FindName("MonitorGroup")
$stackPanel = $window.FindName("Stacked")
$monitorCanvas = $window.FindName("MonitorCanvas")

# Flag to suppress DDC/CI writes during programmatic slider updates
$script:isUpdating = $false

# Enumerate monitors using proper HMONITOR handles via EnumDisplayMonitors
$monitors = @()

Write-Host "Detecting DDC/CI capable monitors..."

$hMonitors = [MonitorControl]::GetAllMonitorHandles()
$unmatchedDisplays = @()

foreach ($hMon in $hMonitors) {
    # Get Windows display number and screen bounds from HMONITOR
    $displayNumber = 0
    $deviceName = [MonitorInfoHelper]::GetMonitorDeviceName($hMon)
    if ($deviceName -match 'DISPLAY(\d+)') {
        $displayNumber = [int]$Matches[1]
    }
    $bounds = [DpiAwareBoundsHelper]::GetMonitorBounds($hMon)

    $ddcMatched = $false
    try {
        $monitorCount = [uint32]0
        [MonitorControl]::GetNumberOfPhysicalMonitorsFromHMONITOR($hMon, [ref]$monitorCount)

        if ($monitorCount -gt 0) {
            $physicalMonitorArray = New-Object MonitorControl+PHYSICAL_MONITOR[] $monitorCount

            if ([MonitorControl]::GetPhysicalMonitorsFromHMONITOR($hMon, $monitorCount, $physicalMonitorArray)) {
                foreach ($physMon in $physicalMonitorArray) {
                    # Test DDC/CI support by reading brightness (VCP 0x10)
                    $current = [uint32]0
                    $maximum = [uint32]0

                    if ([MonitorControl]::GetVCPFeatureAndVCPFeatureReply(
                            $physMon.hPhysicalMonitor, 0x10, [IntPtr]::Zero, [ref]$current, [ref]$maximum)) {

                        $monitors += @{
                            Name           = $physMon.szPhysicalMonitorDescription
                            Type           = 'DDC'
                            DisplayNumber  = $displayNumber
                            Bounds         = $bounds
                            PhysicalHandle = $physMon.hPhysicalMonitor
                        }
                        Write-Host "Found DDC/CI capable monitor: Display $displayNumber - $($physMon.szPhysicalMonitorDescription)"
                        $ddcMatched = $true
                    }
                    else {
                        # Monitor doesn't support DDC/CI — release its handle
                        [MonitorControl]::DestroyPhysicalMonitor($physMon.hPhysicalMonitor) | Out-Null
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Error detecting monitor: $_"
    }

    # Track displays that didn't match DDC/CI (candidates for WMI)
    if (-not $ddcMatched -and $displayNumber -gt 0) {
        $unmatchedDisplays += @{ DisplayNumber = $displayNumber; Bounds = $bounds }
    }
}

# Detect internal displays via WMI (laptop panels)
Write-Host "Detecting WMI brightness-capable internal displays..."
try {
    $wmiInstances = Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightness -ErrorAction Stop
    $wmiIndex = 0
    foreach ($wmiMon in $wmiInstances) {
        $currentWmiBrightness = $wmiMon.CurrentBrightness
        # Assign display number and bounds from a non-DDC display, if available
        $wmiDisplayNumber = 0
        $wmiBounds = @(0, 0, 1920, 1080)
        if ($wmiIndex -lt $unmatchedDisplays.Count) {
            $wmiDisplayNumber = $unmatchedDisplays[$wmiIndex].DisplayNumber
            $wmiBounds = $unmatchedDisplays[$wmiIndex].Bounds
            $wmiIndex++
        }
        $monitors += @{
            Name           = Get-LocalizedString 'InternalDisplay'
            Type           = 'WMI'
            DisplayNumber  = $wmiDisplayNumber
            Bounds         = $wmiBounds
            InstanceName   = $wmiMon.InstanceName
            PhysicalHandle = $null
        }
        Write-Host "Found WMI brightness-capable internal display: Display $wmiDisplayNumber (brightness: $currentWmiBrightness%)"
    }
}
catch {
    Write-Host "No WMI brightness-capable internal displays found"
}

# Sort monitors by Windows display number for natural ordering
$monitors = @($monitors | Sort-Object { $_.DisplayNumber })

# Update ComboBox using the real Windows display numbers
$monitorList.Items.Clear()
for ($i = 0; $i -lt $monitors.Count; $i++) {
    $num = $monitors[$i].DisplayNumber
    $displayName = if ($num -gt 0) { "{0}: {1}" -f $num, $monitors[$i].Name } else { $monitors[$i].Name }
    $monitorList.Items.Add($displayName)
}

if ($monitorList.Items.Count -eq 0) {
    $stackPanel.Children.Remove($monitorGroup)

    $noMonitorsText = New-Object System.Windows.Controls.TextBlock
    $noMonitorsText.Text = Get-LocalizedString 'NoMonitorsFound'
    $noMonitorsText.Foreground = "#FFFF0000"
    $noMonitorsText.FontSize = 14
    $noMonitorsText.TextAlignment = "Center"
    $noMonitorsText.TextWrapping = "Wrap"
    $noMonitorsText.Margin = New-Object System.Windows.Thickness(20)

    $stackPanel.Children.Add($noMonitorsText)
}

# Unified function to read monitor values and update all sliders
function Update-MonitorValues {
    param($selectedMonitor)

    $script:isUpdating = $true

    if ($selectedMonitor.Type -eq 'WMI') {
        # --- WMI internal display: brightness only ---
        try {
            $wmiMon = Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightness -ErrorAction Stop |
                Where-Object { $_.InstanceName -eq $selectedMonitor.InstanceName }
            if ($wmiMon) {
                $brightnessSlider.Maximum = 100
                $brightnessSlider.Value = $wmiMon.CurrentBrightness
                $brightnessSlider.IsEnabled = $true
            }
            else {
                $brightnessSlider.IsEnabled = $false
            }
        }
        catch {
            $brightnessSlider.IsEnabled = $false
        }

        # Contrast, volume, and factory reset are not available for internal displays
        $contrastSlider.IsEnabled = $false
        $volumeSlider.IsEnabled = $false
        $resetButton.IsEnabled = $false
    }
    else {
        # --- DDC/CI external display ---
        $resetButton.IsEnabled = $true

        # Brightness (VCP 0x10)
        $currentBrightness = [uint32]0
        $maxBrightness = [uint32]0
        $hasBrightness = [MonitorControl]::GetVCPFeatureAndVCPFeatureReply(
            $selectedMonitor.PhysicalHandle,
            0x10,
            [IntPtr]::Zero,
            [ref]$currentBrightness,
            [ref]$maxBrightness)
        $brightnessSlider.IsEnabled = $hasBrightness
        if ($hasBrightness) {
            $brightnessSlider.Maximum = $maxBrightness
            $brightnessSlider.Value = $currentBrightness
        }

        # Contrast (VCP 0x12)
        $currentContrast = [uint32]0
        $maxContrast = [uint32]0
        $hasContrast = [MonitorControl]::GetVCPFeatureAndVCPFeatureReply(
            $selectedMonitor.PhysicalHandle,
            0x12,
            [IntPtr]::Zero,
            [ref]$currentContrast,
            [ref]$maxContrast)
        $contrastSlider.IsEnabled = $hasContrast
        if ($hasContrast) {
            $contrastSlider.Maximum = $maxContrast
            $contrastSlider.Value = $currentContrast
        }

        # Volume (VCP 0x62)
        $currentVolume = [uint32]0
        $maxVolume = [uint32]0
        $hasAudio = [MonitorControl]::GetVCPFeatureAndVCPFeatureReply(
            $selectedMonitor.PhysicalHandle,
            0x62,
            [IntPtr]::Zero,
            [ref]$currentVolume,
            [ref]$maxVolume)
        $volumeSlider.IsEnabled = $hasAudio
        if ($hasAudio) {
            $volumeSlider.Maximum = $maxVolume
            $volumeSlider.Value = $currentVolume
        }
    }

    $script:isUpdating = $false
}

# --- Monitor layout diagram ---
function Update-MonitorDiagram {
    $monitorCanvas.Children.Clear()
    if ($monitors.Count -eq 0) { return }

    # Create adjusted bounds for rendering (snap nearly-centered monitors)
    $adj = @()
    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $b = $monitors[$i].Bounds
        $adj += ,@([double]$b[0], [double]$b[1], [double]$b[2], [double]$b[3])
    }

    # Snap horizontal centers: if two monitors are nearly center-aligned, align them exactly
    $snapThresholdX = ($adj | ForEach-Object { $_[2] } | Measure-Object -Maximum).Maximum * 0.05
    for ($i = 0; $i -lt $adj.Count; $i++) {
        for ($j = $i + 1; $j -lt $adj.Count; $j++) {
            $cxi = $adj[$i][0] + $adj[$i][2] / 2.0
            $cxj = $adj[$j][0] + $adj[$j][2] / 2.0
            if ([Math]::Abs($cxi - $cxj) -lt $snapThresholdX) {
                $avg = ($cxi + $cxj) / 2.0
                $adj[$i][0] = $avg - $adj[$i][2] / 2.0
                $adj[$j][0] = $avg - $adj[$j][2] / 2.0
            }
        }
    }

    # Snap vertical centers similarly
    $snapThresholdY = ($adj | ForEach-Object { $_[3] } | Measure-Object -Maximum).Maximum * 0.05
    for ($i = 0; $i -lt $adj.Count; $i++) {
        for ($j = $i + 1; $j -lt $adj.Count; $j++) {
            $cyi = $adj[$i][1] + $adj[$i][3] / 2.0
            $cyj = $adj[$j][1] + $adj[$j][3] / 2.0
            if ([Math]::Abs($cyi - $cyj) -lt $snapThresholdY) {
                $avg = ($cyi + $cyj) / 2.0
                $adj[$i][1] = $avg - $adj[$i][3] / 2.0
                $adj[$j][1] = $avg - $adj[$j][3] / 2.0
            }
        }
    }

    # Recalculate bounding box from adjusted positions
    $minX = ($adj | ForEach-Object { $_[0] } | Measure-Object -Minimum).Minimum
    $minY = ($adj | ForEach-Object { $_[1] } | Measure-Object -Minimum).Minimum
    $maxX = ($adj | ForEach-Object { $_[0] + $_[2] } | Measure-Object -Maximum).Maximum
    $maxY = ($adj | ForEach-Object { $_[1] + $_[3] } | Measure-Object -Maximum).Maximum

    $totalWidth = $maxX - $minX
    $totalHeight = $maxY - $minY
    if ($totalWidth -le 0 -or $totalHeight -le 0) { return }

    # Use actual canvas dimensions for precise centering (available after window renders)
    $canvasWidth = $monitorCanvas.ActualWidth
    $canvasHeight = $monitorCanvas.ActualHeight
    if ($canvasWidth -le 0 -or $canvasHeight -le 0) {
        $canvasWidth = 270.0
        $canvasHeight = 130.0
    }

    # Scale to fit with 15% padding for breathing room
    $scale = [Math]::Min($canvasWidth / $totalWidth, $canvasHeight / $totalHeight) * 0.85

    # Center the scaled layout within the canvas
    $scaledWidth = $totalWidth * $scale
    $scaledHeight = $totalHeight * $scale
    $offsetX = ($canvasWidth - $scaledWidth) / 2
    $offsetY = ($canvasHeight - $scaledHeight) / 2

    $converter = New-Object System.Windows.Media.BrushConverter

    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $mon = $monitors[$i]
        $b = $adj[$i]

        # Exact scaled position from snap-adjusted bounds
        $x = ($b[0] - $minX) * $scale + $offsetX
        $y = ($b[1] - $minY) * $scale + $offsetY
        $w = $b[2] * $scale
        $h = $b[3] * $scale

        # Monitor rectangle (no border thickness so adjacent monitors touch)
        $border = New-Object System.Windows.Controls.Border
        $border.Width = $w
        $border.Height = $h
        $border.CornerRadius = New-Object System.Windows.CornerRadius(4)
        $border.BorderThickness = New-Object System.Windows.Thickness(0)

        if ($i -eq $monitorList.SelectedIndex) {
            $border.Background = $converter.ConvertFromString('#555555')
            # Inset highlight: use a thinner inner border via Padding + nested border
            $border.Padding = New-Object System.Windows.Thickness(2)
            $inner = New-Object System.Windows.Controls.Border
            $inner.Background = $converter.ConvertFromString('#555555')
            $inner.BorderBrush = $converter.ConvertFromString('#AAAAAA')
            $inner.BorderThickness = New-Object System.Windows.Thickness(2)
            $inner.CornerRadius = New-Object System.Windows.CornerRadius(2)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = $mon.DisplayNumber.ToString()
            $label.Foreground = [System.Windows.Media.Brushes]::White
            $label.FontSize = 20
            $label.FontWeight = [System.Windows.FontWeights]::SemiBold
            $label.HorizontalAlignment = 'Center'
            $label.VerticalAlignment = 'Center'
            $inner.Child = $label
            $border.Child = $inner
        }
        else {
            $border.Background = $converter.ConvertFromString('#3A3A3A')

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = $mon.DisplayNumber.ToString()
            $label.Foreground = [System.Windows.Media.Brushes]::White
            $label.FontSize = 20
            $label.FontWeight = [System.Windows.FontWeights]::SemiBold
            $label.HorizontalAlignment = 'Center'
            $label.VerticalAlignment = 'Center'
            $border.Child = $label
        }

        # Position on canvas
        [System.Windows.Controls.Canvas]::SetLeft($border, $x)
        [System.Windows.Controls.Canvas]::SetTop($border, $y)

        # Click to select this monitor
        $border.Tag = $i
        $border.Cursor = [System.Windows.Input.Cursors]::Hand
        $border.Add_MouseLeftButtonDown({
            param($sender, $e)
            $monitorList.SelectedIndex = $sender.Tag
        })

        $monitorCanvas.Children.Add($border)
    }
}

# Set initial monitor selection and read values
if ($monitors.Count -gt 0) {
    $monitorList.SelectedIndex = 0
    Update-MonitorValues $monitors[0]
}

# Defer diagram drawing until the window has rendered (so ActualWidth/ActualHeight are available)
$window.Add_Loaded({
    Update-MonitorDiagram
})

# Handle monitor selection change
$monitorList.Add_SelectionChanged({
    if ($monitorList.SelectedItem) {
        $selectedMonitor = $monitors[$monitorList.SelectedIndex]
        Update-MonitorValues $selectedMonitor
    }
    else {
        $script:isUpdating = $true
        $brightnessSlider.IsEnabled = $false
        $contrastSlider.IsEnabled = $false
        $volumeSlider.IsEnabled = $false
        $script:isUpdating = $false
    }
    Update-MonitorDiagram
})

# Slider change handlers (guarded to prevent writes during programmatic updates)
$brightnessSlider.Add_ValueChanged({
    if (-not $script:isUpdating -and $monitorList.SelectedItem) {
        $selectedMonitor = $monitors[$monitorList.SelectedIndex]
        $brightness = [uint32]$brightnessSlider.Value
        if ($selectedMonitor.Type -eq 'WMI') {
            # Set brightness via WMI for internal displays
            try {
                $wmiMethods = Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightnessMethods -ErrorAction Stop |
                    Where-Object { $_.InstanceName -eq $selectedMonitor.InstanceName }
                if ($wmiMethods) {
                    Invoke-CimMethod -InputObject $wmiMethods -MethodName WmiSetBrightness -Arguments @{ Timeout = 1; Brightness = [byte]$brightness }
                }
            }
            catch {
                Write-Warning "Failed to set WMI brightness: $_"
            }
        }
        else {
            [MonitorControl]::SetMonitorBrightness($selectedMonitor.PhysicalHandle, $brightness)
        }
    }
})

$contrastSlider.Add_ValueChanged({
    if (-not $script:isUpdating -and $monitorList.SelectedItem) {
        $selectedMonitor = $monitors[$monitorList.SelectedIndex]
        $contrast = [uint32]$contrastSlider.Value
        [MonitorControl]::SetMonitorContrast($selectedMonitor.PhysicalHandle, $contrast)
    }
})

$volumeSlider.Add_ValueChanged({
    if (-not $script:isUpdating -and $monitorList.SelectedItem) {
        $selectedMonitor = $monitors[$monitorList.SelectedIndex]
        $volume = [uint32]$volumeSlider.Value
        [MonitorControl]::SetMonitorVolume($selectedMonitor.PhysicalHandle, $volume)
    }
})

# Reset button handler (DDC/CI only — disabled for WMI monitors)
$resetButton.Add_Click({
    if ($monitorList.SelectedItem) {
        $selectedMonitor = $monitors[$monitorList.SelectedIndex]
        if ($selectedMonitor.Type -eq 'WMI') { return }

        $result = [System.Windows.MessageBox]::Show(
            (Get-LocalizedString 'ResetWarningMessage'),
            (Get-LocalizedString 'ResetWarningTitle'),
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)

        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            [MonitorControl]::ResetMonitorFactory($selectedMonitor.PhysicalHandle)

            # Wait for monitor to apply reset
            Start-Sleep -Milliseconds 1000

            # Refresh slider values
            Update-MonitorValues $selectedMonitor
        }
    }
})

# Clean up physical monitor handles on window close (DDC/CI only)
$window.Add_Closing({
    foreach ($mon in $monitors) {
        if ($mon.Type -eq 'DDC' -and $mon.PhysicalHandle) {
            [MonitorControl]::DestroyPhysicalMonitor($mon.PhysicalHandle) | Out-Null
        }
    }
})

# Show window
$window.ShowDialog()