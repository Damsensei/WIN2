# ============================================================
# TOMBOLA WPF V2 - Interface moderne
# Tirage par vagues + gros lots unitaires
# ============================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# 01 - CONFIGURATION
# ============================================================

$script:AppDir = Join-Path $env:APPDATA "TombolaGUI"
$script:ConfigPath = Join-Path $script:AppDir "config.json"
$script:DefaultExportCsv = Join-Path (Get-Location) "resultats_tirage.csv"
$script:DefaultSaveCsv   = Join-Path (Get-Location) "sauvegarde_tirage.csv"

if (-not (Test-Path $script:AppDir)) {
    New-Item -Path $script:AppDir -ItemType Directory -Force | Out-Null
}

$script:Config = [PSCustomObject]@{
    Title = "GRANDE TOMBOLA DE L'ECOLE"
    BannerPath = ""
    ParticipantsCsv = ""
    LotsCsv = ""
    ExportCsv = $script:DefaultExportCsv
    SaveCsv = $script:DefaultSaveCsv
    MaxVagues = 10
    ModeAffichageGagnants = "Groupe"
}

$script:Participants = @()
$script:LotsNormaux = @()
$script:GrosLots = @()
$script:Resultats = @()
$script:Ordre = 1
$script:Vague = 0
$script:NombreVagues = 0

# ============================================================
# 02 - FONCTIONS CONFIG
# ============================================================

function Charger-Config {
    if (Test-Path $script:ConfigPath) {
        try {
            $cfg = Get-Content -Path $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $script:Config.PSObject.Properties.Name) {
                if ($cfg.PSObject.Properties.Name -contains $p -and $null -ne $cfg.$p) {
                    $script:Config.$p = $cfg.$p
                }
            }
        }
        catch {}
    }
}

function Sauver-Config {
    $script:Config | ConvertTo-Json -Depth 4 | Set-Content -Path $script:ConfigPath -Encoding UTF8
}

# ============================================================
# 03 - FONCTIONS METIER
# ============================================================

function Tirer-Element {
    param([array]$Liste)
    $index = Get-Random -Minimum 0 -Maximum $Liste.Count
    return $Liste[$index]
}

function Calculer-NombreVagues {
    param(
        [int]$NombreLotsNormaux,
        [int]$MaxVagues
    )

    if ($NombreLotsNormaux -le 0) { return 0 }
    if ($MaxVagues -lt 1) { $MaxVagues = 1 }

    # Le nombre de vagues est pilote par la configuration.
    # Exemple : 120 lots avec MaxVagues=10 => 10 vagues de 12 lots.
    return [Math]::Min($NombreLotsNormaux, $MaxVagues)
}

function Calculer-TailleVague {
    param([int]$LotsRestants, [int]$VaguesRestantes)
    if ($VaguesRestantes -le 0) { return $LotsRestants }
    return [int][Math]::Ceiling($LotsRestants / $VaguesRestantes)
}

function Sauvegarder-Resultats {
    if ($script:Resultats.Count -gt 0) {
        $script:Resultats | Export-Csv -Path $script:Config.SaveCsv -Delimiter ";" -NoTypeInformation -Encoding UTF8
        $script:Resultats | Export-Csv -Path $script:Config.ExportCsv -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }
}

function Recalculer-Vague-Depuis-Resultats {
    $vaguesNormales = @($script:Resultats | Where-Object { $_.Type -eq "Normal" } | ForEach-Object { [int]$_.Vague })
    if ($vaguesNormales.Count -gt 0) { $script:Vague = ($vaguesNormales | Measure-Object -Maximum).Maximum } else { $script:Vague = 0 }
}

function Charger-Sauvegarde {
    if (-not (Test-Path $script:Config.SaveCsv)) { return }

    $res = [System.Windows.MessageBox]::Show(
        "Une sauvegarde existe. Voulez-vous reprendre le tirage ?",
        "Sauvegarde detectee",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($res -eq [System.Windows.MessageBoxResult]::Yes) {
        $script:Resultats = @(Import-Csv -Path $script:Config.SaveCsv -Delimiter ";")
        $script:Ordre = $script:Resultats.Count + 1
        Recalculer-Vague-Depuis-Resultats

        $ticketsDejaGagnes = @($script:Resultats | ForEach-Object { $_.Ticket })
        $lotsDejaGagnes = @($script:Resultats | ForEach-Object { $_.Lot })

        $script:Participants = @($script:Participants | Where-Object { $_.Ticket -notin $ticketsDejaGagnes })
        $script:LotsNormaux = @($script:LotsNormaux | Where-Object { $_.Description -notin $lotsDejaGagnes })
        $script:GrosLots = @($script:GrosLots | Where-Object { $_.Description -notin $lotsDejaGagnes })
    }
    else {
        Remove-Item $script:Config.SaveCsv -Force -ErrorAction SilentlyContinue
        $script:Resultats = @()
        $script:Ordre = 1
        $script:Vague = 0
    }
}

function Importer-ParticipantsDepuisChemin {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    $script:Participants = @(Import-Csv -Path $Path -Delimiter ";")
    $script:Config.ParticipantsCsv = $Path
    Sauver-Config
    return $true
}

function Importer-LotsDepuisChemin {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    $lots = @(Import-Csv -Path $Path -Delimiter ";")
    $script:LotsNormaux = @($lots | Where-Object { $_.GrosLot -notmatch "^(Oui|O|Yes|True|1)$" })
    $script:GrosLots = @($lots | Where-Object { $_.GrosLot -match "^(Oui|O|Yes|True|1)$" })
    $script:NombreVagues = Calculer-NombreVagues -NombreLotsNormaux $script:LotsNormaux.Count -MaxVagues ([int]$script:Config.MaxVagues)
    $script:Vague = 0
    $script:Config.LotsCsv = $Path
    Sauver-Config
    Charger-Sauvegarde
    return $true
}

# ============================================================
# 04 - XAML INTERFACE WPF
# ============================================================

Charger-Config

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Grande Tombola" MinWidth="1100" MinHeight="760" Width="1350" Height="900"
        WindowStartupLocation="CenterScreen" Background="#F4F7FB">
    <Window.Resources>
        <DropShadowEffect x:Key="SoftShadow" Color="#22000000" BlurRadius="18" ShadowDepth="4" Opacity="0.35"/>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Height" Value="48"/>
            <Setter Property="Padding" Value="18,8"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="18" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SetupLabel" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#334155"/>
            <Setter Property="Margin" Value="0,10,0,6"/>
        </Style>
        <Style x:Key="SetupTextBox" TargetType="TextBox">
            <Setter Property="Height" Value="38"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
    </Window.Resources>

    <DockPanel LastChildFill="True">
        <Grid DockPanel.Dock="Top" Height="230">
            <Grid.RowDefinitions>
                <RowDefinition Height="160"/>
                <RowDefinition Height="70"/>
            </Grid.RowDefinitions>
            <Border Grid.RowSpan="2" Background="#1E3A8A"/>
            <Border Grid.Row="0" ClipToBounds="True">
                <Border.Background>
                    <ImageBrush x:Name="BannerBrush" Stretch="UniformToFill" AlignmentX="Center" AlignmentY="Center"/>
                </Border.Background>
            </Border>
            <Border Grid.Row="1" Background="#CC1E3A8A">
                <TextBlock x:Name="TitleText" Text="GRANDE TOMBOLA" Foreground="White" FontSize="34" FontWeight="Black" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
        </Grid>

        <TabControl x:Name="MainTabs" FontSize="15" Background="#F4F7FB">
            <TabItem Header="Tirage">
                <Grid Margin="18">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="78"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="285"/>
                        <ColumnDefinition Width="18"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Row="0" Grid.Column="0">
                        <Border Background="White" CornerRadius="24" Padding="18" Effect="{StaticResource SoftShadow}">
                            <StackPanel>
                                <TextBlock Text="Tableau de bord" FontSize="22" FontWeight="Black" Foreground="#0F172A" Margin="0,0,0,15"/>
                                <Border Background="#EFF6FF" CornerRadius="18" Padding="15" Margin="0,6">
                                    <StackPanel><TextBlock Text="Tickets restants" Foreground="#475569"/><TextBlock x:Name="StatTickets" Text="0" FontSize="30" FontWeight="Black" Foreground="#2563EB"/></StackPanel>
                                </Border>
                                <Border Background="#F0FDF4" CornerRadius="18" Padding="15" Margin="0,6">
                                    <StackPanel><TextBlock Text="Lots normaux" Foreground="#475569"/><TextBlock x:Name="StatLots" Text="0" FontSize="30" FontWeight="Black" Foreground="#16A34A"/></StackPanel>
                                </Border>
                                <Border Background="#FEF3C7" CornerRadius="18" Padding="15" Margin="0,6">
                                    <StackPanel><TextBlock Text="Gros lots" Foreground="#475569"/><TextBlock x:Name="StatGrosLots" Text="0" FontSize="30" FontWeight="Black" Foreground="#D97706"/></StackPanel>
                                </Border>
                                <Border Background="#F5F3FF" CornerRadius="18" Padding="15" Margin="0,6">
                                    <StackPanel><TextBlock Text="Vague" Foreground="#475569"/><TextBlock x:Name="StatVague" Text="0 / 0" FontSize="30" FontWeight="Black" Foreground="#7C3AED"/></StackPanel>
                                </Border>
                                <Border Background="#FFF1F2" CornerRadius="18" Padding="15" Margin="0,6">
                                    <StackPanel><TextBlock Text="Gagnants" Foreground="#475569"/><TextBlock x:Name="StatGagnants" Text="0" FontSize="30" FontWeight="Black" Foreground="#E11D48"/></StackPanel>
                                </Border>
                            </StackPanel>
                        </Border>
                    </StackPanel>

                    <Border Grid.Row="0" Grid.Column="2" Background="White" CornerRadius="28" Padding="22" Effect="{StaticResource SoftShadow}">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="ResultsHost"/>
                        </ScrollViewer>
                    </Border>

                    <Grid Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,18,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                            <Button x:Name="BtnVague" Style="{StaticResource ModernButton}" Background="#16A34A" Content="LANCER UNE VAGUE" Width="240"/>
                            <Button x:Name="BtnGrosLot" Style="{StaticResource ModernButton}" Background="#D97706" Content="GROS LOT" Width="170"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="BtnExport" Style="{StaticResource ModernButton}" Background="#2563EB" Content="EXPORT CSV" Width="160"/>
                            <Button x:Name="BtnReset" Style="{StaticResource ModernButton}" Background="#DC2626" Content="RESET" Width="130"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="Setup">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="18">
                    <Border Background="White" CornerRadius="28" Padding="28" Effect="{StaticResource SoftShadow}">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="220"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="190"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <TextBlock Grid.Row="0" Grid.ColumnSpan="3" Text="Configuration de la tombola" FontSize="28" FontWeight="Black" Foreground="#0F172A" Margin="0,0,0,25"/>

                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Titre" Style="{StaticResource SetupLabel}"/>
                            <TextBox x:Name="TxtTitle" Grid.Row="1" Grid.Column="1" Style="{StaticResource SetupTextBox}"/>
                            <Button x:Name="BtnSaveTitle" Grid.Row="1" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="SAUVER" Height="38"/>

                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Banniere" Style="{StaticResource SetupLabel}"/>
                            <TextBox x:Name="TxtBanner" Grid.Row="2" Grid.Column="1" Style="{StaticResource SetupTextBox}" IsReadOnly="True"/>
                            <Button x:Name="BtnChooseBanner" Grid.Row="2" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="CHOISIR" Height="38"/>

                            <TextBlock Grid.Row="3" Grid.Column="0" Text="CSV participants" Style="{StaticResource SetupLabel}"/>
                            <StackPanel Grid.Row="3" Grid.Column="1">
                                <TextBox x:Name="TxtParticipants" Style="{StaticResource SetupTextBox}" IsReadOnly="True"/>
                                <TextBlock x:Name="LblParticipantsStatus" Foreground="#64748B" Margin="0,5,0,0"/>
                            </StackPanel>
                            <Button x:Name="BtnImportParticipants" Grid.Row="3" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="IMPORTER" Height="38"/>

                            <TextBlock Grid.Row="4" Grid.Column="0" Text="CSV lots" Style="{StaticResource SetupLabel}"/>
                            <StackPanel Grid.Row="4" Grid.Column="1">
                                <TextBox x:Name="TxtLots" Style="{StaticResource SetupTextBox}" IsReadOnly="True"/>
                                <TextBlock x:Name="LblLotsStatus" Foreground="#64748B" Margin="0,5,0,0"/>
                            </StackPanel>
                            <Button x:Name="BtnImportLots" Grid.Row="4" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="IMPORTER" Height="38"/>

                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Export resultats" Style="{StaticResource SetupLabel}"/>
                            <TextBox x:Name="TxtExport" Grid.Row="5" Grid.Column="1" Style="{StaticResource SetupTextBox}"/>
                            <Button x:Name="BtnChooseExport" Grid.Row="5" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="CHOISIR" Height="38"/>

                            <TextBlock Grid.Row="6" Grid.Column="0" Text="Sauvegarde reprise" Style="{StaticResource SetupLabel}"/>
                            <TextBox x:Name="TxtSave" Grid.Row="6" Grid.Column="1" Style="{StaticResource SetupTextBox}"/>
                            <Button x:Name="BtnChooseSave" Grid.Row="6" Grid.Column="2" Style="{StaticResource ModernButton}" Background="#2563EB" Content="CHOISIR" Height="38"/>

                            <TextBlock Grid.Row="7" Grid.Column="0" Text="Nombre max de vagues" Style="{StaticResource SetupLabel}"/>
                            <TextBox x:Name="TxtMaxVagues" Grid.Row="7" Grid.Column="1" Style="{StaticResource SetupTextBox}"/>

                            <TextBlock Grid.Row="8" Grid.Column="0" Text="Affichage gagnants" Style="{StaticResource SetupLabel}"/>
                            <ComboBox x:Name="CmbModeAffichage" Grid.Row="8" Grid.Column="1" Height="38" FontSize="15" Margin="0,8,15,8" SelectedIndex="0">
                                <ComboBoxItem Content="Groupe"/>
                                <ComboBoxItem Content="Unitaire"/>
                            </ComboBox>

                            <Button x:Name="BtnSaveSetup" Grid.Row="9" Grid.Column="1" HorizontalAlignment="Left" Style="{StaticResource ModernButton}" Background="#16A34A" Content="SAUVEGARDER SETUP" Width="240" Margin="0,28,0,0"/>
                        </Grid>
                    </Border>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ============================================================
# 05 - RECUPERATION CONTROLES
# ============================================================

$BannerBrush = $window.FindName("BannerBrush")
$TitleText = $window.FindName("TitleText")
$MainTabs = $window.FindName("MainTabs")
$ResultsHost = $window.FindName("ResultsHost")

$StatTickets = $window.FindName("StatTickets")
$StatLots = $window.FindName("StatLots")
$StatGrosLots = $window.FindName("StatGrosLots")
$StatVague = $window.FindName("StatVague")
$StatGagnants = $window.FindName("StatGagnants")

$BtnVague = $window.FindName("BtnVague")
$BtnGrosLot = $window.FindName("BtnGrosLot")
$BtnExport = $window.FindName("BtnExport")
$BtnReset = $window.FindName("BtnReset")

$TxtTitle = $window.FindName("TxtTitle")
$TxtBanner = $window.FindName("TxtBanner")
$TxtParticipants = $window.FindName("TxtParticipants")
$TxtLots = $window.FindName("TxtLots")
$TxtExport = $window.FindName("TxtExport")
$TxtSave = $window.FindName("TxtSave")
$TxtMaxVagues = $window.FindName("TxtMaxVagues")
$CmbModeAffichage = $window.FindName("CmbModeAffichage")
$LblParticipantsStatus = $window.FindName("LblParticipantsStatus")
$LblLotsStatus = $window.FindName("LblLotsStatus")

$BtnSaveTitle = $window.FindName("BtnSaveTitle")
$BtnChooseBanner = $window.FindName("BtnChooseBanner")
$BtnImportParticipants = $window.FindName("BtnImportParticipants")
$BtnImportLots = $window.FindName("BtnImportLots")
$BtnChooseExport = $window.FindName("BtnChooseExport")
$BtnChooseSave = $window.FindName("BtnChooseSave")
$BtnSaveSetup = $window.FindName("BtnSaveSetup")

# ============================================================
# 06 - FONCTIONS UI
# ============================================================

function New-TextBlock {
    param([string]$Text, [int]$Size = 16, [string]$Weight = "Normal", [string]$Color = "#0F172A")
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.FontSize = $Size
    $tb.FontWeight = $Weight
    $tb.Foreground = $Color
    $tb.TextWrapping = "Wrap"
    return $tb
}

function New-WinnerCard {
    param($Gagnant)

    $card = New-Object System.Windows.Controls.Border
    $card.Background = "#FFFFFF"
    $card.CornerRadius = "22"
    $card.Padding = "18"
    $card.Margin = "0,0,0,14"
    $card.BorderBrush = "#E2E8F0"
    $card.BorderThickness = 1

    $grid = New-Object System.Windows.Controls.Grid
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "2*" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" }))

    $left = New-Object System.Windows.Controls.StackPanel
    $ticket = New-TextBlock -Text "TICKET $($Gagnant.Ticket)" -Size 24 -Weight "Black" -Color "#2563EB"
    $person = New-TextBlock -Text "$($Gagnant.Prenom) $($Gagnant.Nom)" -Size 18 -Weight "Bold" -Color "#0F172A"
    $classe = New-TextBlock -Text "$($Gagnant.Classe)" -Size 15 -Weight "SemiBold" -Color "#64748B"
    $left.Children.Add($ticket) | Out-Null
    $left.Children.Add($person) | Out-Null
    $left.Children.Add($classe) | Out-Null

    $right = New-Object System.Windows.Controls.StackPanel
    $right.HorizontalAlignment = "Right"
    $right.VerticalAlignment = "Center"
    $lotTitle = New-TextBlock -Text "LOT GAGNE" -Size 13 -Weight "Bold" -Color "#16A34A"
    $lot = New-TextBlock -Text "$($Gagnant.Lot)" -Size 20 -Weight "Black" -Color "#166534"
    $right.Children.Add($lotTitle) | Out-Null
    $right.Children.Add($lot) | Out-Null

    [System.Windows.Controls.Grid]::SetColumn($left, 0)
    [System.Windows.Controls.Grid]::SetColumn($right, 1)
    $grid.Children.Add($left) | Out-Null
    $grid.Children.Add($right) | Out-Null
    $card.Child = $grid
    return $card
}

function Afficher-Accueil {
    $ResultsHost.Children.Clear()
    $ResultsHost.Children.Add((New-TextBlock -Text "Bienvenue dans la tombola" -Size 34 -Weight "Black" -Color "#0F172A")) | Out-Null
    $sub = New-TextBlock -Text "Importez les participants et les lots dans l'onglet Setup, puis lancez les vagues de tirage." -Size 18 -Color "#475569"
    $sub.Margin = "0,18,0,0"
    $ResultsHost.Children.Add($sub) | Out-Null
}

function New-GroupWinnerRow {
    param($Gagnant, [int]$Index)

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "0,0,0,8"
    $grid.Background = "#FFFFFF"

    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "85" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "170" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "170" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "90" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))

    $values = @(
        $Gagnant.Ticket,
        $Gagnant.Nom,
        $Gagnant.Prenom,
        $Gagnant.Classe,
        $Gagnant.Lot
    )

    for ($i = 0; $i -lt $values.Count; $i++) {
        $tb = New-TextBlock -Text ([string]$values[$i]) -Size 16 -Weight "SemiBold" -Color "#0F172A"
        $tb.Margin = "8,10,8,10"
        if ($i -eq 0) {
            $tb.Foreground = "#2563EB"
            $tb.FontWeight = "Black"
        }
        if ($i -eq 4) {
            $tb.Foreground = "#166534"
            $tb.FontWeight = "Black"
        }
        [System.Windows.Controls.Grid]::SetColumn($tb, $i)
        $grid.Children.Add($tb) | Out-Null
    }

    $border = New-Object System.Windows.Controls.Border
    $border.Background = "#FFFFFF"
    $border.BorderBrush = "#E2E8F0"
    $border.BorderThickness = 1
    $border.CornerRadius = "14"
    $border.Margin = "0,0,0,8"
    $border.Child = $grid
    return $border
}

function Afficher-Resultats-Vague {
    param([array]$GagnantsVague, [int]$NumeroVague, [int]$TotalVagues)

    # Mode groupe : un affichage compact sous forme de liste deroulante.
    # Objectif : afficher toute la vague sans empiler de grosses cartes.
    $ResultsHost.Children.Clear()

    $title = New-TextBlock -Text "RESULTATS VAGUE $NumeroVague / $TotalVagues" -Size 32 -Weight "Black" -Color "#0F172A"
    $title.Margin = "0,0,0,8"
    $ResultsHost.Children.Add($title) | Out-Null

    $info = New-TextBlock -Text "$($GagnantsVague.Count) gagnant(s) affiches en mode groupe" -Size 18 -Weight "SemiBold" -Color "#64748B"
    $info.Margin = "0,0,0,20"
    $ResultsHost.Children.Add($info) | Out-Null

    $header = New-Object System.Windows.Controls.Grid
    $header.Margin = "0,0,0,8"
    $header.Background = "#EFF6FF"
    $header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "85" }))
    $header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "170" }))
    $header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "170" }))
    $header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "90" }))
    $header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))

    $headers = @("Ticket", "Nom", "Prenom", "Classe", "Lot gagne")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $tb = New-TextBlock -Text $headers[$i] -Size 15 -Weight "Black" -Color "#1D4ED8"
        $tb.Margin = "8,10,8,10"
        [System.Windows.Controls.Grid]::SetColumn($tb, $i)
        $header.Children.Add($tb) | Out-Null
    }

    $headerBorder = New-Object System.Windows.Controls.Border
    $headerBorder.Background = "#EFF6FF"
    $headerBorder.BorderBrush = "#BFDBFE"
    $headerBorder.BorderThickness = 1
    $headerBorder.CornerRadius = "14"
    $headerBorder.Margin = "0,0,0,8"
    $headerBorder.Child = $header
    $ResultsHost.Children.Add($headerBorder) | Out-Null

    $idx = 1
    foreach ($g in $GagnantsVague) {
        $ResultsHost.Children.Add((New-GroupWinnerRow -Gagnant $g -Index $idx)) | Out-Null
        $idx++
    }
}

function Afficher-Resultats-Vague-Unitaire {
    param([array]$GagnantsVague, [int]$NumeroVague, [int]$TotalVagues)

    # Mode unitaire : une seule carte visible, navigation Precedent / Suivant.
    $script:UnitWinners = @($GagnantsVague)
    $script:UnitIndex = 0
    $script:UnitVague = $NumeroVague
    $script:UnitTotalVagues = $TotalVagues

    function script:Afficher-Gagnant-Unitaire-Courant {
        $ResultsHost.Children.Clear()

        if (-not $script:UnitWinners -or $script:UnitWinners.Count -eq 0) { return }

        $g = $script:UnitWinners[$script:UnitIndex]
        $position = $script:UnitIndex + 1
        $total = $script:UnitWinners.Count

        $title = New-TextBlock -Text "VAGUE $script:UnitVague / $script:UnitTotalVagues - GAGNANT $position / $total" -Size 30 -Weight "Black" -Color "#0F172A"
        $title.Margin = "0,0,0,18"
        $ResultsHost.Children.Add($title) | Out-Null

        $wrap = New-Object System.Windows.Controls.Border
        $wrap.Background = "#EFF6FF"
        $wrap.BorderBrush = "#2563EB"
        $wrap.BorderThickness = 3
        $wrap.CornerRadius = "34"
        $wrap.Padding = "34"
        $wrap.Margin = "0,0,0,22"

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.HorizontalAlignment = "Center"

        $t0 = New-TextBlock -Text "TICKET GAGNANT" -Size 20 -Weight "Black" -Color "#1D4ED8"
        $t0.HorizontalAlignment = "Center"

        $t1 = New-TextBlock -Text "$($g.Ticket)" -Size 66 -Weight "Black" -Color "#2563EB"
        $t1.HorizontalAlignment = "Center"
        $t1.Margin = "0,8,0,6"

        $t2 = New-TextBlock -Text "$($g.Prenom) $($g.Nom)" -Size 32 -Weight "Black" -Color "#0F172A"
        $t2.HorizontalAlignment = "Center"

        $tClasse = New-TextBlock -Text "$($g.Classe)" -Size 20 -Weight "SemiBold" -Color "#64748B"
        $tClasse.HorizontalAlignment = "Center"
        $tClasse.Margin = "0,4,0,0"

        $t3 = New-TextBlock -Text "Lot gagne : $($g.Lot)" -Size 30 -Weight "Black" -Color "#166534"
        $t3.HorizontalAlignment = "Center"
        $t3.Margin = "0,30,0,0"

        $sp.Children.Add($t0) | Out-Null
        $sp.Children.Add($t1) | Out-Null
        $sp.Children.Add($t2) | Out-Null
        $sp.Children.Add($tClasse) | Out-Null
        $sp.Children.Add($t3) | Out-Null
        $wrap.Child = $sp
        $ResultsHost.Children.Add($wrap) | Out-Null

        $nav = New-Object System.Windows.Controls.Grid
        $nav.Margin = "0,10,0,0"
        $nav.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))
        $nav.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))
        $nav.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))
        $nav.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))

        $btnPrev = New-Object System.Windows.Controls.Button
        $btnPrev.Content = "PRECEDENT"
        $btnPrev.Width = 170
        $btnPrev.Height = 52
        $btnPrev.Margin = "8"
        $btnPrev.FontSize = 16
        $btnPrev.FontWeight = "Bold"
        $btnPrev.Background = "#64748B"
        $btnPrev.Foreground = "White"
        $btnPrev.IsEnabled = ($script:UnitIndex -gt 0)

        $btnNext = New-Object System.Windows.Controls.Button
        $btnNext.Content = "SUIVANT"
        $btnNext.Width = 170
        $btnNext.Height = 52
        $btnNext.Margin = "8"
        $btnNext.FontSize = 16
        $btnNext.FontWeight = "Bold"
        $btnNext.Background = "#2563EB"
        $btnNext.Foreground = "White"
        $btnNext.IsEnabled = ($script:UnitIndex -lt ($script:UnitWinners.Count - 1))

        $btnPrev.Add_Click({
            if ($script:UnitIndex -gt 0) {
                $script:UnitIndex--
                Afficher-Gagnant-Unitaire-Courant
            }
        })

        $btnNext.Add_Click({
            if ($script:UnitIndex -lt ($script:UnitWinners.Count - 1)) {
                $script:UnitIndex++
                Afficher-Gagnant-Unitaire-Courant
            }
        })

        [System.Windows.Controls.Grid]::SetColumn($btnPrev, 1)
        [System.Windows.Controls.Grid]::SetColumn($btnNext, 2)
        $nav.Children.Add($btnPrev) | Out-Null
        $nav.Children.Add($btnNext) | Out-Null

        $ResultsHost.Children.Add($nav) | Out-Null
    }

    Afficher-Gagnant-Unitaire-Courant
}

function Afficher-GrosLot {
    param($Gagnant)
    $ResultsHost.Children.Clear()

    $wrap = New-Object System.Windows.Controls.Border
    $wrap.Background = "#FFFBEB"
    $wrap.BorderBrush = "#F59E0B"
    $wrap.BorderThickness = 3
    $wrap.CornerRadius = "32"
    $wrap.Padding = "35"

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.HorizontalAlignment = "Center"

    $t1 = New-TextBlock -Text "GAGNANT DU GROS LOT" -Size 34 -Weight "Black" -Color "#92400E"
    $t1.HorizontalAlignment = "Center"
    $t2 = New-TextBlock -Text "TICKET $($Gagnant.Ticket)" -Size 52 -Weight "Black" -Color "#D97706"
    $t2.HorizontalAlignment = "Center"
    $t2.Margin = "0,24,0,8"
    $t3 = New-TextBlock -Text "$($Gagnant.Prenom) $($Gagnant.Nom) - $($Gagnant.Classe)" -Size 26 -Weight "Bold" -Color "#0F172A"
    $t3.HorizontalAlignment = "Center"
    $t4 = New-TextBlock -Text "$($Gagnant.Lot)" -Size 34 -Weight "Black" -Color "#166534"
    $t4.HorizontalAlignment = "Center"
    $t4.Margin = "0,32,0,0"

    $sp.Children.Add($t1) | Out-Null
    $sp.Children.Add($t2) | Out-Null
    $sp.Children.Add($t3) | Out-Null
    $sp.Children.Add($t4) | Out-Null
    $wrap.Child = $sp
    $ResultsHost.Children.Add($wrap) | Out-Null
}

function Mettre-A-Jour-Stats {
    $StatTickets.Text = "$($script:Participants.Count)"
    $StatLots.Text = "$($script:LotsNormaux.Count)"
    $StatGrosLots.Text = "$($script:GrosLots.Count)"
    $StatVague.Text = "$($script:Vague) / $($script:NombreVagues)"
    $StatGagnants.Text = "$($script:Resultats.Count)"
    $LblParticipantsStatus.Text = "Participants charges : $($script:Participants.Count)"
    $LblLotsStatus.Text = "Lots normaux : $($script:LotsNormaux.Count) | Gros lots : $($script:GrosLots.Count) | Vagues : $($script:NombreVagues) | Mode : $($script:Config.ModeAffichageGagnants)"
}

function Mettre-A-Jour-Boutons {
    $BtnVague.IsEnabled = ($script:Participants.Count -gt 0 -and $script:LotsNormaux.Count -gt 0)
    $BtnGrosLot.IsEnabled = ($script:Participants.Count -gt 0 -and $script:GrosLots.Count -gt 0)
    $BtnExport.IsEnabled = ($script:Resultats.Count -gt 0)
}

function Appliquer-Config-UI {
    $TitleText.Text = $script:Config.Title
    $TxtTitle.Text = $script:Config.Title
    $TxtBanner.Text = $script:Config.BannerPath
    $TxtParticipants.Text = $script:Config.ParticipantsCsv
    $TxtLots.Text = $script:Config.LotsCsv
    $TxtExport.Text = $script:Config.ExportCsv
    $TxtSave.Text = $script:Config.SaveCsv
    $TxtMaxVagues.Text = "$($script:Config.MaxVagues)"
    if ($script:Config.ModeAffichageGagnants -eq "Unitaire") {
        $CmbModeAffichage.SelectedIndex = 1
    }
    else {
        $CmbModeAffichage.SelectedIndex = 0
    }

    if ($script:Config.BannerPath -and (Test-Path $script:Config.BannerPath)) {
        try {
            $uri = New-Object System.Uri($script:Config.BannerPath, [System.UriKind]::Absolute)
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit()
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bmp.UriSource = $uri
            $bmp.EndInit()
            $BannerBrush.ImageSource = $bmp
        }
        catch { $BannerBrush.ImageSource = $null }
    }
}

# ============================================================
# 07 - EVENEMENTS
# ============================================================

$BtnSaveTitle.Add_Click({
    $script:Config.Title = $TxtTitle.Text
    Sauver-Config
    Appliquer-Config-UI
})

$BtnChooseBanner.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Images (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Config.BannerPath = $dlg.FileName
        Sauver-Config
        Appliquer-Config-UI
    }
})

$BtnImportParticipants.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Importer-ParticipantsDepuisChemin -Path $dlg.FileName | Out-Null
        Appliquer-Config-UI; Mettre-A-Jour-Stats; Mettre-A-Jour-Boutons
    }
})

$BtnImportLots.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Importer-LotsDepuisChemin -Path $dlg.FileName | Out-Null
        Appliquer-Config-UI; Mettre-A-Jour-Stats; Mettre-A-Jour-Boutons
    }
})

$BtnChooseExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    $dlg.FileName = [System.IO.Path]::GetFileName($script:Config.ExportCsv)
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Config.ExportCsv = $dlg.FileName
        Sauver-Config
        Appliquer-Config-UI
    }
})

$BtnChooseSave.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    $dlg.FileName = [System.IO.Path]::GetFileName($script:Config.SaveCsv)
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Config.SaveCsv = $dlg.FileName
        Sauver-Config
        Appliquer-Config-UI
    }
})

$BtnSaveSetup.Add_Click({
    $script:Config.Title = $TxtTitle.Text
    $script:Config.ExportCsv = $TxtExport.Text
    $script:Config.SaveCsv = $TxtSave.Text

    $maxVaguesTmp = 10
    if ([int]::TryParse($TxtMaxVagues.Text, [ref]$maxVaguesTmp) -and $maxVaguesTmp -ge 1) {
        $script:Config.MaxVagues = $maxVaguesTmp
    }
    else {
        $script:Config.MaxVagues = 10
        $TxtMaxVagues.Text = "10"
    }

    $selectedMode = "Groupe"
    if ($CmbModeAffichage.SelectedItem -and $CmbModeAffichage.SelectedItem.Content) {
        $selectedMode = [string]$CmbModeAffichage.SelectedItem.Content
    }
    if ($selectedMode -notin @("Groupe", "Unitaire")) { $selectedMode = "Groupe" }
    $script:Config.ModeAffichageGagnants = $selectedMode

    if ($script:LotsNormaux.Count -gt 0) {
        $script:NombreVagues = Calculer-NombreVagues -NombreLotsNormaux $script:LotsNormaux.Count -MaxVagues ([int]$script:Config.MaxVagues)
    }

    Sauver-Config
    Appliquer-Config-UI
    [System.Windows.MessageBox]::Show("Configuration sauvegardee.") | Out-Null
})

$BtnVague.Add_Click({
    if ($script:Participants.Count -eq 0 -or $script:LotsNormaux.Count -eq 0) { return }

    $numeroVagueEnCours = $script:Vague + 1
    $vaguesRestantes = $script:NombreVagues - $script:Vague
    $nbTirages = Calculer-TailleVague -LotsRestants $script:LotsNormaux.Count -VaguesRestantes $vaguesRestantes
    if ($nbTirages -gt $script:Participants.Count) { $nbTirages = $script:Participants.Count }

    $gagnantsVague = @()
    for ($i = 1; $i -le $nbTirages; $i++) {
        $participant = Tirer-Element -Liste $script:Participants
        $lot = Tirer-Element -Liste $script:LotsNormaux

        $ligne = [PSCustomObject]@{
            Ordre  = $script:Ordre
            Vague  = $numeroVagueEnCours
            Type   = "Normal"
            Ticket = $participant.Ticket
            Nom    = $participant.Nom
            Prenom = $participant.Prenom
            Classe = $participant.Classe
            Lot    = $lot.Description
        }

        $script:Resultats += $ligne
        $gagnantsVague += $ligne
        $script:Participants = @($script:Participants | Where-Object { $_.Ticket -ne $participant.Ticket })
        $script:LotsNormaux = @($script:LotsNormaux | Where-Object { $_.Description -ne $lot.Description })
        $script:Ordre++
    }

    $script:Vague = $numeroVagueEnCours
    Sauvegarder-Resultats
    if ($script:Config.ModeAffichageGagnants -eq "Unitaire") {
        Afficher-Resultats-Vague-Unitaire -GagnantsVague $gagnantsVague -NumeroVague $numeroVagueEnCours -TotalVagues $script:NombreVagues
    }
    else {
        Afficher-Resultats-Vague -GagnantsVague $gagnantsVague -NumeroVague $numeroVagueEnCours -TotalVagues $script:NombreVagues
    }
    Mettre-A-Jour-Stats; Mettre-A-Jour-Boutons
    $MainTabs.SelectedIndex = 0
})

$BtnGrosLot.Add_Click({
    if ($script:Participants.Count -eq 0 -or $script:GrosLots.Count -eq 0) { return }

    $participant = Tirer-Element -Liste $script:Participants
    $lot = Tirer-Element -Liste $script:GrosLots

    $ligne = [PSCustomObject]@{
        Ordre  = $script:Ordre
        Vague  = "Gros lot"
        Type   = "Gros lot"
        Ticket = $participant.Ticket
        Nom    = $participant.Nom
        Prenom = $participant.Prenom
        Classe = $participant.Classe
        Lot    = $lot.Description
    }

    $script:Resultats += $ligne
    $script:Participants = @($script:Participants | Where-Object { $_.Ticket -ne $participant.Ticket })
    $script:GrosLots = @($script:GrosLots | Where-Object { $_.Description -ne $lot.Description })
    $script:Ordre++

    Sauvegarder-Resultats
    Afficher-GrosLot -Gagnant $ligne
    Mettre-A-Jour-Stats; Mettre-A-Jour-Boutons
    $MainTabs.SelectedIndex = 0
})

$BtnExport.Add_Click({
    Sauvegarder-Resultats
    [System.Windows.MessageBox]::Show("Export termine : $($script:Config.ExportCsv)") | Out-Null
})

$BtnReset.Add_Click({
    $res = [System.Windows.MessageBox]::Show("Reinitialiser le tirage ? La configuration sera conservee.", "Reset", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($res -eq [System.Windows.MessageBoxResult]::Yes) {
        Remove-Item $script:Config.SaveCsv -Force -ErrorAction SilentlyContinue
        $script:Participants = @(); $script:LotsNormaux = @(); $script:GrosLots = @(); $script:Resultats = @()
        $script:Ordre = 1; $script:Vague = 0; $script:NombreVagues = 0
        if ($script:Config.ParticipantsCsv -and (Test-Path $script:Config.ParticipantsCsv)) { Importer-ParticipantsDepuisChemin -Path $script:Config.ParticipantsCsv | Out-Null }
        if ($script:Config.LotsCsv -and (Test-Path $script:Config.LotsCsv)) { Importer-LotsDepuisChemin -Path $script:Config.LotsCsv | Out-Null }
        Afficher-Accueil; Mettre-A-Jour-Stats; Mettre-A-Jour-Boutons; Appliquer-Config-UI
    }
})

# ============================================================
# 08 - DEMARRAGE
# ============================================================

Appliquer-Config-UI
if ($script:Config.ParticipantsCsv -and (Test-Path $script:Config.ParticipantsCsv)) { Importer-ParticipantsDepuisChemin -Path $script:Config.ParticipantsCsv | Out-Null }
if ($script:Config.LotsCsv -and (Test-Path $script:Config.LotsCsv)) { Importer-LotsDepuisChemin -Path $script:Config.LotsCsv | Out-Null }
Appliquer-Config-UI
Afficher-Accueil
Mettre-A-Jour-Stats
Mettre-A-Jour-Boutons

[void]$window.ShowDialog()
