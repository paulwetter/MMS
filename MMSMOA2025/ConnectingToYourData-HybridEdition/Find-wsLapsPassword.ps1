function Find-wsLapsPassword {
    param ()

        If (-not(Get-module -ListAvailable -name Microsoft.Graph)) {
        write-host "MSGraph module is missing."
        write-host "Run from elevated PS: Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force"
        break
    }
    If (-not(Get-module -ListAvailable -name ActiveDirectory)) {
        write-host "ActiveDirectory module is missing."
        write-host "Run from elevated PS: Install-Module ActiveDirectory -force"
        break
    }

    $AppConfig = @{
        # By Default, enable all sources.
        SourcesEnabled = @{
            AD = $true
            MeId = $true    
        }
    }

    $configRoot = "$([Environment]::GetFolderPath('ApplicationData'))\WettersSource"
    $Configfile = "$configRoot\Find-wsLapsPassword.json"

    
    if (-not (Test-Path -Path $configRoot)) {
        New-Item -Path $configRoot -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $Configfile)) {
        $AppConfig|ConvertTo-Json|Set-Content -Path $Configfile
    } else {
        $content = Get-Content -Path $Configfile -Raw
        $AppConfig = $content | ConvertFrom-Json
    }
    
    #region Make logo image for page  see https://wetterssource.com/ondemandtoast [Update Images] for more details
    $LogoImage = "${Env:Temp}\MMSMOA.png"
    $B64Logo = @'
iVBORw0KGgoAAAANSUhEUgAAAMkAAABUCAYAAAAoPhOYAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAuIgAALiIBquLdkgAAAAd0SU1FB+YJAg0lLQ4k5wQAAAGHaVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49J++7vycgaWQ9J1c1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCc/Pg0KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyI+PHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj48cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0idXVpZDpmYWY1YmRkNS1iYTNkLTExZGEtYWQzMS1kMzNkNzUxODJmMWIiIHhtbG5zOnRpZmY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vdGlmZi8x
LjAvIj48dGlmZjpPcmllbnRhdGlvbj4xPC90aWZmOk9yaWVudGF0aW9uPjwvcmRmOkRlc2NyaXB0aW9uPjwvcmRmOlJERj48L3g6eG1wbWV0YT4NCjw/eHBhY2tldCBlbmQ9J3cnPz4slJgLAAAgnklEQVR4Xu1deZhU1ZX/vdq6qzcWgcYOTZBNUcRviMYJ6jjGSVxGsxjiTOJENDFm+1xGjToGJ4k6zhjQzEQSNWY30WgyETWOCyKCAw4CYkRMlEXZmgaa3pda35k/zjnWrVvv1dILS1O/77tfVb137333vTrnnvXe5xBAKOOIQDeAEIBK+8QhinYAEeN3LwBXxl8l91IqegB0AggCOEo+FR3Sb9g4BgAB63cZwxg1hxGDJACkhYhTQqhpACPle9JuUARiUnoBjLAYROHFEF7HyijjoCMBwAGrOUFDCkbkuGs3yANXpEdM+hgl/ReLMpMcASAAfTITHw4gIWhHxpyUT5WCQSH8YhglKWoUpL0DoE7a
xox6JOcc45iizCRHABIyEx8uxmdKxhyQ767YCqoehYpk+JhIkAoAtdJnUNpHDRunEMpMcgSgVwijP4buQED9ZEy1N1SiVMtvk0nyqVwEoEvuu1YYTG2cKqkTkdJttPEba5lJjgCEDOI4UEgAWAVgXR5izoeIeKKUmF0PJk9Zv2HYH2kxzrVtnzwDU52KSr0+HzVLUWaSIwC1onIcSKQBNAPY1w8m6ZQZPihj1/YmIYc9jO+U2B8By3ul9o39DIIipfoMD5oXnHKcpIyhAAmhBwx1qRhsBPCGEPUpAKYAiAtD1Br10qJSjZC6andVWNdzJd5S5yGJFD1yjZBcw5YqfsxTxjBB0kctGWo4QnClMEgzgDVC1C0ANhg2g02oAcP7FReGiXpcr0ekjh+DwFDD+uwTAvvaZQwjqH4et08cIjBVmLSoZgkhakcM75TUswlV1adOIe5qYRITKZkk7OM2HJFItltYYV+7jGGElBCfnWZxsJESCdFrMEoQwDEAGmXM
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
'@
    [byte[]]$Bytes = [convert]::FromBase64String($B64Logo)
    If (!(Test-Path $LogoImage)) {
        [System.IO.File]::WriteAllBytes($LogoImage,$Bytes)
    }
    #endregion Make Logo    
    Function Write-LapsPassword {
        ## Function writes the bitlocker key information to the table in the WPF UI.
        [CmdletBinding()]
        param (
            [Parameter(Mandatory=$True)]
            [String]
            $Source,
            [Parameter(Mandatory=$True)]
            [String]
            $Account,
            [Parameter(Mandatory=$True)]
            [String]
            $Computer,
            [Parameter(Mandatory=$True)]
            [String]
            $Password,
            [Parameter(Mandatory=$True)]
            [String]
            $PasswordTime
        )
        $window.View1.ItemsSource += [PSCustomObject]@{ Source = "$Source"; Computer = "$Computer"; Account = "$Account"; Password = "$Password"; PasswordTime = "$PasswordTime"}
        #$window.View1.ScrollIntoView($window.View1.items.Item(($window.View1.ItemsSource).count - 1))
        $window.Dispatcher.Invoke([action]{},"Render")
    }
    
    #Region Common Fuctions for building a WPF UI 
    function Convert-XAMLtoWindow {
        param
        (
            [Parameter(Mandatory=$true)]
            [string]
            $XAML
        )
        
        Add-Type -AssemblyName PresentationFramework
        
        $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
        $result = [Windows.Markup.XAMLReader]::Load($reader)
        $reader.Close()
        $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
        while ($reader.Read())
        {
            $name=$reader.GetAttribute('Name')
            if (!$name) { $name=$reader.GetAttribute('x:Name') }
            if($name)
            {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
        }
        $reader.Close()
        $result
    }
    
    function Show-WPFWindow {
        param
        (
            [Parameter(Mandatory)]
            [Windows.Window]
            $Window
        )
        
        $result = $null
        $null = $window.Dispatcher.InvokeAsync{
            $result = $window.ShowDialog()
            Set-Variable -Name result -Value $result -Scope 1
        }.Wait()
        $result
    }
    #EndRegion Common Fuctions for building a WPF UI 
    
    #Region XAML for creating the WPF UI - The only variable that goes into this is the logo image.
    $xaml1 = @"
<Window
 xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
 xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
 Title='Find LAPS Passwords' Width="800"  SizeToContent="Height">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="80" />
                <RowDefinition Height="200" />
                <RowDefinition Height="28" />
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Grid.Column="0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="280"/>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="280"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="42" />
                    <RowDefinition Height="42" />
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" TextAlignment="Center" Margin="5">Computer Name:</TextBlock>
                <TextBox Grid.Row="0" Grid.Column="1" Name="ComputerName" Margin="5" HorizontalAlignment="left" VerticalContentAlignment="center" Width="250" Height="22"/>
                <Button Grid.Row="0" Grid.Column="2" Name="ButFindLaps" MinWidth="80" Height="22" Margin="2" HorizontalAlignment="Left" FontSize="12" Content="Find Password" />
                <Image Width="201" Height="84" Grid.Row="0" Grid.Column="3" Grid.RowSpan="2" HorizontalAlignment="right" Margin="0">
                    <Image.Source>
                        <BitmapImage DecodePixelWidth="201" UriSource="$LogoImage" />
                    </Image.Source>
                </Image>
                <Button Grid.Row="1" Grid.Column="0" Name="ConnectGraph" IsDefault="False" MinWidth="100" Height="22" Margin="2" Padding="2" HorizontalAlignment="Right" Background="Yellow" Foreground="Black" FontSize="11">Connect to Graph</Button>
                <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                <TextBlock Text="Sources:" Margin="50,0,10,0" />
                <CheckBox Content="AD" Margin="10,0,10,0" Name="EnableAD" />
                <CheckBox Content="ME-ID" Margin="10,0,10,0" Name="EnableMEID" />
                </StackPanel>
            </Grid>
            <ListView Grid.Row="1" Grid.Column="0" Name="View1" SelectionMode="Single">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Width="100" Header="Computer" DisplayMemberBinding="{Binding Computer}"/>
                        <GridViewColumn Width="100" Header="Account" DisplayMemberBinding="{Binding Account}"/>
                        <GridViewColumn Width="200" Header="Password" DisplayMemberBinding="{Binding Password}"/>
                        <GridViewColumn Width="100" Header="Source" DisplayMemberBinding="{Binding Source}"/>
                        <GridViewColumn Width="250" Header="Password Time" DisplayMemberBinding="{Binding PasswordTime}"/>
                    </GridView>
                </ListView.View>
            </ListView>
            <Grid Grid.Row="2" Grid.Column="0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="350"/>
                    <ColumnDefinition Width="425"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" HorizontalAlignment="Left" Name='ClearList' MinWidth="80" Margin="3" Content="Clear List" />
                <Button Grid.Column="1" HorizontalAlignment="Right" Name='CopyLapsPass' MinWidth="120" Margin="3" Content="Copy Password" />
            </Grid>
        </Grid>
    </Window>
"@

    # build the window object to manipulate and interact with in powershell.
    $window = Convert-XAMLtoWindow -XAML $xaml1
    $window.View1.ItemsSource = @()

    #region check what sources are enabled
    $window.EnableAD.IsChecked = $AppConfig.SourcesEnabled.AD
    $window.EnableMEID.IsChecked = $AppConfig.SourcesEnabled.MeId
    #endregion check sources

    #region Enable/Disable Sources
    function Save-SourceConfig {
        $AppConfig.SourcesEnabled = @{
            AD = $window.EnableAD.IsChecked
            MeId = $window.EnableMEID.IsChecked
        }
        if (-not (Test-Path -Path $configRoot)) {
            New-Item -Path $configRoot -ItemType Directory -Force
        }
        $AppConfig|ConvertTo-Json|Set-Content -Path $Configfile -Force
    }

    $window.EnableAD.add_Click{
        Save-SourceConfig
    }
    $window.EnableMEID.add_Click{
        Save-SourceConfig
    }

    #endregion Enable/Disable sources

    #Region Execute Search
    #add a Click action to the Find Keys button.
    $window.ButFindLaps.add_Click{
        #change the button to searching....
        $window.ButFindLaps.Content = "Searching..."
        $window.Dispatcher.Invoke([action] {}, "Render")
        #Read the key text from the input box in the dialog
        $ComputerToFind = $window.ComputerName.Text

        If ($AppConfig.SourcesEnabled.MeId -eq $true){
            $lapsPass = Get-LapsAADPassword -DeviceIds $ComputerToFind -IncludePasswords -AsPlainText -ErrorAction SilentlyContinue
            If(-not [string]::IsNullOrEmpty($lapsPass.Password)){
                Write-LapsPassword -Source "Entra ID" -Computer $lapsPass.DeviceName -Account $lapsPass.Account -Password $lapsPass.Password -PasswordTime $lapsPass.PasswordUpdateTime.ToString()
            }    
        }

        if ($AppConfig.SourcesEnabled.AD -eq $true){
            $lapsPass = Get-LapsADPassword -Identity $ComputerToFind -AsPlainText -ErrorAction SilentlyContinue
            If(-not [string]::IsNullOrEmpty($lapsPass.Password)){
                Write-LapsPassword -Source "Active Directory" -Computer $lapsPass.ComputerName -Account $lapsPass.Account -Password $lapsPass.Password -PasswordTime $lapsPass.PasswordUpdateTime.ToString()
            }    
        }
        #set the button back to original text so it shows that its done searching.
        $window.ButFindLaps.Content = "Find Password"
    }
    #EndRegion Execute Search

    #Add click action to button to clear the list of recovered keys.
    $window.ClearList.add_Click{
        $window.View1.ItemsSource = @()
    }

    function Test-GraphConnection {
        $Parameters = @{
            Method = "GET"
            URI = "/v1.0/me"
            OutputType = "HttpResponseMessage"
            ErrorAction = "Stop"
        }
        try {
            $Response = Invoke-GraphRequest @Parameters
            if ($Response) {
                return $true
            }
        } catch {
            return $false
        }    
    }

    #Button to copy the key to the clipboard
    $window.CopyLapsPass.add_Click{
        ($window.View1.SelectedItems).Password | Clip
    }

    # Button connects to graph api using the Graph powershell sdk.
    $window.ConnectGraph.add_Click{
        $Scopes = @(
            'Device.Read.All',
            'DeviceLocalCredential.ReadBasic.All',
            'DeviceLocalCredential.Read.All',
            'DeviceManagementManagedDevices.Read.All'
        )
        Connect-MgGraph -Scopes $Scopes -NoWelcome
        If (Test-GraphConnection){
            $window.ConnectGraph.Background="Green"
            $window.ConnectGraph.Foreground="White"
            $window.ConnectGraph.Content = "Connected to Graph"
        }
    }    

    If (Test-GraphConnection){
        $window.ConnectGraph.Background="Green"
        $window.ConnectGraph.Foreground="White"
        $window.ConnectGraph.Content = "Connected to Graph"
    }

    $null = Show-WPFWindow -Window $window
}
