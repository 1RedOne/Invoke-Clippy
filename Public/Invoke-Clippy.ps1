<#
.Synopsis
    Notify your users or coworkers with a familiar friend

.DESCRIPTION
    Provides a helpful way to interact with users using Clippy, the beloved Office Assistant from Office '97

.EXAMPLE
    Invoke-Clippy -text 'Would you like to install Windows 10?' -Button1Text Yes -Button2Text 'Restart PC'
   
    Provides a helpful Clippy UI to notify users of the wonders of Windows 10, in case they were unaware.   If the user isn't interested, shuts down their PC.

.EXAMPLE
    Invoke-Clippy -text 'Would you like to install Windows 10?' -Button1Text Yes -Button2Text 'Restart PC' -DontHidePowerShell

    Same results as previous example, but the parent PowerShell instance is not hidden from the user.

.LINK
    http://www.foxdeploy.com/powerclippy
    https://github.com/1RedOne/Invoke-Clippy/

.PARAMETER TEXT
    Contains the text you'd like Clippy to display.  Defaults to "Hi! I am Clippy, your office assitant.  Would you like some assistance today?"

.PARAMETER Button1Text
    If specified, creates a button for the user to click.  Add code to line 70 to make the button function.  Include the text you'd like as the value for this param

.PARAMETER Button2Text
    If specified, creates a button for the user to click.  Add code to line 81 to make the button function.  Include the text you'd like as the value for this param

.PARAMETER DontHidePowerShell
    If used, the PowerShell instance will not be hidden from the user. This can be useful if you are calling the function from PowerShell ISE. 
#>

function Invoke-Clippy {
    param (
        [string] $Text = 'Hi! I am Clippy, your office assitant.  Would you like some assistance today?',
        [string] $Button1Text,
        [string] $Button2Text,
        [switch] $DontHidePowershell
    )

    # Add assemblies
    Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms

    # Extract icon from PowerShell to use as the NotifyIcon
    $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$PSHOME\powershell.exe")

    # Create XAML form in Visual Studio, ensuring the ListView looks chromeless
    [xml] $XAML =  '<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="MainWindow" Height="657" Width="525" Background="Transparent" AllowsTransparency="True" WindowStyle="None" Topmost="True">
            <Window.Resources>
                <Style x:Key="ClippyButton" TargetType="Button" >
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="Button">
                                <Border CornerRadius="5" Background="#FFFFFDCF" BorderThickness="1" Padding="2" BorderBrush="Black">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>

                    </Setter>

                </Style>
            </Window.Resources>
            <Grid Name="grid"  Height="200" Width="400">
                <ListView Name="listview" SelectionMode="Single" Margin="0,50,0,0" Foreground="White"
		        Background="Transparent" BorderBrush="Transparent" IsHitTestVisible="False">
                    <ListView.ItemContainerStyle>
                        <Style>
                            <Setter Property="Control.HorizontalContentAlignment" Value="Stretch"/>
                            <Setter Property="Control.VerticalContentAlignment" Value="Stretch"/>
                        </Style>
                    </ListView.ItemContainerStyle>
                    <Image x:Name="image" Height="140" Width="156" Source="$PSScriptRoot\..\Clippy.png"/>
                </ListView>
                <Grid x:Name="gr3id" Margin="150,-190,29,190">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="40"/>
                    </Grid.RowDefinitions>
                    <Rectangle Fill="#FFFFFDCF" Stroke="#FF000000" RadiusX="10" RadiusY="10"/>
                    <Path Fill="#FFFFFDCF" Stretch="Fill" Stroke="#FF000000" HorizontalAlignment="Left" Margin="30,-1.6,0,0" Width="25" Grid.Row="1" 
                Data="M22.166642,154.45381 L29.999666,187.66699 40.791059,154.54395"/>
                    <TextBlock HorizontalAlignment="Center" Name ="ClippyText" VerticalAlignment="Center" FontSize="16" Text="Hi! I am Clippy, your office assitant.  Would you like some assistance today?" TextWrapping="Wrap" FontFamily="Arial" Height="88" Margin="2,20,3,52"/>
                </Grid>
                <Button Name="button1" Style="{StaticResource ClippyButton}" Content="Button" HorizontalAlignment="Left" Height="38" Margin="278,-78,0,0" VerticalAlignment="Top" Width="75"/>
                <Button Name="button2" Style="{StaticResource ClippyButton}" Content="Button" HorizontalAlignment="Left" Height="38" Margin="160,-78,0,0" VerticalAlignment="Top" Width="72"/>
            </Grid>
        </Window>
    '

    # Turn XAML into PowerShell objects
    $Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML))
    $XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Window.FindName($_.Name) -Scope Script }

    # Set textbox content and button text and behavior

    $ClippyText.Text = $Text

    if ($Button1Text) {
        $Button1.Content = $Button1Text
        $Button1.Visibility = 'Hidden'

        $Button1.add_Click({
            # code to execute when the second button is clicked
        })
    }

    if ($Button2Text) {
        $Button2.Content = $Button2Text
        $Button2.Visibility = 'Hidden'

        $Button2.add_Click({
            # code to execute when the second button is clicked
        })
    }
    # Create notifyicon, and right-click -> Exit menu
    $NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $NotifyIcon.Text = 'Clippy'
    $NotifyIcon.Icon = $Icon
    $NotifyIcon.Visible = $true

    $MenuItem = New-Object System.Windows.Forms.MenuItem
    $MenuItem.Text = "Exit"

    $contextmenu = New-Object System.Windows.Forms.ContextMenu
    $NotifyIcon.ContextMenu = $contextmenu
    $NotifyIcon.contextMenu.MenuItems.AddRange($MenuItem)

    # Add a left click that makes the Window appear in the lower right
    # part of the screen, above the notify icon.
    $NotifyIcon.add_Click({
	    if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
			    # reposition each time, in case the resolution or monitor changes
			    $Window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Window.Width)
			    $Window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Window.Height)
			    $Window.Show()
			    $Window.Activate()
	    }
    })

    # Close the window if it's double clicked
    $Window.Add_MouseDoubleClick({
	    $Window.Hide()
    })

    # Close the window if it loses focus
    $Window.Add_Deactivated({
	    Start-Sleep -Milliseconds 500
        $Window.Hide()
    })

    # When Exit is clicked, close everything and kill the PowerShell process
    $MenuItem.add_Click({
	    $NotifyIcon.Visible = $false
	    $Window.Close()
	    Stop-Process $PID
    })

 
    if (-not $DontHidePowerShell) {
        # Make PowerShell Disappear
        $WindowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
        $AsyncWindow = Add-Type -MemberDefinition $WindowCode -Name Win32ShowWindowAsync -namespace Win32Functions -PassThru
        $null = $AsyncWindow::ShowWindowAsync((Get-Process -PID $PID).MainWindowHandle, 0)
    }

    # Force garbage collection just to start slightly lower RAM usage.
    [System.GC]::Collect()

    # Create an application context for it to all run within.
    # This helps with responsiveness, especially when clicking Exit.
    $AppContext = New-Object System.Windows.Forms.ApplicationContext
    [void] [System.Windows.Forms.Application]::Run($AppContext)
}