<#
.Synopsis
   Notify your users or coworkers with a familiar friend
.DESCRIPTION
   Provides a helpful way to interact with users using Clippy, the beloved Office Assistant from Office '97
.EXAMPLE
   Invoke-Clippy -text 'Would you like to install Windows 10?' -Button1Text Yes -Button2Text 'Restart PC'
   
   Provides a helpful Clippy UI to notify users of the wonders of Windows 10, in case they were unaware.   If the user isn't interested, shuts down their PC.
.LINK
   http://www.foxdeploy.com/powerclippy
   https://github.com/1RedOne/Invoke-Clippy/
.PARAMETER TEXT
  Contains the text you'd like Clippy to display.  Defaults to "Hi! I am Clippy, your office assitant.  Would you like some assistance today?"
.PARAMETER Button1Text
  If specified, creates a button for the user to click.  Add code to line 70 to make the button function.  Include the text you'd like as the value for this param
.PARAMETER Button2Text
  If specified, creates a button for the user to click.  Add code to line 81 to make the button function.  Include the text you'd like as the value for this param
#>
Function Invoke-Clippy{
param(
    $text="Hi! I am Clippy, your office assitant.  Would you like some assistance today?",
    $Button1Text,$Button2Text
)
# Add assemblies
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms

# Extract icon from PowerShell to use as the NotifyIcon
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$pshome\powershell.exe")

# Create XAML form in Visual Studio, ensuring the ListView looks chromeless
    [xml]$xaml =  @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="657" Width="525" Background="Transparent" AllowsTransparency="True" WindowStyle="None" Topmost="True">
    <Window.Resources>
        <Style x:Key="ClippyButton"  TargetType="Button"  >
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border CornerRadius="5" Background="#FFFFFDCF" BorderThickness="1" Padding="2" BorderBrush="Black">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
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
            <Image x:Name="image" Height="140" Width="156" Source="$script:ModuleRoot\Clippy.png"/>
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
"@

# Turn XAML into PowerShell objects
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) -Scope Script }

# Set textbox content and button text and behavior

$ClippyText.Text = $text

if ($Button1Text){
$button1.Content = $Button1Text
$button1.Visibility = 'Hidden'

$button1.add_Click({
    #code to execute when the second button is clicked
    })
}

if ($Button2Text){
$button2.Content = $Button2Text
$button2.Visibility = 'Hidden'

$button2.add_Click({
    #code to execute when the second button is clicked
    })
}
# Create notifyicon, and right-click -> Exit menu
$notifyicon = New-Object System.Windows.Forms.NotifyIcon
$notifyicon.Text = "Clippy"
$notifyicon.Icon = $icon
$notifyicon.Visible = $true

$menuitem = New-Object System.Windows.Forms.MenuItem
$menuitem.Text = "Exit"

$contextmenu = New-Object System.Windows.Forms.ContextMenu
$notifyicon.ContextMenu = $contextmenu
$notifyicon.contextMenu.MenuItems.AddRange($menuitem)

# Add a left click that makes the Window appear in the lower right
# part of the screen, above the notify icon.
$notifyicon.add_Click({
	if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
			# reposition each time, in case the resolution or monitor changes
			$window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window.Width)
			$window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window.Height)
			$window.Show()
			$window.Activate()
	}
})

# Close the window if it's double clicked
$window.Add_MouseDoubleClick({
	$window.Hide()
})

# Close the window if it loses focus
$window.Add_Deactivated({
	start-sleep -Milliseconds 500
    $window.Hide()
})

# When Exit is clicked, close everything and kill the PowerShell process
$menuitem.add_Click({
	$notifyicon.Visible = $false
	$window.Close()
	Stop-Process $pid
 })

 

# Make PowerShell Disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Force garbage collection just to start slightly lower RAM usage.
[System.GC]::Collect()

# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)

}