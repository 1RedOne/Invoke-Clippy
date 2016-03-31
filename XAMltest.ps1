    #ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Name="window" WindowStyle="None" Height="657" Width="525" Background="Transparent" BorderBrush="Transparent" 
ResizeMode="NoResize" ShowInTaskbar="False">
    <Grid Name="grid"  Height="200" Width="400" Background="Transparent" >
       <ListView Name="listview" SelectionMode="Single" Margin="0,50,0,0" Foreground="White"
		Background="Transparent" BorderBrush="Transparent" IsHitTestVisible="False">
            <ListView.ItemContainerStyle>
                <Style>
                    <Setter Property="Control.HorizontalContentAlignment" Value="Stretch"/>
                    <Setter Property="Control.VerticalContentAlignment" Value="Stretch"/>
                </Style>
            </ListView.ItemContainerStyle>
            <Image x:Name="image" Height="140" Width="156" Source="C:\git\Invoke-Clippy\Clippy.png"/>
        </ListView>
        <Grid x:Name="gr3id" Margin="150,-190,29,190" Background="Transparent">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="40"/>
            </Grid.RowDefinitions>
            <Rectangle Fill="#FFFFFDCF" Stroke="#FF000000" RadiusX="10" RadiusY="10"/>
            <Path Fill="#FFFFFDCF" Stretch="Fill" Stroke="#FF000000" HorizontalAlignment="Left" Margin="30,-1.6,0,0" Width="25" Grid.Row="1" 
        Data="M22.166642,154.45381 L29.999666,187.66699 40.791059,154.54395"/>
            <TextBlock HorizontalAlignment="Center" Name ="ClippyText" VerticalAlignment="Center" FontSize="16" Text="Hi! I am Clippy, your office assitant.  Would you like some assistance today?" TextWrapping="Wrap" FontFamily="Arial"/>
        </Grid>

    </Grid>
    

</Window>
"@ 
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch [System.Management.Automation.MethodInvocationException] {
    Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
    write-host $error[0].Exception.Message -ForegroundColor Red
    if ($error[0].Exception.Message -like "*button*"){
        write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"}
}
catch{#if it broke some other way <span class="wp-smiley wp-emoji wp-emoji-bigsmile" title=":D">:D</span>
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
        }
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
    # Use this space to add code to the various form elements in your GUI
    #===========================================================================
                                                                    
     
    #Reference 
 
    #Adding items to a dropdown/combo box
      #$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
     
    #Setting the text of a text box to the current PC name    
      #$WPFtextBox.Text = $env:COMPUTERNAME
     
    #Adding code to a button, so that when clicked, it pings a system
    # $WPFbutton.Add_Click({ Test-connection -count 1 -ComputerName $WPFtextBox.Text
    # })
    #===========================================================================
    # Shows the form
    #===========================================================================
write-host "To show the form, run the following" -ForegroundColor Cyan
'$Form.ShowDialog() | out-null'
 
 
 