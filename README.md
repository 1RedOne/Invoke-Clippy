# Invoke-Clippy

######Excerpted from [PowerClippy on FoxDeploy.com](https://foxdeploy.wordpress.com/?p=2646&preview=true)

Back with the release of Office ’97 in November of 1996, Microsoft introduced a beloved new helper tool to assist users in navigating through the daunting changes in Microsoft Office,

Microsoft Clippy!


![Clippy](https://foxdeploy.files.wordpress.com/2016/03/clippy.png?w=584&h=542)


For too many years, Clippy has gone missing in Windows, and admins were left with horrible solutions, like sending e-mails or talking to their users face-to-face.

But now he’s back!

I am proud to share with you my newest slap-dash project,  ‘Invoke-Clippy!‘

###Installation

Install the module by downloading from here, or using the PowerShell Gallery

######Manual steps

* Copy the "PowerClippy" folder into your module path. Note: You can find an
appropriate directory by running `$ENV:PSModulePath.Split(';')`.
 * Run `Import-Module PowerCippy` from your PowerShell command prompt.

######Using the PowerShell Gallery (WMF 5 and up!)

>Find-Module -Name PowerClippy | Install-Module


###Syntax

Syntax is simple!

![](https://foxdeploy.files.wordpress.com/2016/03/clippy01.png?w=1272&h=114)

![](https://foxdeploy.files.wordpress.com/2016/03/clippy02.png?w=1272&h=114)

>Invoke-Clippy -text 'Would you like to install Windows 10?' -Button1Text Yes -Button2Text 'Restart PC'
   
>   >Provides a helpful Clippy UI to notify users of the wonders of Windows 10, in case they were unaware.   If the user isn't interested, shuts down their PC.

##Clippy, how I've missed you!

Today, it pretty much just pops up the familiar character.  Being a very lazy retreading of Chrissy’s code from her ‘Hey Scripting Guy’ article, it also features the PowerShell logo in the system tray to end the code!  AND it runs hidden!

You have the option of specifying -Button1 or -Button2 to add additional buttons.  If you’d like the buttons to do anything, add some code for them to the empty script blocks on line 71 and line 80.

Consider this a framework to use to ~~annoy~~ notify your coworkers with helpful reminders.

##Suggestions

Scheduled Task on your coworkers machine every 15 minutes to remind them to check the ticket queue
Add two buttons, and make the second button spawn another instance of Clippy (consider reversing the X,Y values to make Clippy appear on the other side)
Use this as a nice and professional way to communicate mandatory reboots to your end users
No matter what you come up with, share it with the class!  Did you find a way to make this appear interactively on a remote session?  Did you add -ComputerName support (If you did, AWESOME!).

Either comment here or make your own fork and send me a Pull Request.  I’d love to see what you come up with.

##References

Pretty much everything here I learned on the spot thanks to Stack Overflow.  Also big big thanks to Chrissy Lemaire in her excellent Scripting Guys article, ‘How to Create Popups’ in PowerShell.  Most of the code for window sizing comes from her work!

How do I apply a style to all buttons
How to make rounded corners on a button Corner Rounded Buttons in WPF
Creating a custom template in WPF
