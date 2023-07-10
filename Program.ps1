Add-Type -AssemblyName System.Windows.Forms
$Icon = New-Object system.drawing.icon (".\Files\Logo.ico")
$Font = New-Object System.Drawing.Font("Times New Roman",9)

<# Update to Powershell 3.0 if neccesary

$Version=(Get-Host).version.major
if ($Version -lt 3)
{
		$Update = New-Object System.Windows.Forms.Form 
		$Update.Text = "Update Powershell"
		$Update.Size = New-Object System.Drawing.Size(290,148) 
		$Update.StartPosition = "CenterScreen"
		$Update.Topmost = $True
		$Update.Icon = $Icon
		$Update.Font = New-Object System.Drawing.Font("Times New Roman",10)
		$Update.FormBorderStyle = 'FixedDialog'
		
		$UpdateLabel = New-Object System.Windows.Forms.Label
		$UpdateLabel.Location = New-Object System.Drawing.Size(5,7) 
		$UpdateLabel.Size = New-Object System.Drawing.Size(275,40) 
		$UpdateLabel.Text = "Your version of Powershell is Version: $Version.00 
		You need an update in order to use this program."
		
		$UpdateButton = New-Object System.Windows.Forms.Button
		$UpdateButton.Location = New-Object System.Drawing.Size(60,80)
		$UpdateButton.Size = New-Object System.Drawing.Size(75,23)
		$UpdateButton.Text = "Update"
		$UpdateButton.Add_Click({$Update.Close();$Global:UpdateStatus=$true})
		
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(135,80)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$Update.Close()})
		
		$Update.Controls.Add($CancelButton) 
		$Update.Controls.Add($UpdateLabel) 
		$Update.Controls.Add($UpdateButton)
		
		$Update.Add_Shown({$Update.Activate()})
		[void] $Update.ShowDialog()
if ($Global:UpdateStatus -eq $true)
{ 
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$command = "`"" + "$executingScriptDirectory\Files\Update.msu" + "`""
$parameters = $command + " /quiet /norestart"
$install = [System.Diagnostics.Process]::Start("wusa", $command)
Exit
}
}
 #>
$TANYRHealthcare = New-Object system.Windows.Forms.Form
$TANYRHealthcare.Text = "TANYR Healthcare"
$TANYRHealthcare.TopMost = $false
$TANYRHealthcare.Width = 297
$TANYRHealthcare.Height = 280
$TANYRHealthcare.Icon = $Icon
$TANYRHealthcare.Font = $Font
$TANYRHealthcare.FormBorderStyle = 'FixedDialog'


#Write-Host $Version -foregroundcolor "Green"

##----------------------------MAIN PAGE GUI----------------------------##

$Button_Task1 = New-Object system.windows.Forms.Button
$Button_Task1.Text = "Sharepoint"
$Button_Task1.Width = 126
$Button_Task1.Height = 35
$Button_Task1.location = new-object system.drawing.point(10,16)
$Button_Task1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task1)

$Button_Task2 = New-Object system.windows.Forms.Button
$Button_Task2.Text = "Outlook"
$Button_Task2.Width = 126
$Button_Task2.Height = 35
$Button_Task2.location = new-object system.drawing.point(147,16)
$Button_Task2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task2)

$Button_Task3 = New-Object system.windows.Forms.Button
$Button_Task3.Text = ""
$Button_Task3.Width = 126
$Button_Task3.Height = 35
$Button_Task3.location = new-object system.drawing.point(10,63)
$Button_Task3.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task3)

$Button_Task4 = New-Object system.windows.Forms.Button
$Button_Task4.Text = ""
$Button_Task4.Width = 126
$Button_Task4.Height = 35
$Button_Task4.location = new-object system.drawing.point(147,63)
$Button_Task4.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task4)

$Button_Task5 = New-Object system.windows.Forms.Button
$Button_Task5.Text = ""
$Button_Task5.Width = 126
$Button_Task5.Height = 35
$Button_Task5.location = new-object system.drawing.point(10,110)
$Button_Task5.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task5)

$Button_Task6 = New-Object system.windows.Forms.Button
$Button_Task6.Text = ""
$Button_Task6.Width = 126
$Button_Task6.Height = 35
$Button_Task6.location = new-object system.drawing.point(147,110)
$Button_Task6.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task6)

$Button_Task7 = New-Object system.windows.Forms.Button
$Button_Task7.Text = ""
$Button_Task7.Width = 126
$Button_Task7.Height = 35
$Button_Task7.location = new-object system.drawing.point(10,157)
$Button_Task7.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task7)

$Button_Task8 = New-Object system.windows.Forms.Button
$Button_Task8.Text = ""
$Button_Task8.Width = 126
$Button_Task8.Height = 35
$Button_Task8.location = new-object system.drawing.point(147,157)
$Button_Task8.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task8)

$Button_Admin = New-Object system.windows.Forms.Button
$Button_Admin.Text = ""
$Button_Admin.Width = 75
$Button_Admin.Height = 12
$Button_Admin.location = new-object system.drawing.point(187,217)
$Button_Admin.Font = "Microsoft Sans Serif,10"
$TANYRHealthcare.controls.Add($Button_Admin)
$Button_Admin.TabStop = $false;
$Button_Admin.FlatStyle = "Flat";
$Button_Admin.FlatAppearance.BorderSize = 0;
$Button_Admin.FlatAppearance.MouseDownBackColor = "Transparent"
$Button_Admin.FlatAppearance.MouseOverBackColor = "Transparent"
$Button_Admin.BackColor = "Transparent"

#ComboBox - Change Page
$ChangePage = New-Object system.windows.Forms.ComboBox
if([string]::IsNullOrEmpty($ChangePage.Text))
{
$ChangePage.Items.AddRange(("Main Page","Excel Commands"));
$ChangePage.Text = "Main Page"
$ChangePage.Width = 173
$ChangePage.Height = 20
$ChangePage.location = new-object system.drawing.point(7,212)
$ChangePage.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($ChangePage)
$ChangePage.DropDownStyle= 'DropDownList'
$ChangePage.SelectionStart= 0
}

$HiddenLabel = New-Object System.Windows.Forms.Label
$HiddenLabel.Location = New-Object System.Drawing.Size(1,1) 
$HiddenLabel.Size = New-Object System.Drawing.Size(1,1) 
$HiddenLabel.Text = ""
$TANYRHealthcare.controls.Add($HiddenLabel)

##----------------------------EXCEL COMMANDS GUI----------------------------##

$Button_Task_Excel_1 = New-Object system.windows.Forms.Button
$Button_Task_Excel_1.Text = "Excel Link-Creation"
$Button_Task_Excel_1.Width = 126
$Button_Task_Excel_1.Height = 35
$Button_Task_Excel_1.location = new-object system.drawing.point(10,16)
$Button_Task_Excel_1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_1)
$Button_Task_Excel_1.Visible = $false

$Button_Task_Excel_2 = New-Object system.windows.Forms.Button
$Button_Task_Excel_2.Text = "Excel Edit"
$Button_Task_Excel_2.Width = 126
$Button_Task_Excel_2.Height = 35
$Button_Task_Excel_2.location = new-object system.drawing.point(147,16)
$Button_Task_Excel_2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_2)
$Button_Task_Excel_2.Visible = $false

$Button_Task_Excel_3 = New-Object system.windows.Forms.Button
$Button_Task_Excel_3.Text = "Browser Edit"
$Button_Task_Excel_3.Width = 126
$Button_Task_Excel_3.Height = 35
$Button_Task_Excel_3.location = new-object system.drawing.point(10,63)
$Button_Task_Excel_3.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_3)
$Button_Task_Excel_3.Visible = $false

$Button_Task_Excel_4 = New-Object system.windows.Forms.Button
$Button_Task_Excel_4.Text = ""
$Button_Task_Excel_4.Width = 126
$Button_Task_Excel_4.Height = 35
$Button_Task_Excel_4.location = new-object system.drawing.point(147,63)
$Button_Task_Excel_4.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_4)
$Button_Task_Excel_4.Visible = $false

$Button_Task_Excel_5 = New-Object system.windows.Forms.Button
$Button_Task_Excel_5.Text = ""
$Button_Task_Excel_5.Width = 126
$Button_Task_Excel_5.Height = 35
$Button_Task_Excel_5.location = new-object system.drawing.point(10,110)
$Button_Task_Excel_5.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_5)
$Button_Task_Excel_5.Visible = $false

$Button_Task_Excel_6 = New-Object system.windows.Forms.Button
$Button_Task_Excel_6.Text = ""
$Button_Task_Excel_6.Width = 126
$Button_Task_Excel_6.Height = 35
$Button_Task_Excel_6.location = new-object system.drawing.point(147,110)
$Button_Task_Excel_6.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_6)
$Button_Task_Excel_6.Visible = $false

$Button_Task_Excel_7 = New-Object system.windows.Forms.Button
$Button_Task_Excel_7.Text = "Save As Excel"
$Button_Task_Excel_7.Width = 126
$Button_Task_Excel_7.Height = 35
$Button_Task_Excel_7.location = new-object system.drawing.point(10,157)
$Button_Task_Excel_7.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_7)
$Button_Task_Excel_7.Visible = $false

$Button_Task_Excel_8 = New-Object system.windows.Forms.Button
$Button_Task_Excel_8.Text = "Test Email"
$Button_Task_Excel_8.Width = 126
$Button_Task_Excel_8.Height = 35
$Button_Task_Excel_8.location = new-object system.drawing.point(147,157)
$Button_Task_Excel_8.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_8)
$Button_Task_Excel_8.Visible = $false

##----------------------------ADMIN COMMANDS GUI----------------------------##

$Button_Task_Admin_1 = New-Object system.windows.Forms.Button
$Button_Task_Admin_1.Text = "List Users"
$Button_Task_Admin_1.Width = 126
$Button_Task_Admin_1.Height = 35
$Button_Task_Admin_1.location = new-object system.drawing.point(10,16)
$Button_Task_Admin_1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_1)
$Button_Task_Admin_1.Visible = $false

$Button_Task_Admin_2 = New-Object system.windows.Forms.Button
$Button_Task_Admin_2.Text = "List Groups"
$Button_Task_Admin_2.Width = 126
$Button_Task_Admin_2.Height = 35
$Button_Task_Admin_2.location = new-object system.drawing.point(147,16)
$Button_Task_Admin_2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_2)
$Button_Task_Admin_2.Visible = $false

$Button_Task_Admin_3 = New-Object system.windows.Forms.Button
$Button_Task_Admin_3.Text = "Usage Reports"
$Button_Task_Admin_3.Width = 126
$Button_Task_Admin_3.Height = 35
$Button_Task_Admin_3.location = new-object system.drawing.point(10,63)
$Button_Task_Admin_3.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_3)
$Button_Task_Admin_3.Visible = $false

$Button_Task_Admin_4 = New-Object system.windows.Forms.Button
$Button_Task_Admin_4.Text = "Mail Rules"
$Button_Task_Admin_4.Width = 126
$Button_Task_Admin_4.Height = 35
$Button_Task_Admin_4.location = new-object system.drawing.point(147,63)
$Button_Task_Admin_4.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_4)
$Button_Task_Admin_4.Visible = $false

$Button_Task_Admin_5 = New-Object system.windows.Forms.Button
$Button_Task_Admin_5.Text = "Azure Portal"
$Button_Task_Admin_5.Width = 126
$Button_Task_Admin_5.Height = 35
$Button_Task_Admin_5.location = new-object system.drawing.point(10,110)
$Button_Task_Admin_5.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_5)
$Button_Task_Admin_5.Visible = $false

$Button_Task_Admin_6 = New-Object system.windows.Forms.Button
$Button_Task_Admin_6.Text = "Exchange Admin Center"
$Button_Task_Admin_6.Width = 126
$Button_Task_Admin_6.Height = 35
$Button_Task_Admin_6.location = new-object system.drawing.point(147,110)
$Button_Task_Admin_6.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_6)
$Button_Task_Admin_6.Visible = $false

$Button_Task_Admin_7 = New-Object system.windows.Forms.Button
$Button_Task_Admin_7.Text = "Security Admin Center"
$Button_Task_Admin_7.Width = 126
$Button_Task_Admin_7.Height = 35
$Button_Task_Admin_7.location = new-object system.drawing.point(10,157)
$Button_Task_Admin_7.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_7)
$Button_Task_Admin_7.Visible = $false

$Button_Task_Admin_8 = New-Object system.windows.Forms.Button
$Button_Task_Admin_8.Text = "Sharepoint Admin Center"
$Button_Task_Admin_8.Width = 126
$Button_Task_Admin_8.Height = 35
$Button_Task_Admin_8.location = new-object system.drawing.point(147,157)
$Button_Task_Admin_8.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_8)
$Button_Task_Admin_8.Visible = $false

##----------------------------MAIN PAGE COMMANDS----------------------------##

    #SharePoint Home Button
    $Button_Task1.Add_Click(
        {    
		Start-Process -FilePath "https://tanyrhealthcare.sharepoint.com/_layouts/15/sharepoint.aspx"
        }
    )
	
    #Outlook Home Button
    $Button_Task2.Add_Click(
        {    
		Start-Process -FilePath "https://outlook.office365.com/owa/?realm=tanyrhealthcare.com&exsvurl=1&ll-cc=1033&modurl=0"
        }
    )

    #Tbd
    $Button_Task3.Add_Click({})
	
	#Tbd
    $Button_Task4.Add_Click({})
	
	#Tbd
    $Button_Task5.Add_Click({})
	
	#Tbd
    $Button_Task6.Add_Click({})

	#Tbd
    $Button_Task7.Add_Click({})
	
	#Tbd
    $Button_Task8.Add_Click({})		
	
	#Admin Button
    $Button_Admin.Add_Click(
        {    
		$Index = $ChangePage.Items.Count
		if ($Index -eq 2){
		$ChangePage.Items.Add(("Admin Commands"));
		}
        }
    )

	#Change Page
	$ChangePage.Add_SelectedIndexChanged({
		if ($ChangePage.SelectedIndex -eq 0)
		{
		#Main Page Visible
		$Button_Task1.Visible = $true
		$Button_Task2.Visible = $true
		$Button_Task3.Visible = $true
		$Button_Task4.Visible = $true
		$Button_Task5.Visible = $true
		$Button_Task6.Visible = $true
		$Button_Task7.Visible = $true
		$Button_Task8.Visible = $true
		
		#Excel Commands Invisible
		$Button_Task_Excel_1.Visible = $false
		$Button_Task_Excel_2.Visible = $false
		$Button_Task_Excel_3.Visible = $false
		$Button_Task_Excel_4.Visible = $false
		$Button_Task_Excel_5.Visible = $false
		$Button_Task_Excel_6.Visible = $false
		$Button_Task_Excel_7.Visible = $false
		$Button_Task_Excel_8.Visible = $false
		
		#Admin Commands Invisible
		$Button_Task_Admin_1.Visible = $false
		$Button_Task_Admin_2.Visible = $false
		$Button_Task_Admin_3.Visible = $false
		$Button_Task_Admin_4.Visible = $false
		$Button_Task_Admin_5.Visible = $false
		$Button_Task_Admin_6.Visible = $false
		$Button_Task_Admin_7.Visible = $false
		$Button_Task_Admin_8.Visible = $false
		
		$HiddenLabel.Focus()
		}
})
	$ChangePage.Add_SelectedIndexChanged({
		if ($ChangePage.SelectedIndex -eq 1)
		{
		#Main Page Invisible
		$Button_Task1.Visible = $false
		$Button_Task2.Visible = $false
		$Button_Task3.Visible = $false
		$Button_Task4.Visible = $false
		$Button_Task5.Visible = $false
		$Button_Task6.Visible = $false
		$Button_Task7.Visible = $false
		$Button_Task8.Visible = $false
		
		#Excel Commands Visible
		$Button_Task_Excel_1.Visible = $true
		$Button_Task_Excel_2.Visible = $true
		$Button_Task_Excel_3.Visible = $true
		$Button_Task_Excel_4.Visible = $true
		$Button_Task_Excel_5.Visible = $true
		$Button_Task_Excel_6.Visible = $true
		$Button_Task_Excel_7.Visible = $true
		$Button_Task_Excel_8.Visible = $true
		
		#Admin Commands Invisible
		$Button_Task_Admin_1.Visible = $false
		$Button_Task_Admin_2.Visible = $false
		$Button_Task_Admin_3.Visible = $false
		$Button_Task_Admin_4.Visible = $false
		$Button_Task_Admin_5.Visible = $false
		$Button_Task_Admin_6.Visible = $false
		$Button_Task_Admin_7.Visible = $false
		$Button_Task_Admin_8.Visible = $false
		
		$HiddenLabel.Focus()
		}
})
	$ChangePage.Add_SelectedIndexChanged({
		if ($ChangePage.SelectedIndex -eq 2)
		{
		#Main Page Invisible
		$Button_Task1.Visible = $false
		$Button_Task2.Visible = $false
		$Button_Task3.Visible = $false
		$Button_Task4.Visible = $false
		$Button_Task5.Visible = $false
		$Button_Task6.Visible = $false
		$Button_Task7.Visible = $false
		$Button_Task8.Visible = $false
		
		#Excel Commands Invisible
		$Button_Task_Excel_1.Visible = $false
		$Button_Task_Excel_2.Visible = $false
		$Button_Task_Excel_3.Visible = $false
		$Button_Task_Excel_4.Visible = $false
		$Button_Task_Excel_5.Visible = $false
		$Button_Task_Excel_6.Visible = $false
		$Button_Task_Excel_7.Visible = $false
		$Button_Task_Excel_8.Visible = $false
		
		#Admin Commands Visible
		$Button_Task_Admin_1.Visible = $true
		$Button_Task_Admin_2.Visible = $true
		$Button_Task_Admin_3.Visible = $true
		$Button_Task_Admin_4.Visible = $true
		$Button_Task_Admin_5.Visible = $true
		$Button_Task_Admin_6.Visible = $true
		$Button_Task_Admin_7.Visible = $true
		$Button_Task_Admin_8.Visible = $true
		
		$HiddenLabel.Focus()
		}
})

##----------------------------EXCEL COMMANDS----------------------------##
    #Choose File
    $Button_Task_Excel_1.Add_Click(
        {   
		$FileLocator = New-Object System.Windows.Forms.Form 
		$FileLocator.Text = "Data Entry Form"
		$FileLocator.Size = New-Object System.Drawing.Size(325,190) 
		$FileLocator.StartPosition = "CenterScreen"
		$FileLocator.Topmost = $True
		$FileLocator.Font = $Font
		$FileLocator.FormBorderStyle = 'FixedDialog'
		
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,13) 
		$objLabel.Size = New-Object System.Drawing.Size(305,25) 
		$objLabel.Text = "Please enter the URL of the document in the space below:"
		
		$objTextBox = New-Object System.Windows.Forms.TextBox 
		$objTextBox.Location = New-Object System.Drawing.Size(10,30) 
		$objTextBox.Size = New-Object System.Drawing.Size(285,20) 
		
		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(40,120)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"
		$FileLocator.AcceptButton = $OKButton
		$OKButton.Add_Click({$global:FileAdded=$true;$global:FilePath=$objTextBox.Text;$FileLocator.Close()})
		
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(115,120)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$FileLocator.Close()})
		
		$HelpButton = New-Object System.Windows.Forms.Button
		$HelpButton.Location = New-Object System.Drawing.Size(190,120)
		$HelpButton.Size = New-Object System.Drawing.Size(75,23)
		$HelpButton.Text = "Need Help?"
		$HelpButton.Add_Click({		
		$HelpCopyPaste = New-Object System.Windows.Forms.Form 
		$HelpCopyPaste.Size = New-Object System.Drawing.Size(770,700) 
		$HelpCopyPaste.StartPosition = "CenterScreen"
		$HelpCopyPaste.Topmost = $True
		$HelpCopyPaste.Visible =$True
		$HelpCopyPaste.FormBorderStyle = 'FixedDialog'
		
		$HelpPicture1 = New-Object system.windows.Forms.PictureBox
		$Image1 = [System.Drawing.Image]::Fromfile((get-item .\Files\Help1.png));
		$HelpPicture1.Image = $Image1
		$HelpPicture1.location = new-object system.drawing.point(10,16)
		$HelpPicture1.Width = $Image1.Width
		$HelpPicture1.Height = $Image1.Height
		
		$HelpPicture2 = New-Object system.windows.Forms.PictureBox
		$Image2 = [System.Drawing.Image]::Fromfile((get-item .\Files\Help2.png));
		$HelpPicture2.Image = $Image2
		$HelpPicture2.location = new-object system.drawing.point(10,335)
		$HelpPicture2.Width = $Image2.Width
		$HelpPicture2.Height = $Image2.Height
		
		$HelpCopyPaste.ShowIcon = $false
		$HelpCopyPaste.controls.Add($HelpPicture1)
		$HelpCopyPaste.controls.Add($HelpPicture2)
		})
		
		$FileLocator.Icon = $Icon
		$FileLocator.Controls.Add($CancelButton)
		$FileLocator.Controls.Add($objTextBox) 
		$FileLocator.Controls.Add($objLabel) 
		$FileLocator.Controls.Add($OKButton)
		$FileLocator.Controls.Add($HelpButton)
		
		
		$FileLocator.Add_Shown({$FileLocator.Activate();$objTextBox.focus()})
		[void] $FileLocator.ShowDialog()
		
			if ($global:FileAdded -eq $true)
	{
		$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $false
	[String]$PageURL = $global:FilePath
	$ie.navigate($PageURL)
	While ($ie.Busy)
	{
    Sleep -Milliseconds 100
}
	sleep -milliseconds 2000
	$Document = $ie.document
	sleep -milliseconds 2000
	$HTML=Invoke-Webrequest -Uri $PageURL
	$data = $HTML.AllElements | Where {$_.class -eq "BreadcrumbItem"}
	#$Name=$Document.documentElement.getElementsByClassName('BreadcrumbItem')[1].href
	Write-Host $data.href -Separator `n -foregroundcolor "Green"
	
	
	
<# 	$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $true
	[String]$PageURL = $global:FilePath
	$ie.navigate($PageURL)
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
sleep -milliseconds 2000
$Document = $ie.document
sleep -milliseconds 5000 #>

	}
		
		}
    )

    #Edit Excel Program
    $Button_Task_Excel_2.Add_Click({
	$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
	$excel.Cells.Item(1,1)="Help"
	$excel.Cells.Item(1,2)="Me"
	$excel.Cells.Item(1,7).Select()
	
  }
)

#Browser Edit
$Button_Task_Excel_3.Add_Click({
	$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $true
	$PageURL = 'https://www.google.com/'
	$ie.navigate($PageURL)
	
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
Sleep -Milliseconds 2000

#OPEN GOOGLE DON'T FORGET TO PAUSE
$Document = $ie.document
#$CommandsICanUseHere = $ie | Get-Member -MemberType method | Out-GridView

$SearchBar=$Document.getElementByID('lst-ib')
$WhatToGoogle ="How do I code in HTML"
$SearchBar.value= $WhatToGoogle
$GoogleFormat = $WhatToGoogle -replace " ","+"
$SearchURL="https://www.google.com/#q=$GoogleFormat"

$ie.navigate($SearchURL)

	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
Sleep -Milliseconds 2000

#$Document = $ie.document
$Search=$Document.getElementByID('_fZl')
$Search.click()

Sleep -Milliseconds 1000
$NodeList=$ie.document.getElementsByTagName("h3") | Where-Object -FilterScript { ($_.className -eq "r") }

Write-Host $NodeList.length -Separator `n -foregroundcolor "Green"
for($i=0; $i -lt $NodeList.Length; $i++)
{
Write-Host $NodeList[$i].innertext -Separator `n -foregroundcolor "Green"
}
	}
)

#Save as Excel
$Button_Task_Excel_7.Add_Click({
	$ie = New-Object -ComObject 'InternetExplorer.Application'
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $false

$WhatToGoogle ="How do I code in HTML"
$GoogleFormat = $WhatToGoogle -replace " ","+"
$SearchURL="https://www.google.com/#q=$GoogleFormat"
$ie.navigate($SearchURL)
	While ($ie.Busy)
{
    Sleep -Milliseconds 100
}
$Document=$ie.document
Sleep -Milliseconds 2000

$NodeList=$Document.getElementsByTagName("h3") | Where-Object -FilterScript { ($_.className -eq "r") }

$Export = @("")
$Object = New-Object psobject
for($i=0;$i -lt $NodeList.Length; $i++)
{
$Export=($NodeList[$i].innerHTML -replace '".+?" |{"' -replace '">.+?a>' -replace '<a.+?="') #-replace '<.+?>|{"' #-replace '":"?','=' -replace '"' -join "`n"
#Write-Host $Export -Separator `n -foregroundcolor "Green"
$Object | add-member -memberType NoteProperty -name "$i" -value "$Export"
}
$Object | ConvertTo-Csv -NoTypeInformation -Delimiter "`n" | select -Skip 1 | Set-Content .\Lists.csv
}
)

#Send Email
$Button_Task_Excel_8.Add_Click({

$Credential  = Get-Credential 'connor.kaiser@tanyrhealthcare.com'

$param = @{
    SmtpServer = 'smtp.office365.com'
    Port = 587
    UseSsl = $true
    Credential  = $Credential
    From = 'connor.kaiser@tanyrhealthcare.com'
    To = 'connor.kaiser@tanyrhealthcare.com'
    Subject = 'Data found'
    Body = 'Data Here'
    #Attachments = 'C:\attachment.csv'
}
Send-MailMessage @param

	}
)



##----------------------------ADMIN COMMANDS----------------------------##
    #List Users
    $Button_Task_Admin_1.Add_Click(
        {    
		Start-Process -FilePath "https://portal.office.com/adminportal/home#/users"
        }
    )

    #List Groups
    $Button_Task_Admin_2.Add_Click(
        {    
		Start-Process -FilePath "https://portal.office.com/adminportal/home#/groups"
        }
    )
	
    #Usage Reports
    $Button_Task_Admin_3.Add_Click(
        {    
		Start-Process -FilePath "https://portal.office.com/adminportal/home#/homepage"
        }
    )
	
	#Mail Rules
    $Button_Task_Admin_4.Add_Click(
        {    
		Start-Process -FilePath "https://outlook.office365.com/ecp/?p=TransportRules&rfr=admin_o365&exsvurl=1&Realm=tanyrhealthcare.com&RpsCsrfState=4e8f3243-ce6b-07e5-e9ba-eafb758fa906&wa=wsignin1.0"
        }
    )
	
    #Azure Portal
    $Button_Task_Admin_5.Add_Click(
        {    
		Start-Process -FilePath "https://portal.azure.com/#dashboard/private/65b245f2-d671-48f2-b65d-bc3a26fe200f"
        }
    )
	
	#Exchange Admin
    $Button_Task_Admin_6.Add_Click(
        {    
		Start-Process -FilePath "https://outlook.office365.com/ecp/?exsvurl=1&rfr=admin_o365&Realm=tanyrhealthcare.com"
        }
    )
	   
	#Security & Compliance Admin
    $Button_Task_Admin_7.Add_Click(
        {    
		Start-Process -FilePath "https://protection.office.com/#/homepage"
        }
    )

	#Sharepoint Admin Center
	$Button_Task_Admin_8.Add_Click(
        {    
		Start-Process -FilePath "https://tanyrhealthcare-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx"
        }
    )
[void]$TANYRHealthcare.ShowDialog()
$TANYRHealthcare.Dispose()