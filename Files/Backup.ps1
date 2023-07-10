Add-Type -AssemblyName System.Windows.Forms
$Version=(Get-Host).version.major
$Icon = New-Object system.drawing.icon (".\Files\Logo.ico")
$Font = New-Object System.Drawing.Font("Times New Roman",9)
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
$command = "`"" + "$executingScriptDirectory\Files\Update1.msu" + "`""
$parameters = $command + " /quiet /norestart"
$install = [System.Diagnostics.Process]::Start("wusa", $command)
Exit
}
}

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
#Button - Task 1
$Button_Task1 = New-Object system.windows.Forms.Button
$Button_Task1.Text = "Sharepoint"
$Button_Task1.Width = 126
$Button_Task1.Height = 35
$Button_Task1.location = new-object system.drawing.point(10,16)
$Button_Task1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task1)
$TANYRHealthcare.AcceptButton = $Button_Task1

#Button - Task 2
$Button_Task2 = New-Object system.windows.Forms.Button
$Button_Task2.Text = "Outlook"
$Button_Task2.Width = 126
$Button_Task2.Height = 35
$Button_Task2.location = new-object system.drawing.point(147,16)
$Button_Task2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task2)
$TANYRHealthcare.AcceptButton = $Button_Task2

#Button - Admin
$Button_Admin = New-Object system.windows.Forms.Button
$Button_Admin.Text = ""
$Button_Admin.Width = 75
$Button_Admin.Height = 12
$Button_Admin.location = new-object system.drawing.point(187,217)
$Button_Admin.Font = "Microsoft Sans Serif,10"
$TANYRHealthcare.controls.Add($Button_Admin)
$TANYRHealthcare.AcceptButton = $Button_Admin
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
}


##----------------------------EXCEL COMMANDS GUI----------------------------##
#Button - Task 1
$Button_Task_Excel_1 = New-Object system.windows.Forms.Button
$Button_Task_Excel_1.Text = "Choose Document"
$Button_Task_Excel_1.Width = 126
$Button_Task_Excel_1.Height = 35
$Button_Task_Excel_1.location = new-object system.drawing.point(10,16)
$Button_Task_Excel_1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_1)
$TANYRHealthcare.AcceptButton = $Button_Task_Excel_1
$Button_Task_Excel_1.Visible = $false

#Button - Task 2
$Button_Task_Excel_2 = New-Object system.windows.Forms.Button
$Button_Task_Excel_2.Text = "Updated?"
$Button_Task_Excel_2.Width = 126
$Button_Task_Excel_2.Height = 35
$Button_Task_Excel_2.location = new-object system.drawing.point(147,16)
$Button_Task_Excel_2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Excel_2)
$TANYRHealthcare.AcceptButton = $Button_Task_Excel_2
$Button_Task_Excel_2.Visible = $false

##----------------------------ADMIN COMMANDS GUI----------------------------##
#Button - Task 1
$Button_Task_Admin_1 = New-Object system.windows.Forms.Button
$Button_Task_Admin_1.Text = "List Users"
$Button_Task_Admin_1.Width = 126
$Button_Task_Admin_1.Height = 35
$Button_Task_Admin_1.location = new-object system.drawing.point(10,16)
$Button_Task_Admin_1.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_1)
$TANYRHealthcare.AcceptButton = $Button_Task_Admin_1
$Button_Task_Admin_1.Visible = $false


#Button - Task 2
$Button_Task_Admin_2 = New-Object system.windows.Forms.Button
$Button_Task_Admin_2.Text = "List Groups"
$Button_Task_Admin_2.Width = 126
$Button_Task_Admin_2.Height = 35
$Button_Task_Admin_2.location = new-object system.drawing.point(147,16)
$Button_Task_Admin_2.Font = "Arial,10,style=Bold"
$TANYRHealthcare.controls.Add($Button_Task_Admin_2)
$TANYRHealthcare.AcceptButton = $Button_Task_Admin_2
$Button_Task_Admin_2.Visible = $false

$Button_Task_Admin_3 = New-Object system.windows.Forms.Button
$Button_Task_Admin_3.Text = "Azure Portal"
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
$Button_Task_Admin_5.Text = ""
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
$Button_Task_Admin_7.Text = ""
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
	$ChangePage.Add_TextChanged({
		if ($ChangePage.Text -eq "Main Page")
		{
		#Main Page Visible
		$Button_Task1.Visible = $true
		$Button_Task2.Visible = $true
		#Excel Commands Invisible
		$Button_Task_Excel_1.Visible = $false
		$Button_Task_Excel_2.Visible = $false
		
		#Admin Commands Invisible
		$Button_Task_Admin_1.Visible = $false
		$Button_Task_Admin_2.Visible = $false
		$Button_Task_Admin_3.Visible = $false
		$Button_Task_Admin_4.Visible = $false
		$Button_Task_Admin_5.Visible = $false
		$Button_Task_Admin_6.Visible = $false
		$Button_Task_Admin_7.Visible = $false
		$Button_Task_Admin_8.Visible = $false
		}
})
	$ChangePage.Add_TextChanged({
		if ($ChangePage.Text -eq "Excel Commands")
		{
		#Main Page Invisible
		$Button_Task1.Visible = $false
		$Button_Task2.Visible = $false
		
		#Excel Commands Visible
		$Button_Task_Excel_1.Visible = $true
		$Button_Task_Excel_2.Visible = $true
		
		#Admin Commands Invisible
		$Button_Task_Admin_1.Visible = $false
		$Button_Task_Admin_2.Visible = $false
		$Button_Task_Admin_3.Visible = $false
		$Button_Task_Admin_4.Visible = $false
		$Button_Task_Admin_5.Visible = $false
		$Button_Task_Admin_6.Visible = $false
		$Button_Task_Admin_7.Visible = $false
		$Button_Task_Admin_8.Visible = $false
		}
})
	$ChangePage.Add_TextChanged({
		if ($ChangePage.Text -eq "Admin Commands")
		{
		#Main Page Invisible
		$Button_Task1.Visible = $false
		$Button_Task2.Visible = $false
		
		#Excel Commands Invisible
		$Button_Task_Excel_1.Visible = $false
		$Button_Task_Excel_2.Visible = $false
		
		#Admin Commands Visible
		$Button_Task_Admin_1.Visible = $true
		$Button_Task_Admin_2.Visible = $true
		$Button_Task_Admin_3.Visible = $true
		$Button_Task_Admin_4.Visible = $true
		$Button_Task_Admin_5.Visible = $true
		$Button_Task_Admin_6.Visible = $true
		$Button_Task_Admin_7.Visible = $true
		$Button_Task_Admin_8.Visible = $true
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
		
		$FileLocator.Add_Shown({$FileLocator.Activate()})
		[void] $FileLocator.ShowDialog()
		
	if ($global:FileAdded -eq $true)
	{
	$ie = New-Object -ComObject InternetExplorer.Application
	$ie.AddressBar = $false
	$ie.MenuBar = $false
	$ie.StatusBar = $false
	$ie.ToolBar = $false
	$ie.Visible = $true
	[String]$PageURL = $global:FilePath
	$ie.navigate($PageURL)
	While ($ie.LocationURL -notlike "https://tanyrhealthcare.sharepoint.com/*")
{
    Sleep -Milliseconds 100
}
$ButtonName="ewa-bb-entry-button" 
$Document = $ie.Document
$ButtonID= $document.getElementByID($ButtonName)
$ButtonID.Click()
	}
		}
    )

	
	

    #Updated?
    $Button_Task_Excel_2.Add_Click(
        {
	Add-PSSnapin Microsoft.Sharepoint.Powershell		
	if (Get-Command Get-SPWeb -errorAction SilentlyContinue)
	{
	}
	else
	{
		$UpdateMANAGEMENT  = New-Object System.Windows.Forms.Form 
		$UpdateMANAGEMENT.Text = "Data Entry Form"
		$UpdateMANAGEMENT.Size = New-Object System.Drawing.Size(300,200) 
		$UpdateMANAGEMENT.StartPosition = "CenterScreen"
		$UpdateMANAGEMENT.Topmost = $True
		
		$MANAGEMENTUpdateLabel = New-Object System.Windows.Forms.Label
		$MANAGEMENTUpdateLabel.Location = New-Object System.Drawing.Size(10,30) 
		$MANAGEMENTUpdateLabel.Size = New-Object System.Drawing.Size(280,40) 
		$MANAGEMENTUpdateLabel.Text = "It appears you do not have the files needed to auotmate excel. Would you like to update these files?"
		
		$MANAGEMENTUpdateButton = New-Object System.Windows.Forms.Button
		$MANAGEMENTUpdateButton.Location = New-Object System.Drawing.Size(75,120)
		$MANAGEMENTUpdateButton.Size = New-Object System.Drawing.Size(75,23)
		$MANAGEMENTUpdateButton.Text = "Update"
		$MANAGEMENTUpdateButton.Add_Click({$global:Help="true"; $UpdateMANAGEMENT.Close()})
		
		$MANAGEMENTCancelButton = New-Object System.Windows.Forms.Button
		$MANAGEMENTCancelButton.Location = New-Object System.Drawing.Size(150,120)
		$MANAGEMENTCancelButton.Size = New-Object System.Drawing.Size(75,23)
		$MANAGEMENTCancelButton.Text = "Cancel"
		$MANAGEMENTCancelButton.Add_Click({$UpdateMANAGEMENT.Close()})
		
		$UpdateMANAGEMENT.Controls.Add($MANAGEMENTCancelButton) 
		$UpdateMANAGEMENT.Controls.Add($MANAGEMENTUpdateLabel) 
		$UpdateMANAGEMENT.Controls.Add($MANAGEMENTUpdateButton)
		
		$UpdateMANAGEMENT.Add_Shown({$UpdateMANAGEMENT.Activate()})
		[void] $UpdateMANAGEMENT.ShowDialog()
		
	if ($global:Help -eq $true)
{ 
		[System.Windows.Forms.Application]::Exit($null)
		Start-process '.\Files\Module.msi'
}
	}
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
	
    #Azure Portal
    $Button_Task_Admin_3.Add_Click(
        {    
		Start-Process -FilePath "https://portal.azure.com/#dashboard/private/65b245f2-d671-48f2-b65d-bc3a26fe200f"
        }
    )
	
	#Mail Rules
    $Button_Task_Admin_4.Add_Click(
        {    
		Start-Process -FilePath "https://outlook.office365.com/ecp/?p=TransportRules&rfr=admin_o365&exsvurl=1&Realm=tanyrhealthcare.com&RpsCsrfState=4e8f3243-ce6b-07e5-e9ba-eafb758fa906&wa=wsignin1.0"
        }
    )
	
    #
    $Button_Task_Admin_5.Add_Click(
        {    
		#Start-Process -FilePath ""
        }
    )
	
	#Exchange Admin
    $Button_Task_Admin_6.Add_Click(
        {    
		Start-Process -FilePath "https://outlook.office365.com/ecp/?exsvurl=1&rfr=admin_o365&Realm=tanyrhealthcare.com"
        }
    )
	   
	#
    $Button_Task_Admin_7.Add_Click(
        {    
		#Start-Process -FilePath ""
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