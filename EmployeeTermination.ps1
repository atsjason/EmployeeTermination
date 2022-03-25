Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$global:finalvalue = @()
$global:output = @()


Function OpenWordDoc($documentPath){
    Write-host $documentpath
    try{
       $documents = $documentPath.ToString()
        $word = New-Object -ComObject Word.application
        $document = $word.Documents.Open($documents,$null,$true)
        $array = New-Object System.Collections.Generic.List[System.Object]
        $document.shapes |Where-Object {$_. type -eq 17} | ForEach-Object{
        $array.add($_.TextFrame.TextRange.text)}
        $username = $array[0]
        write-host $username
        if(($username -eq $null) -or ($username -eq [String]::Empty)){
            return $null
        }
        else{
            $username = $username.Trim()
            if($username -eq [String]::Empty){
                return $null
            }
        $array.clear()
        $document.Close($Word.WdSaveOptions.wdDoNotSaveChanges)
        $word.Quit()
        return ("*" + $username + "*")
        }

    } catch [System.Runtime.InteropServices.COMException] {

        Write-host "Sorry, we couldn't find your file. Was it moved, renamed, or deleted?"
        return $null
      }

}


Function isExtension($document){
$extn = [IO.Path]::GetExtension($document)
    if ($extn -eq ".docx" ) #will add future acceptable file extenions. have only tested docx for the meantime
    {
       return $true
    }
    return $false
}


Function DoesUserExist($username){
     Write-host "in method" $username
     if($username){
         $users = get-azureaduser -all $true | Where-Object{$_.DisplayName -like $username}
         return $users
     }
     return $null
}

Function ExportToArray($user){
    $item = New-Object PSObject
    $item | Add-Member -type NoteProperty -Name 'Email' -Value $user.UserPrincipalName
    $item | Add-Member -type NoteProperty -Name 'Display Name' -Value $user.DisplayName
    return $item
}

Function terminate($user){
    Write-host "In Terminate"
    #forceUserSignout($user)
    #blockUserAccount($user)
    #removeUserFromGroups($user)
    #getOneDriveLink($username) #Can be commented out if user using script has SharePoint access
    #if(ifSharedMailbox($user)){setSharedMailbox($user)}
    $global:output += ExportToArray($user)
    Write-host $global:output "printing"
}

Function forceUserSignOut($username){
    Revoke-AzureADUserAllRefreshToken -objectid $username.objectid
}

Function blockUserAccount($username){
    Set-AzureADUser -ObjectID $username.ObjectID -AccountEnabled $false
}

Function removeUserFromGroups($username){
    $Membership = Get-AzureADUserMembership -ObjectId $username.ObjectId
    if($Membership){    
        ForEach($group in $Membership){
            try{
             Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $username.ObjectId
            } catch [Microsoft.Open.AzureAD16.Client.ApiException] {
                Remove-DistributionGroupMember -Identity $group.displayname -Member $username.ObjectId -Confirm:$False
            }
        }
    }
}

Function ifSharedMailbox($username){
    $isShared = (Get-Mailbox -Identity $username.UserPrincipalName).RecipientTypeDetails
    return $isShared -eq "UserMailBox"
}

Function setSharedMailbox($username){
    Set-Mailbox $username.UserPrincipalName -Type Shared
}

Function Dropbox(){
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell GUI"
$form.Size = '260,320'
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size
$form.MaximizeBox = $False
$form.Topmost = $True

$button = New-Object System.Windows.Forms.Button
$button.Location = '5,5'
$button.Size = '75,23'
$button.Width = 120
$button.Text = "Print list to console"
 
$checkbox = New-Object Windows.Forms.Checkbox
$checkbox.Location = '140,8'
$checkbox.AutoSize = $True
$checkbox.Text = "Clear afterwards"
 
$label = New-Object Windows.Forms.Label
$label.Location = '5,40'
$label.AutoSize = $True
$label.Text = "Drop files or folders here:"
 
$listBox = New-Object Windows.Forms.ListBox
$listBox.Location = '5,60'
$listBox.Height = 200
$listBox.Width = 240
$listBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$listBox.IntegralHeight = $False
$listBox.AllowDrop = $True
 
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = "Ready"
 
$form.SuspendLayout()
$form.Controls.Add($button)
$form.Controls.Add($checkbox)
$form.Controls.Add($label)
$form.Controls.Add($listBox)
$form.Controls.Add($statusBar)
$form.ResumeLayout()

$button_Click = {
    write-host "Listbox contains:" -ForegroundColor Yellow
    #$documentPath = @()
	foreach ($item in $listBox.Items)
    {
        $global:finalvalue += $item
	}
    if($checkbox.Checked -eq $True)
    {
        $listBox.Items.Clear()
    }
    
    $statusBar.Text = ("List contains $($listBox.Items.Count) items")
    $form.Close()
}

$listBox_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}
	
$listBox_DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
    {
                #Write-host $i
                #$filenames += $filename.toString()
        $i = Get-Item -LiteralPath $filename
		$listBox.Items.Add($i)
                #$filenames += $filename.toString()
	}
    $statusBar.Text = ("List contains $($listBox.Items.Count) items")
}
 
$form_FormClosed = {
	try
    {
        $listBox.remove_Click($button_Click)
		$listBox.remove_DragOver($listBox_DragOver)
		$listBox.remove_DragDrop($listBox_DragDrop)
        $listBox.remove_DragDrop($listBox_DragDrop)
		$form.remove_FormClosed($Form_Cleanup_FormClosed)
	}
	catch [Exception]
    { }
}
 
$button.Add_Click($button_Click)
$listBox.Add_DragOver($listBox_DragOver)
$listBox.Add_DragDrop($listBox_DragDrop)
$form.Add_FormClosed($form_FormClosed)

[void] $form.ShowDialog()
}

Function SelectUserBox($UserList){

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select the Correct User if possible'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a user:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

$userlist | ForEach {[void] $listBox.Items.Add($_.UserPrincipalName)}


$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem
    return $x
}

    return $null

}

Function SendEmail($Output, $TicketNumber){
    $smtpserver = "somatus-com.mail.protection.outlook.com"
    #$from = "ats.jason@somatus.com"
    $from = (Get-AzureADCurrentSessionInfo).account.id
    #$emailaddress = "jason@myaligned.com"
    $emailaddress = helpdesk@myalignedit.com
    $subject = ""

    if($ticketnumber -eq [String]::Empty){
        $subject= "Office 365 Contact Info Update"
        }
        else{
        $subject= "#" + $ticketnumber + ": Office 365 Contact Info Update"
        }

    Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $Output

}

 function OutputFormToUser(){
    $out = $global:output | Out-String
    $WindowTitle = "List of users terminated" 
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

    # Create the Label.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = "Final Result"

    # Create the TextBox used to capture the user's text.
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(10,40)
    $textBox.Size = New-Object System.Drawing.Size(575,200)
    #$textBox.AcceptsReturn = $true
    #$textBox.AcceptsTab = $false
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $out
    $textBox.ReadOnly = $true
    
    # Create ticketbox Label.
    $ticketlabel = New-Object System.Windows.Forms.Label
    $ticketlabel.Location = New-Object System.Drawing.Size(185,252)
    $ticketlabel.Size = New-Object System.Drawing.Size(280,20)
    $ticketlabel.AutoSize = $true
    $ticketlabel.Text = "Ticket Number?"

    # Create the TicketBox
    $ticketbox = New-Object System.Windows.Forms.TextBox
    $ticketbox.Location = New-Object System.Drawing.Point(300,252)
    $ticketbox.Size = New-Object System.Drawing.Size(100,50)

    # Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(405,250)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "Send Email"
    $okButton.Add_Click({ $form.Tag = $textBox.Text; sendEmail $out $ticketbox.Text; $form.Close() })
    #$okButton.Add_Click({ $form.Tag = $textBox.Text; Write-Host $ticketbox.text; $form.Close() })

    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(510,250)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Close"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })

    # Create the form.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(610,320)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true

    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($ticketlabel)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.Controls.Add(($ticketbox))

    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null  # Trash the text of the button that was clicked.
    #$form.ShowDialog()
    # Return the text that the user entered.
    return $form.Tag
}


Function isModuleInstalled(){
    $ADModule=Get-Module -Name AzureAD -ListAvailable  

    if($ADModule.count -eq 0) {   
        $Confirm= Read-Host These following module prompts are required to run program. Install? [Y] Yes [N] No 
        if($Confirm -match "[yY]") { 
            Install-Module AzureAD 
            Import-Module AzureD
        } 
        else 
        { 
            Write-Host AzureAD module is required to connect.
            Exit
        }
    }
    else{
    try{
    Connect-AzureAD -ErrorAction -Stop
    } catch [System.AggregateException] {
        Write-host "Connecting to AzureAD is required to use this application"
        Exit
      }
    }


    $EOModule=Get-Module -Name ExchangeOnlineManagement -ListAvailable  
    
    if($EOModule.count -eq 0) {   
        $Confirm= Read-Host These following module prompts are required to run program. Install? [Y] Yes [N] No 
        if($Confirm -match "[yY]") { 
            Set-ExecutionPolicy RemoteSigned
            Install-module ExchangeOnlineManagement
            Import-module ExchangeOnlineManagement
        } 
        else 
        { 
            Write-Host ExchangeOnline module is required.
            Exit
        }
    }
    else{
        try{
    Connect-ExchangeOnline -ErrorAction -Stop
    } catch [System.AggregateException] {
        Write-host "Connecting to ExchangeOnline is required to use this application"
        Exit
      }
    }
    <#$SPModule=Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable 
    if($SPModule.count -eq 0) {   
        $Confirm= Read-Host These following module prompts are required to run program. Install? [Y] Yes [N] No 
        if($Confirm -match "[yY]") { 
                Install-Module -Name  Microsoft.Online.SharePoint.PowerShell
        } 
        else 
        { 
            Write-Host SPOnline module is required.
            Exit
        }
    }
    else{
        try{
        Connect-SPOService -URL https://somatusoffice365-admin.sharepoint.com/ -ErrorAction -Stop
    } catch [System.AggregateException] {
        Write-host "Connecting to SPOnline is required to use this application"
        Exit
      }
    }#>
    
}

function Read-InputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle, $DefaultText)
}

#isModuleInstalled

Dropbox
foreach($path in $global:finalvalue){
    if(isExtension($path)){
    $displayname = OpenWordDoc($path)
    #check displayname null here
    $username = DoesUserExist($displayname)

    if($username -ne $null){
        if($username.count -eq 1){
            Write-host "in here count"
            terminate($username) 
        }
        elseif($username.count -gt 1){ 
            $userSelection = SelectUserBox($username)
            if($userSelection){
                write-host $userSelection "test"
                $selectedUser = Get-AzureADUser -ObjectId $userSelection
                terminate($selectedUser)
            }
        }
        else{
            Write-Host "Username not found"
            $textEntered = Read-InputBoxDialog -Message "User could not be matched, please enter the user's email address" -WindowTitle "No User Found!"
            if ($textEntered -eq $null) { Write-Host "You clicked CANCEL or left the field Blank!" }
            elseif ($textEntered.trim().length -eq 0) { Write-host "Field was left empty" }
            else { 
                Write-Host "Looking for $textEntered" 
                $searchedUser = Get-AzureADUser -ObjectID $textEntered
                if($searchedUser){
                    terminate($searchedUser)
                }
                else{ Write-Host "User $textEntered could not be found" }
            
            }    
        }
    }
        Write-Host "No Name foud in Document"
    }
    else{
        Write-Host "File: " $path " is of wrong extension"
    }

}

OutputFormToUser


