[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function OpenWordDoc($documentPath){
    Write-host $documentpath
    $documents = $documentPath.ToString()
    $word = New-Object -ComObject Word.application
    $document = $word.Documents.Open($documents,$null,$true)
    $array = New-Object System.Collections.Generic.List[System.Object]
    $document.shapes |Where-Object {$_. type -eq 17} | ForEach-Object{
    $array.add($_.TextFrame.TextRange.text)}
    $username = $array[0]
    #write-host $username
    #$array.ToArray()
    $array.clear()
    $document.Close($Word.WdSaveOptions.wdDoNotSaveChanges)
    $word.Quit()
    return ('*' + $username + '*')
}

Function DoesUserExists($username){
     if($username){
         $users = get-azureaduser -all $true | Where-Object{$_.DisplayName -like $username}
         return $users
     }
     return $null
}

Function terminate($user){
    forceUserSignout($user)
    blockUserAccount($user)
    removeUserFromGroups($user)
    #getOneDriveLink($username) #Can be commented out if user using script has SharePoint access
    if(ifSharedMailbox($user)){setSharedMailbox($user)}
}

#Exchange
Function ifSharedMailbox($username){
    $isShared = (Get-Mailbox -Identity $username.UserPrincipalName).RecipientTypeDetails
    return $isShared -eq "UserMailBox"
}

#Exchange
Function setSharedMailbox($username){
    Set-Mailbox $username.UserPrincipalName -Type Shared
}

#Azure
Function forceUserSignOut($username){
    #Get-AzureADUser -SearchString $username | Revoke-AzureADUserAllRefreshToken
    $username | Revoke-AzureADUserAllRefreshToken
}

#Azure
Function blockUserAccount($username){
    Set-AzureADUser -ObjectID $username.ObjectID -AccountEnabled $false
}

#Still have not tested.
Function removeUserFromGroups($username){
    $Membership = Get-AzureADUserMembership -ObjectId $username.ObjectId
    ForEach($group in $Membership){
        try{
            Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $username.ObjectId
        } catch [Microsoft.Open.AzureAD16.Client.ApiException] {
            Remove-DistributionGroupMember -Identity $group.displayname -Member $username.ObjectId -Confirm:$False
        }
    }
}

Function getOneDriveLink($username){
    return (Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Owner -like '$username.UserPrincipalName'").url
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
}

isModuleInstalled

### Create form ###
 
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell GUI"
$form.Size = '260,320'
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size
$form.MaximizeBox = $False
$form.Topmost = $True
 
 
### Define controls ###
 
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
 
 
### Add controls to form ###
 
$form.SuspendLayout()
$form.Controls.Add($button)
$form.Controls.Add($checkbox)
$form.Controls.Add($label)
$form.Controls.Add($listBox)
$form.Controls.Add($statusBar)
$form.ResumeLayout()
 
 
### Write event handlers ###

$button_Click = {
    write-host "Listbox contains:" -ForegroundColor Yellow
	foreach ($item in $listBox.Items)
    {
        if($item -is [System.IO.DirectoryInfo])
        {
            write-host ("`t" + $item.Name + " [Directory]")
            
        }
        else
        {
            $username = OpenWordDoc($item)
            Write-host $username
            $user = DoesUserExist($Username)

            if($user.count -eq 1){ terminate($user) }
            if($user.count -gt 1){ 
               $userSelection = SelectUserBox($User)
               if($userSelection){terminate($user)}
            }
            else{
             Write-Host "Username not found"
             #will add entry field option.
            }

            #write-host ("`t" + $item + " [" + [math]::round($i.Length/1MB, 2) + " MB]")

        }
	}
 
    if($checkbox.Checked -eq $True)
    {
        $listBox.Items.Clear()
    }
    
    $statusBar.Text = ("List contains $($listBox.Items.Count) items")
    #Disconnect-AzureAD
    #Disconnect-ExchangeOnline
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
 
 
### Wire up events ###
 
$button.Add_Click($button_Click)
$listBox.Add_DragOver($listBox_DragOver)
$listBox.Add_DragDrop($listBox_DragDrop)
$form.Add_FormClosed($form_FormClosed)
 
 
#### Show form ###
 
[void] $form.ShowDialog()


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


