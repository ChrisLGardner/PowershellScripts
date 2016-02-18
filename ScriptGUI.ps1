<#
.Synopsis
   A simple GUI to allow for the running of scripts outside the commandline and with no knowledge of the parameters required.
.DESCRIPTION
   This GUI will allow users to run scripts from a specified home directory or from a module of choice. It will allow the user to select the Script or command to run and
   then call Show-Command for it to produce a simple GUI with all the parameter options available. 

   Selecting the script from a dropdown list will launch the Show-Command window for that script. Selecting a Module from it's dropdown list will populate the Command list
   with all the available commands from that module, selecting one of these will launch the Show-Command window for it.

   The "include system modules" checkbox will also populate the Module list with all the modules currently available on the system.

.EXAMPLE
   .\ScriptGUI.ps1

   This will launch the GUI and all other functionality takes place within it.

#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


#######################
# Gets all files in target
# location and all files
# one layer below that.
# Returns array of all files
# including folder if needed
######################
function Get-Folders 
{

    $ScriptStore = $args[0]
    $FileType = $Args[1]

	$AllItems = Get-ChildItem -Path $ScriptStore -Include $FileType -Recurse | Select Fullname

	Foreach ($item in $AllItems) 
    {
        $Item.FullName = ($Item.FullName).Replace($ScriptStore,"")
	}
	return $AllItems
}


#######################
# Some variables used throughout
# the application are
# defined here
########################
$FolderLocation = "<SourceFolder>"
$ScriptSource = Get-Folders $FolderLocation "*.ps1"
$ModuleSource = Get-Folders $FolderLocation "*.psm1"

###########################################################
#
# Beginning of form layout
#
###########################################################
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Script Selection Form"
$objForm.Size = New-Object System.Drawing.Size(500,300) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {
        $x=$objTextBox.Text;$objForm.Close()
    }
})
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
    {
        $objForm.Close()
    }
})

##############################
#
# Creating controls for the form
#
##############################

$ScriptComboBox = New-Object System.Windows.Forms.ComboBox
$ScriptComboBox.DropDownStyle = "DropDownList"
$ScriptComboBox.Location = New-Object System.Drawing.Size(40,40)
$ScriptComboBox.Size = New-Object System.Drawing.Size(400,20)
foreach($file in $ScriptSource) 
{
	$ScriptComboBox.Items.Add($file.Fullname)
}
$ScriptComboBox.Add_SelectedIndexChanged({
    Show-Command -Name ($FolderLocation + $ScriptSource[$ScriptComboBox.SelectedIndex].FullName)
})

$ModuleComboBox = New-Object System.Windows.Forms.ComboBox
$ModuleComboBox.DropDownStyle = "DropDownList"
$ModuleComboBox.Location = New-Object System.Drawing.Size(40,110)
$ModuleComboBox.Size = New-Object System.Drawing.Size(400,20)
foreach($file in $ModuleSource) 
{
	$ModuleComboBox.Items.Add($file.Fullname)
}
$ModuleComboBox.Add_SelectedIndexChanged({
    $CommandComboBox.Items.Clear()
    if ((($ModuleComboBox.SelectedItem).split("."))[-1] -eq "psm1") {
        try 
        {
            Import-Module ($FolderLocation + $ModuleSource[$ModuleComboBox.SelectedIndex].FullName)
            $module = ((($ModuleSource[$ModuleComboBox.SelectedIndex].FullName).Split("\"))[-1]).replace(".psm1","")
        }
        catch [System.Exception]
        {
        Write-Error "Could not import module. Please ensure it is a valid module file."
        }
    }
    else
    {
        $module = $ModuleComboBox.SelectedItem
        Import-Module $Module
    }
    foreach ($Command in Get-Command -Module $module)
    {
        $CommandComboBox.Items.Add($Command.Name)
    }
})

$CommandComboBox = New-Object System.Windows.Forms.ComboBox
$CommandComboBox.DropDownStyle = "DropDownList"
$CommandComboBox.Location = New-Object System.Drawing.Size(40,170)
$CommandComboBox.Size = New-Object System.Drawing.Size(400,20)
$CommandComboBox.Add_SelectedIndexChanged({
    Show-Command -Name $CommandComboBox.SelectedItem
})

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(213,230)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Exit"
$CancelButton.Add_Click({$objForm.Close()})

$ScriptLabel = New-Object System.Windows.Forms.Label
$ScriptLabel.Location = New-Object System.Drawing.Size(10,20) 
$ScriptLabel.Size = New-Object System.Drawing.Size(400,20) 
$ScriptLabel.Text = "Either select a script from the dropdown:"

$ModuleLabel = New-Object System.Windows.Forms.Label
$ModuleLabel.Location = New-Object System.Drawing.Size(10,80) 
$ModuleLabel.Size = New-Object System.Drawing.Size(400,20) 
$ModuleLabel.Text = "Or select a module from the dropdown:"

$CommandLabel = New-Object System.Windows.Forms.Label
$CommandLabel.Location = New-Object System.Drawing.Size(10,140) 
$CommandLabel.Size = New-Object System.Drawing.Size(400,20) 
$CommandLabel.Text = "Then select a command for that module from the dropdown:"

$SystemModuleCheckBox = New-Object System.Windows.Forms.CheckBox
$SystemModuleCheckBox.Location = New-Object System.Drawing.Size(10,200)
$SystemModuleCheckBox.Size = New-Object System.Drawing.Size(170,25)
$SystemModuleCheckBox.Text = "Include System Modules?"
$SystemModuleCheckBox.Add_CheckedChanged({
    if ($SystemModuleCheckBox.Checked) {
        Foreach ($module in (Get-Module -ListAvailable)) {
            $ModuleComboBox.Items.Add($Module.Name)
        }
    }
    else {
        Foreach ($module in (Get-Module -ListAvailable)) {
            $ModuleComboBox.Items.Remove($Module.Name)
        }
    }

})

######################################
#
# Adding all the Controls to the form.
#
######################################


$objForm.Controls.Add($ScriptComboBox)
$objForm.Controls.Add($ModuleComboBox)
$objForm.Controls.Add($CommandComboBox)
$objForm.Controls.Add($CancelButton)
$objForm.Controls.Add($ScriptLabel)
$objForm.Controls.Add($ModuleLabel) 
$objForm.Controls.Add($CommandLabel)
$objForm.Controls.Add($SystemModuleCheckBox)

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()