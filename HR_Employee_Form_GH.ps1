<#

        .NOTES
        =========================================================================
         Created With:  Windows PowerShell ISE
         Created On:    04/23/2021
        =========================================================================
        .DESCRIPTION
        This script was designed for HR to be able to upload new
        employee information to the system without having MIS, or anyone with system
        administration access do it instead. Unfortunately, due to the nature
        of the system, the password section does not 'technically' create the password
        and therefore requires a sysadmin to manually input it, defeating 
        the purpose of this script.
         
        The script creates two forms. The first one is for employee information.
        The second one is for which groups the user should be apart of.
        =========================================================================
        .PERMISSIONS
        This required certain permissions to function. It was required to log
        in to the database server, find each of the above tables within SSMS 
        and add a specific user to its' permission group, then to be granted 
        'INSERT' permission and NOTHING ELSE. This is to ensure that no other tables
        are inadvertently modified. No other table within the database 
        contains this user for remote insert modification or any other permissions. 
        =========================================================================
        .AFTERTHOUGHTS
        An array would have been much better to use for certain aspects of the
        pre-populated areas.
        
        Changed 'Lucinda' to 'Lucida'
        
        INSERT might not be the best way to update. Rather, look into TRANSACTION.
        This ensures that ALL aspects are successful before publishing changes to the
        database.
        
         
#>



[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


## BUILDING THE FORM
##############################################################################################################################



## Position Variables
$okXPosition = 110
$okYPosition = 780

#

$cancelXPosition = 260
$cancelYPosition = 780

#

$formWidth = 500
$formHeight = 900

#

$loginLabelXPosition = 47
$loginLabelYPosition = 20

$loginFieldXPosition = 50
$loginFieldYPosition = 55

#

$empFirstLabelXPosition = 47
$empFirstLabelYPosition = 95

$empFirstXPosition = 50
$empFirstYPosition = 130

#

$empLastLabelXPosition = 47
$empLastLabelYPosition = 170

$empLastXPosition = 50
$empLastYPosition = 210

#

$laborClassLabelXPosition = 47
$laborClassLabelYPosition = 250

$laborClassFieldXPosition = 50
$laborClassFieldYPosition = 290

#

$empTypeLabelXPosition = 47
$empTypeLabelYPosition = 330

$empTypeMenuXPosition = 50
$empTypeMenuYPosition = 370

#

$roleLabelXPosition = 47
$roleLabelYPosition = 510

$roleMenuXPosition = 50
$roleMenuYPosition = 600

#

$timeTypeLabelXPosition = 47
$timeTypeLabelYPosition = 330

$timeTypeMenuXPosition = 50
$timeTypeMenuYPosition = 370

#

$shopLabelXPosition = 47
$shopLabelYPosition = 460

$shopMenuXPosition = 50
$shopMenuYPosition = 500





## Form itself
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "AiM Employee Form"
$objForm.Size = New-Object System.Drawing.Size($formWidth,$formHeight) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
        $objForm.Close()
    }
})



## OK Button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point($okXPosition,$okYPosition)
$OKButton.Size = New-Object System.Drawing.Size(110,60)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$objForm.AcceptButton = $OKButton
$objForm.Controls.Add($OKButton)



###



## Cancel Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point($cancelXPosition,$cancelYPosition)
$CancelButton.Size = New-Object System.Drawing.Size(110,60)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
#$CancelButton.Add_Click({$objForm.Close()})
$objForm.CancelButton = $CancelButton
$objForm.Controls.Add($CancelButton)


###



## A Number/Login label
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($loginLabelXPosition,$loginLabelYPosition) 
$objLabel.Size = New-Object System.Drawing.Size(325,30) 
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Employee ID"
$objForm.Controls.Add($objLabel) 






## Login field
$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size($loginFieldXPosition,$loginFieldYPosition) 
$objTextBox.Size = New-Object System.Drawing.Size(400,20) 
#$objTextBox.Multiline = $true
$objTextBox.Font = New-Object System.Drawing.Font("Lucida",16,[System.Drawing.FontStyle]::Regular)
#$objTextBox.Text = "()"
$objForm.Controls.Add($objTextBox) 

###


## Employee first name label
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($empFirstLabelXPosition,$empFirstLabelYPosition) 
$objLabel.Size = New-Object System.Drawing.Size(325,30) 
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Employee First Name"
$objForm.Controls.Add($objLabel)






## Employee first Name field
$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size($empFirstXPosition,$empFirstYPosition) 
$objTextBox2.Size = New-Object System.Drawing.Size(400,20)
$objTextBox2.Font = New-Object System.Drawing.Font("Lucida",16,[System.Drawing.FontStyle]::Regular) 
#$objTextBox2.Text = "(Employee Name)"
$objForm.Controls.Add($objTextBox2) 


###


## Employee last name label
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($empLastLabelXPosition,$empLastLabelYPosition) 
$objLabel.Size = New-Object System.Drawing.Size(325,30) 
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Employee Last Name"
$objForm.Controls.Add($objLabel)




## Employee last Name field
$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size($empLastXPosition,$empLastYPosition) 
$objTextBox3.Size = New-Object System.Drawing.Size(400,20)
$objTextBox3.Font = New-Object System.Drawing.Font("Lucida",16,[System.Drawing.FontStyle]::Regular) 
#$objTextBox2.Text = "(Employee Name)"
$objForm.Controls.Add($objTextBox3) 


###


## Labor Class Label
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($laborClassLabelXPosition,$laborClassLabelYPosition) 
$objLabel.Size = New-Object System.Drawing.Size(325,30) 
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Labor Class"
$objForm.Controls.Add($objLabel)


## Labor Class Field
$objTextBox4 = New-Object System.Windows.Forms.TextBox 
$objTextBox4.Location = New-Object System.Drawing.Size($laborClassFieldXPosition,$laborClassFieldYPosition) 
$objTextBox4.Size = New-Object System.Drawing.Size(400,20)
$objTextBox4.Font = New-Object System.Drawing.Font("Lucida",16,[System.Drawing.FontStyle]::Regular) 
#$objTextBox2.Text = "(Employee Name)"
$objForm.Controls.Add($objTextBox4) 

###

## Time type label 
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($timeTypeLabelXPosition,$timeTypeLabelYPosition)
$objLabel.Size = New-Object System.Drawing.Size(325,30)
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Select Time Type"
$objForm.Controls.Add($objLabel)




## Time type menu 
$listbox2 = New-Object System.Windows.Forms.ListBox
$listbox2.Location = New-Object System.Drawing.Point($timeTypeMenuXPosition,$timeTypeMenuYPosition)
$listbox2.Size = New-Object System.Drawing.Size(400,20)
$listbox2.Font = New-Object System.Drawing.Font("Lucida",12,[System.Drawing.FontStyle]::Regular)
$listbox2.Height = 90

[void] $listBox2.Items.Add('REGULAR')
[void] $listBox2.Items.Add('OVERTIME')
[void] $listBox2.Items.Add('COMP EARNED')
[void] $listBox2.Items.Add('OVERDBL')
[void] $listBox2.Items.Add('SAFETY OVT')
[void] $listBox2.Items.Add('SAFETY COMP')
[void] $listBox2.Items.Add('COMPDBL')


$objForm.Controls.Add($listBox2)





## Help/label prompt for SHOP
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($shopLabelXPosition,$shopLabelYPosition) 
$objLabel.Size = New-Object System.Drawing.Size(325,40) 
$objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Select Shop"
$objForm.Controls.Add($objLabel)






## Drop down menu for SHOP
$listBox1 = New-Object System.Windows.Forms.ListBox
$listBox1.Location = New-Object System.Drawing.Point($shopMenuXPosition,$shopMenuYPosition)
$listBox1.Size = New-Object System.Drawing.Size(400,20)
$listBox1.Font = New-Object System.Drawing.Font("Lucida",12,[System.Drawing.FontStyle]::Regular)
$listBox1.Height = 270

#$listBox.SelectionMode = 'MultiExtended'

## Could be put into an array rather than individual lines
[void] $listBox1.Items.Add('ADMIN')
[void] $listBox1.Items.Add('BUSSERV')
[void] $listBox1.Items.Add('CARPENTERS')
[void] $listBox1.Items.Add('CEP')
[void] $listBox1.Items.Add('COMMISSION')
[void] $listBox1.Items.Add('CUSTSERV')
[void] $listBox1.Items.Add('D&C')
[void] $listBox1.Items.Add('DIRECTOR')
[void] $listBox1.Items.Add('DISTRIBUTE')
[void] $listBox1.Items.Add('ELECTRICAL')
[void] $listBox1.Items.Add('ENERGYMGMT')
[void] $listBox1.Items.Add('EQUIPOPER')
[void] $listBox1.Items.Add('EVENTUTILS')
[void] $listBox1.Items.Add('FINISHES')
[void] $listBox1.Items.Add('FM')
[void] $listBox1.Items.Add('HIGHVOLT')
[void] $listBox1.Items.Add('HVAC')
[void] $listBox1.Items.Add('LABGAS')
[void] $listBox1.Items.Add('LIFTMAINT')
[void] $listBox1.Items.Add('LOAM')
[void] $listBox1.Items.Add('LOCKSMITH')
[void] $listBox1.Items.Add('MAINTADMIN')
[void] $listBox1.Items.Add('MIS')
[void] $listBox1.Items.Add('MOVING')
[void] $listBox1.Items.Add('PARTEAM')
[void] $listBox1.Items.Add('PLANNING')
[void] $listBox1.Items.Add('PLUMBERS')
[void] $listBox1.Items.Add('PROJ&ENGR')
[void] $listBox1.Items.Add('PROJECTS')
[void] $listBox1.Items.Add('PURCHASING')
[void] $listBox1.Items.Add('RCDEMGR')
[void] $listBox1.Items.Add('RECEIVING')
[void] $listBox1.Items.Add('RECPOSTAGE')
[void] $listBox1.Items.Add('RECYCLING')
[void] $listBox1.Items.Add('SAFETY')
[void] $listBox1.Items.Add('SECURITY')
[void] $listBox1.Items.Add('SNOW REMOVAL')
[void] $listBox1.Items.Add('STORAGE')
[void] $listBox1.Items.Add('SUPPTADMIN')
[void] $listBox1.Items.Add('SURPLUS')
[void] $listBox1.Items.Add('UTILADMIN')
[void] $listBox1.Items.Add('UTILITIES')
[void] $listBox1.Items.Add('WAREHOUSE')
[void] $listBox1.Items.Add('WASTEMGMT')
[void] $listBox1.Items.Add('WATERQLTY')


$objForm.Controls.Add($listBox1)

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
#[void]$objForm.ShowDialog()


$result = $objForm.ShowDialog()


if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Login = $objTextBox.Text ## Emp ID
    $FirstName = $objTextBox2.Text ## First Name
    $LastName = $objTextBox3.Text ## Last Name
    $LaborClass = $objTextBox4.Text ## Labor Class
    $Description = $objTextBox2.Text + " " + $objTextBox3.Text ## Description = first name and last name
    $EmployeeID = $objTextBox.Text ## Emp ID
    $Password = "P@ssw0rd" ## Password
    $EmployeeType = "S"
    $TimeType = $listbox2.SelectedItem
    $Shop = $listBox1.SelectedItem ## Employee Shop
    #$Shop = $objTextBox4.Text
    #$Role = @($listBox.SelectedItems)
    
    ## Confirm variables. Used for testing
    <#
    $Login
    $FirstName
    $LastName
    $Description
    $Password
    $LaborClass
    $EmployeeID
    $EmployeeType
    $TimeType
    $Shop
    #$Shop
    #$Role
    #$Role.GetType()#>

    ## Sql Statement to insert employee information
    $EmployeeInfo = "INSERT INTO table1 (login, description, password, employee_id, shop, default_org, active)
    VALUES ('$Login', '$Description', '$Password', '$EmployeeID', '$Shop', 'N', 'Y')
    "

    ## Pass query to SQL Server & return results
    $Pass = Invoke-Sqlcmd -Query $EmployeeInfo -ServerInstance "SERVER IP" -Username "USERNAME" -Password "PASSWORD"
    #$Pass




    ## Sql statement to insert values into HR table
    $HRTable = "INSERT INTO table2 (shop_person, fname, lname, time_type, labor_class, active, emp_type)
    VALUES ('$Login', '$FirstName', '$LastName', '$TimeType', '$LaborClass', 'Y', '$EmployeeType')
    "


    ## Pass query to SQL
    $EmployeeProfileUpdate = Invoke-Sqlcmd -Query $HRTable -ServerInstance "SERVER IP" -Username "USERNAME" -Password "PASSWORD"


} ## End 'IF' Statement

## ROLES Prompt now that the user has been added
##############################################################################################################################

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    
    ## Perform this task until the "finish" button is pressed
    do
    {

        ## Position Variables
        $doneXPosition = 360
        $doneYPosition = 400

        $cancelXPosition = 360
        $cancelYPosition = 400

        $addXPosition = 50
        $addYPosition = 400

        $formWidth = 500
        $formHeight = 530

        $LabelXPosition = 100
        $LabelYPosition = 20

        $roleMenuXPosition = 50
        $roleMenuYPosition = 120



        ## Form itself
        $objForm = New-Object System.Windows.Forms.Form 
        $objForm.Text = "AiM Employee Roles"
        $objForm.Size = New-Object System.Drawing.Size($formWidth,$formHeight) 
        $objForm.StartPosition = "CenterScreen"

        $objForm.KeyPreview = $True
        $objForm.Add_KeyDown({
            if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
                $objForm.Close()
            }
        })


        ## Done Button
        $DoneButton = New-Object System.Windows.Forms.Button
        $DoneButton.Location = New-Object System.Drawing.Point($doneXPosition,$doneYPosition)
        $DoneButton.Size = New-Object System.Drawing.Size(90,35)
        $DoneButton.Text = 'Finish'
        $DoneButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $objForm.AcceptButton = $DoneButton
        $objForm.Controls.Add($DoneButton)

        <#

        
        ## Cancel Button
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size($cancelXPosition,$cancelYPosition)
        $CancelButton.Size = New-Object System.Drawing.Size(90,35)
        $CancelButton.Text = "Cancel"
        $CancelButton.Add_Click({$objForm.Close()})
        $objForm.Controls.Add($CancelButton)
        
        ###>


        ## Add Button
        $AddButton = New-Object System.Windows.Forms.Button
        $AddButton.Location = New-Object System.Drawing.Point($addXPosition,$addYPosition)
        $AddButton.Size = New-Object System.Drawing.Size(90,35)
        $AddButton.Text = 'Add'
        $AddButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $objForm.AcceptButton = $AddButton
        $objForm.Controls.Add($AddButton)




        ## Label/help 
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size($LabelXPosition,$LabelYPosition) 
        $objLabel.Size = New-Object System.Drawing.Size(325,80) 
        $objLabel.Font = New-Object System.Drawing.Font("Lucida",14,[System.Drawing.FontStyle]::Regular)
        $objLabel.Text = "Select one Role at a time then press 'Add'. Press 'Finish' when done."
        $objForm.Controls.Add($objLabel) 




        ## Drop down menu for ROLES
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point($roleMenuXPosition,$roleMenuYPosition)
        $listBox.Size = New-Object System.Drawing.Size(400,20)
        $listBox.Font = New-Object System.Drawing.Font("Lucida",12,[System.Drawing.FontStyle]::Regular)
        $listBox.Height = 260

        #$listBox.SelectionMode = 'MultiExtended'

        ## This could be modified into an array rather than individual lines
        [void] $listBox.Items.Add('USU_OM_SHOP_FOREMAN')
        [void] $listBox.Items.Add('USU_OM_SHOP_PERSON')
        [void] $listBox.Items.Add('USU_OM_VIEW_ONLY')
        [void] $listBox.Items.Add('EXEMPT')
        [void] $listBox.Items.Add('HOURLY')
        [void] $listBox.Items.Add('NON-EXEMPT')
        [void] $listBox.Items.Add('GROUP MANAGER')


        $objForm.Controls.Add($listBox)


        $objForm.Topmost = $True

        $objForm.Add_Shown({$objForm.Activate()})
        #[void]$objForm.ShowDialog()


        $result2 = $objForm.ShowDialog()



        ## If the user presses "add", add roles to specified user
        if ($result2 -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $Role = @($listBox.SelectedItems)
            #$Role

            ## Sql Statement to add Roles to specified user
            $EmployeeRoles = "INSERT INTO table3 (role_id, login)
            VALUES ('$Role', '$Login')
            SELECT t3.role_id, t3.login
            FROM
            table3 t3
            INNER JOIN
            table1 t1 ON (t3.login = t1.login)
            WHERE
            sec.login = '$Login'
            "
        
            $ExecuteRoleQuery = Invoke-Sqlcmd -Query $EmployeeRoles -ServerInstance "SERVER IP" -Username "USERNAME" -Password "PASSWORD"

        }



    } while ($result2 -ne [System.Windows.Forms.DialogResult]::Cancel)

}
#>