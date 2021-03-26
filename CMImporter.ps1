#Version:        1.0
#Author:         Kim Andersen
#Creation Date:  25/03/2021
#Purpose/Change: Initial script development
#
#DESCRIPTION
#
#
#test
#-----------------------------------------------------------[Functions]------------------------------------------------------------

#Loads the Present
Add-Type -AssemblyName PresentationFramework

#Active Directory

#Function to check if provided ADComputerName value exists in Active Directory
function adcomputerexists {
    param (
        [string]$Comp,
        [string]$existcheck = $false
    )
    
    #Checks if the ADComputerName field is populated
    if ($Comp) {

        $output = $null
        $var_console.AppendText("Computer is $comp`n")

        try {
            #Uses the simple get-adcomputer command to check if exists in AD
            #Also grabs the distinguished name property from the AD object, if it exists
            $ADobject = Get-ADComputer $comp -Properties distinguishedname

            $var_console.AppendText("Found $comp in Active Directory`n")

            #Populates the output variable with "Exists"
            $output = "Exists"
            
        }
        catch {
            #If the command cannot find a object in AD with the provided value
            $ADobject = $null
            $var_console.AppendText("Could not find $comp in Active Directory`n")

            #Populates the output variable with "Does not Exists"
            $output = "Does not Exists"
        }
        
        #Populates the validation field with the $output variable
        $var_adCompQueryresult.Text = $output

        #Locks the recreate and create new button
        $var_recreateADcomp.IsEnabled = $false
        $var_createADcomp.IsEnabled = $false
        
        #This part grabs the distinguished name from the initial AD query and correctly formats it
        if (($ADobject) -and ($existcheck -eq $false)) {

            #If the adobject was found, the string is formatted to remove the adobject name from the beginning of the DN
            $distinguishedname = $ADobject.DistinguishedName
            $adpath = $distinguishedname.Substring($distinguishedname.IndexOf(",") + 1)
            $var_console.AppendText("DistinguishedName is $adpath`n")

            #Returns the formatted DN to the DN output field in the application
            $var_adDN.Text = $adpath
        }
        elseif (($ADobject -eq $null) -and ($existcheck -eq $true)) {
            $var_console.AppendText("Computercheck is set to $existcheck, DN won't be removed`n")
        } 
        else {
            #If no ADobject was found, no DN is returned and the DN output field is cleared in the application
            $var_console.AppendText("DistinguishedName not automatically added`n")
            $var_adDN.Text = $null
            $adpath = $null
        }
    }
    else {
        #If the ADComputername field is empty, all buttons are locked
        $var_console.AppendText("ADComputerName field is empty!`n")
        $var_recreateADcomp.IsEnabled = $false
        $var_createADcomp.IsEnabled = $false
        $var_CCMrecreatebtn.IsEnabled = $false
        $var_CCMCreateNewbtn.IsEnabled = $false
        $var_adDN.Text = $null
        $var_adCompQueryresult.Text = $null
        $var_CCMUUIDtxt.Text = $null
        $var_CCMMACtxt.Text = $null
        $var_CCMCheckbtnresult.Text = $null
    }

    $var_console.ScrollToEnd()

    #Returns the DistinguishedName value to be used in the DN check function
    return $adpath   
}

#Function to check provided DN is valid
function distinguishednameexists {
    param (
        [string]$adpath
    )

    #Checks if the $adpath variable is populated
    #$adpath is generated in the adcomputerexists function
    if ($adpath) {
        try {
            
            #Simple command that queries AD for the provided value 
            $DNexists = Get-ADOrganizationalUnit $adpath
            
            #If the query is successfull, it populates the validation variable with "Exists"
            $var_console.AppendText("DistinguishedName is valid`n")
            $DNoutput = "Exists"
        }
        #If the Get-ADOrganizationalUnit command fails, re-create button is locked
        catch {
            $DNexists = $null
            $var_console.AppendText("DistinguishedName is not valid`n")
            $DNoutput = "Does not Exists"
            $var_recreateADcomp.IsEnabled = $false
        }

        #Validate DN text box is populated
        $var_adDNQueryresult.Text = $DNoutput

        #Checks if the computername value "Exists", and the DN path is valid, then unlocks the re-create button
        if (($var_adDNQueryresult.Text -eq "Exists") -and ($var_adCompQueryresult.Text -eq "Exists")) {
            $var_recreateADcomp.IsEnabled = $true
            $var_createADcomp.IsEnabled = $false
        }
        #Checks if the computername value is "Does not Exists", and the DN path is valid, then unlocks the create new button
        elseif (($var_adDNQueryresult.Text -eq "Exists") -and ($var_adCompQueryresult.Text -eq "Does not Exists")) {
            $var_createADcomp.IsEnabled = $true
            $var_recreateADcomp.IsEnabled = $false
        }
        #If either of the above checks fails, both buttons are locked
        else {
            $var_createADcomp.IsEnabled = $false
            $var_recreateADcomp.IsEnabled = $false
        }
    }
    #If no DN value is provided, both buttons are locked
    else {
        $var_console.AppendText("DistinguishedName is empty!`n")
        $var_recreateADcomp.IsEnabled = $false
        $var_createADcomp.IsEnabled = $false
    }

    
    $var_console.ScrollToEnd()
}

#Function to recreate the ADObject
function recreateADobject {
    param (
        [Parameter(Mandatory)]
        [string]$adpath,
        [string]$comp,
        [string]$recreate
    )

    #If the recreate value is $true, the existing object in SCCM will be deleted.
    #If the recreate value is $false, the deletion step is skipped

    if ($var_txtComputer.Text) {

        if ($recreate -eq $true) {
        
            $var_console.AppendText("Attempting to delete $comp from Active Directory`n")
        
            try {
                Write-Warning "$comp will be deleted from AD"
    
                #Deletes the provided ADComputerName value from Active Directory
                Get-ADComputer $comp | Remove-ADComputer -Confirm:$false
                
                #Starts a loop that queries AD every 5 seconds for the ADComputerName, until no longer detected
                #This is useful for larger AD forests with multiple domain controllers, where replication can take a while
                do {
                    try {
                        $Check = Get-ADComputer $comp
                        Write-Host "$comp still exists, please wait..." -Verbose
                        Start-Sleep 5    
                    }
                    catch {
                        $Check = $null
                        write-host "$comp removed, continuing...." -Verbose
                    }
                }
                until($null -eq $Check)
                $var_console.AppendText("$comp successfully deleted`n")
    
            }
            catch {
                Write-Error $error[0].Exception.Message
                $var_console.AppendText($error[0].Exception.Message + "`n")
            }
        
        }
        #Creates the object in Active Directory
        try {
            $var_console.AppendText("Attempting to create $comp`n")
            Write-Warning "New AD computer will have name $comp, and added to $adpath"
    
            #Runs the new-adcomputer command with provided name and DN path
            New-ADComputer -Name $comp -SAMAccountName $comp -Path $adpath -Enabled $true -Confirm:$false
            
            #Starts a loop that queries AD every 5 seconds for the ADComputerName, until the computer is detected
            #This is useful for larger AD forests with multiple domain controllers, where replication can take a while
            do {
                try {
                    $Check = Get-ADComputer $comp
                    write-verbose "$comp created, continuing..." -Verbose
                         
                }
                catch {
                    $Check = $null
                    write-warning "$comp does not yet exists, please wait..."
                    Start-Sleep 5 
                }
            }
            until($null -ne $Check)
        
            $var_console.AppendText("$comp successfully created`n")
        
        }
        catch {
            Write-Error $error[0].Exception.Message
            $var_console.AppendText($error[0].Exception.Message + "`n")
        }
    }
    else {
        $var_console.AppendText("ADComputerName field is empty!`n")
        $var_recreateADcomp.IsEnabled = $false
        $var_createADcomp.IsEnabled = $false
        $var_CCMrecreatebtn.IsEnabled = $false
        $var_CCMCreateNewbtn.IsEnabled = $false
        $var_adDN.Text = $null
        $var_adCompQueryresult.Text = $null
        $var_CCMUUIDtxt.Text = $null
        $var_CCMMACtxt.Text = $null
        $var_CCMCheckbtnresult.Text = $null
        $var_CCMCreateNewbtn.IsEnabled = $false
        $var_CCMrecreatebtn.IsEnabled = $false
    }

    $var_console.ScrollToEnd()
    
}

#SCCM Module import
function importccm {

    #Copy paste your config manager PS module import
    #Can be generated by the config manager console by going to configmgrconsole menu/Connect via Windows Powershell ISE


    #Leave this as it scrolls the output text in the application
    $var_console.ScrollToEnd()
}

#ActiveDirectory Module import
function importAD {

    $var_console.AppendText("Checking if ActiveDirectory module is loaded`n")

    if (Get-Module ActiveDirectory) {
        $var_console.AppendText("Active Directory module already loaded`n")
    }
    else {
        try {
            $var_console.AppendText("Active Directory module not detected`n")
            $var_console.AppendText("Importing, please wait...`n")

            Import-Module "Path to module"

            $var_console.AppendText("Active Directory module loaded`n")
        }
        catch {
            $var_console.AppendText("Failed to load Active Directory module`n")
            Write-Error $error[0].Exception.Message
            $var_console.AppendText($error[0].Exception.Message + "`n")
        }
    }
    $var_console.ScrollToEnd()
}

#Function to get a list of SCCM collections
function ccmcollectionsrefresh {

    #Currently a list
    $CMcollectioninput = @(

    )




    #A test to see if the SCCM query has already been run
    #This is a step to try speeding up the script, as refreshing the list can take a while
    #Currently reduntent and list will always query SCCM
    if ($Null -eq $CMCollections) {

        #Starts a loop that checks if a collection exists in SCCM that matches provided values in text file
        $CMCollections = foreach ($collection in $CMcollectioninput) {

            try {
                #Returns collection name and collectionID
                #CollectionID is currently not used
                Get-CMCollection -Name $collection | Select-Object CollectionID, Name

                #Writes each result to the application console
                $var_console.AppendText("Found $collection`n")
            }
            catch {
                $var_console.AppendText("Could not find $collection`n")
            }
             
        }
    }

    #Clears the list of collections
    $var_ColctionScroll.Items.Clear()

    #Writes each collections found in SCCM to the list
    foreach ($item in $CMCollections) {
        $var_ColctionScroll.Items.Add($item.Name)
    }

    #Unlocks the validate button
    $var_CMValidate.IsEnabled = $true
    $var_console.ScrollToEnd()

}

#Function that checks if the object already exists in SCCM
function testcmdevice {
    param (
        [string]$Comp
    )

    #Checks if a value has been provided in the ADComputerName field
    #If no value has been provided, no check is executed
    if ($Comp) {
        $var_console.AppendText("Checking if $comp exists in CCM`n")

        #Runs the simple get-cmdevice command to query SCCM for the device name
        $CMcomptest = Get-CMDevice -Name $Comp

        #Checks if a result was returned in the $CMcomptest variable
        if ($CMcomptest) {
            $var_console.AppendText("Found $comp in CCM`n")

            #Populates the $CMtestoutput variable with "Exists"
            $cmtestoutput = "Exists"

            #Populates the UUID field with the UUID value returned from the get-cmdevice command, if it exists
            $var_CCMUUIDtxt.Text = $CMcomptest.SMBIOSGUID
            $var_console.AppendText("Found UUID: " + $CMcomptest.SMBIOSGUID + "`n")
            
            #Populates the MAC address field with the MAC value returned from the get-cmdevice command, if it exists
            $var_CCMMACtxt.Text = $CMcomptest.MACAddress
            $var_console.AppendText("Found MAC Address: " + $CMcomptest.MACAddress + "`n")
        }
        #If no results returned from the get-cmdevice command, UUID and MAC are not populated
        else {
            $var_console.AppendText("$comp does not exist in CCM`n")
            $cmtestoutput = "Does not exists"

            #Clears the UUID and MAC address field and unlocks incase they been locked
            $var_CCMUUIDtxt.Text = $null
            $var_CCMUUIDtxt.IsEnabled = $true

            $var_CCMMACtxt.Text = $null
            $var_CCMMACtxt.IsEnabled = $true


        }
        $var_CCMCheckbtnresult.Text = $cmtestoutput

    }
    #If no value has been provided in the ADComputerName field, check is skipped and buttons locked
    else {
        $var_console.AppendText("ADComputerName field is empty!`n")
        $var_recreateADcomp.IsEnabled = $false
        $var_createADcomp.IsEnabled = $false
        $var_CCMrecreatebtn.IsEnabled = $false
        $var_CCMCreateNewbtn.IsEnabled = $false
        $var_adDN.Text = $null
        $var_adCompQueryresult.Text = $null
        $var_CCMUUIDtxt.Text = $null
        $var_CCMMACtxt.Text = $null
        $var_CCMCheckbtnresult.Text = $null
    }
    $var_CCMCreateNewbtn.IsEnabled = $false
    $var_CCMrecreatebtn.IsEnabled = $false
    $var_console.ScrollToEnd()

}

#Function to recreate, or create new SCCM object
function recreatecmdevice {
    param (
        [string]$comp,
        [string]$recreate,
        [string]$UUID,
        [string]$MAC,
        [string]$coll
    )

    #If the recreate value is $true, the existing object in SCCM will be deleted.
    #If the recreate value is $false, the deletion step is skipped
    if ($recreate -eq $true) {
        try {
            $var_console.AppendText("Attempting to delete object from SCCM`n")

            #Disables buttons before script is run
            $var_CCMrecreatebtn.IsEnabled = $false
            $var_CCMCreateNewbtn.IsEnabled = $false
            $var_CCMUUIDtxt.IsEnabled = $false
            $var_CCMMACtxt.IsEnabled = $false

            #Deletes the SCCM object
            Remove-CMDevice -Name $comp -Force
            Write-Host "$comp removed from SCCM, continuing..." -Verbose
            $var_console.AppendText("$comp successfully deleted from SCCM`n")

            #Sleeps for 5 seconds before proceeding to re-create the object
            Start-Sleep 5
        }
        catch {
            Write-Error $error[0].Exception.Message
            $var_console.AppendText($error[0].Exception.Message + "`n")
        }
        
    }

    #Creates the object in SCCM
    try {
        Write-Warning "Importing $comp to SCCM..."
        $var_console.AppendText("Importing $comp to SCCM`n")

        #Checks if UUID and Mac address has been provided
        if ($UUID -and $MAC) {
            $var_console.AppendText("MAC Address and UUID detected`n")

            #Imports the object with provided values
            Import-CMComputerInformation -ComputerName $comp -SMBiosGuid $UUID -MacAddress $MAC -CollectionName $coll
        }
        #Checks if only UUID has been provided
        elseif ($UUID -and (-not ($MAC))) {
            $var_console.AppendText("UUID detected`n")

            #Imports the object with provided values
            Import-CMComputerInformation -ComputerName $comp -SMBiosGuid $UUID -CollectionName $coll
        }
        #Imports the device with MAC only
        else {
            $var_console.AppendText("MAC Address detected`n")

            #Imports the object with provided values
            Import-CMComputerInformation -ComputerName $comp -MacAddress $MAC -CollectionName $coll
        }
        
        
        $var_console.AppendText("Successfully imported $comp to SCCM`n")
        $var_console.AppendText("$comp added to $coll`n")
        Write-Host "Added $comp to SCCM, please update memberhip for "$coll"" -Verbose
    }
    catch {
        Write-Error $error[0].Exception.Message
        $var_console.AppendText($error[0].Exception.Message + "`n")
    }

    #Unlocks and clears the UUID and MAC field
    $var_CCMUUIDtxt.IsEnabled = $true
    $var_CCMMACtxt.IsEnabled = $true
    $var_CCMUUIDtxt.Clear()
    $var_CCMMACtxt.Clear()
    $var_console.ScrollToEnd()

}


#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Provided by June Castillote Article
#https://adamtheautomator.com/powershell-gui/

# where is the XAML file?
$xamlFile = "$PSScriptRoot\MainWindow.xaml"

#create window
$inputXML = Get-Content $xamlFile -Raw
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXML

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Warning $_.Exception
    throw
}

# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    }
    catch {
        throw
    }
}
Get-Variable var_*

#Disable buttons at program startup

$CMCollections = $null
$var_recreateADcomp.IsEnabled = $false
$var_createADcomp.IsEnabled = $false
$var_CCMrecreatebtn.IsEnabled = $false
$var_CCMCreateNewbtn.IsEnabled = $false
$var_CMValidate.IsEnabled = $false

#Write to program console
$var_console.AppendText("Program has started`n")

#Importing SCCM module
importccm

#Importing AD moduel
importAD


#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Active Directory Buttons and Actions

#Query provided Computername with AD
$var_adCompQuery.Add_Click( {
        Write-Host "Button adCompQuery has been clicked"
        $var_adDNQueryresult.Text = $Null
        adcomputerexists -Comp $var_txtComputer.Text
    }
)

#Query auto generated or provided DN with AD
$var_adDNQuery.Add_Click( {
        Write-Host "Button adDNQuery has been clicked"
        $var_console.AppendText("Checking if DistinguishedName is valid`n")
        $var_console.ScrollToEnd()
        distinguishednameexists -adpath $var_adDN.Text
    }
)

#Recreate provided AD computer in AD
$var_recreateADcomp.Add_Click( {

        Write-Host "Button recreateADcomp has been clicked"

        #Reruns the DN query function
        $var_console.AppendText("Double-checking dinstinguishedname is valid`n")
        distinguishednameexists -adpath $var_adDN.Text
        
        #If the DN path is not valid, or the computer object does not exist, buttons are disabled and operation cancelled
        #If DN and computer is valid, script runs
        if ($var_adDNQueryresult.Text -eq "Exists") {

            $var_console.AppendText("Double-checking object exists in AD`n")
            adcomputerexists -Comp $var_txtComputer.Text

            if ($var_adCompQueryresult.Text -eq "Exists") {

                $var_console.AppendText("DN still valid`n")
                $var_console.AppendText("Computername still valid`n")
                $var_console.AppendText("Recreating ADObject`n")
                $var_console.AppendText("Computername is " + $var_txtComputer.Text + "`n")
                $var_console.AppendText("Path is " + $var_adDN.Text + "`n")
                $var_console.AppendText("WARNING: Console might freeze during this process, please be patient`n")

                recreateADobject -recreate $true -comp $var_txtComputer.Text -adpath $var_adDN.Text

                $var_console.ScrollToEnd()
            }
            else {
                $var_console.AppendText("Computername does not exist in AD, operation cancelled`n")
                $var_adDNQueryresult.Text = $null
            }
        }
        else {
  
            $var_console.AppendText("DistinguishedName is not valid, operation cancelled`n")
            
        }
    }
)

#Creates a new AD Object with the value provided
$var_createADcomp.Add_Click( {

        Write-Host "Button createADcomp has been clicked"

        $var_console.AppendText("Double-checking dinstinguishedname is valid`n")
        distinguishednameexists -adpath $var_adDN.Text
        
        #If the DN path is not valid, or the computer object does not exist, buttons are disabled and operation cancelled
        #If DN and computer is valid, script runs
        if ($var_adDNQueryresult.Text -eq "Exists") {

            $var_console.AppendText("Double-checking object does not exists in AD`n")
            adcomputerexists -Comp $var_txtComputer.Text -existcheck $true

            if ($var_adCompQueryresult.Text -eq "Does not Exists") {

                $var_console.AppendText("DN still valid`n")
                $var_console.AppendText("Computername still valid`n")
                $var_console.AppendText("Recreating ADObject`n")
                $var_console.AppendText("Computername is " + $var_txtComputer.Text + "`n")
                $var_console.AppendText("Path is " + $var_adDN.Text + "`n")
                $var_console.AppendText("WARNING: Console might freeze during this process, please be patient`n")

                recreateADobject -recreate $false -comp $var_txtComputer.Text -adpath $var_adDN.Text

                $var_console.ScrollToEnd()
            }
            else {
                $var_console.AppendText("Computername exists in AD, operation cancelled`n")
                $var_adDNQueryresult.Text = $null
            }
        }
        else {
  
            $var_console.AppendText("DistinguishedName is not valid, operation cancelled`n")
            
        }
    }
)


#Config Manager Buttons and Actions

#Refreshes the list of collections provided in text file
$var_ColctionScrollrefreshbtn.Add_Click( {
        Write-Host "Button ColctionScrollrefreshbtn has been clicked"
        $var_console.AppendText("Refreshing SCCM deployment list`n")
        $var_console.AppendText("WARNING: This can take up to 1minute, please be patient!`n")
        $var_console.ScrollToEnd()
        ccmcollectionsrefresh
        
    }
)

#Checks whether provided computer name exists in SCCM
$var_CCMCheckbtn.Add_Click( {
        Write-Host "Button CCMCheckbtn has been clicked"
        testcmdevice -Comp $var_txtComputer.Text
    }
)

#Validates wether required values has been provided
$var_CMValidate.Add_Click( {

        Write-Host "Button CMValidate has been clicked"
        $var_console.AppendText("Validating provided values`n")

        #Checks the status of the CCMCheck field, if UUID/MAC has been provided and a collection has been selected, unlocks the re-create button
        if (($var_CCMCheckbtnresult.Text -eq "Exists") -and ($var_CCMUUIDtxt.Text -or $var_CCMMACtxt.Text) -and ($var_ColctionScroll.SelectedValue -ne $null)) {
            $var_CCMrecreatebtn.IsEnabled = $true
            $var_CCMCreateNewbtn.IsEnabled = $false
            $var_console.AppendText("Device exists, UUID/MAC provided, re-create device allowed`n")
        }
        #Checks the status of the CCMCheck field, if UUID/MAC has been provided and a collection has been selected, unlocks the create new button
        elseif (($var_CCMCheckbtnresult.Text -eq "Does not exists") -and ($var_CCMUUIDtxt.Text -or $var_CCMMACtxt.Text) -and ($var_ColctionScroll.SelectedValue -ne $null)) {
            $var_CCMCreateNewbtn.IsEnabled = $true
            $var_CCMrecreatebtn.IsEnabled = $false
            $var_console.AppendText("Device does not exists, UUID/MAC provided, create new device allowed`n")
        }
        #If no values have been provided, both buttons are locked
        else {
            $var_CCMrecreatebtn.IsEnabled = $false
            $var_CCMCreateNewbtn.IsEnabled = $false
            $var_console.AppendText("Required values not provided!`n")
        }
        $var_console.ScrollToEnd()
    }
)

#Recreates the provided computer in SCCM
$var_CCMrecreatebtn.Add_Click( {
        Write-Host "Button CCMrecreatebtn has been clicked"
        recreatecmdevice -Comp $var_txtComputer.Text -recreate $true -UUID $var_CCMUUIDtxt.Text -MAC $var_CCMMACtxt.Text -coll $var_ColctionScroll.SelectedValue
    }
)

#Creates a new object with the provided computer name in SCCM
$var_CCMCreateNewbtn.Add_Click( {
        Write-Host "Button CCMCreateNewbtn has been clicked"
        recreatecmdevice -Comp $var_txtComputer.Text -recreate $false -UUID $var_CCMUUIDtxt.Text -MAC $var_CCMMACtxt.Text -coll $var_ColctionScroll.SelectedValue
    }
)

$Null = $window.ShowDialog()



