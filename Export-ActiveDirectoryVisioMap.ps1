#############################################################################
# Editor  : Kyle Schuler
# Version : 1.3.1
# Modified : 09/09/2024
#
#
# Original Author  : Tyler Cox
# Created : 11/2/2021
#
# Purpose : This script will build an inventory of all GPOs and their links.
#
# Requirements: Computer running the script must have the ActiveDirectory, GroupPolicy, and Visio modules installed.
#               ActiveDirectory and GroupPolicy modules can be installed via the RSAT installer or via Add Windows Features.
#               Visio module can be installed via the PS Gallery, using the command "Install-Module -Name Visio"
#             
# Change Log: Ver 1.0    - Initial release
#             Ver 1.1    - Fixed Visio Cmdlet Parameters,
#                        - Adjusted for Azure AD joined devices
#                        - Fixed issue with importing Visio module
#                        - Reduced output to console
#             Ver 1.2    - Added more error handling and output, refactored, reformatted
#             Ver 1.3    - Added options for user to include or exclude GPOs, and choose the direction of the layout
#             Ver 1.3.1a - (WIP) Refactoring again and code cleanup
#
#############################################################################


#Requires -Modules ActiveDirectory, GroupPolicy, Visio
Clear-Host
Write-Output "Starting up..."

################################################################
# Adjust the following variables to suit your environment
# Set up options
$IncludeGPOs = $true
$LayoutDirection = "TopToBottom"
################################################################





# Get user input
do{
    if($IncludeGPOs){ $IncludeGPOsInput = "Y" }
    else { $IncludeGPOsInput = "N" }
    Write-Output ""
    Write-Output "Would you like to include GPOs in the map? Y/N"
    $IncludeGPOsInput = Read-Host "Type Y or N and press Enter, or press Enter to accept the default of $IncludeGPOsInput"
    if($IncludeGPOsInput.ToUpper() -eq "Y") { $IncludeGPOs = $true; break }
    elseif($IncludeGPOsInput.ToUpper() -eq "N") { $IncludeGPOs = $false; break }
    else { Write-Output "Invalid input. Please try again." }
} while($true)

do{
    if($LayoutDirection -eq "TopToBottom") { $LayoutDirectionInput = "1" }
    else { $LayoutDirectionInput = "2" }
    Write-Output ""
    Write-Output "Choose the direction of the layout:"
    Write-Output "1. Top to Bottom"
    Write-Output "2. Left to Right"
    $LayoutDirectionInput = Read-Host "Type 1 or 2 and press Enter, or press Enter to accept the default of $LayoutDirectionInput"
    if($LayoutDirectionInput -eq "1") { $LayoutDirection = "TopToBottom"; break }
    elseif($LayoutDirectionInput -eq "2") { $LayoutDirection = "LeftToRight"; break }
    else { Write-Output "Invalid input. Please try again." }
} while($true)


try {
    Write-Output "Creating the Visio Document"
    #Create the Visio Application
    New-VisioApplication
    #Create the Visio Document
    $VisioDoc = New-VisioDocument
    #Create the Visio Page
    $Page = $VisioDoc.Pages[1]
    #Create the Visio Point at 1,1
    $Point_1_1 = New-VisioPoint -X 1.0 -Y 1.0
}
catch {
    Write-Error "Error creating the Visio document or page $_"
    Read-Host "Press any key to exit"
    exit
}

#Set our counters
$nodeCount = 0
$conCount = 0
$gpoCount = 0

#Get our root domain from the current logged on user
$DNSDomain = $env:USERDNSDOMAIN 

if($null -eq $DNSDomain) {
    Write-Warning "Unable to get the DNS Domain. Please ensure you are logged in to a domain joined computer, or have your default domain set"
    Read-Host "Press any key to continue"
}

Write-Output "Getting the OUs from the domain $DNSDomain"
#Get all OUs except LostAndFound
try {
    $OUs = Get-ADOrganizationalUnit -Server $DNSDomain -Filter 'Name -like "*"' -Properties Name, DistinguishedName, CanonicalName, LinkedGroupPolicyObjects | `
            Where-Object {$_.canonicalname -notlike "*LostandFound*"} | Select-Object Name, Canonicalname, DistinguishedName, LinkedGroupPolicyObjects | `
            Sort-Object CanonicalName # | Select -First 50
}
catch {
    Write-Error "Error getting the OUs from the domain $DNSDomain $_"
    Read-Host "Press any key to exit"
    exit
}

try {
    #Gather our shapes from Visio's stencils
    $ADO_u = Open-VisioDocument "ADO_U.vss"
    $connectors = Open-VisioDocument "Connectors.vss"
    $masterOU = Get-VisioMaster "Organizational Unit" -Document $ADO_u
    $connector = Get-VisioMaster "Dynamic Connector" -Document $Connectors
    $masterDomain = Get-VisioMaster "Domain" -Document $ADO_u
    $masterGPO = Get-VisioMaster "Policy" -Document $ADO_u
}
catch {
    Write-Error "Error getting the Visio shapes $_"
    Read-Host "Press any key to exit"
    exit
}

try {
    #Create our first shape. This is the root domain node
    $n0 = New-VisioShape -Master $MasterDomain -Position $Point_1_1
    #Set shape properties
    $n0.Text = $DNSDomain
    $n0.Name = "n" + $DNSDomain
}
catch {
    Write-Error "Error creating the root domain shape $_"
    Read-Host "Press any key to exit"
    exit
}

if ($IncludeGPOs) {
    Write-Output "Getting the GPOs linked to the root domain $DNSDomain"
    #Get Root Domain linked GPOs and process them accordingly
    try {
        $RootGPOs = Get-ADObject -Server $DNSDomain -Identity (Get-ADDomain -Identity $DNSDomain).distinguishedName -Properties name, distinguishedName, gPLink, gPOptions
        
    }
    catch {
        Write-Error "Error getting the GPOs linked to the root domain $DNSDomain $_"
        Read-Host "Press any key to exit"
        exit
    }#Loop through each root GPO
    Write-Output "Creating the GPO shapes and connecting them to the root domain"
    ForEach ($gpolink in $RootGPOs.gPlink -split "\]\[") {
        #Add to our counters (for naming)
        $gpoCount += 1 
        $conCount += 1 
        #get only the GUID of the gpo
        $gpoGUID = ([Regex]::Match($gpoLink, '{[a-zA-Z0-9]{8}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{12}}')).Value 
        #pull details for the GPO based on the GUID
        try {
            $gpo = Get-GPO -Guid $gpoGUID -Domain $DNSDomain
        }
        catch {
            Write-Warning "Error getting the GPO with GUID $gpoGUID $_"
            Write-Warning "Skipping this GPO"
            Continue
        }
        #declare what we'll call the gpo shape 
        $shapename = "g" + $gpoCount 
        #Create the GPO shape
        $shapeGPO = New-VisioShape -Master $MasterGPO -Position $Point_1_1 
        #Set the shape properties
        $ShapeGPO.Text = $GPO.DisplayName
        $ShapeGPO.Name = $shapename
        #Set the shape's custom properties
        $GUID = "{" + $gpo.id.guid + "}"
        If ($GPO.DisplayName) {
            Set-VisioCustomProperty -Shape $ShapeGPO -Name "GPOName" -Value $GPO.DisplayName
        }
        If ($GPO.Description) {
            Set-VisioCustomProperty -Shape $ShapeGPO -Name "Description" -Value $GPO.Description
        }
        If ($GPO.ID.Guid) {
            Set-VisioCustomProperty -Shape $shapeGPO -Name "GUID" -Value $GUID
        }
        If ($GPO.GPOStatus) {
            Set-VisioCustomProperty -Shape $shapeGPO -Name "Status" -Value $GPO.GpoStatus.ToString()
        }
        If ($GPO.CreationTime) {
            Set-VisioCustomProperty -Shape $shapeGPO -Name "CreationTime" -Value $GPO.CreationTime.ToString()
        }
        If ($GPO.ModificationTime) {
            Set-VisioCustomProperty -Shape $shapeGPO -Name "ModifiedTime" -Value $GPO.ModificationTime.ToString()
        }
        If ($GPO.WmiFilter) {
            Set-VisioCustomProperty -Shape $shapeGPO -Name "WMIFilterName" -Value $GPO.WMIFilter.Name
        }
        #Create the shape's connections
        $con = Connect-VisioShape -From $n0 -To $shapeGPO -Master $connector 
        #Set the connections custom properties
        $con.text = "GPO"
        $con.name = "gcon" + $conCount #We name it like this so that later we can identify all GPO connections for formatting of the connector's text
        $con_cells = New-VisioShapeCells
        $con_cells.LineColor = "rgb(0,175,240)"
        $con_cells.LineEndArrowSize = "3"
        $con_cells.LineBeginArrowSize = "2"
        $con_cells.LineEndArrow = "42"
        $con_cells.LineBeginArrow = "4"
        $con_cells.CharColor = "rgb(0,175,240)"
        #Set the shape properties
        Set-VisioShapeCells -Cells $con_cells -Shape $con     
    }
}



Write-Output "Creating the OU shapes and connecting them to the root domain"
#Loop through each OU
ForEach ($ou in $OUs) {
    #Add to our counters
    $nodeCount += 1
    $conCount += 1
     
    #Massage the OU details to get the name
    $OUName = $OU.Name
    #Massage the OU details to get the Canonical name. We use this to get the previous OU name
    $OUConName = $OU.Canonicalname
    $nameSplit = $ou.CanonicalName -split '(?<!\\)/'
    $nameRecombined = $nameSplit[0..($nameSplit.length - 2)] -join "/"
    #If the previous OU name is the root domain..
    If ($nameSplit[$index - 2] -eq $DNSDomain) {
        #declare what we'll call the shape
        $shapename = "n" + $OUConName
        #Create the new shape
        $shape = New-VisioShape -Master $MasterOU -Position $Point_1_1
        #Set the shape details
        $Shape.Text = $OUName
        $Shape.Name = $shapename
                
        #Set custom properties of the shape
        Set-VisioCustomProperty -Shape $shape -Name "OU_Name" -Value $OU.Name
        Set-VisioCustomProperty -Shape $shape -Name "DistinguishedName" -Value $OU.DistinguishedName
        if ($IncludeGPOs) {
            Set-VisioCustomProperty -Shape $shape -Name "Linked_GPOs" -Value $OU.LinkedGroupPolicyObjects.Count
        }
                
        #Connect the shape to the root domain shape
        Connect-VisioShape -From $n0 -To $shape -Master $connector | Out-Null

    }
    #If it's not the root domain, then do this..
    else {
        #Set the name of the previous shape 
        $prevOUName = "n" + $nameRecombined

        #Get the previous shape from Visio based on the name
        $prevOUshape = Get-VisioShape -Name * | Where-Object {$_.Nameu -eq $prevOUName}

        #Set the name of the new shape
        $shapename = "n" + $OUConName
        #Create the new shape
        $shape = New-VisioShape -Master $MasterOU -Position $Point_1_1
        #Set the shape properties
        $Shape.Text = $OUName
        $Shape.Name = $shapename

        #Set custom properties of the shape
        Set-VisioCustomProperty -Shape $shape -Name "OU_Name" -Value $OU.Name
        Set-VisioCustomProperty -Shape $shape -Name "DistinguishedName" -Value $OU.DistinguishedName
        if ($IncludeGPOs) {
            Set-VisioCustomProperty -Shape $shape -Name "Linked_GPOs" -Value $OU.LinkedGroupPolicyObjects.Count
        }

        #Connect the shape to the previous shape
        Connect-VisioShape -From $prevOUshape -To $shape -Master $connector | Out-Null
    }

    #If the OU had linked GPOs..
    If ($OU.LinkedGroupPolicyObjects -and $IncludeGPOs) {
        #Loop through each GPO
        Foreach ($gpoLink in $OU.LinkedGroupPolicyObjects) {
            #increase our counters
            $gpoCount += 1
            $conCount += 1
            #get only the GUID of the gpo
            $gpoGUID = ([Regex]::Match($gpoLink, '{[a-zA-Z0-9]{8}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{12}}')).Value
            #Create the GPO shape
            try {
                $gpo = Get-GPO -Guid $gpoGUID -Domain $DNSDomain
            }
            catch {
                Write-Warning "Error getting the GPO with GUID $gpoGUID $_"
                Write-Warning "Skipping this GPO"
                Continue
            }

            #declare what we'll call the gpo shape 
            $shapename = "g" + $gpoCount
            #Create the GPO shape
            $shapeGPO = New-VisioShape -Master $MasterGPO -Position $Point_1_1
            #Set the shape properties
            $ShapeGPO.Text = $GPO.DisplayName
            $ShapeGPO.Name = $shapename
            $GUID = "{" + $gpo.id.guid + "}"
            If ($GPO.DisplayName) {
                Set-VisioCustomProperty -Shape $ShapeGPO -Name "GPOName" -Value $GPO.DisplayName
            }
            If ($GPO.Description) {
                Set-VisioCustomProperty -Shape $ShapeGPO -Name "Description" -Value $GPO.Description
            }
            If ($GPO.ID.Guid) {
                Set-VisioCustomProperty -Shape $shapeGPO -Name "GUID" -Value $GUID
            }
            If ($GPO.GPOStatus) {
                Set-VisioCustomProperty -Shape $shapeGPO -Name "Status" -Value $GPO.GpoStatus.ToString()
            }
            If ($GPO.CreationTime) {
                Set-VisioCustomProperty -Shape $shapeGPO -Name "CreationTime" -Value $GPO.CreationTime.ToString()
            }
            If ($GPO.ModificationTime) {
                Set-VisioCustomProperty -Shape $shapeGPO -Name "ModifiedTime" -Value $GPO.ModificationTime.ToString()
            }
            If ($GPO.WmiFilter) {
                Set-VisioCustomProperty -Shape $shapeGPO -Name "WMIFilterName" -Value $GPO.WMIFilter.Name
            }

            #Create the shape's connections
            $con = Connect-VisioShape -From $shape -To $shapeGPO -Master $connector
            $con.text = "GPO"
            $con.Name = "gcon" + $conCount #We name it like this so that later we can identify all GPO connections for formatting of the connector's text
            $con_cells = New-VisioShapeCells
            $con_cells.LineColor = "rgb(0,175,240)"
            $con_cells.LineEndArrowSize = "3"
            $con_cells.LineBeginArrowSize = "2"
            $con_cells.LineEndArrow = "42"
            $con_cells.LineBeginArrow = "4"
            $con_cells.CharColor = "rgb(0,175,240)"
            #Set the shape properties
            Set-VisioShapeCells -Cells $con_cells -Shape $con                      
        }
    }
}


try {
    Write-Output "Formatting the Visio Page"
    #Create a new layout object
    $ls = New-Object VisioAutomation.Models.LayoutStyles.hierarchyLayoutStyle
    #set object properties (this is how we format the page)
    $ls.AvenueSizeX = 1
    $ls.AvenueSizeY = 1
    $ls.LayoutDirection = $LayoutDirection
    $ls.ConnectorStyle = "Simple"
    $ls.ConnectorAppearance = "Straight"
    $ls.horizontalAlignment = "Left"
    $ls.verticalAlignment = "Top"
    
    #Apply the layout object to the page
    Format-VisioPage -LayoutStyle $ls 
    
    #Change the page's size to match the new data
    Format-VisioPage -FitContents -BorderWidth 1.0 -BorderHeight 1.0
    
    #This section is to set text for the GPO shapes based on the length of the line. We had to move the shapes around first before we could run this part.
    #Create a new Shape Cell Object
    $con_cells = New-VisioShapeCells
    #Set the location of the text based on the length of the line
    $con_cells.TextFormPinX = "=POINTALONGPATH(Geometry1.Path,1)"
    $con_cells.TextFormPinY = "=POINTALONGPATH(Geometry1.Path,.75)"
    #Get all gpo connections
    $gpoShapes = Get-VisioShape -Name * | Where-Object {$_.Nameu -like "gcon*"}
    #Loop through each connection
    ForEach($shape in $gpoShapes) {
        #Set the shape from the shape cell object
        Set-VisioShapeCells -Cells $con_cells -Shape $shape   
    }
    Write-Output "Visio Page formatted"
    Write-Output "Visio Document created"
}
catch {
    Write-Warning "Error formatting the Visio page $_"
    Write-Output "Unless there were errors, the Visio document should be created, but may not be formatted correctly"
}

# Powershell garbage collection
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()