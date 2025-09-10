#######################################################################
#
# Powershell functions for operate Visio Drawing
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  14.09.2007
# Purpose/Change: Initial script development.
# Create function New-VisioApplication.
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    14.09.2007
# Purpose/Change: Create function New-VisioDocument.
#
# Version:        0.3
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    14.09.2007
# Purpose/Change: Create function Set-VisioPage.
#
# Version:        0.4
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    15.09.2007
# Purpose/Change: Create function Add-VisioStensil.
#
# Version:        0.5
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    15.09.2007
# Purpose/Change: Create function Set-VisioStensilMasterItem.
#
# Version:        0.6
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    15.09.2007
# Purpose/Change: Create function Draw-VisioItem.
#
# Version:        0.7
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    15.09.2007
# Purpose/Change: Create function Draw-VisioLine.
#
# Version:        0.8
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  16.09.2007
# Purpose/Change: Change function Draw-VisioItem.
#
# Version:        0.9
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    16.09.2007
# Purpose/Change: Create function Draw-VisioIcon.
#
# Version:        0.10
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Change function Draw-VisioIcon.
#
# Version:        0.11
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Change function Draw-VisioLine.
#
# Version:        0.12
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Change function Set-VisioStensilMasterItem.
#
# Version:        0.13
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Change function Draw-VisioItem.
#
# Version:        0.14
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Add function Resize-VisioPageToFitContents.
#
# Version:        0.15
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    17.09.2007
# Purpose/Change: Add function Save-VisioDocument.
#
#
# Version:        0.16
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    18.09.2007
# Purpose/Change: Add function Draw-VisioText.
#
# Version:        0.17
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    18.09.2007
# Purpose/Change: Add function Draw-VisioPolyLine.
#
# Version:        0.18
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Change Date:    19.09.2007
# Purpose/Change: Change function Draw-VisioText.
#
# Run:
#   
# . .\PSVisio.ps1
#
#######################################################################

# Set Variables
$Shape=0
$Line=0
$Icon=0
$Text=0
$PolyLine=0

Function New-VisioApplication {

#######################################################################
#
# Powershell function for create Visio Application
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  14.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# New-VisioApplication
#
#######################################################################

# Create Visio Object
$Script:Application = New-Object -ComObject Visio.Application
$Script:Application.Visible = $True
}

Function New-VisioDocument {

#######################################################################
#
# Powershell function for create Visio Document
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  14.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# New-VisioDocument
#
#######################################################################

# Create Document from Blank Template
$Script:Documents = $Script:Application.Documents
$Script:Document=$Script:Application.Documents.Add('')
}

Function Set-VisioPage {

#######################################################################
#
# Powershell function for Set Active Visio Document Page
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  14.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Set-VisioPage
#
#######################################################################

# Set Visio Active Page
$Script:Page=$Script:Application.ActivePage
$Script:Application.ActivePage.PageSheet
}

Function Add-VisioStensil ($Name, $File) {

#######################################################################
#
# Powershell function for Add Visio Stensil
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  15.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Add-VisioStensil -Name "Basic" -File "BASIC_M.vss"
#
#######################################################################

# Set Expression and Add Visio Stensil
$Expression = '$Script:' + $Name + ' = $Script:Application.Documents.Add("' + $File +'")'
Invoke-Expression $Expression
}

Function Set-VisioStensilMasterItem ($Stensil, $Item) {

#######################################################################
#
# Powershell function for Set Stensil Master Item
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  15.09.2007
# Purpose/Change: Initial script development
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Reorganize Variables
#
# Run:
#   
# Set-VisioStensilMasterItem -Stensil "Basic" -Item "Rectangle"
#
#######################################################################

# Set Expression And Set Masters Item Rectangle
$ItemWithoutSpace = $Item -replace " ",""
$Expression = '$Script:' + $ItemWithoutSpace + ' = $Script:' + $Stensil + '.Masters.Item("' + $Item + '")'
Invoke-Expression $Expression
}

Function Draw-VisioItem ($Master, $X, $Y, $Width, $Height, $FillForegnd, $LinePattern, $Text, $VerticalAlign, $ParaHorzAlign, $CharSize, $CharColor) {

#######################################################################
#
# Powershell function for Draw Visio Item
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  15.09.2007
# Purpose/Change: Initial script development
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  16.09.2007
# Purpose/Change: Add Flow Control Input Parameters
#
# Version:        0.3
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Reorganize Variables
#
# Run:
#   
# Draw-VisioItem -Master "Rectangle" -X 6.375 -Y 7.125 -Width 12.2501 -Height 7.25 -FillForegnd "RGB(0,153,204)"`
# -LinePattern 0 -Text "Microsoft Virtual Machine Manager Architecture" -VerticalAlign 0 -ParaHorzAlign 0`
# -CharSize "20 pt" -CharColor "RGB(255,255,255)"
#
#######################################################################

$Script:Shape++
$Master = $Master -replace " ",""

# Set Expression And Draw Item
$Expression = '$Script:Shape' + $Script:Shape + ' = $Script:Page.Drop(' + '$' + $Master + ',' + $X + ',' + $Y + ')'
Invoke-Expression $Expression

# Set Item Width Properties
If ($Width)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Width").Formula = ' + $Width
		Invoke-Expression $Expression
	}

# Set Item Height Properties
If ($Height)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Height").Formula = ' + $Height
		Invoke-Expression $Expression
	}

# Set Item FillForegnd Properties
If ($FillForegnd)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("FillForegnd").Formula = "=' +  $FillForegnd + '"'
		Invoke-Expression $Expression
	}

# Set Item LinePattern Properties
If ($LinePattern)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("LinePattern").Formula = ' + $LinePattern
		Invoke-Expression $Expression
	}

# Set Item Text
If ($Text)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Item VerticalAlign Properties
If ($VerticalAlign)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("VerticalAlign").Formula = ' + $VerticalAlign
		Invoke-Expression $Expression
	}

# Set Item HorzAlign Properties
If ($ParaHorzAlign)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Para.HorzAlign").Formula = ' + $ParaHorzAlign
		Invoke-Expression $Expression
	}

# Set Item Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}

# Set Item Char.Color Properties
If ($CharColor)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Char.Color").Formula = "=' +  $CharColor + '"'
		Invoke-Expression $Expression
	}
}

Function Draw-VisioLine ($BeginX, $BeginY, $EndX, $EndY, $LineWeight, $LineColor, $BeginArrow, $EndArrow) {

#######################################################################
#
# Powershell function for Draw Visio Line
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  15.09.2007
# Purpose/Change: Initial script development
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Add Flow Control Input Parameters
#
# Run:
#   
# Draw-VisioLine -BeginX 0.3125 -BeginY 10.3438 -EndX 12.4948 -EndY 10.3438 -LineWeight "1 pt"`
# -LineColor "RGB(255,255,255)" -BeginArrow 4 -EndArrow 4
#
#######################################################################

$Script:Line++

# Set Expression And Draw Line
$Expression = '$Script:Line' + $Script:Line + ' = $Script:Page.DrawLine(' + $BeginX + ',' + $BeginY + ',' + $EndX + ',' + $EndY + ')'
Invoke-Expression $Expression

# Set Line Width Properties
If ($LineWeight)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("LineWeight").Formula = "' + $LineWeight + '"'
		Invoke-Expression $Expression
	}

# Set Line Color Properties
$Expression = '$Script:Line' + $Script:Line + '.Cells("LineColor").Formula = "=' +  $LineColor + '"'
Invoke-Expression $Expression

# Set Line Begin Arrow Properties
If ($BeginArrow)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("BeginArrow").Formula = ' + $BeginArrow
		Invoke-Expression $Expression
	}

# Set Line End Arrow Properties
If ($EndArrow)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("EndArrow").Formula = ' + $EndArrow
		Invoke-Expression $Expression
	}
}

Function Draw-VisioIcon ($IconPath, $Width, $Height, $PinX, $PinY, $Text, $CharSize) {

#######################################################################
#
# Powershell function for Draw Visio Icon
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  16.09.2007
# Purpose/Change: Initial script development
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Add Flow Control Input Parameters
#
# Run:
#   
# Draw-VisioIcon -IconPath "c:\!\powershell.png" -Width 0.9843 -Height 0.9843 -PinX 2.5547 -PinY 9.2682`
# -Text "Windows Powershell" -CharSize "10 pt"
#
#######################################################################

$Script:Icon++

# Import Icon Item
$Expression = '$Script:Icon' + $Script:Icon + ' = $Script:Page.Import("' + $IconPath + '")'
Invoke-Expression $Expression

# Set Icon Width Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Width").Formula = ' + $Width
Invoke-Expression $Expression

# Set Icon Height Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Height").Formula = ' + $Height
Invoke-Expression $Expression

# Set Icon PinX Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("PinX").Formula = ' + $PinX
Invoke-Expression $Expression

# Set Icon PinY Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("PinY").Formula = ' + $PinY
Invoke-Expression $Expression

# Set Icon Text
If ($Text)
	{
		$Expression = '$Script:Icon' + $Script:Icon + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Icon Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}
}

Function Resize-VisioPageToFitContents {

#######################################################################
#
# Powershell function for Resize Active Visio Document Page to Fit Contents
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Resize-VisioPageToFitContents
#
#######################################################################

# Resize Page to Fit Contents
$Script:Page.ResizeToFitContents()
}

Function Save-VisioDocument ($File) {

#######################################################################
#
# Powershell function for save Visio Document
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Save-VisioDocument -File 'C:\!\MsSCVMM2007Arch.vsd'
#
#######################################################################

# Save Document
$Expression = '$Script:Document.SaveAs("' + $File + '")'
Invoke-Expression $Expression
}

Function Close-VisioApplication {

#######################################################################
#
# Powershell function for create Visio Application
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  17.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Close-VisioApplication
#
#######################################################################

# Close Visio Application
$Script:Application.Quit()
}

Function Draw-VisioText ($X, $Y, $Width, $Height, $FillForegnd, $LinePattern, $Text, $VerticalAlign, $ParaHorzAlign, $CharSize, $CharColor, $CharStyle, $FillForegndTrans) {

#######################################################################
#
# Powershell function for Draw Visio Text Label
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  18.09.2007
# Purpose/Change: Initial script development
#
# Version:        0.2
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  19.09.2007
# Purpose/Change: Add CharStyle Parameter
#
# Run:
#   
# Draw-VisioText -X 4.25 -Y 8.875 -Width 1.3751 -Height 0.375 -Text "Microsoft Virtual Machine Manager Architecture"`
# -LinePattern "0" -FillForegndTrans "100%" -CharStyle 17
#
#######################################################################

$Script:Text++
$Master = "Rectangle"

# Set Expression And Draw Text
$Expression = '$Script:Text' + $Script:Text + ' = $Script:Page.Drop(' + '$' + $Master + ',' + $X + ',' + $Y + ')'
Invoke-Expression $Expression

# Set Item Width Properties
If ($Width)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Width").Formula = ' + $Width
		Invoke-Expression $Expression
	}

# Set Item Height Properties
If ($Height)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Height").Formula = ' + $Height
		Invoke-Expression $Expression
	}

# Set Item FillForegnd Properties
If ($FillForegnd)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("FillForegnd").Formula = "=' +  $FillForegnd + '"'
		Invoke-Expression $Expression
	}

# Set Item LinePattern Properties
If ($LinePattern)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("LinePattern").Formula = ' + $LinePattern
		Invoke-Expression $Expression
	}

# Set Item Text
If ($Text)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Item VerticalAlign Properties
If ($VerticalAlign)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("VerticalAlign").Formula = ' + $VerticalAlign
		Invoke-Expression $Expression
	}

# Set Item HorzAlign Properties
If ($ParaHorzAlign)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Para.HorzAlign").Formula = ' + $ParaHorzAlign
		Invoke-Expression $Expression
	}

# Set Item Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}

# Set Item Char.Color Properties
If ($CharColor)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Color").Formula = "=' +  $CharColor + '"'
		Invoke-Expression $Expression
	}
	
# Set Item Char.Style Properties
If ($CharStyle)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Style").Formula = "' + $CharStyle + '"'
		Invoke-Expression $Expression
	}
	
# Set Item FillForegndTrans Properties
If ($FillForegndTrans)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("FillForegndTrans").Formula = "' + $FillForegndTrans + '"'
		Invoke-Expression $Expression
	}		
}

Function Draw-VisioPolyLine ($PolyLine, $LineWeight, $LineColor, $BeginArrow, $EndArrow) {

#######################################################################
#
# Powershell function for Draw Visio PolyLine
#  
# Version:        0.1
# Author:         Andrii Romanenko
# Website:        blogs.airra.net
# Creation Date:  18.09.2007
# Purpose/Change: Initial script development
#
# Run:
#   
# Draw-VisioPolyLine -Polyline 1.0938,9.0625,1.4063,9.0625,1.4063,8.6563,1.0938,8.6563 -LineWeight "0.5 pt"`
# -LineColor "RGB(255,255,255)" -BeginArrow 1 -EndArrow 1
#
#######################################################################

$Script:PolyLine++
[double[]]$PolyLineCoordinates=@()
$PolyLineCoordinates += $Polyline

# Set Expression And Draw PolyLine
$Expression = '$Script:PolyLine' + $Script:PolyLine + ' = $Script:Page.DrawPolyLine(($PolyLineCoordinates),0)'
Invoke-Expression $Expression

# Set Line Width Properties
If ($LineWeight)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("LineWeight").Formula = "' + $LineWeight + '"'
		Invoke-Expression $Expression
	}

# Set Line Color Properties
$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("LineColor").Formula = "=' +  $LineColor + '"'
Invoke-Expression $Expression

# Set Line Begin Arrow Properties
If ($BeginArrow)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("BeginArrow").Formula = ' + $BeginArrow
		Invoke-Expression $Expression
	}

# Set Line End Arrow Properties
If ($EndArrow)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("EndArrow").Formula = ' + $EndArrow
		Invoke-Expression $Expression
	}
}
