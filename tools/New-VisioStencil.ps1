#
# Import SVG files as Visio stencils
# New-VisioStencil (Get-ChildItem "*.svg" -Recurse) -StencilPath "..\CiscoSafe.vssx"
#
using namespace Microsoft.Office.Interop.Visio

function Initialize-VisioStencil {    # Stupid Americanisation Get-Verb
    [CmdletBinding()]
    param(
        [string]$StencilType = "vssx"    # Visio stencil file format, default vssx
    )

    try {
        $visioApplication = New-Object -ComObject Visio.Application    # Create new instance of Visio Application
    }
    catch {
        Write-Error "Could not create Visio Application instance, Visio is installed?"
        throw "Initialisation Error: Could not create Visio Application instance."
    }

    try {
        $visioStencil = $visioApplication.Documents.Add($StencilType)    # Add new stencil document
    }
    catch {
        Write-Error "Failed to add new stencil."
        $visioApplication.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visioApplication) | Out-Null
        Remove-Variable visioApplication
        throw "Initialisation Error: Failed to add new stencil."
    }

    return $visioApplication, $visioStencil
}

function Close-ComObject($comObject) {
    if ($null -ne $comObject -and $comObject -is [System.__ComObject]) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
    }
}

function Import-SvgFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [System.IO.FileInfo]$svgFile,

        [Parameter(Mandatory = $true)]
        [ref]$VisioMaster
    )

    try {
        $masterName = $svgFile.BaseName
        $addMaster = $VisioMaster.Value.Add()
        $addMaster.Name = $masterName
        $importShape = $addMaster.Import($svgFile.FullName)

        Set-ShapeSize -Shape $importShape

        try {
            Set-ShapeData -Shape $importShape
        }
        catch {
            Write-Warning "Failed to set properties for shape: $masterName from file: $($svgFile.FullName) - $_"
        }
    }
    catch {
        throw "Failed to import SVG file: $($svgFile.FullName) - $_"
    }
}

function Set-ShapeSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $Shape
    )

    $scaleFactor = 0.5    # Define constant for the scale factor to adjust shape size
    $oldWidth = $Shape.CellsU("Width").ResultIU    # Retrieve the current width of shape
    $Shape.CellsU("Width").ResultIU = $scaleFactor    # Scale the width of shape by the defined scale factor
    $Shape.CellsU("Height").ResultIU *= $scaleFactor / $oldWidth    # Adjust the height to maintain the aspect ratio based on the new width
    $Shape.CellsSrc([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowGroup, [VisCellIndices]::visGroupSelectMode).FormulaU = "0"    # Set the group select mode to none (0) to prevent the shape from being selected as a group
}

function Set-ShapeData {
    param(
        $Shape
    )

    # https://learn.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa200961(v=office.10)

    # Add connection points
    $Shape.AddSection([VisSectionIndices]::visSectionConnectionPts) | Out-Null

    # Position for connection points
    $leftRightPositions = @('0.25', '0.5', '0.75')
    $topBottomPositions = @('0.25', '0.5', '0.75')

    # Add connection points left and right sides
    foreach ($pos in $leftRightPositions) {
        # Left
        $Shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
        $rowIndex = $Shape.RowCount([VisSectionIndices]::visSectionConnectionPts) - 1
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctX).FormulaU = "0"
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctY).FormulaU = "$pos*Height"

        # Right
        $Shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
        $rowIndex = $Shape.RowCount([VisSectionIndices]::visSectionConnectionPts) - 1
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctX).FormulaU = "Width"
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctY).FormulaU = "$pos*Height"
    }

    # Add connection points top and bottom sides
    foreach ($pos in $topBottomPositions) {
        # Top
        $Shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
        $rowIndex = $Shape.RowCount([VisSectionIndices]::visSectionConnectionPts) - 1
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctX).FormulaU = "$pos*Width"
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctY).FormulaU = "Height"

        # Bottom
        $Shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
        $rowIndex = $Shape.RowCount([VisSectionIndices]::visSectionConnectionPts) - 1
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctX).FormulaU = "$pos*Width"
        $Shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, $rowIndex, [VisCellIndices]::visCnnctY).FormulaU = "0"
    }

    # Make dynamic connector connect to the bounding box instead of the shape's geometry
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowShapeLayout, [VisCellIndices]::visSLOConFixedCode).FormulaU = [VisCellVals]::visSLOFixedNoFoldToShape

    # Add a control point centered under the shape to allow control of shape's text position
    $Shape.AddSection([VisSectionIndices]::visSectionControls) | Out-Null
    $Shape.AddRow([VisSectionIndices]::visSectionControls, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlX).FormulaU = "Width*0.5"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlY).FormulaU = "-0.5*(ABS(SIN(Angle))*TxtWidth+ABS(COS(Angle))*TxtHeight)"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlXDyn).FormulaU = "Width*0.5"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlYDyn).FormulaU = "Height*0.5"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlXCon).FormulaU = "(Controls.Row_1>Width*0.5)*2+2+IF(OR(HideText,STRSAME(SHAPETEXT(TheText),`"`")),5,0)"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlYCon).FormulaU = "(Controls.Row_1.Y>Height*0.5)*2+2"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlGlue).FormulaU = "TRUE"
    $Shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlTip).FormulaU = "`"Reposition text`""

    # Set shapes's text font and color
    $Shape.CellsSRC([VisSectionIndices]::visSectionCharacter, 0, [VisCellIndices]::visCharacterFont).FormulaU = "FONT(""Tahoma"")"
    $Shape.CellsSRC([VisSectionIndices]::visSectionCharacter, 0, [VisCellIndices]::visCharacterSize).FormulaU = "8 pt"
    $Shape.CellsSRC([VisSectionIndices]::visSectionCharacter, 0, [VisCellIndices]::visCharacterColor).FormulaU = "IF(LUM(THEMEVAL())>200,0,THEMEVAL(`"TextColor`",0))"

    # Set shape's text transformation properties
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormWidth).FormulaU = "IF(TextDirection=0,TEXTWIDTH(TheText),TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormHeight).FormulaU = "IF(TextDirection=1,TEXTWIDTH(TheText),TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormAngle).FormulaU = "IF(BITXOR(FlipX,FlipY),1,-1)*Angle"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormPinX).FormulaU = "Controls.Row_1"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormPinY).FormulaU = "Controls.Row_1.Y"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormLocPinX).FormulaU = "TxtWidth*0.5"
    $Shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormLocPinY).FormulaU = "TxtHeight*0.5"
}

function New-VisioStencil {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [ValidateScript({
            if (-Not (Test-Path $_ -PathType Leaf)) {
                throw "Path $_ does not exist or is not a file."
            }
            if ($_ -notmatch '\.svg$') {
                throw "File $_ is not an SVG file."
            }
            return $true
        })]
        [string[]]$svgPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$stencilPath
    )

    Begin {
        try {
            $visioApplication, $visioStencil = Initialize-VisioStencil
            $visioMaster = $visioStencil.Masters
            $i = 0
        }
        catch {
            Write-Error "Failed to initialise Visio: $_"
            return
        }
    }

    Process {
        foreach ($svgfile in $svgPath) {
            if ($svgPath.Count -gt 1) {
                Write-Progress -ParentId 1 -Id 100 -PercentComplete ($i / $svgPath.Count * 100) -Activity "Creating Visio stencil" -Status "Importing SVG $($i + 1) of $($svgPath.Count)"
            }

            try {
                Import-SvgFile -svgFile $svgfile -VisioMaster ([ref]$visioMaster)
            }
            catch {
                Write-Warning "Failed to process SVG file '$svgfile': $_"
            }

            $i++
        }
    }

    End {
        if ($svgPath.Count -gt 1) { Write-Progress -Id 100 -Activity "Importing SVG" -Completed }

        try {
            $resolvedStencilPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($stencilPath)
            $visioStencil.SaveAs($resolvedStencilPath) | Out-Null
            Write-Host "Stencil saved: '$resolvedStencilPath'."
        }
        catch {
            Write-Error "Failed to save stencil: $_"
        }
        finally {
            $visioStencil.Close()
            $visioApplication.Quit()

            Close-ComObject $visioStencil
            Close-ComObject $visioApplication

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
