#
# Convert SVG to plain SVG. Use inkscape 0.48.5
#

$directory = "."
$inkscapePath = 'C:\Program Files (x86)\Inkscape\inkscape.exe'
$exportDirectory = Join-Path $directory "export"

if (-not (Test-Path -Path $exportDirectory)) {
    New-Item -ItemType Directory -Path $exportDirectory | Out-Null
}

$svgFiles = Get-ChildItem -Path $directory -Recurse -Filter "*.svg"

foreach ($file in $svgFiles) {
    $exportPath = Join-Path $exportDirectory $file.Name
    & $inkscapePath --export-plain-svg=$exportPath $file.FullName
}
