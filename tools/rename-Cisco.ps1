#
# Rename Cisco SAFE SVG files
#

$directory = "."
$svgFiles = Get-ChildItem -Path $directory -Recurse -Filter "*.svg" -File

foreach ($file in $svgFiles) {
    $newName = $file.Name -replace '[\w\s]+_\d+_', '' -replace '[-_]', ' '
    Rename-Item -Path $file.FullName -NewName $newName
}
