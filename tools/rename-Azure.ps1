#
# Rename Azure SVG files
#

$directory = "."
$svgFiles = Get-ChildItem -Path $directory -Recurse -Filter "*.svg"

foreach ($file in $svgFiles) {
    $newName = $file.Name -replace '^\d+-icon-service-', '' -replace '[-_]', ' '
    Rename-Item $file.FullName -NewName $newName
}
