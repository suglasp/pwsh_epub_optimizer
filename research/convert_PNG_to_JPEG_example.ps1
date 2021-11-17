
#
# Pieter De Ridder
# Example how to convert a png to a jpg image format.
# https://github.com/suglasp/pwsh_epub_optimizer.git
#
# Created : 16/11/2021
# Updated : 16/11/2021
#

Add-Type -AssemblyName System.Drawing
#Add-Type -AssemblyName System.Drawing.Imaging

[System.Collections.ArrayList]$PNGFiles = @(Get-ChildItem -Path .\ -Filter *.png -Recurse)

If ($PNGFiles.Count -gt 0) {
    ForEach($PNGFile In $PNGFiles) {
        If (Test-Path -Path $PNGFile.FullName) {
            # Open PNG and save as a JP(E)G
            Write-Host "Converting PNG2JPG : $($PNGFile)..."
            [System.Drawing.Image]$png = [System.Drawing.Image]::FromFile($PNGFile.FullName)
            $png.Save($PNGFile.FullName.Replace(".png", ".jpg"), [System.Drawing.Imaging.ImageFormat]::JPEG)
            $png.Dispose()

            # Clean PNG file
            Remove-Item -Path $PNGFile.FullName -Confirm:$false
        }
    }
} Else {
    Write-Warning "Nothing to do!"
}

Exit(0)
