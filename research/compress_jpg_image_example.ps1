
#
# Pieter De Ridder
# Code loads jpg file and adjusts compression level and save it to a new jpg file.
# https://github.com/suglasp/pwsh_epub_optimizer.git
#
# Created : 13/11/2021
# Updated : 14/11/2021
#
# Ported example from Microsoft to Powershell code
# ref : https://docs.microsoft.com/en-us/dotnet/desktop/winforms/advanced/how-to-set-jpeg-compression-level?view=netframeworkdesktop-4.8
#

Add-Type -AssemblyName System.Drawing

#
# Function : Get-ImageEncoder
# Get the .NET System imaging codecs
#
Function Get-ImageEncoder
{  
    param(
        [System.Drawing.Imaging.ImageFormat]$format
    )

    [System.Drawing.Imaging.ImageCodecInfo[]]$codecs = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()
    ForEach ($codec In $codecs)  
    {  
        If ($codec.FormatID -eq $format.Guid)  
        {  
            Return $codec
        }
    }

    Return $null
}


# in and out file
[string]$inFile = "$($WorkDir)\TestPhoto.jpg"
[string]$outFile = "$($WorkDir)\TestPhotoQualityReduction.jpg"

# load in file and save to new out file
If (Test-Path -Path $inFile) {
    [System.Drawing.Bitmap]$bmp = New-Object System.Drawing.Bitmap($inFile)

    If ($bmp -ne $null) {
        [System.Drawing.Imaging.ImageCodecInfo]$jpgEncoder = Get-ImageEncoder([System.Drawing.Imaging.ImageFormat]::Jpeg)
        [System.Drawing.Imaging.Encoder]$myEncoder = [System.Drawing.Imaging.Encoder]::Quality
        [System.Drawing.Imaging.EncoderParameters]$myEncoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
        #[System.Drawing.Imaging.EncoderParameter]$myEncoderParameter = New-Object  System.Drawing.Imaging.EncoderParameter($myEncoder, 100L) # 100% (best) compression
        [System.Drawing.Imaging.EncoderParameter]$myEncoderParameter = New-Object  System.Drawing.Imaging.EncoderParameter($myEncoder, 50L)  # 50% compression
        #[System.Drawing.Imaging.EncoderParameter]$myEncoderParameter = New-Object  System.Drawing.Imaging.EncoderParameter($myEncoder, 0L)   # 0% (worst) compression
        $myEncoderParameters.Param[0] = $myEncoderParameter
        $bmp.Save($outFile, $jpgEncoder, $myEncoderParameters)
        $bmp.Dispose()
    }
}

Exit(0)
