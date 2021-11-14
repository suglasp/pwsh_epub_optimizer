
#
# Pieter De Ridder
# epub ebook optimizer
# https://github.com/suglasp/pwsh_epub_optimizer.git
#
# Created : 14/11/2021
# Updated : 14/11/2021
#
# Google Books only supports epub files up to 100Mb in size.
# This created the need for me to optimize epub books.
# It does no more, then opening a epub file and scan for jp(e)g files and compress them some bit more
# to create a smaller resulting epub file as an end result.
# Does not work on epub files containing .gif, .svg and .png files.
#
# Usage:
# .\epub_ebook_optimizer.ps1 <path\myebookfile.epub>  # single argument is the same as providing arg -epub <path\myebookfile.epub>
# .\epub_ebook_optimizer.ps1 -epub <path\myebookfile.epub> [-limit <size as bytes>] [-compression <0..100 as procent>]
#


#region Assemblies needed
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.XML
#endregion


#region Global vars
[UInt32]$Global:DefaultSizeLimit   = (100*1024*1024)  # Default epub files larger then 100Mb (converted to bytes)
[UInt32]$Global:DefaultCompression = 50L              # Default 50% JPG compression
#endregion


#region Methods for handling image formats
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


#
# Function : Optimize-JPEGImageFile
# Optimize images for jp(e)g or png
#
Function Optimize-JPEGImageFile
{  
    Param(
        [string]$JPEGImageFilePath,
        [Int32]$CompressionLevel = 50L
    )

    # load in file and save to new out file
    If (Test-Path -Path $JPEGImageFilePath) {
        [string]$Extension = ([System.IO.Path]::GetExtension($JPEGImageFilePath))
        [string]$newJPEGImageFilePath = $JPEGImageFilePath.Replace($Extension, "_new$($Extension)")
        [System.Drawing.Bitmap]$bmp = New-Object System.Drawing.Bitmap($JPEGImageFilePath)

        If ($bmp -ne $null) {
            [System.Drawing.Imaging.ImageCodecInfo]$jpgEncoder = Get-ImageEncoder([System.Drawing.Imaging.ImageFormat]::Jpeg)
            [System.Drawing.Imaging.Encoder]$myEncoder = [System.Drawing.Imaging.Encoder]::Quality
            [System.Drawing.Imaging.EncoderParameters]$myEncoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
            [System.Drawing.Imaging.EncoderParameter]$myEncoderParameter = New-Object  System.Drawing.Imaging.EncoderParameter($myEncoder, $CompressionLevel)
            $myEncoderParameters.Param[0] = $myEncoderParameter
            $bmp.Save($newJPEGImageFilePath, $jpgEncoder, $myEncoderParameters)  # overwrite existing file
            $bmp.Dispose()
            $bmp = $null
        }

        # swap out old image with new one
        Remove-Item -Path $JPEGImageFilePath -Force -ErrorAction SilentlyContinue  # delete original
        Move-Item -Path $newJPEGImageFilePath -Destination $JPEGImageFilePath -ErrorAction SilentlyContinue # rename old one with new one
    }
}
#endregion

#region Methods for handling the epub file format
#
# Function : Verify-EpubSignature
# Check the epub file it is compressed as ZIP
#
Function Verify-EpubSignature
{
    Param(
        [string]$EpubFile
    )

    [Bool]$IsCompressedEpub = $false

    If ((Test-Path -Path $EpubFileInfo.FullName) -and ([System.IO.Path]::GetExtension($EpubFileInfo.FullName) -eq ".epub")) {
        # check if the file is ZIP compressed (and possibly has epub file content)
        # Get the first 2 bytes of the epub file
        $epubFileHandle = [System.IO.File]::OpenRead($EpubFile)
        $epubFileReader = New-Object System.IO.BinaryReader($epubFileHandle, [System.Text.Encoding]::ASCII)
        [string]$epubCode = $([System.Text.Encoding]::ASCII.GetString($epubFileReader.ReadBytes(2)))
        $epubFileReader.Close()
        $epubFileReader.Dispose()
        $epubFileReader = $null
        $epubFileHandle.Close()
        $epubFileHandle.Dispose()
        $epubFileHandle = $null

        If ($epubCode.Equals("PK")) {
            $IsCompressedEpub = $true
        }
    }

    Return $IsCompressedEpub
}


#
# Function : Unpack-EpubToFiles
# Unpack a single epub file into the original internal files
#
Function Unpack-EpubToFiles
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo
    )
    
    Write-Host "Verify and extracting epub..."

    If (-not ($EpubFileInfo -eq $null)) {
        If (Verify-EpubSignature -EpubFile $EpubFileInfo.FullName) {
            [string]$extractFolder = "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName)"

            # create unpack folder
            If (-not (Test-Path -Path $extractFolder)) {
                New-Item -Path $EpubFileInfo.Directory -Value $EpubFileInfo.BaseName -ItemType Directory -Force -Confirm:$false | Out-Null
            }

            # unpack the epub (zip file format)
            If (-not (Test-Path -Path $extractFolder)) {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($EpubFileInfo.FullName, $extractFolder)
            }
        } Else {
            Write-Warning "[!] Not a valid epub file."
        }
    }
}


#
# Function : Unpack-FilesToEpub
# Pack a group of files into a epub file
#
Function Pack-FilesToEpub
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo
    )

    Write-Host "Packing new optimized epub file..."

    If (-not ($EpubFileInfo -eq $null)) {
        # rename original epub file
        #If (Test-Path -Path $EpubFileInfo.FullName) {
        #    Rename-Item -Path $EpubFileInfo.FullName -NewName "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName).bck" -Force -Confirm:$false
        #    Write-Host "[i] Renamed old epub file to $($EpubFileInfo.BaseName).bck."
        #}

        [string]$Extension = ([System.IO.Path]::GetExtension($EpubFileInfo.FullName))

        # build extract folder path
        [string]$extractFolder = "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName)"
        [string]$epubOptimizedFileName = $EpubFileInfo.FullName.Replace($Extension, "_optimized$($Extension)")
        
        # delete old optimized epub file
        If (Test-Path -Path $epubOptimizedFileName) {
            Remove-Item $epubOptimizedFileName -Force -Confirm:$false
        }

        # pack the content to a new .epub file
        [System.IO.Compression.ZipFile]::CreateFromDirectory($extractFolder, $epubOptimizedFileName)

        # delete temporary content folder(s)
        # remove extracted epub files
        Get-ChildItem -Path $extractFolder -Recurse | Remove-Item -Force -Recurse -Confirm:$false
        
        # remove extract epub folder
        Remove-Item $extractFolder -Force -Confirm:$false

        Write-Host "[i] Created optimized epub $($epubOptimizedFileName)."
    }
}


#
# Function : Unpack-FilesToEpub
# Pack a group of files into a epub file
#
Function Optimize-EpubBody
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo,
        [string]$EpubImageTypes = "jpg",
        [Int32]$ImageCompressionLevel = 50L
    )

    If (-not ($EpubFileInfo -eq $null)) {
        # build extract folder
        [string]$extractFolder = "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName)"

        If (Test-Path -Path $extractFolder) {
            # find the package.opf META file
            [string[]]$imgFiles = @(Get-ChildItem -Path $extractFolder -Include "*.$($EpubImageTypes)" -Recurse)

            # got image file(s)?
            If ($imgFiles.Length -gt 0) {
                Write-Host "Optimizing $($EpubImageTypes) type images in epub file..."
                ForEach($imgFile In $imgFiles) {
                    Optimize-JPEGImageFile -JPEGImageFilePath $imgFile -CompressionLevel $ImageCompressionLevel
                }
            } Else {
                Write-Warning "[i] Nothing to optimize for $($EpubImageTypes) image type format."
            }
        }
    }
}


#
# Function : Get-EpubRawDetails
# Get OS epub file details
#
Function Get-EpubRawDetails
{
    Param(
        [string]$EpubFile
    )

    [System.IO.FileInfo]$epubOSInfo = New-Object System.IO.FileInfo($EpubFile)
    Return $epubOSInfo
}


#
# Function : Get-EpubMetaDetails
# Get META data from the epub file
#
Function Get-EpubMetaDetails
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo
    )

    Write-Host "Fetching META data..."

    # create a epub META template object
    [PSObject]$epubMetaHeader = New-Object PSObject
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name FileName -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name Title -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name SubTitle -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name Author -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name ISBN -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name Publisher -Value ([string]::Empty)
    Add-Member -InputObject $epubMetaHeader -MemberType NoteProperty -Name PublisherDate -Value ([string]::Empty)

    If (-not ($EpubFileInfo -eq $null)) {
        # okay, we got the filename then?
        $epubMetaHeader.FileName = $EpubFileInfo.BaseName

        # build extract folder
        [string]$extractFolder = "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName)"

        If (Test-Path -Path $extractFolder) {
            # find the package.opf META file
            [string[]]$opfFiles = @(Get-ChildItem -Path $extractFolder -Include "*.opf" -Recurse)

            # got META file(s)?
            If ($opfFiles.Length -gt 0) {            
                try {
                    # try xml reading the FIRST opf file
                    [System.IO.StreamReader]$opfRawDataStream = New-Object System.IO.StreamReader($opfFiles[0]) # read first META file. Mostly, there is only one.
                    [System.Xml.XmlReader]$opfXmlReader = [System.Xml.XmlReader]::Create($opfRawDataStream.BaseStream)
                    [System.Xml.XmlDocument]$opfXmlData = New-Object System.Xml.XmlDocument

                    $opfXmlData.Load($opfXmlReader)

                    # most opf files have an xml attribute called #text
                    # not all epub files support this structure. So we need to check the base type.
                    #$title = $($opfXmlData.package.metadata.title).'#text'
                    #$creator = $($opfXmlData.package.metadata.creator).'#text'
                    #$isbn = $($opfXmlData.package.metadata.identifier).'#text'
                    #publisher = $($opfXmlData.package.metadata.publisher).'#text'

            
                    # What we do here below, is a very fast approach of parsing the XML data.
                    # I don't want to read a whole book in XML parsing in .NET, before coding this in Powershell.
                    # I just want a few data fields as fast as possible in string readable format. We use datatype checking and convert if possible.

                    # collect the data fields as an array
                    # sometimes it will contain one record with a string value.
                    # other times it will contain more records like a hash value.
                    [System.Array]$epubTitleData = @($opfXmlData.package.metadata.title)
                    [System.Array]$epubCreatorData = @($opfXmlData.package.metadata.creator)
                    [System.Array]$epubISBNData = @($opfXmlData.package.metadata.identifier)
                    [System.Array]$epubPublisherData = @($opfXmlData.package.metadata.publisher)
                    [System.Array]$epubDateData = @($opfXmlData.package.metadata.date)

                    # extract the data from array. Fields we want are always on the end of the array.
                    [System.Object]$title = [string]::Empty
                    [System.Object]$subtitle = [string]::Empty
                    Switch($epubTitleData.Length) {
                        1 { $title = $epubTitleData[$epubTitleData.Length -1] }
                        2 { $title = $epubTitleData[$epubTitleData.Length -2]; $subtitle = $epubTitleData[$epubTitleData.Length -1] }  
                    }

                    [System.Object]$creator = $epubCreatorData[$epubCreatorData.Length -1]
                    [System.Object]$isbn = $epubISBNData[$epubISBNData.Length -1]
                    [System.Object]$publisher = $epubPublisherData[$epubPublisherData.Length -1]
                    [System.Object]$date = $epubDateData[$epubDateData.Length -1]
            
                    # override if the type is NOT string
                    If ($title -Is [System.Xml.XmlElement]) {
                        $title = ($title).InnerText
                    }

                    If ($subtitle -Is [System.Xml.XmlElement]) {
                        $subtitle = ($subtitle).InnerText
                    }

                    If ($creator -Is [System.Xml.XmlElement]) {
                        $creator = ($creator).InnerText
                    }

                    If ($isbn -Is [System.Xml.XmlElement]) {
                        $isbn = ($isbn).InnerText
                        #[string]$isbn = $isbn.'#text'
                    }

                    If ($publisher -Is [System.Xml.XmlElement]) {
                        $publisher = ($publisher).InnerText
                    }

                    If ($date -Is [System.Xml.XmlElement]) {
                        $date = ($date).InnerText
                    }
            
                    # ISBN is sometimes something like "urn:isbn:xxx-x-xxxx-xxxx-x"
                    # We split the string value, and again, fetch the last part of the chunck.
                    If ($isbn -Is [String]) {
                        If ($isbn.Contains(':')) {
                            $isbn = $isbn.Split(':')
                            $isbn = $isbn[$isbn.Length -1]
                        }
                    }

                    # now, format the epub meta data for output to stdout and store it in a object
                    $epubMetaHeader.Title = $title
                    $epubMetaHeader.SubTitle = $subtitle
                    $epubMetaHeader.Author = $creator
                    $epubMetaHeader.ISBN = $isbn
                    $epubMetaHeader.Publisher = $publisher
                    $epubMetaHeader.PublisherDate = $date
                } catch {
                    Write-Warning "[!] Error parsing epub META data!"
                } finally {
                    # close handles
                    $opfXmlData = $null
                    $opfXmlReader.Close()
                    $opfXmlReader.Dispose()
                    $opfXmlReader = $null
                    $opfRawDataStream.Close()
                    $opfRawDataStream.Dispose()
                    $opfRawDataStream = $null
                }
            } Else {
                Write-Warning "[!] Don't know where the META data is?!"
            }
        }
    }

    Return $epubMetaHeader
}
#endregion


#region Main Method
#
# Function : Main
# The main function of the script
#
Function Main
{
    Param(
        [string[]]$Arguments
    )

    # print script title
    Write-Host ""
    Write-Host "----------------------------"
    Write-Host " epub ebook optimizer tool"
    Write-Host "----------------------------"
    Write-Host ""

    # private parameters
    [string]$TargetFile       = [string]::Empty            # default no file
    [UInt32]$TargetSizeLimit  = $Global:DefaultSizeLimit   # default to 100Mb
    [Int32]$TargetCompression = $Global:DefaultCompression # default JPG compression level   

    $TargetFile = "E:\epub\Mastering_VMware_Horizon8.epub"

    # process script cli arguments
    If ($Arguments) {
        If ($Arguments.Length -eq 1) {
            # single argument. Accept is as parameter "-epub <file>".
            $TargetFile = $Arguments[0]
        } Else {
            # more then one argument. Parse them as arguments list.
            For($i = 0; $i -lt $Arguments.Length; $i++) {
                #Write-Host "DEBUG : Arg #$($i.ToString()) is $($Arguments[$i])"

                # default, a PWSH Switch statement on a String is always case insensitive
                Switch ($Arguments[$i].ToLowerInvariant()) {
                    "-epub" { If (($i +1) -le $Arguments.Length) { $TargetFile = $Arguments[$i +1] } }                                      # -epub <path_and_file_name>
                    "-limit" { If (($i +1) -le $Arguments.Length) { [UInt32]::TryParse($Arguments[$i +1], [ref]$TargetSizeLimit) } }        # -limit <size_in_bytes>
                    "-compression" { If (($i +1) -le $Arguments.Length) { [Int32]::TryParse($Arguments[$i +1], [ref]$TargetCompression) } } # -compression <ratio_as_procent>
                    default {}
                }
            }
        }
    }


    # process epub target file
    If (-not ([string]::IsNullOrEmpty($TargetFile))) {
        Write-Host "Target File Name : $($TargetFile)"
        Write-Host "Target Limit Size : $($TargetSizeLimit.ToString()) bytes"
        Write-Host "Target Compression : $($TargetCompression.ToString())%"
        Write-Host ""

        If (Test-Path -Path $TargetFile) {
            # check if we have a epub file and size
            $epubOSInfo = Get-EpubRawDetails -EpubFile $TargetFile

            If ($epubOSInfo.Extension.ToLowerInvariant() -eq ".epub") {
                If ($TargetSizeLimit -lt $epubOSInfo.Length) {
                    # unpack the epub file
                    Unpack-EpubToFiles -EpubFileInfo $epubOSInfo

                    # fetch epub META data details
                    $epubMETA = Get-EpubMetaDetails -EpubFileInfo $epubOSInfo

                    Write-Host ""
                    Write-Host "-- Meta info --"
                    Write-Host "epub File      : $($epubMETA.FileName)"
                    Write-Host "book Title     : $($epubMETA.Title)"
                    Write-Host "book Sub-Title : $($epubMETA.SubTitle)"
                    Write-Host "book Author    : $($epubMETA.Author)"
                    Write-Host "book ISBN      : $($epubMETA.ISBN)"
                    Write-Host "Publisher      : $($epubMETA.Publisher)"
                    Write-Host "Publisher Date : $($epubMETA.PublisherDate)"
                    Write-Host ""

                    # optimize the epub images we support (*.jpg and *.jpeg)
                    #Optimize-EpubBody -EpubFileInfo $epubOSInfo -EpubImageTypes "jpg"
                    #Optimize-EpubBody -EpubFileInfo $epubOSInfo -EpubImageTypes "jpeg"

                    # repackage
                    #Pack-FilesToEpub -EpubFileInfo $epubOSInfo

                    # free
                    $epubOSInfo = $null

                    Write-Host "Done."
                } Else {
                    Write-Warning "[!] No epub action needed! Size is okay."
                } 
            } Else {
                Write-Warning "[!] Not a epub file?"
            }
        } Else {
            Write-Warning "[!] File $($TargetFile) not found!"
        }
    } Else {
        Write-Warning "[!] Provide a target epub file using parameter -epub <file>, and optionally -limit or -compression!"
        Write-Host ""
        Write-Host "Usage : .\$(Split-Path -Path $MyInvocation.ScriptName -Leaf) <path\ebookfile.epub>"
        Write-Host "Usage : .\$(Split-Path -Path $MyInvocation.ScriptName -Leaf) -epub <path\ebookfile.epub> [-limit <size as bytes>] [-compression <0..100 as procent>]"
        Write-Host "[i] Default -limit size is $($Global:DefaultSizeLimit) bytes."
        Write-Host "[i] Default -compression ratio value is $($Global:DefaultCompression)%."
        Write-Host ""
    }


    # gracefully exit
    Exit(0)
}
#endregion




# --------------------




# call main in c-style
Main -Arguments $args

