
#
# Pieter De Ridder
# epub ebook optimizer
# https://github.com/suglasp/pwsh_epub_optimizer.git
#
# Created : 14/11/2021
# Updated : 17/11/2021
#
# Google Books only supports epub files up to ~100Mb in size.
# This created the need for me to optimize epub books.
# It does no more, then opening a epub file and scan .png files. convert .png to .jpg. And then for all jp(e)g files, increase compression for them a bit more
# to create a smaller resulting epub file as an end result.
# Does not work on epub files containing .gif and .svg image files.
#
# Usage:
# .\epub_ebook_optimizer.ps1 <path\myebookfile.epub>  # single argument is the same as providing arg -epub <path\myebookfile.epub>
# .\epub_ebook_optimizer.ps1 -epub <path\myebookfile.epub> [-limit <size in bytes>] [-compression <0..100 in procent>]
#
# Earlier versions only compressed the .jp(e)g files. This resulted with some test epub files containing only jpg files in about a 20%-30% smaller epub file.
# Newer version also converts .png files to .jp(e)g files. This resulted with some test epub files containing a mix of jpg and png files in about a 50%-60% smaller epub file.
# The size gain, does result that images will look less clear (because they are Lossy jpeg's compressed default at 50%).
#


#Requires -Version 5.1

#region Assemblies needed
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Text.Encoding
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.XML
#endregion


#region Global vars
[UInt32]$Global:DefaultSizeLimit     = (98*1024*1024)   # Default epub files larger then 98Mb (converted to bytes). Ideal is 100Mb, but we set it a bit smaller
[Long]$Global:DefaultJPEGCompression = 50               # Default 50% JPG compression
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
        [Long]$CompressionLevel = 50
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

        # swap out old image with new one on disk
        Remove-Item -Path $JPEGImageFilePath -Force -ErrorAction SilentlyContinue  # delete original
        Move-Item -Path $newJPEGImageFilePath -Destination $JPEGImageFilePath -ErrorAction SilentlyContinue # rename old one with new one
    }
}


#
# Function : Convert-PNGtoJPGFormat
# Convert png files to jp(e)g images
#
Function Convert-PNGtoJPGFormat
{
    Param(
        [String]$FolderToSearch
    )

    [System.Collections.ArrayList]$PNGFiles = @(Get-ChildItem -Path $FolderToSearch -Filter *.png -Recurse)

    If ($PNGFiles.Count -gt 0) {
        ForEach($PNGFile In $PNGFiles) {
            If (Test-Path -Path $PNGFile.FullName) {
                # Open PNG and Save as a JP(E)G
                Write-Host "Converting PNG2JPG : $($PNGFile)..."
                [System.Drawing.Image]$png = [System.Drawing.Image]::FromFile($PNGFile.FullName)
                $png.Save($PNGFile.FullName.Replace(".png", ".jpg"), [System.Drawing.Imaging.ImageFormat]::JPEG)
                $png.Dispose()

                # Clean PNG file
                Remove-Item -Path $PNGFile.FullName -Confirm:$false
            }
        }
    } Else {
    #    Write-Warning "Nothing to do!"
    }

    Return $PNGFiles
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
# Unpack/unzip a single epub archive file into it's original data files
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
# Function : Cleanup-UnpackedEpubFiles
# Cleanup the extracted/unzipped epub archive files
#
Function Cleanup-UnpackedEpubFiles
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo
    )

    If (-not ($EpubFileInfo -eq $null)) {
        [string]$extractFolder = "$($EpubFileInfo.Directory)\$($EpubFileInfo.BaseName)"

        # remove extracted epub files
        Get-ChildItem -Path $extractFolder -Recurse | Remove-Item -Force -Recurse -Confirm:$false
        
        # remove extract epub folder
        Remove-Item $extractFolder -Force -Confirm:$false
    }
}


#
# Function : Pack-FilesToEpub (DEPRICATED)
# Pack/zip a group of files into a new epub archive file
# DEPRICATED => .NET zip files are packes as Deflate.
# Seems like most epub archive files are some other type of compression and not all epub readers can handle the Deflate
# archive. Seems like Google Android Google Books app can't read or import them anyway.
# So my solution : clone the original epub file and re-inject the replacement files.
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
        Cleanup-UnpackedEpubFiles

        Write-Host "[i] Created optimized epub $($epubOptimizedFileName)."
    }
}


#
# Function : Replace-FileInEpubFile
# Replace a file inside a epub archive file
#
# Notice : Every time you call this Method, the epub file (or zip file let's say) is opened and written to.
# Ideal, you would open once the file, edit it all the way until it's done and then close it (using a array as input with files te replace).
# Because we work here with a scripting language and can any time be interrupted by the user,
# we do it the less I/O efficient, but safer approach.
#
Function Replace-FileInEpubFile
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo,
        [string]$ReplacementFile
    )

    If ( (-not ($EpubFileInfo -eq $null)) -and (-not ([string]::IsNullOrEmpty($ReplacementFile))) ) {
        # Open epub archive and find the particular content file (assumes only one is inside the epub file archive)
        $epubArchive =  [System.IO.Compression.ZipFile]::Open($EpubFileInfo.FullName, [System.IO.Compression.ZipArchiveMode]::Update)
        #$epubArchive =  [System.IO.Compression.ZipFile]::Open($EpubFileInfo.FullName, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::Default)
        $epubContentFiles = $epubArchive.Entries.Where({$_.name -eq $(Split-Path -Path $ReplacementFile -Leaf)})

        If ($epubContentFiles.Count -gt 0) {
            # Update the epub archive contents
            [System.IO.StreamReader]$OptimizedImgFileStream = [System.IO.StreamReader]($ReplacementFile)
                
            $desiredFile = [System.IO.StreamWriter]($epubContentFiles).Open()
            If ($desiredFile -ne $null) {
                $desiredFile.BaseStream.SetLength(0)
                $desiredFile.BaseStream.Position = 0
                $OptimizedImgFileStream.BaseStream.CopyTo($desiredFile.BaseStream)
                $desiredFile.Flush()
                $desiredFile.Close()
            } Else {
                Write-Warning "[!] Error while updating data lump in archive!"
            }

            $OptimizedImgFileStream.Close()
        } Else {
            Write-Warning "[!] Unable to update data lump in archive!"
        }
        
        
        # Write the changes and close the epub archive file
        $epubArchive.Dispose()
    }
}


#
# Function : Rename-FileInEpubFile
# Rename a file inside a epub archive file
# Creates a new file entry and deletes the old one.
# The new entry will have no data, we don't care anyway because in a later stage, we will fill it up with new data from external.
#
# Notice : Every time you call this Method, the epub file (or zip file let's say) is opened and written to.
# Ideal, you would open once the file, edit it all the way until it's done and then close it (using a array as input with files te replace).
# Because we work here with a scripting language and can any time be interrupted by the user,
# we do it the less I/O efficient, but safer approach.
#
Function Rename-FileInEpubFile
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo,
        [string]$SearchFile,
        [string]$ReplaceFile
    )

    If ( (-not ($EpubFileInfo -eq $null)) -and (-not ([string]::IsNullOrEmpty($SearchFile))) -and (-not ([string]::IsNullOrEmpty($ReplaceFile))) ) {
        # Open epub archive and find the particular content file (assumes only one is inside the epub file archive)
        $epubArchive =  [System.IO.Compression.ZipFile]::Open($EpubFileInfo.FullName, [System.IO.Compression.ZipArchiveMode]::Update)
        #$epubArchive =  [System.IO.Compression.ZipFile]::Open($EpubFileInfo.FullName, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::Default)
        $epubContentFiles = $epubArchive.Entries.Where({$_.name -eq $(Split-Path -Path $SearchFile -Leaf)})

        If ($epubContentFiles.Count -gt 0) {
            ForEach($archiveFile In $epubContentFiles) {
                $newFile = $epubArchive.CreateEntry("$(Split-Path -Path $archiveFile.FullName -Parent)\$($ReplaceFile)")
                # Normally, we would copy the data stream from 'old entry' to the 'new entry'.
                # We skip this step. Later in the process, we inject the .jpg data stream from external to the new entries.
                $archiveFile.Delete()
            }
        } Else {
            Write-Warning "[!] Unable to update data lump in archive!"
        }        
        
        # Write the changes and close the epub archive file
        $epubArchive.Dispose()
    }
}


#
# Function : Clone-EpubToEpubCopy
# Clone (copy) a epub archive file
#
Function Clone-EpubToEpubCopy
{
    Param(
        [System.IO.FileInfo]$EpubFileInfo
    )

    Write-Host ""
    Write-Host ">> Cloning epub file..."

    $epubCloneOSInfo = $null

    If (-not ($EpubFileInfo -eq $null)) {
        # get original file extension
        [string]$Extension = ([System.IO.Path]::GetExtension($EpubFileInfo.FullName))
                
        # build new filename
        [string]$epubOptimizedFileName = $EpubFileInfo.FullName.Replace($Extension, "_optimized$($Extension)")
        
        # delete old optimized epub file
        If (Test-Path -Path $epubOptimizedFileName) {
            Remove-Item $epubOptimizedFileName -Force -Confirm:$false
        }

        # copy the original to the new file
        Copy-Item -Path $EpubFileInfo.FullName -Destination $epubOptimizedFileName -Force -Confirm:$false

        # return file OS info
        $epubCloneOSInfo = Get-EpubRawDetails -EpubFile $epubOptimizedFileName

        Write-Host "[i] Created epub clone file $($epubOptimizedFileName)."
    } Else {
        Write-Warning "[!] Cloning epub file failed!"
        Exit(-1)
    }

    Return $epubCloneOSInfo
}


#
# Function : Optimize-EpubBody
# Optimize the unpacked data files of a epub archive file
#
Function Optimize-EpubBodyLossyImages
{
    Param(
        [System.IO.FileInfo]$EpubOriginalFileInfo,
        [string]$EpubImageTypes = "jpg",
        [Int32]$ImageCompressionLevel = 50L
    )

    Write-Host ""
    Write-Host ">> File type $($EpubImageTypes):"

    If (-not ($EpubOriginalFileInfo -eq $null)) {
        # Build extract folder
        [string]$extractFolder = "$($EpubOriginalFileInfo.Directory)\$($EpubOriginalFileInfo.BaseName)"

        If (Test-Path -Path $extractFolder) {
            # Find the package.opf META file
            [string[]]$imgFiles = @(Get-ChildItem -Path $extractFolder -Include "*.$($EpubImageTypes)" -Recurse)

            # Got image file(s)?
            If ($imgFiles.Length -gt 0) {
                Write-Host "Optimizing $($EpubImageTypes) type images in epub file..."
                Write-Host "Range : $($imgFiles.Length) files for optimization"

                ForEach($imgFile In $imgFiles) {
                    # Compress Image
                    Optimize-JPEGImageFile -JPEGImageFilePath $imgFile -CompressionLevel $ImageCompressionLevel
                }
            } Else {
                Write-Warning "[i] Nothing to optimize for $($EpubImageTypes) image type format."
            }
        }
    } Else {
        Write-Warning "[!] Optimization failed!"
    }
}


#
# Function : Calc-FolderHashList
# Calculate hashes of all files in a folder
#
Function Calc-FolderHashList
{
    Param (
        [string]$TargetFolder
    )

    Write-Host ""
    Write-Host "Indexing $($TargetFolder)..."

    [hashtable]$FileHashList = @{}

    If (Test-Path -Path $TargetFolder) {
        [string[]]$allFiles = @(Get-ChildItem -Path $TargetFolder -Include "*.*" -Recurse)

        ForEach($someFile In $allFiles) {
            # calculate file signature
            [bool]$success = $false
            try {
                [string]$fileSig = (Get-FileHash -Path $someFile -Algorithm SHA256).Hash
                [void]$FileHashList.Add($fileSig, $someFile)
                $success = $true
            } catch {
            }

            # retry with signature of filepath
            If (-not ($success)) {                
                [string]$fileSig = (Get-FileHash -InputStream $([IO.MemoryStream]::new([byte[]][char[]]$someFile)) -Algorithm SHA256).Hash
                [void]$FileHashList.Add($fileSig, $someFile)
                $success = $true
            }
        }
    }
    
    Return $FileHashList
}


#
# Function : Replace-StringInFiles
# Replace a string in a file with some other string
#
Function Replace-StringInFiles
{
    Param (
        [string]$FolderToSearch,
        [string]$Find,
        [string]$Replace
    )
    
    #Write-Host "REPLACE : $($FolderToSearch) :: $($Find) => $($Replace)"
    Write-Host "Updating png reference $($Find) => $($Replace) in epub body..."
    
    If (Test-Path -Path $FolderToSearch) {
        # Find all files
        #[string[]]$dataFiles = @(Get-ChildItem -Path $FolderToSearch -Include @("*.*") -Exclude @("*.jpg, *.jpeg, *.svg, *.gif") -Recurse)
        #[string[]]$dataFiles = @(Get-ChildItem -Path $FolderToSearch -Include @("*.opf", "*.*htm*", "*.css") -Exclude @("*.jpg, *.jpeg, *.svg, *.gif") -Recurse)
        [string[]]$dataFiles = @(Get-ChildItem -Path $FolderToSearch -Include @("*.opf", "*.*htm*", "*.css", "*.xml", "*.ncx") -Recurse)

        ForEach($dataFile In $dataFiles) {
            Try {
                (Get-Content -Path $dataFile).Replace($Find, $Replace) | Set-Content -Path $dataFile
            } Catch {
                Write-Warning "[!] Could not update content in file $($dataFile)."
            }
        }
    }
}


#
# Function : Get-EpubRawDetails
# Get OS epub archive file details
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

    Write-Host ""
    Write-Host ">> Fetching META data..."

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
            # find the package.opf META file (from disk, not live in the epub archive)
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
    [string]$TargetFile       = [string]::Empty                # default no file
    [UInt32]$TargetSizeLimit  = $Global:DefaultSizeLimit       # default to 98Mb (just a bit smaller then 100Mb)
    [Long]$TargetCompression  = $Global:DefaultJPEGCompression # default JPG compression level

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
                    "-compression" { If (($i +1) -le $Arguments.Length) { [Long]::TryParse($Arguments[$i +1], [ref]$TargetCompression) } }  # -compression <ratio>
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

                    If ($epubMETA -ne $null) {
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
                    }

                    # calc hashes of files
                    [hashtable]$originalHashList = Calc-FolderHashList -TargetFolder "$($epubOSInfo.Directory)\$($epubOSInfo.BaseName)"

                    # Clone the original epub file
                    $epubCloneOSInfo = Clone-EpubToEpubCopy -EpubFileInfo $epubOSInfo

                    # first, convert PNG files to JPG. This makes helps to make the epub file smaller.
                    # we do this on the original uncompressed epub data
                    #$convertedPNGFiles = Convert-PNGtoJPGFormat -FolderToSearch $(Split-Path -Path $epubOSInfo.FullName -Parent)
                    $convertedPNGFiles = Convert-PNGtoJPGFormat -FolderToSearch "$($epubOSInfo.Directory)\$($epubOSInfo.BaseName)"

                    # update the table of content file (mostly called Package.opf) and also *.xhtml body files
                    # we do this on the original uncompressed epub data
                    If ($convertedPNGFiles.Length -gt 0) {
                        ForEach($convertedPNGFile In $convertedPNGFiles) {
                            #Replace-StringInFiles -FolderToSearch $(Split-Path -Path $convertedPNGFile.FullName -Parent) -Find $(Split-Path -Path $convertedPNGFile.FullName -Leaf) -Replace $((Split-Path -Path $convertedPNGFile.FullName -Leaf).Replace(".png", ".jpg"))
                            Replace-StringInFiles -FolderToSearch "$($epubOSInfo.Directory)\$($epubOSInfo.BaseName)" -Find $(Split-Path -Path $convertedPNGFile.FullName -Leaf) -Replace $((Split-Path -Path $convertedPNGFile.FullName -Leaf).Replace(".png", ".jpg"))
                        }
                    }

                    # optimize the epub lossy images we support (*.jpg and *.jpeg)
                    # we do this on the original uncompressed epub data
                    If ($epubCloneOSInfo -ne $null) {
                        # rename filename .png to .jpg inside the new epub archive
                        If ($convertedPNGFiles.Length -gt 0) {
                            ForEach($convertedPNGFile In $convertedPNGFiles) {
                                Rename-FileInEpubFile -EpubFileInfo $epubCloneOSInfo -SearchFile $(Split-Path -Path $convertedPNGFile.FullName -Leaf) -ReplaceFile $((Split-Path -Path $convertedPNGFile.FullName -Leaf).Replace(".png", ".jpg"))
                                #Replace-FileInEpubFile -EpubFileInfo $epubCloneOSInfo -ReplacementFile $((Split-Path -Path $convertedPNGFile.FullName -Leaf).Replace(".png", ".jpg"))
                            }
                        }

                        # optimize the jpg and jpeg files
                        Optimize-EpubBodyLossyImages -EpubOriginalFileInfo $epubOSInfo -EpubImageTypes "jpg"
                        Optimize-EpubBodyLossyImages -EpubOriginalFileInfo $epubOSInfo -EpubImageTypes "jpeg"
                    }

                    # repackage (DEPRICATED)
                    #Pack-FilesToEpub -EpubFileInfo $epubOSInfo
                    
                    # re-calc hashes of files
                    [hashtable]$changedHashList = Calc-FolderHashList -TargetFolder "$($epubOSInfo.Directory)\$($epubOSInfo.BaseName)"

                    # if the file was changed, replace it in the optimize epub archive file
                    Write-Host ""
                    Write-Host ">> Comparing index..."
                    [System.Array]$compareResult = Compare-Object -ReferenceObject $($originalHashList.Keys) -DifferenceObject $($changedHashList.Keys) -IncludeEqual

                    # Repackage directly into the cloned epub archive file
                    Write-Host ""
                    Write-Host ">> Repacking $(Split-Path -Path $epubCloneOSInfo.FullName -Leaf)..."
                    ForEach($r In $compareResult) {
                        #If ($r.SideIndicator -eq "==") {
                        #    [string]$changedFileName = $changedHashList[$r.InputObject]
                        #    Write-Host "EQUAL : $($changedFileName)"
                        #}

                        #If ($r.SideIndicator -eq "<=") {
                        #    [string]$changedFileName = $originalHashList[$r.InputObject]
                        #
                        #}
                        
                        If ($r.SideIndicator -eq "=>") {
                            [string]$changedFileName = $changedHashList[$r.InputObject]
                            #Write-Host "CHANGED : $($changedFileName)"
                            If ([System.IO.Path]::GetExtension($changedFileName) -like ".jp*g") {
                                Write-Host "Optimizing -> $(Split-Path -Path $changedFileName -Leaf)"
                            } Else {
                                Write-Host "Updating   -> $(Split-Path -Path $changedFileName -Leaf)"
                            }

                            Replace-FileInEpubFile -EpubFileInfo $EpubOptimizedFileInfo -ReplacementFile $changedFileName
                        }
                    }

                    # clean up temporary unzipped epub files
                    Cleanup-UnpackedEpubFiles -EpubFileInfo $epubOSInfo

                    # output result filename for user friendliness (copy/paste functionality)
                    Write-Host ""
                    Write-Host ""
                    Write-Host "// ----------- RESULT -------------- //"
                    Write-Host ""
                    Write-Host ""
                    Write-Host "Original file : $($epubOSInfo.FullName)"

                    If ($epubCloneOSInfo -ne $null) {
                        # refresh de cloned file info (it should be same size or smaller)
                        $epubCloneOSInfo = Get-EpubRawDetails -EpubFile $epubCloneOSInfo.FullName

                        # calculate percentage of file size reduction
                        If ($epubCloneOSInfo.Length -lt $epubOSInfo.Length) {
                            [Double]$FileSizeDiff = [Math]::Round((($epubOSInfo.Length - $epubCloneOSInfo.Length) / 1024 / 1024))
                            [Int32]$procentReduction = (($FileSizeDiff / 100) * [Math]::Round(($epubOSInfo.Length / 1024 / 1024)))

                            Write-Host "Optimized file : $($epubCloneOSInfo.FullName) [$($procentReduction)% size reduction]"
                        } Else {
                            Write-Host "Optimized file : $($epubCloneOSInfo.FullName) [no size reduction gain]"
                        }

                        Write-Host ""
                    } Else {
                        Write-Host "Optimized file : Unknown [???]"
                    }

                    # free
                    $epubOSInfo = $null
                    $epubCloneOSInfo = $null

                    Write-Host "-- Done"
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
        Write-Host "Usage : .\$(Split-Path -Path $MyInvocation.ScriptName -Leaf) -epub <path\ebookfile.epub> [-limit <size in bytes>] [-compression <0..100 in %>]"
        Write-Host "[i] Default -limit size is $($Global:DefaultSizeLimit) bytes."
        Write-Host "[i] Default -compression ratio value is $($Global:DefaultJPEGCompression)%."
        Write-Host ""
    }


    # gracefully exit
    Exit(0)
}
#endregion




# --------------------




# call main in C-style
Main -Arguments $args

