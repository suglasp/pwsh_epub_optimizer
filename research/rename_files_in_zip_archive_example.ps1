
#
# update a file in a zip file archive -- example code in Powershell
# Notice this example replaces a text file, or a string in face inside a file in the zip
# For binary files, you need to use System.IO.StreamReader class to read the original file and then copy the original file's stream to the $desiredfile.BaseStream
#


#Requires -Version 5.1

# Assemblies needed
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Text.Encoding

# Some Parameters we need
$zipfileName  = "$($PSScriptRoot)\example.zip"
$fileToRename = "TestPhoto.jpg"
$fileToName   = "TestPhoto.png"


# -- overwrite a file inside a zip archive
# Open zip and find the particular file (assumes only one is inside the Zip file archive)
$zip =  [System.IO.Compression.ZipFile]::Open($zipfileName, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::Default)
$archiveFiles = $zip.Entries.Where({$_.name -eq $fileToRename})

ForEach($archiveFile In $archiveFiles) {
    $newFile = $zip.CreateEntry($fileToName)
#    $oldFile = $archiveFile.Open()
#    $oldFile.CopyTo($newFile)
#    $oldFile.Close()
#    $newFile.Close()   
    $archiveFile.Delete()
}

# Update the contents of the file
#$desiredFile = [System.IO.StreamWriter]($archiveFiles).Open()
#$desiredFile.BaseStream.SetLength(0)
#$desiredFile.Write($contents)
#$desiredFile.Flush()
#$desiredFile.Close()

# Write the changes and close the zip file
$zip.Dispose()
Write-Host "zip file updated"

