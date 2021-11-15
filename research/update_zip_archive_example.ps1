
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
$zipfileName = "$($PSScriptRoot)\example.zip"
$fileToEdit = "replace.txt"
$contents = "replacement file"


# -- overwrite a file inside a zip archive
# Open zip and find the particular file (assumes only one is inside the Zip file archive)
$zip =  [System.IO.Compression.ZipFile]::Open($zipfileName, [System.IO.Compression.ZipArchiveMode]::Update, [System.Text.Encoding]::Default)
$archiveFiles = $zip.Entries.Where({$_.name -eq $fileToEdit})

# Update the contents of the file
$desiredFile = [System.IO.StreamWriter]($archiveFiles).Open()
$desiredFile.BaseStream.SetLength(0)
$desiredFile.Write($contents)
$desiredFile.Flush()
$desiredFile.Close()

# Write the changes and close the zip file
$zip.Dispose()
Write-Host "zip file updated"


# -- verify file is overwitten inside a zip archive
# Open zip and find the particular file (assumes only one is inside the Zip file archive)
$zip =  [System.IO.Compression.ZipFile]::Open($zipfileName, [System.IO.Compression.ZipArchiveMode]::Update)
$archiveFiles = $zip.Entries.Where({$_.name -eq $fileToRead})

# Read the contents of the file
$desiredFile = [System.IO.StreamReader]($archiveFiles).Open()
$text = $desiredFile.ReadToEnd()
$desiredFile.Close()
$desiredFile.Dispose()

# Close the zip file
$zip.Dispose()

# Output the contents
$text

Write-Host "zip file verify"

