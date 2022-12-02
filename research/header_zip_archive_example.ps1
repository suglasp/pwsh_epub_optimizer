
#
# example code to check if the archive is a PK (Zip) archive.
# Pieter De Ridder
#

#Requires -Version 5.1

# Assemblies needed
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Text.Encoding

# Some Parameters we need
$zipfileName = "$($PSScriptRoot)\example.zip"

If (Test-Path -Path $zipfileName) {
	$PKarchive = [System.IO.File]::OpenRead($zipfileName)
	$PKreader = New-Object System.IO.BinaryReader($PKarchive, [System.Text.Encoding]::ASCII)
	$PKreader.BaseStream.Position = 0
	[string]$PKheader = [System.Text.Encoding]::ASCII.GetString($PKreader.ReadBytes(2))

	If ($PKheader -eq "PK") {
		Write-Host "ZIP Archive"
	} Else {
		Write-Host "Not ZIP Archive"
	}

	If ($PKreader) {
		$PKreader.Close()
		$PKreader = $null
	}

	If ($PKarchive) {
		$PKarchive.Close()
		$PKarchive = $null
	}
} Else {
	Write-Warning "No file $($zipfileName)!"	
}

Exit(0)
