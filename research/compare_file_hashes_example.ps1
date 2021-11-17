
#
# Pieter De Ridder
# Compare file hashes and output last changes
#

# (fictive) create a hash table with file names and hashes
$hash1 = @{}
$hash1.Add("134598713425", "file1.txt")
$hash1.Add("903453234365", "file2.txt")
$hash1.Add("662034592345", "file3.txt")
$hash1.Add("754698324594", "file4.txt")

# (fictive) create a hash tabel where file1 and file3 hashes differ.
# file2 and file4 hashes are equal.
$hash2 = @{}
$hash2.Add("134598713436", "file1.txt") # changed
$hash2.Add("903453234365", "file2.txt") # equal (unchanged)
$hash2.Add("662034592356", "file3.txt") # changed
$hash2.Add("754698324594", "file4.txt") # equal (unchanged)


# compare hash tables
[System.Array]$compareResult = Compare-Object -ReferenceObject $($hash1.Keys) -DifferenceObject $($hash2.Keys) -IncludeEqual
#$compareResult
#$compareResult.GetType()

# hash2 contains the "last changed files".
# we need to output the result and find those that are changed.
ForEach($r In $compareResult) {
    #$r.SideIndicator
    #$r.InputObject
    If ($r.SideIndicator -eq "=>") {
        $hash2[$r.InputObject]
    }
}

