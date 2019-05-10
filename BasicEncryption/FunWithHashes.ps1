

$intHS1 = New-Object -TypeName System.Collections.Generic.HashSet[int]
1..5 | %{[void]($intHS1.Add($_))}
[void]($intHS2 = New-Object -TypeName System.Collections.Generic.HashSet[int])
3..7 | %{[void]($intHS2.Add($_))}

Write-Verbose "`$intHS1 =  $intHS1" -Verbose
Write-Verbose "`$intHS2 =  $intHS2" -Verbose

Write-Warning "Remove all instances where items from set 2 intersect with set 1" 
$intHS1.IntersectWith($intHS2)
Write-Verbose "`$intHS1 =  $intHS1" -Verbose 
#get the intersection of two distinct sets
 

# Join only the unique values from two distinct sets
$intHS1 = New-Object -TypeName System.Collections.Generic.HashSet[int]
1..5 | %{[void]($intHS1.Add($_))}
[void]($intHS2 = New-Object -TypeName System.Collections.Generic.HashSet[int])
3..7 | %{[void]($intHS2.Add($_))}

Write-Verbose "`$intHS1 =  $intHS1" -Verbose
Write-Verbose "`$intHS2 =  $intHS2" -Verbose

Write-Warning "Remove all instances where items from set 2 intersect with set 1" 
$intHS1.UnionWith($intHS2)
Write-Verbose "`$intHS1 =  $intHS1" -Verbose 
 
