#$inputFile = "in.txt"
param($inputFile, $outputFile)
#$outputFile = "test.txt"
$i = 0
$j = 0


$f = (Get-Content $inputFile) -as [string[]]
"確認" | % { $_.ToString() } > $outputFile

foreach ($l in $f) {
    while ( $i -le $l.Length - 8 )
    {
	    if($l.Substring($i, 8) -match " [0-9][0-9][0-9][0-9][0-9][0-9][0-9]"){
		    $l.Substring($j, $i - $j)  | % { $_.ToString() } >> $outputFile
            $j = $i + 1
	    }
	    $i++
    }
    $l.Substring($j, ($i + 7) - $j)  | % { $_.ToString() } >> $outputFile
}
