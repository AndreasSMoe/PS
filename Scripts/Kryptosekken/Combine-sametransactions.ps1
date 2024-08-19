# Combine all transactions with same To/From currency and gebyr to one entry pr year
# Import all transactions to kryptosekken
# Export all transactions from kryptosekken for a specific year (https://www.kryptosekken.no/regnskap/eksport?grense=0)
# Import $ImportReady to Kryptosekken 
¤ Example script, use with caution
#Files
$FileName         = "transactions-2023"
$PathToFile       = "c:\temp\Krytosekken\"
$ImportFile       = "$($PathToFile)$($FileName).csv"
$ImportFileErverv = "$($PathToFile)$($FileName)_erverv.csv"
$ImportReady      = "$($PathToFile)$($FileName)_ReadyForImport.csv"

#Combine inntekt and erverv
(get-content $ImportFile).replace('inntekt','erverv')|set-content -Path $ImportFileErverv

#Remove old result
if(Test-Path $ImportReady){    Remove-Item $ImportReady -Force }

# Import file with combined inntekt and erverv
$Export         = import-csv $ImportFileErverv

#Create resultfile
Add-Content $ImportReady "Tidspunkt,Type,Inn,Inn-Valuta,Ut,Ut-Valuta,Gebyr,Gebyr-Valuta,Marked,Notat"
#add column for Year and month
Foreach($Line in $Export){
    $Line|add-member -NotePropertyName "Year" -NotePropertyValue $($Line.Tidspunkt.substring(0,4))
    $Line|add-member -NotePropertyName "Month" -NotePropertyValue $($Line.Tidspunkt.substring(5,2))
}
# Group on Year,inn-valuta,ut-valuta and Gebyr-valuta
Foreach($Year in $export|Group-Object Year,Type,Inn-valuta,Ut-valuta,Gebyr-Valuta){ 
    $Yearinfo = $Year.name.Split(',').trim()
    write-host " From $($Yearinfo[3]) to $($Yearinfo[2]) "
    $In = 0
    $Out = 0
    $Gebyr = 0
    $Marked = ""
    $antall = 0
    Foreach($Group in $Year.group){
        $Out+=   $(($group|Measure-Object -Property Ut -Sum).sum)
        $In+=    $(($group|Measure-Object -Property Inn -Sum).sum)
        $Gebyr+= $(($group|Measure-Object -Property Gebyr -Sum).sum)
        if($Marked -notmatch $($group.Marked)){
            if($marked -ne ""){             $Marked = "$Marked - $($group.Marked)"            }
            else{                           $Marked = "$Marked $($group.Marked)"            }
        }
        $antall++
    }
    #Use firstdate as time for transaction
    if(($Year.group.tidspunkt|measure-object).count -eq 1){        $tidspunkt = $Year.group.tidspunkt    }
    else{                                                          $tidspunkt = $Year.group.tidspunkt[0]    }
    #If single transaction, keep note, if not replace with info
    if($($Year.Group.notat|measure-object).count -eq 1){           $Notat = $Year.Group.notat    }
    else{                                                           $Notat = "Sum alle $($Yearinfo[2]) mot $($Yearinfo[3]) i $($Yearinfo[0])"    }
    #write-host "$Tidspunkt,$($Yearinfo[1]),$In,$($Yearinfo[2]),$Out,$($Yearinfo[3]),$gebyr,$($Yearinfo[4]),$($Yearinfo[5]),$notat - $marked, $antall"
    Add-Content $ImportReady "$Tidspunkt,$($Yearinfo[1]),$In,$($Yearinfo[2]),$Out,$($Yearinfo[3]),$gebyr,$($Yearinfo[4]),$marked,$notat"
}
