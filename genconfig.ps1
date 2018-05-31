$sobjects = get-process | force describe -t=metadata -n=CustomObject
foreach($sobj in $sobjects){
    $sobjname=$sobj -replace " - CustomObject", ""
    $sobjname = $sobjname.trim()
    $descr = Get-Process | force describe -t=sobject -n="$sobjname"
    $csvop="API Name,Label,Length,Type,Is Nillable,Reference To,Is Custom,Is External Id,Is Unique,Scale,Precision,Picklist Values"
    foreach($f in $desc.fields){
        if($f.permissionable){
            $pickval=","
            $op=""
            $op="$($f.name),$($f.label),$($f.length),$($f.type),$($f.nillable),$($f.referenceTo),$($f.custom),$($f.idLookup),$($f.unique),$($f.scale),$($f.precision)"
            if($f.type -eq "picklist"){
                foreach($val in $f.picklistValues){
                    $pickval="$($pickval)|$($val.label)|"
                }
                $op="$($op)$($pickval)"
            }
            $csvop = "$($csvop)`n$($op)"
        }
    }
    Write-Output $csvop | out-file "$($sobjname).csv" -encoding utf8
}