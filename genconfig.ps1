if($args.Length -gt 0){
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $sobjects = $args[0].Split(",",$option)
}else{
    $sobjects = get-process | force describe -t=metadata -n=CustomObject
}
$configexcel = New-Object -ComObject excel.application
$configexcel.visible = $true
$configbook = $configexcel.Workbooks.Add()
$isNewWorkbook = $true

foreach($sobj in $sobjects){

    $sobjname=$sobj -replace " - CustomObject", ""
    $sobjname = $sobjname.trim()
    $row=1
    $column=1
    $sheetsNotAdded = $true
    $descr = Get-Process | force describe -t=sobject -n="$($sobjname)" | ConvertFrom-Json
    $csvop="API Name,Label,Length,Type,Is Nillable,Reference To,Is Custom,Is External Id,Is Unique,Scale,Precision,Picklist Values" 
    
    foreach($f in $descr.fields){
        
        if($f.permissionable -Or $f.calculated -Or $f.createable -Or $f.updateable -Or $f.idLookup -Or $f.externalId -Or $f.nameField){

            if($sheetsNotAdded){
                if($isNewWorkbook){
                    $mysheet = $configbook.Worksheets.Item(1)
                    $isNewWorkbook = $false
                }else{
                    $mysheet = $configbook.Worksheets.Add()
                }
                $mysheet.name="$($sobjname.substring(0,[System.Math]::Min(31,$sobjname.Length)))"
                $mysheet.Cells.Item($row,$column++)='API Name'
                $mysheet.Cells.Item($row,$column++)='Label'
                $mysheet.Cells.Item($row,$column++)='Length'
                $mysheet.Cells.Item($row,$column++)='Type'
                $mysheet.Cells.Item($row,$column++)='Is Nillable'
                $mysheet.Cells.Item($row,$column++)='Reference To'
                $mysheet.Cells.Item($row,$column++)='Is External Id'
                $mysheet.Cells.Item($row,$column++)='Is Unique'
                $mysheet.Cells.Item($row,$column++)='Scale'
                $mysheet.Cells.Item($row,$column++)='Precision'
                $mysheet.Cells.Item($row,$column)='Picklist Values'
                $mysheet.UsedRange.Font.Bold=$True
                $sheetsNotAdded = $false
            }
            
            $row++
            $column=1
            $pickval="|"
            $op="$($f.name),$($f.label),$($f.length),$($f.type),$($f.nillable),$($f.referenceTo),$($f.custom),$($f.idLookup),$($f.unique),$($f.scale),$($f.precision)"

            $mysheet.Cells.Item($row,$column++)="$($f.name)"
            $mysheet.Cells.Item($row,$column++)="$($f.label)"
            $mysheet.Cells.Item($row,$column++)="$($f.length)"
            $mysheet.Cells.Item($row,$column++)="$($f.type)"
            $mysheet.Cells.Item($row,$column++)="$($f.nillable)"
            $mysheet.Cells.Item($row,$column++)="$($f.referenceTo)"
            $mysheet.Cells.Item($row,$column++)="$($f.idLookup)"
            $mysheet.Cells.Item($row,$column++)="$($f.unique)"
            $mysheet.Cells.Item($row,$column++)="$($f.scale)"
            $mysheet.Cells.Item($row,$column++)="$($f.precision)"

            if($f.type -eq "picklist"){
                foreach($val in $f.picklistValues){
                    $pickval="$($pickval)$($val.label)|"
                }
                $op="$($op)$($pickval)"
                $mysheet.Cells.Item($row,$column++)="$($pickval)"
            }
            
            $usedRange = $mysheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
            $csvop = "$($csvop)`n$($op)"
        }
    }
    <# Write-Output $csvop | out-file "$($sobjname).csv" -encoding utf8 #>
    
}

$configbook.SaveAs(".\ConfigBook.xlsx")
$configexcel.Workbooks.Close()
$configexcel.Quit()