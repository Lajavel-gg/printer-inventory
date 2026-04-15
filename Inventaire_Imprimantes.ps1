# ============================================
# CONFIGURATION — à modifier avant utilisation
# ============================================
$SearchBase = "OU=TON_OU,OU=Worldwide_Sites,OU=TON_DOMAINE,DC=TON_DOMAINE,DC=pri"
$xlsxPath   = "$env:USERPROFILE\Desktop\Inventaire_Imprimantes_Final.xlsx"
# ============================================

$dejaDansCSV = [System.Collections.Generic.HashSet[string]]::new()
$toutes = [System.Collections.Generic.List[object]]::new()

if (Test-Path $xlsxPath) {
    Write-Host "Fichier existant trouvé, chargement..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($xlsxPath)
    $sheet = $workbook.Sheets.Item(1)
    $row = 2
    while ($sheet.Cells.Item($row, 1).Value2 -ne $null) {
        $nomImprimante = $sheet.Cells.Item($row, 1).Value2.Trim().ToLower()
        $postes = $sheet.Cells.Item($row, 2).Value2 -split ", "
        foreach ($poste in $postes) {
            $poste = $poste.Trim()
            $cle = "$poste|$nomImprimante"
            if ($dejaDansCSV.Add($cle)) {
                $toutes.Add([PSCustomObject]@{
                    Imprimante = $nomImprimante
                    Poste      = $poste
                })
            }
        }
        $row++
    }
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "$($toutes.Count) entrées chargées depuis le fichier existant" -ForegroundColor Yellow
} else {
    Write-Host "Pas de fichier existant, scan complet en cours..." -ForegroundColor Yellow
}

Write-Host "Scan des postes en cours..." -ForegroundColor Yellow

$computers = Get-ADComputer -Filter * -SearchBase $SearchBase | 
    Select-Object -ExpandProperty Name

Write-Host "$($computers.Count) postes à scanner" -ForegroundColor Yellow

$results = Invoke-Command -ComputerName $computers -ThrottleLimit 50 -ErrorAction SilentlyContinue -ScriptBlock {
    $users = Get-ChildItem "REGISTRY::HKEY_USERS" | Where-Object {$_.Name -notlike "*_Classes" -and $_.Name -like "*S-1-5-21*"}
    foreach ($user in $users) {
        $path = "$($user.PSPath)\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
        if (Test-Path $path) {
            $printers = Get-ItemProperty $path
            $printers.PSObject.Properties | Where-Object {$_.Name -like "*zebra*" -or $_.Name -like "*printer*"} | 
            Select-Object @{N="Poste";E={$env:COMPUTERNAME}}, @{N="Imprimante";E={$_.Name}}
        }
    }
}

$exclure = @('redirection', 'redirected', 'brother', 'canon', 'ededoc',
             'ipp printer', 'first floor', 'second floor', 'etiquette',
             '?tiquette', 'zebra grande', 'zebra moyenne', 'impri\',
             'printers\', 'srv-printer\')

foreach ($f in $results) {
    $nomPropre = $f.Imprimante -replace "(?i)\\\\[\w\.\-]+\\", ""
    $nomPropre = $nomPropre.ToLower().Trim().Trim('\')

    $exclureLigne = $false
    foreach ($ex in $exclure) {
        if ($nomPropre -like "*$ex*") { $exclureLigne = $true; break }
    }

    if (-not $exclureLigne -and ($nomPropre -like "*zebra*" -or $nomPropre -like "*printer*")) {
        $cle = "$($f.Poste)|$nomPropre"
        if ($dejaDansCSV.Add($cle)) {
            Write-Host "NOUVEAU : $($f.Poste) -> $nomPropre" -ForegroundColor Green
            $toutes.Add([PSCustomObject]@{
                Imprimante = $nomPropre
                Poste      = $f.Poste
            })
        }
    }
}

$final = $toutes | Group-Object Imprimante | Sort-Object Name | ForEach-Object {
    [PSCustomObject]@{
        "Nom imprimante" = $_.Name
        "Postes"         = ($_.Group.Poste | Sort-Object -Unique) -join ", "
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Sheets.Item(1)
$sheet.Name = "Inventaire"

$sheet.Cells.Item(1,1) = "Nom imprimante"
$sheet.Cells.Item(1,2) = "Postes"

$row = 2
foreach ($ligne in $final) {
    $sheet.Cells.Item($row, 1) = $ligne.'Nom imprimante'
    $sheet.Cells.Item($row, 2) = $ligne.Postes
    $row++
}

$sheet.Columns.AutoFit() | Out-Null
$sheet.ListObjects.Add(1, $sheet.Range("A1:B$($row-1)"), $null, 1) | Out-Null

if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
$workbook.SaveAs($xlsxPath)
$workbook.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "TERMINE ! $($final.Count) imprimantes dans $xlsxPath" -ForegroundColor Green
