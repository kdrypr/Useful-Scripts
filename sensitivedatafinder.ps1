$Machines = @('192.168.2.76', '192.168.2.5')
$Disks = @('C', 'D')
$TodayDate = Get-Date -Format "dd-MM-yyyy"
$word = New-Object -ComObject Word.Application
$SearchKeywords = ('pass','şifre','parola','kullanıcı','kredi','kart','kimlik','pasaport','cep','telefon')
$TCNO = '[1-9]{1}[0-9]{9}[0-9]{1}'
$KrediKarti = '[0-9]{4}\s?[0-9]{4}\s?[0-9]{4}\s?[0-9]{4}'
$DogumTarihi = '[0-9]{2}\.?\/?\-?[0-9]{2}\.?\/?\-?(19|20)[0-9]{2}'
$CepTel = '05[0-9]{2}\s?[0-9]{3}\s?[0-9]{2}\s?[0-9]{2}'
$RegexPatterns = @($TCNO, $KrediKarti, $DogumTarihi, $CepTel)

Get-Date -DisplayHint Date | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
foreach ($Machine in $Machines) {

    foreach ($Disk in $Disks) {
        "$Machine - $Disk Sonuclari:" | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
        "======================================== " | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
        if (Test-Path \\$Machine\$Disk$) {
            Write-Host Checking $Machine Disk $Disk
            $docs = Get-ChildItem -Path \\$Machine\$Disk$\ -include ('*.docx', '*.doc', '*.txt', '*.xls', '*.xlsx', '*.csv') -Recurse
            foreach ($doc in $docs)
            {

                if (($doc.FullName -Like "*.docx") -or ($doc.FullName -Like "*.doc"))
                {
                    $document = $word.Documents.Open($doc.FullName,$false,$true)
                    $range = $document.Paragraphs

                    ForEach($para in $range){
                        ForEach($regex in $RegexPatterns){
                            if($para.Range.Text -match $regex){
                                $para.Range.Text + " in $doc" | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                            }
                        }
                    }
                    
                    ForEach ($keyword in @($SearchKeywords)){
                        if ($word.Documents.Open($doc.FullName).Content.Find.Execute($keyword))
                        {
                            $word.Application.ActiveDocument.Close()
                            "$doc contains ' $keyword '" | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                        }
                        else
                        {
                            $word.Application.ActiveDocument.Close()
                        }
                    }
        
                }elseif (($doc.FullName -Like "*.xls") -or ($doc.FullName -Like "*.xlsx"))
                {
                    $excel = New-Object -ComObject Excel.Application
                    $filePath = $doc.FullName
                    $workbook = $excel.Workbooks.Open($filePath)
                    $sheetSize = $workbook.Sheets.Count
                    for ($i = 1; $i -le $sheetSize; $i++){
                        $sheet = $workbook.Sheets.Item($i)
                        for ($column = 0; $column -lt 500; $column++)
                        {
                            for ($row = 0; $row -lt 500; $row++)
                            {
                                try {
                                    if(![string]::IsNullOrEmpty($sheet.Cells.Item($column,$row).Text)){
                                        if($sheet.Cells.Item($column,$row).Text | Select-String -Pattern $SearchKeywords -AllMatches){
                                           $sheet.Cells.Item($column,$row).Text + " in " + $doc.FullName + " - " + $sheet.Name | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                                        }
                                         
                                        ForEach ($regex in @($RegexPatterns)){
                                            if ($sheet.Cells.Item($column,$row).Text -match $regex){
                                                $sheet.Cells.Item($column,$row).Text + " in $doc - " + $sheet.Name | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                                            }
                                        }
                                    }
                                }catch {}
                            }
                        }
                    }
                }elseif (($doc.FullName -Like "*.txt") -or ($doc.FullName -Like "*.csv"))
                {
                    Get-ChildItem $doc.FullName | Select-String -Pattern $SearchKeywords -AllMatches | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                    ForEach ($regex in @($RegexPatterns)){
                         Get-ChildItem $doc.FullName | Select-String -Pattern $regex -AllMatches | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
                    }

                }
    
            }
        } else {
            "\\$Machine\\$Disk$ path not found!" | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
        }
        "######################################## " | Out-File $HOME\Desktop\cikti\$TodayDate.txt -Append
    }
}
