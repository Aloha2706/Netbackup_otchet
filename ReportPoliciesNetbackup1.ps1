
$serverName = "SourServerName"
$ShowOnlyActivePolicies="no" # Так будут собираться только активные политики.
#$ShowOnlyActivePolicies="абракадабра" # А так и активные и нет. 

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$WorkBook = $Excel.Workbooks.Add()
#$WorkBook = $Excel.Workbooks.Open("\D:\DST\Скриптики\NetbacupOtchet.xlsx")
$sheet1 = $WorkBook.Worksheets.Item(1)




#Здесь мы рисуем предварительные красивости в таблице 
Function MakeTable{

    
    $sheet1.range('t:hc').columnwidth = 1
    $sheet1.range('t:aq').Interior.ColorIndex = 15
    $sheet1.range('bp:cm').Interior.ColorIndex = 15
    $sheet1.range('dl:ei').Interior.ColorIndex = 15
    $sheet1.range('fh:ge').Interior.ColorIndex = 15
    $Range = $sheet1.Range("A1","hc400")
    7..12 | ForEach-Object `
    {
        $Range.Borders.Item($_).LineStyle = 1
        $Range.Borders.Item($_).Weight = 2
        $Range.Borders.Item($_).ColorIndex = 24

    }
    $sheet1.Cells.Item('1', 'a')="Policy Name"
    $sheet1.Cells.Item('1', 'b')="Policy Type"
    #$sheet1.Cells.Item('1', 'c')="Policy Type"
    $sheet1.Cells.Item('1', 'c')="Active"
    $sheet1.Cells.Item('1', 'd')="Residence"
    $sheet1.Cells.Item('1', 'e')="Volume Pool"
    $sheet1.Cells.Item('1', 'f')="Client"
    $sheet1.Cells.Item('1', 'g')="Schedule"
    

    $sheet1.range('t1:aq1').Merge()
    $sheet1.range('ar1:bo1').Merge()
    $sheet1.range('bp1:cm1').Merge()
    $sheet1.range('cn1:dk1').Merge()
    $sheet1.range('dl1:ei1').Merge()
    $sheet1.range('ej1:fg1').Merge()
    $sheet1.range('fh1:ge1').Merge()
    $sheet1.range('gf1:hc1').Merge()
    $sheet1.Cells.Item('1', 't')="Воскресенье"
    $sheet1.Cells.Item('1', 'ar')="Понедельник"
    $sheet1.Cells.Item('1', 'bp')="Вторник"
    $sheet1.Cells.Item('1', 'cn')="Среда"
    $sheet1.Cells.Item('1', 'dl')="Четверг"
    $sheet1.Cells.Item('1', 'ej')="Пятница"
    $sheet1.Cells.Item('1', 'fh')="Суббота"
    $sheet1.Cells.Item('1', 'gf')="Воскресенье"

}

# Вызов предварительной информации о политике
Function PolicyFindString($PolicyName){
    $PolicyData = Invoke-Command -ComputerName $serverName -ScriptBlock {bppllist $Using:PolicyName -U}
    $i=$global:i
    ForEach ($String in $PolicyData){
        $i=$global:i
        if( $String -match "Policy Name:"){$sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()}
        if( $String -match "Policy Type:"){$sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()}
        # Здесь мы доходим до строки Active: проверяем активна она или нет, и выходим из функции если неактивна.
        if( $String -match "Active:"){
            
            $sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()
            if( $String.Split(":")[1].Trim() -match $ShowOnlyActivePolicies){
            Return 0}

        }
        if( $String -match "Residence:"){$sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()}
        if( $String -match "Volume Pool:"){$sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()}
        if( $String -match "HW/OS/Client:"){$sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()}
        # доходим до строки Schedule: и выходим из функции
        if( $String -match "Schedule:"){
            $global:j=$j
            #ShedFindString($PolicyName )
            Return  1
            }

        #Получи красивый вывод обработки
        #$String
    }
    $global:j=$j
}



# Вызов информации о расписании и вызов отрисовки
Function ShedFindString($PolicyName){
    $SchedData = Invoke-Command -ComputerName $serverName -ScriptBlock {bpplsched $Using:PolicyName -L}
    
    $SchedData
    #$i=$global:i
    #$j=$global:j
    ForEach ($String in $SchedData){
        if( $String -match "Schedule:"){
        $i++
        if( $j -gt 7){
            $j=$global:j
            #$i++
            $global:i=$i 
             }
        $sheet1.Cells.Item($i,$j++) = $String.Split(":")[1].Trim()
        }
        if( $String -match "Type:"){$sheet1.Cells.Item($i,$j++) = $String}
        if( $String -match "Frequency:"){$sheet1.Cells.Item($i,$j++) = $String}
        if( $String -match "Retention Level:"){$sheet1.Cells.Item($i,$j++) = $String}
        if( $String -match "Residence:"){$sheet1.Cells.Item($i,$j++) = $String}
        if( $String -match "Volume Pool:"){$sheet1.Cells.Item($i,$j++) = $String}




        if( $String  -match '[0-9][0-9][0-9]:[0-9][0-9]:[0-9][0-9]\s{2}[0-9][0-9][0-9]:[0-9][0-9]:[0-9][0-9]\s{3}[0-9][0-9][0-9]:[0-9][0-9]:[0-9][0-9]\s{2}[0-9][0-9][0-9]:[0-9][0-9]:[0-9][0-9]'){
            $TimeStart, $TimeEnd = Pars($string)
            $TimeStart = $TimeStart + 20
            $TimeEnd = $TimeEnd + 20
            MarkCells $TimeStart $TimeEnd 
            #$Range = $sheet1.Range(($global:i -f  $TimeStart),($global:i -f $Offset + $TimeEnd))
            
            #$Range = $sheet1.Range(.Cells($global:i, $TimeStart) , .Cells($global:i, $TimeEnd))

            #$Range.Select()
            #$Range.Interior.ColorIndex = 6
            #$sheet1.Cells.Item($i,$j++)
    
        }
    }
    $global:i=$i
    $global:j=$j
}
# попытка получить день недели выраженный через часы
Function Pars($string){
    $arr = [Regex]::Matches($string,'[0-9][0-9][0-9]')
    $TimeStart = [int]$arr[2].Value
    $TimeEnd = [int]$arr[3].Value
    Return $TimeStart, $TimeEnd
    }

# закрашиваем ячейки в таблице
Function MarkCells($TimeStart, $TimeEnd){
    for($j=$TimeStart;$j -le $TimeEnd ; $j++){
    $sheet1.Cells.Item($i,$j).Interior.ColorIndex = 5
    }
    }



# основной запуск Main
$ListOfAllPolicies = Invoke-Command -ComputerName $serverName -ScriptBlock {bppllist -l} 
    $global:i=1
    $global:j=1
    MakeTable
ForEach ($PolicyName in $ListOfAllPolicies){
        
        $global:i++
        $Active = PolicyFindString($PolicyName )
        #$global:i++
        if ($Active){ShedFindString($PolicyName)}
        #$global:i++
        $global:j=1
        
        }


#$WorkBook.Save()
#$WorkBook.close($true)
#$Excel.Quit()