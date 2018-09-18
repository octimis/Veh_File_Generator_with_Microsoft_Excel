Attribute VB_Name = "Main"
Sub genVeh()
    Dim StartTime As Double
    StartTime = Timer
    
    Dim importSheet As Worksheet
    Dim tableSheet As Worksheet
    
    Set importSheet = Sheets("Sheet1")
    Set tableSheet = Sheets("Sheet2")
    
    Dim sourceVehCell As Range
    Dim importFilePrefCell As Range
    Dim pitGroupPrefCell As Range
    Dim descriptionPrefCell As Range
    Dim classesCell As Range
        
    'Dim fileAttributeRange As String
    
    Dim filePath As String
    Dim fileName As String
    Dim exportPath As String
    Dim importContent As String
    
    Dim lastRowOfTableSheet As Integer
    lastRowOfTableSheet = tableSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Dim tableIsBuilt As Boolean: tableIsBuilt = False
       
    'Dim arr As Variant
    'tblSheetColumns = attributeNamesArray
    
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    '| Source veh # | Veh # | Import file Prefix | Export file Prefix | PitGroup Prefix | PitGroup Suffix | Driver | Description Prefix | Classes | Category |
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    '|            1 |     2 |                  3 |                  4 |               5 |               6 |      7 |                  8 |       9 |       10 |
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    
    Set sourceVehCell = tableSheet.Cells.Find(What:="Source Veh #", LookIn:=xlFormulas, lookat:=xlPart, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                        SearchFormat:=False).Offset(1, 0)                                 '<<<<<<<<<<<<<< A9
                              
    Set importFilePrefCell = tableSheet.Cells.Find(What:="Import file Prefix", LookIn:=xlFormulas, lookat:=xlPart, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                        SearchFormat:=False).Offset(1, 0)

    Set pitGroupPrefCell = tableSheet.Cells.Find(What:="PitGroup Prefix", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(1, 0)
                    
    Set descriptionPrefCell = tableSheet.Cells.Find(What:="Description Prefix", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(1, 0)
                    
    Set classesCell = tableSheet.Cells.Find(What:="Classes", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(1, 0)
                    
    filePath = importSheet.Range("B1").Value
    exportPath = importSheet.Range("B2").Value
    'fileName = importFilePrefCell.Value & sourceVehCell.Value
    
    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath + "\"
    End If
    
    Call fillTableSheet(tableSheet, sourceVehCell, sourceVehCell.Offset(0, 1), importFilePrefCell, _
                                    pitGroupPrefCell.Offset(0, 1), lastRowOfTableSheet)                     'sourceVeh, vehNum, imFilePrefix '& pitGroupSuffix cells
    
    tableIsBuilt = True
    
    Dim proceedToExport As Boolean: proceedToExport = False

    '<<<<<<<<<<<<<<<<<<<<<<<<<< Break or Pause Program Here, then give user option to resume and export files >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    Call fillImportSheet(importSheet, tableSheet, sourceVehCell, importFilePrefCell, _
                        pitGroupPrefCell, descriptionPrefCell, classesCell, lastRowOfTableSheet, _
                        filePath, exportPath) '10 args
    
    If tableSheet.Cells.Find(What:="Export Table", _
        LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value = "Yes" Then
        Call exportTable(tableSheet, exportPath, lastRowOfTableSheet)
    End If
    
    tableSheet.Range("F1").Value = Round(Timer - StartTime, 2) & " seconds"
    
End Sub


Function fillImportSheet(importSheet As Worksheet, tableSheet As Worksheet, sourceVehCell As Range, importFilePrefCell As Range, _
                        pitGroupPrefCell As Range, descriptionPrefCell As Range, classesCell As Range, _
                        lastRow As Integer, filePath As String, exportPath As String) 'As Variant
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    '| Source veh # | Veh # | Import file Prefix | Export file Prefix | PitGroup Prefix | PitGroup Suffix | Driver | Description Prefix | Classes | Category |
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    '|            1 |     2 |                  3 |                  4 |               5 |               6 |      7 |                  8 |       9 |       10 |
    '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    
    '10 columns' variable to hold value for export
    Dim sourceVeh As String, vehNum As String
    Dim impFilePrefix As String, expFilePrefix As String
    Dim pitGroupPrefix As String, pitGroupSuffix As String
    Dim driver As String
    Dim descriptionPrefix As String
    Dim team As String, classes As String, category As String
    
    'Vba class modules don't support static variables
    Dim sourceVeh_default As String
    Dim impFilePrefix_default As String, expFilePrefix_default As String
    Dim pitGroupPrefix_default As String
    Dim team_default As String
    Dim descriptionPrefix_default As String
    Dim classes_default As String, category_default As String
    
    Dim currentCellReading As Range, descPrefixDelimetersCell As Range
    
    pitGroupPrefix_default = "" 'Failure to specify pitgroup prefix & suffix will result in using existing value in file
    team_default = ""
    descriptionPrefix_default = ""
    classes_default = ""
    category_default = ""

    sourceVeh_default = sourceVehCell.Value
    impFilePrefix_default = importFilePrefCell.Value
    expFilePrefix_default = importFilePrefCell.Offset(, 1).Value
    
    Set descPrefixDelimetersCell = tableSheet.Range("B4")
    
    Set currentCellReading = sourceVehCell '.Offset(1, 0)   '.Offset(1, 0)  is for testing next row before for loop implementation
    Dim lastColumn As Long              'currentCellReading will traverse columns
    lastColumn = tableSheet.Cells(sourceVehCell.Row, Columns.Count).End(xlToLeft).Column
        
    Dim descriptionDelimiter As String
    descriptionDelimiter = ""
    
    Dim arrOfDelimeters As Variant
    arrOfDelimeters = delimArray(descPrefixDelimetersCell.Value)
    
    Dim delimCounter As Integer
    Dim containsDelimiter As Boolean
    
    Dim loopCounter As Integer      'Due to possible blank vehNums
    loopCounter = 1
    
    Dim vCollection As New Collection
    Dim vR As vehRows
    
    Dim differentImpFilePrefix As Boolean
    differentImpFilePrefix = False              'Let's loop know, its gotten to a row with a different car model
    Dim diffExpFilePrefix As Boolean
    diffExpFilePrefix = False
    
    'tableSheet.Range("G1").Value = lastRow
    
    Dim importSheetLastRow As Integer
    importSheetLastRow = importSheet.Range("A" & importSheet.Rows.Count).End(xlUp).Row
    
    'tableSheet.Range("H1").Value = "importSheetLastRow: " & importSheetLastRow
    
    If importSheetLastRow >= 4 Then
        Call clearImportSheet(importSheet, importSheetLastRow)
        'tableSheet.Range("G1").Value = "Ln 140 Is called"   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End If
    
    Dim specifiedExportPrefix As Boolean
    specifiedExportPrefix = False
   
    For i = sourceVehCell.Row To lastRow
    
        Set vR = New vehRows
        
        'Be sure to reset currentCellReading
        Set currentCellReading = sourceVehCell.Cells(i - sourceVehCell.Row + 1, 1)
        
        'Starting at PitGroup Prefix, attributes may start being acquired directly from the file, not from table sheet
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
        '| Source veh # | Veh # | Import file Prefix | Export file Prefix | PitGroup Prefix*| PitGroup Suffix | Driver | Description Prefix | Classes | Category |
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
        '|            1 |     2 |                  3 |                  4 |               5 |               6 |      7 |                  8 |       9 |       10 |
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+

        'Column 1: Source veh #
        If currentCellReading.Value = "" Or IsEmpty(currentCellReading) Then
            sourceVeh = sourceVeh_default
            'tableSheet.Range("L" & i).Value = "is called, iter: " & i & " sourceVeh default: " & sourceVeh_default
        Else
            sourceVeh = currentCellReading.Value
            sourceVeh_default = sourceVeh
        End If
        
        'tableSheet.Range("N" & i).Value = sourceVeh
        
        'Reposition to cell to the right of
        Set currentCellReading = currentCellReading.Offset(, 1)
     
        'Column 2: Veh #
'        If currentCellReading.Value = "" Or IsEmpty(currentCellReading) Then
'            'tableSheet.Range("K" & i).Value = "veh <> called: " & i
'            'GoTo NextIteration
'            'vehNum = "ether"
'        Else
            vehNum = currentCellReading.Value
'        End If
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 3: Import file Prefix
        If currentCellReading = "" Or currentCellReading = impFilePrefix_default Then
            impFilePrefix = impFilePrefix_default
        Else
            impFilePrefix = currentCellReading.Value
            impFilePrefix_default = impFilePrefix
            differentImpFilePrefix = True
            specifiedExportPrefix = False
        End If
        
        '**********************************************************************************
        'Reset defaults for Team, Desc, Classes, Category
        If differentImpFilePrefix Then
            'pitGroupPrefix_default = "" 'Failure to specify pitgroup prefix & suffix will result in using existing value in file
            team_default = ""
            descriptionPrefix_default = ""
            'classes_default = ""           'It's possible for different car models to share the same class.
            category_default = ""
        End If
        '**********************************************************************************
        'tableSheet.Range("K" & loopCounter + 8).Value = "impFilePrefix: " & impFilePrefix & ", impDefault: " & impFilePrefix_default
        
        Set currentCellReading = currentCellReading.Offset(, 1)

        'Column 4: Export file Prefix
        
        'Will be reset to false each time, impFilePrefix changes
        If currentCellReading.Value <> "" Then
            specifiedExportPrefix = True
        End If
        
        If currentCellReading.Value = "" Or IsEmpty(currentCellReading) Then
            If expFilePrefix_default = "" Then
                If differentImpFilePrefix Then
                    expFilePrefix = impFilePrefix_default
                    expFilePrefix_default = expFilePrefix
                    differentImpFilePrefix = False
                    'tableSheet.Range("M" & loopCounter + 8).Value = "ln 222, counter: " & loopCounter & ", impFilePrefix_default: " & impFilePrefix_default & ", expFilePref: " & expFilePrefix
                Else
                    expFilePrefix = impFilePrefix_default
                    expFilePrefix_default = expFilePrefix
                End If
            Else
                'tableSheet.Range("N" & i).Value = "ln 234, expFilePrefix: " & expFilePrefix & ", impFilePrefix_default: " & impFilePrefix_default
                If differentImpFilePrefix And Not (specifiedExportPrefix) Then
                    expFilePrefix = impFilePrefix_default
                    expFilePrefix_default = expFilePrefix
                    differentImpFilePrefix = False
                    'tableSheet.Range("M" & loopCounter + 8).Value = "ln 235, counter: " & loopCounter & ", impFilePrefix_default: " & impFilePrefix_default & ", expFilePref: " & expFilePrefix
                Else
                    expFilePrefix = expFilePrefix_default
                End If
            End If
        Else
            expFilePrefix = currentCellReading.Value
            expFilePrefix_default = expFilePrefix
        End If
                
        
        'tableSheet.Range("L" & loopCounter + 8).Value = "ln 246 expFilePrefix: " & expFilePrefix & ", expDefault: " & expFilePrefix_default

        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 5: PitGroup Prefix
        If currentCellReading = "" Then
            pitGroupPrefix = pitGroupPrefix_default
        Else
            pitGroupPrefix = currentCellReading.Value
            pitGroupPrefix_default = pitGroupPrefix
        End If
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 6: PitGroup Suffix
        pitGroupSuffix = currentCellReading
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 7: Driver, class module will use "import sheet's" driver cell value
        driver = currentCellReading
        
        Set currentCellReading = currentCellReading.Offset(, 1)
               
        'Column 8: Team
        If currentCellReading = "" Or IsEmpty(currentCellReading) Then
            team = team_default
        Else
            team = currentCellReading
            team_default = team
        End If
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 9: Description Prefix
        If currentCellReading = "" Or IsEmpty(currentCellReading) Then
            descriptionPrefix = descriptionPrefix_default       'set to ""
        Else
            descriptionPrefix = currentCellReading
            descriptionPrefix_default = descriptionPrefix
        End If
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 10: Classes
        If currentCellReading = "" Or IsEmpty(currentCellReading) Then
            classes = classes_default
        Else
            classes = currentCellReading
            classes_default = classes
        End If
        
        Set currentCellReading = currentCellReading.Offset(, 1)
        
        'Column 11: Category
        If currentCellReading = "" Or IsEmpty(currentCellReading) Then
            category = category_default
        Else
            category = currentCellReading
            category_default = category
        End If
        
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
        '| Source veh # | Veh # | Import file Prefix | Export file Prefix | PitGroup Prefix | PitGroup Suffix | Driver | Description Prefix | Classes | Category |
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
        '|            1 |     2 |                  3 |                  4 |               5 |               6 |      7 |                  8 |       9 |       10 |
        '+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
                
        delimCounter = 0
        containsDelimiter = False
        'examine beginning or end of string for '#' character
        Do While containsDelimiter = False And delimCounter < arrayLen(arrOfDelimeters)
            'If arrOfDelimeters(delimCounter) = Right(descriptionPrefix, 1) Or arrOfDelimeters(delimCounter) = Left(descriptionPrefix, 1) Then
            If InStr(descriptionPrefix, descriptionDelimiter_c) > 0 Then
                containsDelimiter = True
                descriptionDelimiter = arrOfDelimeters(delimCounter)
            End If
            delimCounter = delimCounter + 1
        Loop
        
        'This increased the processing time from 9 to 12 seconds, when generating 32 files
        If (tableSheet.Range("D4").Value = "Yes") Then
            descriptionDelimiter = arrOfDelimeters(0)
        End If
        
        'tableSheet.Range("L" & loopCounter).Value = "contains delimiter: " & containsDelimiter & ", loop#: " & loopCounter & ", desc: " & descriptionPrefix   '<<<<<<<<<<<<<<<<<<<<<<<
        
        If vehNum <> "" Then
            'vR instances will likely need to extract unspecified attribute from the text file
            '**************************************************************************************
                                                                                                 '|
            If (Left(impFilePrefix, 1) = "_") Then
                Call importVehData(importSheet, filePath, sourceVeh + impFilePrefix + ".veh") '  '|
            Else
                Call importVehData(importSheet, filePath, impFilePrefix + sourceVeh + ".veh")    '|
            End If
                                                                                                 '|
            '**************************************************************************************
            
            vR.InitializeAttributes sourceVeh_c:=sourceVeh, vehNum_c:=vehNum, impFilePrefix_c:=impFilePrefix, expFilePrefix_c:=expFilePrefix, _
                                    pitGroupPrefix_c:=pitGroupPrefix, pitGroupSuffix_c:=pitGroupSuffix, driver_c:=driver, team_c:=team, descriptionPrefix_c:=descriptionPrefix, _
                                    classes_c:=classes, category_c:=category, descriptionDelimiter_c:=descriptionDelimiter, importSheet_c:=importSheet, _
                                    tableSheet_c:=tableSheet, index_c:=loopCounter, containsDelimiter_c:=containsDelimiter
            
            'Prevent from having to reload text files again, for each vR instance
            '*******************************************************************************************************************************************
                    
            vR.saveSheet = importSheet.Range(importSheet.Cells(1, 1), importSheet.Cells(importSheet.Range("A" & Rows.Count).End(xlUp).Row, 3)).Value  '|
            
            '*******************************************************************************************************************************************
                    
            vCollection.Add vR
            
            Call clearImportSheet(importSheet, vR.lastRowImportSheet)
            loopCounter = loopCounter + 1   'index counter
        End If
'NextIteration:
    Next i
    
    Call transferToImportSheet(importSheet, vR, vCollection, exportPath)
    
    If tableSheet.Cells.Find(What:="Export PitGroup Order", _
        LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value = "Yes" Then
        Call exportPitGrouping(tableSheet, vR, vCollection, exportPath)
    End If
End Function

Function exportTable(tableSheet As Worksheet, exportPath As String, lastRow As Integer)
    Dim lastColumn As Integer
    With tableSheet
        lastColumn = .Cells(8, .Columns.Count).End(xlToLeft).Column
    End With
    
    Dim exportTarget As String
    
    exportTarget = exportPath + "Table_Sheet_Entries" + ".xlsx"
    
    Dim Rng As Range
    Set Rng = tableSheet.Range(Cells(1, 1), Cells(lastRow, lastColumn))
    
    Set xls = CreateObject("Excel.Application")
    xls.DisplayAlerts = False
    
    Dim NewBook As Workbook
    Set NewBook = xls.Workbooks.Add
    
    Rng.Copy
    NewBook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValues)
    NewBook.SaveAs fileName:=exportTarget
    NewBook.Close (True)
End Function

Function clearImportSheet(importSheet As Worksheet, lastRowOfSheet As Integer)
    importSheet.Range("A4:C" & lastRowOfSheet).ClearContents
End Function

Function arrayLen(regularArray As Variant) As Integer
    arrayLen = UBound(regularArray) - LBound(regularArray) + 1
End Function

Function transferToImportSheet(importSheet As Worksheet, vR As vehRows, ByVal vCollection As Collection, exportPath As String)
    For Each vR In vCollection
        
        importSheet.Range("A1:C" & vR.lastRowImportSheet).Value = vR.saveSheet
        
        With importSheet.Cells
            .Find(What:="DefaultLivery=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    vR.defaultLivery(vR.expFilePrefixStr)
            .Find(What:="Number=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    "'" & vR.vehNum()
            .Find(What:="PitGroup=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    "'" & vR.pitGroup(vR.pitGroupPrefix)
            .Find(What:="Driver=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                   "'" & vR.driver()
            .Find(What:="Team=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                   "'" & vR.team()
            .Find(What:="Description=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    "'" & vR.description(vR.descriptionPrefix, vR.descriptionDelimiter)
            .Find(What:="Classes=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    "'" & vR.classes()
            .Find(What:="Category=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = _
                    "'" & vR.category()
        End With
        
        vR.saveSheet = importSheet.Range(importSheet.Cells(1, 1), importSheet.Cells(importSheet.Range("A" & Rows.Count).End(xlUp).Row, 3))
        
        Call exportVehRow(importSheet, vR, importSheet.Range("A" & Rows.Count).End(xlUp).Row, exportPath)
        
        Call clearImportSheet(importSheet, vR.lastRowImportSheet)
    Next vR
End Function

Function exportVehRow(importSheet As Worksheet, vehRow As vehRows, lastRow As Integer, exportPath As String)
    Dim fileName As String, exportTarget As String
    'Dim FF As Integer
    
    Dim lastColumn As Integer
    lastColumn = 3 'importSheet.Range("A4").Offset(, 2).Column
    
    fileName = vehRow.exportFileName(vehRow.expFilePrefixStr)
    
    exportTarget = exportPath + fileName
    
    Dim X As Long, Z As Long, FF As Long, TextOut As String

    Const StartColumn As Long = 1   'Column A
    Const startRow As Long = 4
    
    FF = FreeFile
    Open exportTarget For Output As #FF
    For X = startRow To lastRow
        TextOut = ""
        For Z = StartColumn To lastColumn
          TextOut = TextOut & importSheet.Cells(X, Z).Value
        Next
        Print #FF, TextOut
    Next
    Close #FF
End Function

Function exportPitGrouping(tableSheet As Worksheet, vR As vehRows, vCollection As Collection, exportPath As String)
    Dim currentPitGroup As String
    currentPitGroup = ""
    
    Dim pgCounter As Integer, totalGroups As Integer
    pgCounter = 0
    totalGroups = 1
    
    exportTarget = exportPath + "Pit_Group_Order.txt"
    
    FF = FreeFile
    Dim TextOut As String
    
    Open exportTarget For Output As #FF
    
    TextOut = "PitGroupOrder" & vbNewLine & "{"
    Print #FF, TextOut
    
    For Each vR In vCollection
        TextOut = ""
        If currentPitGroup <> "" And currentPitGroup <> vR.pitGroup(vR.pitGroupPrefix) Then
            TextOut = vbTab & "PitGroup = " & pgCounter & ", " & Replace(currentPitGroup, """", "")  'Drop quotations
            Print #FF, TextOut
            pgCounter = 1
            totalGroups = totalGroups + 1
        Else
            pgCounter = pgCounter + 1
        End If
        
        currentPitGroup = vR.pitGroup(vR.pitGroupPrefix)
        
        If vR.index = vCollection.Count Then
            TextOut = vbTab & "PitGroup = " & pgCounter & ", " & Replace(currentPitGroup, """", "") & vbNewLine & "}" & " // Total PitGroups: " & totalGroups
            Print #FF, TextOut;
        End If
    Next vR
    Close #FF
End Function

Function delimArray(delimChars As String) As Variant
    Dim rawArray() As String
    
    If InStr(delimChars, " ") > 0 Then
        Do While (InStr(delimChars, " ") > 0)
            delimChars = Replace(delimChars, " ", "")
        Loop
    End If
    
    rawArray = Split(delimChars, ",")
    
    delimArray = rawArray
End Function
Function fillTableSheet(tableSheet As Worksheet, sourceVehCell As Range, VehNumCell As Range, importFilePrefCell As Range, pitGroupSufxCell As Range, lastRowOfSourceVeh As Integer) As Variant
    Dim sourceVehVal As String
    Dim displayVNSC As String
    
    Dim sourceVehTeamsCounter As Integer
    
    Dim genVehOption As String:
    
    
    If tableSheet.Cells.Find(What:="Guess Veh #", _
        LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value <> "No" Then
        
        'tableSheet.Range("L2").Value = "ln 545"
        
        Call guessVehNum(sourceVehCell, VehNumCell, lastRowOfSourceVeh)
    End If
    
    If tableSheet.Cells.Find(What:="Generate Veh #", _
        LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value <> "No" Then
        'Call populateVehNum(tableSheet, sourceVehCell, VehNumCell, lastRowOfSourceVeh, CInt(sourceVehCell.Value))       '3rd arg
        Call populateVehNum(tableSheet, sourceVehCell, VehNumCell, lastRowOfSourceVeh, sourceVehCell.Value)       '3rd arg
    End If
    
    Dim pitGroupOption As String: pitGroupOption = tableSheet.Cells.Find(What:="Generate Pitgroup", LookIn:=xlValues, lookat:=xlWhole, _
                                                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value
    
    If Not (tableSheet.Cells.Find(What:="Generate Pitgroup", _
        LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value = "No") Then
        Call populatePitGroupSufx(tableSheet, sourceVehCell, VehNumCell, importFilePrefCell, pitGroupSufxCell, lastRowOfSourceVeh)     '5th arg
    End If
End Function

Function guessVehNum(sourceVehCell As Range, VehNumCell As Range, lastRowOfSourceVeh As Integer)
    Dim startRow As Integer: startRow = sourceVehCell.Row
    
    For i = startRow To lastRowOfSourceVeh
    
        If (sourceVehCell.Value <> "") And (CStr(sourceVehCell.Offset(-1, 0).Value) <> CStr(sourceVehCell.Value)) Then
        
            VehNumCell.Value = extractVehNum(sourceVehCell.Value, "")
            
            If (InStr(LCase(VehNumCell.Value), "aston") > 0) And (InStr(LCase(VehNumCell.Value), "martin") > 0) And _
                Left(VehNumCell.Value, 1) = "0" Then
                VehNumCell.Value = "0" + VehNumCell.Value
            End If
        End If
        
        Set sourceVehCell = sourceVehCell.Offset(1, 0)
        Set VehNumCell = VehNumCell.Offset(1, 0)
    Next i
    
    Set sourceVehCell = sourceVehCell.Offset(startRow - lastRowOfSourceVeh - 1, 0)     'Must reset position, caused problems in populate vehNum
    Set VehNumCell = VehNumCell.Offset(startRow - lastRowOfSourceVeh - 1, 0)
    
'    Dim tableSheet As Worksheet: Set tableSheet = Sheets("Sheet2")
'    tableSheet.Range("L1").Value = "Ln 585 is called, VehNumCell.Row: " & VehNumCell.Row & ", value: " & VehNumCell.Value
    
End Function
Function extractVehNum(sourceVeh As String, rightOf_sVeh As String) As String
    If Len(sourceVeh) = 0 Or Len(rightOf_sVeh) = 3 Or (Len(rightOf_sVeh) = 2 And Right(sourceVeh, 1) = 0) Then
        extractVehNum = rightOf_sVeh
    Else
        Do While Not (IsNumeric(Right(sourceVeh, 1)))
            sourceVeh = Left(sourceVeh, Len(sourceVeh) - 1)
        Loop
        extractVehNum = extractVehNum(Left(sourceVeh, Len(sourceVeh) - 1), Right(sourceVeh, 1) + rightOf_sVeh)
    End If
End Function

Function populateVehNum(tableSheet As Worksheet, sourceVehCell As Range, currentCellTraversing As Range, lastRow As Integer, currentSourceVehNum As String) 'cCT is VehNumCell
    Dim sourceVehCellPVN As Range           'The row # of sourceVehCell at the end of this function, will have an unwanted effect on other functions
    Set sourceVehCellPVN = sourceVehCell    'e.g. other functions will start sourceVeh cell at row 40 instead of row 8
    
    Dim vehNum As String: vehNum = ""
    Dim sourceVehTeamsCounter As Integer
    
    sourceVehTeamsCounter = 0
    
    'tableSheet.Range("L3").Value = "Ln 606, is called, currentCellTraversing.Row: " & currentCellTraversing.Row & ", value: " & currentCellTraversing.Value
    
    Dim lastVehNum As String: lastVehNum = ""
    
    Dim doWhileCounter_diagnostics As Integer: doWhileCounter_diagnostics = 0
        
    For i = currentCellTraversing.Row To lastRow
            'tableSheet.Range("L2").Value = "ln 612, currentCellTraversing.Value: " & currentCellTraversing.Value & ", row: " & currentCellTraversing.Row & ", col: " & currentCellTraversing.Column
                  
        'Assign last vehNum
        If currentCellTraversing.Value <> "" Then
            lastVehNum = currentCellTraversing.Value
            
            'tableSheet.Range("L" & i).Value = "Ln 616 called, iteration: " & Iteration & ", lastVehNum: " & lastVehNum
        ElseIf IsNumeric(currentSourceVehNum) Then
            
            If tableSheet.Range("B1").Value = "No" Then
                lastVehNum = CInt(currentSourceVehNum)
            Else
                lastVehNum = CInt(currentSourceVehNum) - 1
            End If
            
            If (isLenDifferentFromDerivedVehNum(currentSourceVehNum, lastVehNum)) Then
                lastVehNum = PadVehNumLength(Len(currentSourceVehNum), lastVehNum)
            End If
        Else
            MsgBox "Auto-generate Veh# num Error." & vbNewLine & "Please examine row: " & i - 1
            End
        End If
        
        Do While (currentCellTraversing.Offset(, -1).Value = currentSourceVehNum Or currentCellTraversing.Offset(, -1).Value = "") _
            And Not (currentCellTraversing.Row > lastRow)
            
            If tableSheet.Range("B1").Value = "No" And sourceVehTeamsCounter = 0 Then       'may substiture with boolean
                currentCellTraversing.Value = ""
            Else
                If currentCellTraversing.Value = "" Then
                    vehNum = CStr(CInt(lastVehNum) + 1)            'be sure to update lastVehNum
                    
                    If (isLenDifferentFromDerivedVehNum(lastVehNum, vehNum)) Then
                        vehNum = PadVehNumLength(Len(lastVehNum), vehNum)
                    End If
                    
                    currentCellTraversing.Value = vehNum
                    lastVehNum = vehNum
                End If
                
            End If
        
            Set currentCellTraversing = currentCellTraversing.Offset(1, 0)      'Reposition to cell below current cell
            
            sourceVehTeamsCounter = sourceVehTeamsCounter + 1
            i = currentCellTraversing.Row
            
            doWhileCounter_diagnostics = doWhileCounter_diagnostics + 1
            
        Loop
        'tableSheet.Range("O" & i).Value = "Ln 647, iteration: " & i & ", sourceVehNum: " & currentSourceVehNum & ", sourceVehTeamsCounter: " & sourceVehTeamsCounter & _
                                    ", doWhileCounter: " & doWhileCounter_diagnostics & ", currentCellTraversing.Row: " & currentCellTraversing.Row
           
        Set sourceVehCellPVN = currentCellTraversing.Offset(, -1)               'Select cell to the left of current traversing
        currentSourceVehNum = sourceVehCellPVN.Value
        
        lastVehNum = ""
        sourceVehTeamsCounter = 0
    Next i
End Function

Function isLenDifferentFromDerivedVehNum(sourceVehNum As String, vehNum As String) As Boolean
    If Len(vehNum) < Len(sourceVehNum) Then
        isLenDifferentFromDerivedVehNum = True
    Else
        isLenDifferentFromDerivedVehNum = False
    End If
End Function
Function PadVehNumLength(sourceVehNumLen As Integer, vehNum As String) As String
    Dim LenDifference As Integer
    LenDifference = sourceVehNumLen - Len(vehNum)
    vehNum = Format(vehNum, String(LenDifference + 1, "0"))
    PadVehNumLength = vehNum
End Function

Function populatePitGroupSufx(tableSheet As Worksheet, sourceVehCell As Range, VehNumCell As Range, impFilePrefixCell As Range, currentCellTraversing As Range, lastRow As Integer)
    Dim currentPitPrefix As String
    
    Dim currentSourceVehNum As String: currentSourceVehNum = sourceVehCell.Value
    Dim currentImpFilePrefix As String: currentImpFilePrefix = impFilePrefixCell.Value
    
    Dim suffixCounter As Integer
    Dim sourceVehCol As Integer, vehNumCol As Integer
    
    suffixCounter = 1
    sourceVehCol = sourceVehCell.Column - currentCellTraversing.Column '1 - 6 = -5
    vehNumCol = VehNumCell.Column - currentCellTraversing.Column '2-6 = -4
    importFilePrefCol = impFilePrefixCell.Column - currentCellTraversing.Column '3-6 = -3
    
    currentPitPrefix = currentCellTraversing.Offset(, -1).Value
    
    Dim doWhileCounter As Integer
    
    Dim withinGroupingRange As Boolean: withinGroupingRange = True       'Set to false if next row doesn't share same source veh or same import file prefix, depending on cell B3
    
    Dim pitGroupingOption As String: pitGroupingOption = tableSheet.Cells.Find(What:="Generate Pitgroup", LookIn:=xlValues, lookat:=xlWhole, _
                                                        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value
    
    For i = currentCellTraversing.Row To lastRow
        
        doWhileCounter = 1
        
        Dim vehNum As String
        
        vehNum = currentCellTraversing.Offset(0, vehNumCol)
        
'       tableSheet.Range("K" & I).Value = "Ln 619, iteration: " & I & ", sourceVehNum: " & currentSourceVehNum & ", suffixCounter: " & suffixCounter & _
'                                        ", doWhileCounter: " & doWhileCounter & ", currentCellTraversing ROW: " & currentCellTraversing.Row
                                                
        i = currentCellTraversing.Row       'may want to place this on the outside

        Do While (withinGroupingRange Or _
                        IsEmpty(currentCellTraversing.Offset(, sourceVehCol).Value) Or _
                        isShareSameFillColor(currentCellTraversing.Offset(-1, sourceVehCol), currentCellTraversing.Offset(, sourceVehCol))) And _
                        (doWhileCounter <= CInt(tableSheet.Range("D3").Value) Or isShareSameFillColor(currentCellTraversing.Offset(-1, sourceVehCol), currentCellTraversing.Offset(, sourceVehCol))) And _
                        Not (currentCellTraversing.Row > lastRow)
'           tableSheet.Range("L" & I).Value = "Ln 625, iteration: " & I & ", sourceVehNum: " & currentSourceVehNum & ", suffixCounter: " & suffixCounter & _
'                                        ", doWhileCounter: " & doWhileCounter & ", currentCellTraversing ROW: " & currentCellTraversing.Row
            
            If currentCellTraversing.Offset(, vehNumCol) <> "" Then
                currentCellTraversing.Value = suffixCounter
                doWhileCounter = doWhileCounter + 1
            End If
            
            Set currentCellTraversing = currentCellTraversing.Offset(1, 0)      'Reposition to cell below current cell
            
            'Loop is terminated based on user's selection and either of the following set to false
            Select Case pitGroupingOption
                Case "Per Source Veh"
                    If currentCellTraversing.Offset(, sourceVehCol).Value <> currentSourceVehNum And currentCellTraversing.Offset(, sourceVehCol).Value <> "" Then
                        withinGroupingRange = False
                    End If
                Case "Per Import Prefix"
                    If currentCellTraversing.Offset(, importFilePrefCol).Value <> currentImpFilePrefix And currentCellTraversing.Offset(, importFilePrefCol).Value <> "" Then
                        withinGroupingRange = False
                    End If
            End Select
        Loop
        
        withinGroupingRange = True
        
        suffixCounter = suffixCounter + 1       'This stays the same, if the sourceVeh/impFilePref AND pitgroup prefix stays the same
       
'        tableSheet.Range("N" & I).Value = "Ln 651, iteration: " & I & ", sourceVehNum: " & currentSourceVehNum & ", suffixCounter: " & suffixCounter & _
'                                        ", doWhileCounter: " & doWhileCounter & ", currentCellTraversing ROW: " & currentCellTraversing.Row
        'Re-check sourceVehNum
        'If Not (IsEmpty(currentCellTraversing.Offset(, sourceVehCol))) And CInt(currentCellTraversing.Offset(, sourceVehCol).Value) Then
        If Not (IsEmpty(currentCellTraversing.Offset(, sourceVehCol))) Then     '<<<<<<<<<<<< Include an statement that will anylyze drop down option at cell B3
            currentSourceVehNum = currentCellTraversing.Offset(, sourceVehCol).Value
        End If                                                                  '<<<<<<<<<<<< ElseIf check to see ImpFilePrefix is the same as above, if not re-set sVehNum
     
        'Reset importFilePrefCell
        If Not (IsEmpty(currentCellTraversing.Offset(, importFilePrefCol))) Then     '<<<<<<<<<<<< Include an statement that will anylyze drop down option at cell B3
            currentImpFilePrefix = currentCellTraversing.Offset(, importFilePrefCol).Value
        End If
     
        'Check if PitGroup Prefix has changed
        If Not (IsEmpty(currentCellTraversing.Offset(, -1))) And currentCellTraversing.Offset(, -1).Value <> currentPitPrefix Then
            currentPitPrefix = currentCellTraversing.Offset(, -1).Value
            suffixCounter = 1
        End If
        
'        tableSheet.Range("O" & I).Value = "Ln 665, iteration: " & I & ", sourceVehNum: " & currentSourceVehNum & ", suffixCounter: " & suffixCounter & _
'                                        ", doWhileCounter: " & doWhileCounter & ", currentCellTraversing ROW: " & currentCellTraversing.Row
    Next i
End Function
Function isShareSameFillColor(previousCell As Range, currentCell As Range) As Boolean
    If (previousCell.Interior.Color <> 16777215 And previousCell.Interior.Color = currentCell.Interior.Color) Then
        isShareSameFillColor = True
    Else
        isShareSameFillColor = False
    End If
End Function


Function importVehData(importSheet As Worksheet, filePath As String, fileName As String)
    'Windows explorer address bar doesn't include '\' at the end of path
    If Right(filePath, 1) <> "\" Then
        filePath = filePath + "\"
    End If
    
    If fileName = "C6R_14.veh" Then
        Stop
    End If
    
    Dim fileTarget As String
    fileTarget = filePath + fileName

    If FileExists(fileTarget) Then
        With importSheet '"Sheet1"
            .Rows("4:" & Rows.Count).Clear
            With .QueryTables.Add(Connection:="TEXT;" & fileTarget, Destination:=.Range("A4"))
                .Refresh BackgroundQuery:=False
                .Delete
            End With
        End With
    Else
        MsgBox fileName & ", doesn't exist."
        End
    End If
    
    Call arrangeContent(importSheet)
End Function

Public Function FileExists(ByVal path_ As String) As Boolean
    FileExists = (Len(Dir(path_)) > 0)
End Function

Function arrangeContent(importSheet As Worksheet) As Variant
    Dim b4_formula As String
    Dim c4_formula As String
    Dim d4_formula As String
    
    Dim lastRowNum As Integer
    lastRowNum = importSheet.Range("A" & Rows.Count).End(xlUp).Row
        
    b4_formula = "=IF(ISBLANK(A4),"""",IF(NOT(ISNUMBER(FIND(""="",A4))),A4,IF(NOT(OR(ISNUMBER(FIND(""//"",A4)),ISNUMBER(FIND("" "",A4)))),LEFT(A4,SEARCH(""="",A4)),IF(NOT(ISNUMBER(FIND(""//"",A4))), LEFT(SUBSTITUTE(A4, "" "", """"), SEARCH(""="", SUBSTITUTE(A4, "" "", """"))),IF(FIND(""="",A4)<FIND(""//"",A4), LEFT(SUBSTITUTE(A4, "" "", """"), SEARCH(""="",SUBSTITUTE(A4, "" "", """"))), A4)))))"
    c4_formula = "=IF(OR(ISNUMBER(FIND(""//"",B4)),NOT(ISNUMBER(FIND(""="",B4)))),"""",IF(NOT(ISNUMBER(FIND(""//"",A4))),RIGHT(A4,LEN(A4)-SEARCH(""="",A4)),IF(AND(ISNUMBER(FIND(CHAR(34),A4,1)),IFERROR(FIND(CHAR(34),A4,1)<FIND(""//"",A4,1),FALSE)),MID(A4,FIND(CHAR(34),A4,1),FIND(CHAR(34),A4,FIND(CHAR(34),A4,1)+1)-FIND(CHAR(34),A4,1)+1),IF(AND(ISNUMBER(FIND("" "",A4,FIND(""="",A4,1))),IFERROR(FIND("" "",A4,FIND(""="",A4,1))<FIND(""//"",A4,FIND(""="",A4,1)),FALSE)),MID(A4,FIND(""="",A4,1)+1,FIND("" "",A4,FIND(""="",A4,1))-FIND(""="",A4,1)-1),MID(A4,FIND(""="",A4,1)+1,FIND(""//"",A4,FIND(""="",A4,1))-FIND(""="",A4,1)-1)))))"
    d4_formula = "=IF(OR(ISNUMBER(FIND(""//"", B4)), NOT(AND(ISNUMBER(FIND(""//"",A4,1)), IFERROR(FIND(""="",A4,1) < FIND(""//"",A4,FIND(""="",A4,1)),FALSE)))), """", IF(LEN(C4) = 0, RIGHT(A4, LEN(A4)-FIND(""//"", A4, 1)+1), RIGHT(A4, LEN(A4)-FIND(C4, A4, 1)+1-LEN(C4))))"

    
    For i = 4 To lastRowNum
        With importSheet
            .Range("B" & i).Value = WorksheetFunction.Substitute(b4_formula, "A4", "A" & i)
            .Range("B" & i).Value = importSheet.Range("B" & i).Text 'Store displayed value, not formula
            .Range("C" & i).Value = WorksheetFunction.Substitute(WorksheetFunction.Substitute(c4_formula, "A4", "A" & i), "B4", "B" & i)
            .Range("C" & i).NumberFormat = "@" 'prevent leading zeroes from being dropped
            .Range("C" & i).Value = importSheet.Range("C" & i).Text
            .Range("D" & i).Value = WorksheetFunction.Substitute(WorksheetFunction.Substitute(WorksheetFunction.Substitute(d4_formula, "A4", "A" & i), "B4", "B" & i), "C4", "C" & i)
            .Range("D" & i).Value = importSheet.Range("D" & i).Text
            
            'Remove trailing spaces in c4_formula result
            While Right(.Range("C" & i).Value, 1) = " "
                .Range("C" & i).Value = Left(.Range("C" & i).Value, Len(.Range("C" & i).Value) - 1)
            Wend
        End With
    Next i
    
    For i = 4 To lastRowNum
        importSheet.Range("A" & i).Delete Shift:=xlToLeft
    Next i
End Function
