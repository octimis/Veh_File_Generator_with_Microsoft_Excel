Attribute VB_Name = "ReadFileDir"
Sub get_SourceVehs()
    Dim tableSheet As Worksheet: Set tableSheet = Sheets("Sheet2")
    Dim importSheet As Worksheet: Set importSheet = Sheets("Sheet1")
       
    Dim sourceVehCell As Range, importFilePrefCell As Range, importFilePathCell As Range
        
    Set sourceVehCell = tableSheet.Cells.Find(What:="Source veh #", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(1, 0)
                    
    Set importFilePrefCell = tableSheet.Cells.Find(What:="Import file Prefix", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(1, 0)

    Set importFilePathCell = importSheet.Cells.Find(What:="Import Path", LookIn:=xlFormulas, lookat:=xlPart, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).Offset(0, 1)
    
    'Much more dynamic than: lastRowOfTableSheet = tableSheet.Range("C" & Rows.Count).End(xlUp).Row
    Dim lastRowOfSheet As Integer: lastRowOfSheet = tableSheet.Cells(Rows.Count, importFilePrefCell.Column).End(xlUp).Row
    'Dim lastRowOfSheet As Integer: lastRowOfSheet = importFilePrefCell.Row + fileNamesColl.Count - 1
    
    'We don't know amount of unique impFile Prefix hence a dynamic collection
    Dim impPrefixColl As New Collection: Set impPrefixColl = buildPrefixCollection(tableSheet, importFilePrefCell, lastRowOfSheet)
    
    If (impPrefixColl.Count = 0) Then
        MsgBox "You haven't specified any " & vbNewLine & "Import File Prefixes..."
        End
    End If
    
    'Dim filePath As String: filePath = tableSheet.Range("B1").Value + "\"
    Dim fileNamesColl As New Collection: Set fileNamesColl = buildFileNameCollection(tableSheet, importFilePathCell.Value)
        
    Dim matchingPrefixFound As Boolean: matchingPrefixFound = False
    Dim mpCounter As Integer: mpCounter = 1
    
    Do While Not (matchingPrefixFound) And mpCounter <= impPrefixColl.Count
        For i = 1 To fileNamesColl.Count
            If InStr(fileNamesColl(i), impPrefixColl(mpCounter)) > 0 Then
                matchingPrefixFound = True
                GoTo ExitForLoop
            End If
        Next i
ExitForLoop:
        mpCounter = mpCounter + 1
    Loop
    
    If Not (matchingPrefixFound) Then
        MsgBox "No matching, 'File Prefix', found." & vbNewLine & "Macro terminated."
        End
    End If
    
    Dim sviColl As New Collection   'Without 'New', 'object variable not set error' originates
    Dim svim As sVehImport
        
    Dim fileNameStr As Variant
    
    For i = 1 To fileNamesColl.Count        'collection of filenames
        Set svim = New sVehImport
        svim.InitializeAttributes tableSheet_c:=tableSheet, fileName_c:=fileNamesColl(i), possiblePrefixes:=impPrefixColl, index_c:=CInt(i)
        sviColl.Add svim
    Next i
    
    Call assignToCell(tableSheet, impPrefixColl, sviColl, sourceVehCell, importFilePrefCell, lastRowOfSheet)
End Sub

Function assignToCell(tableSheet As Worksheet, impPrefixColl As Collection, sviColl As Collection, _
                                    sourceVehCell As Range, importFilePrefCell As Range, lastRowOfSheet As Integer)
                                    
    Dim rowToStartFilling As Integer
    Dim allocatedRows As Integer, dwCounter1 As Integer, sviCollIndex As Integer
        
    'Dim q As New queue
    'Dim q As Object
    'Set q = CreateObject("System.Collections.Queue") 'Create the Queue, required .NET Framework
    Dim q As New Queue      'sourceVehs
    Dim q2 As New Queue     'vehNums
    
    'Diagnostics
    'Dim sVeh_searchRng As Range: Set sVeh_searchRng = tableSheet.Range(tableSheet.Cells(impFilePrefix_Cell.Row, 1), tableSheet.Cells(lastRowOfTableSheet, 1))
    Dim prefix_SearchRng As Range: Set prefix_SearchRng = tableSheet.Range(tableSheet.Cells(importFilePrefCell.Row, importFilePrefCell.Column), tableSheet.Cells(lastRowOfSheet, importFilePrefCell.Column))
    
    'Diagnostics
    Dim prefixStartCell As Range: Set prefixStartCell = prefix_SearchRng.Find(What:=impPrefixColl(1), LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False)
    Dim currentCell As Range: Set currentCell = importFilePrefCell

    Dim vehNumColl As New Collection
    Dim vehNumDifference As Integer
    Dim remainingRows As Integer, maxOffset As Integer                            'Allocated spaces remaining divided by remaining sources (integer division rounds down)
    Dim sourceVehLocation As Integer: sourceVehLocation = sourceVehCell.Column - importFilePrefCell.Column
    
    Dim selected_IF As String       'Debugging
    
    For iMain = 1 To impPrefixColl.Count
        allocatedRows = 0
        maxOffset = 0
        
        sviCollIndex = 1
        'After:=prefix_SearchRng(prefix_SearchRng.Count) will allow retrieval first, not 2nd, occurence of the prefix, not the
        rowToStartFilling = prefix_SearchRng.Find(What:=impPrefixColl(iMain), LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False, After:=prefix_SearchRng(prefix_SearchRng.Count)).Row         'Perform a search range & set it to impPrefixColl(i)
        Set currentCell = tableSheet.Cells(rowToStartFilling, importFilePrefCell.Column)
        
        'If iMain = impPrefixColl.Count Then Stop      '<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Debugging
        
        'Count rows allocated to present prefix
        Do While (currentCell.Value = impPrefixColl(iMain) Or currentCell.Value = "") And Not currentCell.Row > lastRowOfSheet
            
            Set currentCell = currentCell.Offset(1, 0)
            allocatedRows = allocatedRows + 1
        Loop
        dwCounter1 = 1
        
        'Reset traversing cell to initial slot of impFilePrefix
        Set currentCell = currentCell.Offset(rowToStartFilling - currentCell.Row, 0)
  
        Do While q.Count <= allocatedRows And dwCounter1 <= sviColl.Count
            
            If sviColl(sviCollIndex).impPrefix(impPrefixColl) = impPrefixColl(iMain) Then       'Will use q.Count to space out the dequeues evenly (where applicable)
                q.Enqueue sviColl(sviCollIndex).sourceVeh()
                q2.Enqueue sviColl(sviCollIndex).possibleVehNum()
                
                vehNumColl.Add sviColl(sviCollIndex).possibleVehNum()
            End If
            
            sviCollIndex = sviCollIndex + 1
            dwCounter1 = dwCounter1 + 1
        Loop
        dwCounter1 = 1
        
        'If iMain = 2 Then Stop      '<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Debugging
        
        Do While q.Count > 0 And dwCounter1 <= allocatedRows
            vehNumDifference = 0
            
            remainingRows = allocatedRows - totalOffset 'dwCounter1 + 1    '20 -  0 = 20 at first iteration         '12 -  0 = 12
                                                                           '20 -  2 = 18                            '12 -  1 = 11
                                                                           '20 -  4 = 16                            '12 -  2 = 10
                                                                           '20 -  6 = 14                            '12 -  3 = 9
                                                                           '20 -  8 = 12                            '12 -  4 = 8
                                                                           '20 - 12 = 8                             '12 -  5 = 7
                                                                           '20 - 14 = 6                             '12 -  7 = 5
                                                                           '20 - 18 = 2                             '12 -  9 = 3
                                                                           
            maxOffset = remainingRows \ q.Count                            '20 \ 8 = 2.5 = 2,                       '12 \ 8 = 1.5 = 1
                                                                           '18 \ 7 = 2.7 = 2                        '11 \ 7 = 1.5 = 1
                                                                           '16 \ 6 = 2.6 = 2                        '10 \ 6 = 1.6 = 1
                                                                           '14 \ 5 = 2.8 = 2                        '9  \ 5 = 1.8 = 1  4
                                                                           '12 \ 4 = 3 -> 4                         '8  \ 4 = 2 <<<<
                                                                           ' 8 \ 3 = 2                              '7  \ 3 = 2
                                                                           ' 6 \ 2 = 3 -> 4                         '5  \ 2 = 2.5 = 2
                                                                           ' 2 \ 1 = 2                              '3  \ 1 = 3
            
            If q.Count > 1 Then
                vehNumDifference = Abs(CInt(vehNumColl(dwCounter1 + 1)) - CInt(vehNumColl(dwCounter1)))
                If vehNumDifference = 0 Then vehNumDifference = 1
            End If
                        
            currentCell.Offset(, sourceVehLocation).Value = "'" & q.Dequeue      'First argument of .Offset will likely be replaced with 'maxOffset'
            currentCell.Offset(, -1).Value = "'" & q2.Dequeue
            
            If (vehNumDifference < maxOffset And q.Count > 0) Then             'On last iteration, last source veh's maxOffset takes precedence vehNum difference
                maxOffset = vehNumDifference
                selected_IF = "1st Ln 134"
                
            ElseIf (maxOffset > 2 And maxOffset Mod 2 <> 0 And q.Count > 0 And remainingRows > q.Count And vehNumDifference > maxOffset) Then       'Make even where applicable; eg 4 cars per sveh instead of 3
                maxOffset = maxOffset + 1                                                                                                           'E.g. 5,6,7,8,10 instead of 5,6,7,8, ,10
                selected_IF = "2nd Ln 138"
            
            'For diagnostics purpose I will not combine 2nd & 3rd if statements
           
            ElseIf (dwCounter1 Mod 2 = 0 And (remainingRows - 2 * q.Count) >= 3 And (vehNumDifference >= maxOffset + 1) And _
                                                                                currentCell.Offset(-1, sourceVehLocation).Value <> "") Then         '10,11,,,14 instead of 10,11,,14
                'prevVehNumDifference = Abs(CInt(vehNumColl(dwCounter1)) - CInt(vehNumColl(dwCounter1 + 1)))
                
                maxOffset = maxOffset + 1
                selected_IF = "3rd Ln 150"
                
            Else
                maxOffset = maxOffset
                selected_IF = "4th Ln 154"
            End If
            
            'Last row of comlumn 1 must be specfied, for the automation algorithm of 'Main Module' to function properly
            If (iMain = impPrefixColl.Count And q.Count = 0 And currentCell.Row <> lastRowOfSheet) Then
                tableSheet.Cells(lastRowOfSheet, sourceVehCell.Column) = currentCell.Offset(, sourceVehLocation).Value
            End If
                        
            Set currentCell = currentCell.Offset(maxOffset, 0)
            
            totalOffset = totalOffset + maxOffset
            
            Dim ccRow As Integer: ccRow = currentCell.Row
            
            'vehNumDifference = 0
            dwCounter1 = dwCounter1 + 1
        Loop
        totalOffset = 0
        dwCounter1 = 1
        
        q.Clear
        Set vehNumColl = Nothing              'Clear vehNumColl for next group of car models sharing prefix
    Next iMain

End Function

Function buildFileNameCollection(tableSheet As Worksheet, filePath As String) As Collection
    Set buildFileNameCollection = New Collection
    
    If Right(filePath, 1) <> "\" Then
        filePath = filePath + "\"
    End If
    
    Dim MyFile As Variant

    'Loop through all the files in the directory by using Dir$ function
    MyFile = Dir$(filePath & "*.veh")
    Do While MyFile <> ""
        buildFileNameCollection.Add (Left(MyFile, Len(MyFile) - 4))     'Discard file extension
        MyFile = Dir$
    Loop
End Function

Function buildPrefixCollection(tableSheet As Worksheet, importFilePrefCell As Range, lastRowOfSheet As Integer) As Collection
    Set buildPrefixCollection = New Collection
    
    Dim lastImpPrefix As String: lastImpPrefix = ""
    
    Dim currentCell As Range: Set currentCell = importFilePrefCell

    Dim dwC_bp As Integer: dwC_bp = 1   'debugging

    Do While currentCell.Row <= lastRowOfSheet
        If currentCell.Value <> "" And currentCell <> lastImpPrefix Then
            
            buildPrefixCollection.Add (currentCell.Value)
            lastImpPrefix = currentCell.Value
        Else
            Set currentCell = currentCell.Offset(1, 0)
        End If
        dwC_bp = dwC_bp + 1
    Loop
    dwC_bp = 0
End Function


