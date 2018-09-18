VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vehRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private importSheet As Worksheet
Private tableSheet As Worksheet

Public indexNum As Integer

'These are within Propert Get and must be Public, not Private
Public expFilePrefixStr As String
Public impFilePrefixStr As String

Public pitGroupPrefix As String
Public descriptionPrefix As String

Public descriptionDelimiter As String

Private sourceVeh As String
Private sourceVeh_original As String     'Initialized when sourceVeh chages

Private derivedSourceVeh As String  'Unique Default Livery Set to No
Private pitGroupSuffix As String

Private defaultLiveryStr As String

Public vehNumStr As String

Private impFileNameStr As String
Private expFileNameStr As String
Private pitGroupStr As String
Private driverStr As String
Private teamStr As String
Private descriptionStr As String
Private classesStr As String
Private categoryStr As String

Private containsDelimiter As Boolean
Private srcRange As Variant

Private lastRowImportSheet_Num As Integer   'importSheet
Private lastRowOfTableSheet As Integer

Private sourceVehCell_vehNum As String      'vehNum of the first occurence of sourceVeh in column A; vital to automating 3 attributes

'+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
'| Source veh # | Veh # | Import file Prefix | Export file Prefix | PitGroup Prefix | PitGroup Suffix | Driver | Description Prefix | Classes | Category |
'+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
'|            1 |     2 |                  3 |                  4 |               5 |               6 |      7 |                  8 |       9 |       10 |
'+--------------+-------+--------------------+--------------------+-----------------+-----------------+--------+--------------------+---------+----------+
    
'+-----------------------+
'| Attributes of vehFile |
'+-----------------------+
'| DefaultLivery=        |
'| Number=               |
'| PitGroup=             |
'| Driver=               |
'| Description=          |
'| Classes=              |
'| Category=             |
'+-----------------------+

'constructor
Private Sub Class_Initialize()
    
End Sub
Public Sub InitializeAttributes(sourceVeh_c As String, vehNum_c As String, impFilePrefix_c As String, expFilePrefix_c As String, _
                             pitGroupPrefix_c As String, pitGroupSuffix_c As String, driver_c As String, team_c As String, _
                             descriptionPrefix_c As String, classes_c As String, category_c As String, descriptionDelimiter_c As String, _
                             importSheet_c As Worksheet, tableSheet_c As Worksheet, index_c As Integer, containsDelimiter_c As Boolean)
    Set importSheet = importSheet_c
    Set tableSheet = tableSheet_c
    
    'If (impFilePrefix_c = "488_GT3_") Then Stop
    
    sourceVeh = sourceVeh_c
    derivedSourceVeh = sourceVeh_c
    expFilePrefixStr = expFilePrefix_c
    impFilePrefixStr = impFilePrefix_c
    expFilePrefixStr = expFilePrefix_c
    pitGroupPrefix = pitGroupPrefix_c
    pitGroupSuffix = pitGroupSuffix_c
    descriptionPrefix = descriptionPrefix_c
    teamStr = team_c
    descriptionDelimiter = descriptionDelimiter_c
    containsDelimiter = containsDelimiter_c
    indexNum = index_c
    
    'Property Let/Get, may be called outside of class
    vehNum = vehNum_c
    
    team = team_c
    classes = classes_c
    category = category_c       'adding new columns instead of replacing
            
    pitGroup(pitGroupPrefix) = pitGroupSuffix
    
    importFileName(impFilePrefixStr) = sourceVeh
    index = indexNum
    
    lastRowImportSheet = importSheet.Range("A" & Rows.Count).End(xlUp).Row
    lastRowOfTableSheet = tableSheet.Range("A" & Rows.Count).End(xlUp).Row      '<<<<<<<<<<<<<<<<<<< May need to be accessible from outside the class AND re-called after/in driver = driver_c
    
    sourceVehCell_vehNum = extractOrginalSourceVehNum        'Original vehNum of derived SourceVeh File, must be placed after 'lastRowOfTableSheet ='
    exportFileName(expFilePrefixStr) = vehNum                'This MUST come after sourceVehCell_vehNum, since it calls a function dependent on sourceVehCell_vehNum
    defaultLivery(expFilePrefixStr) = vehNum
    description(descriptionPrefix, descriptionDelimiter) = vehNum   'May rely on orignal source vehNum
    driver = driver_c   'Must be placed after lastRowImportSheet =
    
End Sub

'This will prevent us from having to reload each sourceVeh text files a 2nd time, when updating desired individual attributes, and finally exporting them
'Will be set outside of class module, after attributes have beeb initialized
'Property Let is intended for primitive data types & String reference type
Public Property Let saveSheet(copyRange_m As Variant)
    srcRange = copyRange_m
    'tableSheet.Range("L1").Value = lastRow
End Property
Public Property Get saveSheet() As Variant
    'sourceVeh files may have varying lines
    'importSheet.UsedRange.ClearContents
    saveSheet = srcRange
End Property

Public Property Let lastRowImportSheet(lastRow_m As Integer)
    lastRowImportSheet_Num = lastRow_m
End Property
Public Property Get lastRowImportSheet() As Integer
    Let lastRowImportSheet = lastRowImportSheet_Num
End Property

'+-----------------------+  *: may have to derive attribute from imported text file
'| Attributes of vehFile |
'+-----------------------+
'| DefaultLivery=        |
'| Number=               |
'| PitGroup=             |
'| *Driver=              |
'| *Description=         |
'| *Classes=             |
'| *Category=            |
'+-----------------------+

Private Property Let defaultLivery(expFilePrefix_m As String, vehNum_m As String)
    Dim liveryOption As String
    liveryOption = tableSheet.Cells.Find(What:="Unique DefaultLivery", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value
    
    Dim exportOption As String
    exportOption = tableSheet.Cells.Find(What:="Export Format", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value
    
    Select Case liveryOption
        Case "Yes"
            Select Case exportOption
                Case "Type A"
                
                    defaultLiveryStr = """" + expFilePrefix_m + vehNum_m + ".DDS"""
                Case "Type B"   'SourceVeh and VehNum columns are of clearly different char sizes/composition.
                
                    defaultLiveryStr = """" + expFilePrefix_m + ".DDS"""
            End Select
        Case "No"       'User intends to use exact same livery, with identical car # displayed on them
            Select Case exportOption
                Case "Type A"
                    defaultLiveryStr = """" + expFilePrefix_m + sourceVehCell_vehNum + ".DDS"""
                Case "Type B"
                    If Left(expFilePrefix_m, 1) = "_" Then
                        defaultLiveryStr = """" + sourceVeh_original + Replace(expFilePrefix_m, sourceVeh, "", 1) + ".DDS"""
                    Else
                        defaultLiveryStr = """" + Replace(expFilePrefix_m, sourceVeh, "", 1) + sourceVeh_original + ".DDS"""
                    End If
            End Select
    End Select
End Property
Public Property Get defaultLivery(expFilePrefix_m As String) As String
    Let defaultLivery = defaultLiveryStr
End Property

Private Property Let vehNum(vehNum_m As String)
    vehNumStr = vehNum_m
End Property
Public Property Get vehNum() As String
    Let vehNum = vehNumStr
End Property

Private Property Let pitGroup(pitGroupPrefix As String, pitGroupSuffix As String)
    pitGroupStr = pitGroupPrefix + pitGroupSuffix
        
    If (pitGroupStr = "") Then
        pitGroupStr = importSheet.Cells.Find(What:="PitGroup=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)
        pitGroupStr = Mid(pitGroupStr, 2, Len(pitGroupStr) - 2)
    End If
    
    pitGroupStr = """" + pitGroupStr + """"
End Property
Public Property Get pitGroup(pitGroupPrefix As String) As String
    Let pitGroup = pitGroupStr
End Property

Private Property Let driver(driver_m As String)

    Dim driverCell As Range
    Set driverCell = importSheet.Cells.Find(What:="Driver=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)

    If InStr(driver_m, "_") > 0 Then    'User specified multiple drivers
        Dim underScorePos As Integer
        
        underScorePos = InStr(driver_m, "_")
        
        Dim coDriverStr As String: coDriverStr = Right(driver_m, Len(driver_m) - underScorePos)
        
        driver_m = Left(driver_m, underScorePos - 1)
        
        Dim driverArray() As String: driverArray = Split(coDriverStr, "_")
        
        importSheet.Range("D5").Value = driverArray(0)
        
        For i = LBound(driverArray) To UBound(driverArray)
            Call insertNameToImportSheet(driverArray(i), driverCell.Row + 1, CInt(i))
        Next i
        'Save sheet is called in main module, after this instance is done being made
    ElseIf (driverName = "") Then
    
        driver_m = Mid(driverCell.Value, 2, Len(driverCell.Value) - 2)      'Omit quotations, hence 2
'    Else
'    'If (driver_m = "") Then
'        driver_m = importSheet.Cells.Find(What:="Driver=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)
'        driver_m =
    End If
    
    driverStr = """" + driver_m + """"
End Property
Public Property Get driver() As String
    Let driver = driverStr
End Property

Private Function insertNameToImportSheet(driverName As String, rowToInsertTo As Integer, driverNum As Integer)
    
    Dim searchString As String: searchString = "Driver" & (1 + driverNum) & "="
    
    Dim coDriver_SearchRng As Range: Set coDriver_SearchRng = importSheet.Range("A1:A" & lastRowImportSheet)
    Dim coDriverCell As Range: Set coDriverCell = coDriver_SearchRng.Find(What:=searchString, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False)
    
    'Dim impFP_searchRng As Range: Set impFP_searchRng = tableSheet.Range("C1:C" & lastRowOfTableSheet)
    'Dim impFilePrefix_Cell As Range: Set impFilePrefix_Cell = impFP_searchRng.Find(What:=impFilePrefixStr, lookat:=xlWhole, After:=impFP_searchRng(impFP_searchRng.Count))
      
    Dim exportCell As Range
    
    If Not coDriverCell Is Nothing Then           'searchString exists in importSheet
    
        importSheet.Cells.Find(What:=searchString, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False).Offset(, 1).Value = """" + driverName + """"
    Else
        importSheet.Cells(rowToInsertTo + driverNum, 1).EntireRow.Insert Shift:=xlDown
        
        Set exportCell = importSheet.Cells(rowToInsertTo + driverNum, 1)
        
        exportCell.Value = "Driver" & driverNum + 1 & "="
        exportCell.Offset(0, 1).Value = """" + driverName + """"
    End If
        
    lastRowImportSheet = importSheet.Range("A" & Rows.Count).End(xlUp).Row
    
End Function


Private Property Let team(team_m As String)
    If (team_m = "") Then
        team_m = importSheet.Cells.Find(What:="Team=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)
        team_m = Mid(team_m, 2, Len(team_m) - 2)
    End If
    
    If tableSheet.Cells.Find(What:="Overwrite FT Name", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Value = "Yes" Then
        importSheet.Cells.Find(What:="FullTeamName=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1).Value = team_m
    End If
    
    teamStr = """" + team_m + """"
End Property
Public Property Get team() As String
    Let team = teamStr
End Property

Private Property Let description(descriptionPrefix_c As String, descriptionDelimiter_c As String, vehNum_c As String)
    
    Dim selected_IF As String           'Debuging
    
    'If user specified description with with delimiter, none of these If/Else statement should be executed.
    Dim rightOf As String: rightOf = ""
    
    If (descriptionPrefix_c = "") Then
        descriptionPrefix_c = importSheet.Cells.Find(What:="Description=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(0, 1)
        descriptionPrefix_c = Replace(descriptionPrefix_c, """", "")
        If descriptionDelimiter_c = "" Or InStr(descriptionPrefix_c, descriptionDelimiter_c) = 0 Then
            'Guess on user's behalf where delimeter and vehNum are located at
            rightOf = trimAndCompare(descriptionPrefix_c, rightOf, sourceVehCell_vehNum)
            descriptionPrefix_c = Left(descriptionPrefix_c, Len(descriptionPrefix_c) - Len(rightOf)) & Replace(rightOf, sourceVehCell_vehNum, vehNum)
            
            selected_IF = "If-> If, Ln 297"
        Else    'Delimiter specified in B4 & D5 set to yes
            'In case there's space chars between # and vehNum, split string into 2 separate ones, then replace & combine
            Dim leftOf As String    ', rightOf As String
            Dim posOfDelim As Integer, posOfVehNum As Integer
            
            posOfDelim = InStr(descriptionPrefix_c, descriptionDelimiter_c)
            'posOfVehNum = InStr(posOfDelim, descriptionPrefix_c, vehNum_c)        'VehNum always comes after delimeter, hence 1st argument
            
            leftOf = Left(descriptionPrefix_c, posOfDelim)
            'rightOf = Right(descriptionPrefix_c, Len(descriptionPrefix_c) - posOfDelim)
            'rightOf = Replace(rightOf, sourceVehCell_vehNum, vehNum)
            rightOf = Replace(Right(descriptionPrefix_c, Len(descriptionPrefix_c) - posOfDelim), sourceVehCell_vehNum, vehNum)
            
            descriptionPrefix_c = leftOf + rightOf
            
            selected_IF = "If-> Else, Ln 313"
        End If
    Else 'Add another if-else statement checking delimiter already existed
        If Not (containsDelimiter) Then   'Add space char at end
            descriptionPrefix_c = descriptionPrefix_c + " "
            
            selected_IF = "Else->If, Ln 319"
            
            If tableSheet.Cells.Find(What:="Force assign delimiter", _
                LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value = "Yes" Then
                    descriptionPrefix_c = descriptionPrefix_c + descriptionDelimiter_c + vehNum_c
                    
                    selected_IF = "Else-> If-> If, Ln 325"
            End If
        
        Else
            descriptionPrefix_c = descriptionPrefix_c + " " + descriptionDelimiter_c + vehNum_c
        End If
    End If
    
    descriptionStr = """" + descriptionPrefix_c + """" 'may need to change vehNum_c back to vehNum
    
End Property
Public Property Get description(descriptionPrefix_c As String, descriptionDelimiter_c As String) As String
    Let description = descriptionStr
End Property
'Trim description, and hand it over to rightOf
Private Function trimAndCompare(analyzedString As String, rightOf As String, numStr As String) As String
    If Len(analyzedString) = 0 Or InStr(rightOf, numStr) > 0 Then    'Exit if no char is left in analyzedString or vehNum is found
        trimAndCompare = rightOf
    Else
        trimAndCompare = trimAndCompare(Left(analyzedString, Len(analyzedString) - 1), Right(analyzedString, 1) + rightOf, numStr)
    End If
End Function

Private Property Let classes(classes_m As String)
    If (classes_m = "") Then
        classes_m = importSheet.Cells.Find(What:="Classes=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)
        classes_m = Mid(classes_m, 2, Len(classes_m) - 2)
    End If
    classesStr = """" + classes_m + """"
End Property
Public Property Get classes() As String
    Let classes = classesStr
End Property

Private Property Let category(category_m As String)
    If (category_m = "") Then
        category_m = importSheet.Cells.Find(What:="Category=", LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Offset(, 1)
        category_m = Mid(category_m, 2, Len(category_m) - 2)
    End If
    
    categoryStr = """" + category_m + """"
End Property
Public Property Get category() As String
    Let category = categoryStr
End Property

'+-----------------------+
'| Attributes of vehFile |
'+-----------------------+
'| DefaultLivery=        |
'| Number=               |
'| PitGroup=             |
'| Driver=               |
'| Description=          |
'| Classes=              |
'| Category=             |
'+-----------------------+

Private Property Let importFileName(impFilePrefix_m As String, sourceVeh As String)
    impFileNameStr = impFilePrefix_m + sourceVeh
End Property
Public Property Get importFileName(impFilePrefix_m As String) As String
    Let importFileName = impFileNameStr
End Property

Private Property Let exportFileName(expFilePrefix_m As String, vehNum_m As String)
    Select Case tableSheet.Cells.Find(What:="Export Format", _
                LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Offset(, 1).Value
    Case "Type A"
        If Left(expFilePrefix_m, 1) = "_" Then
            expFilePrefix_m = vehNum_m + expFilePrefix_m ' + ".veh"
        Else
            expFilePrefix_m = expFilePrefix_m + vehNum_m ' + ".veh"
        End If
    Case "Type B"
        If sourceVehCell_vehNum <> vehNum Then
            Call updateSourceVeh
        End If
        
        If Left(expFilePrefix_m, 1) = "_" Then
            expFilePrefix_m = sourceVeh + expFilePrefix_m ' + ".veh"
        Else
            expFilePrefix_m = expFilePrefix_m + sourceVeh ' + ".veh"
        End If
    End Select
    
    expFileNameStr = expFilePrefix_m + ".veh"
End Property
Public Property Get exportFileName(expFilePrefix_m As String) As String
    Let exportFileName = expFileNameStr
End Property

Private Property Let index(num As Integer)
    indexNum = num
End Property
Public Property Get index() As Integer
    Let index = indexNum
End Property

Function extractOrginalSourceVehNum() As String
    Dim impFP_searchRng As Range: Set impFP_searchRng = tableSheet.Range("C1:C" & lastRowOfTableSheet)
    'It's possible for 2 different car models to have same SourceVeh ID
    'Search for 1st impFilePrefix row, then search for 1st occurence of sourceVeh and extract original vehNum
    Dim impFilePrefix_Cell As Range: Set impFilePrefix_Cell = impFP_searchRng.Find(What:=impFilePrefixStr, lookat:=xlWhole, After:=impFP_searchRng(impFP_searchRng.Count))
    
    'Allows us to start searching at the first instance where impfilePrefix was obtained.
    Dim sVeh_searchRng As Range: Set sVeh_searchRng = tableSheet.Range(tableSheet.Cells(impFilePrefix_Cell.Row, 1), tableSheet.Cells(lastRowOfTableSheet, 1))   '1 is columnA
    
    Dim sourceVehCell As Range: Set sourceVehCell = sVeh_searchRng.Find(What:=sourceVeh, lookat:=xlWhole)
        
    sourceVeh_original = sourceVehCell.Value
    
    extractOrginalSourceVehNum = sourceVehCell.Offset(, 1).Value
End Function


'Type B' selected in TableSheet, will trigger this function
'This is updated, because a veh file may have a suffix totally different from "VehNum"
Private Function updateSourceVeh() As String
    'If vehNum matches orgiVehNum, don't run

    If Len(sourceVehCell_vehNum) <> Len(vehNum) Then
        'VehNum was derived from sourVeh, right... right?
        sourceVehCell_vehNum = enlargeOrigVehNum(sourceVeh, sourceVehCell_vehNum, vehNum)
    End If
    
    'Replace origVehNum in sourceVeh with vehNum, by extracting values left of, and the appending it.
    Dim tempSourceVeh As String: tempSourceVeh = truncateSourceVeh(sourceVeh, sourceVehCell_vehNum) + vehNum
    
    sourceVeh = tempSourceVeh + Right(sourceVeh, Len(sourceVeh) - Len(tempSourceVeh))
End Function

Function enlargeOrigVehNum(sourceVeh As String, os As String, vn As String) As String
    If Right(sourceVeh, Len(os)) = os Then
        enlargeOrigVehNum = Right(sourceVeh, Len(vn))
    Else
        enlargeOrigVehNum = enlargeOrigVehNum(Left(sourceVeh, Len(sourceVeh) - 1), os, vn)
    End If
End Function

Function truncateSourceVeh(sourceVeh As String, os As String) As String
    If Right(sourceVeh, Len(os)) = os Then
        truncateSourceVeh = Left(sourceVeh, Len(sourceVeh) - Len(os))
    Else
        truncateSourceVeh = truncateSourceVeh(Left(sourceVeh, Len(sourceVeh) - 1), os)
    End If
End Function

