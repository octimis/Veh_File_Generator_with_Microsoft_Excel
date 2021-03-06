VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sVehImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myWs As Worksheet

Private fileName As String
Private sourceVehStr As String
Private vehNumStr As String
Private impportPrefixStr As String
Private indexNum As Integer
Private isExported_b As Boolean


'constructor
Private Sub Class_Initialize()
    
End Sub


Public Sub InitializeAttributes(tableSheet_c As Worksheet, fileName_c As String, possiblePrefixes As Collection, index_c As Integer)
    Set myWs = tableSheet_c
    
    fileName = fileName_c       'global variable
    
    'impPrefix(fileName) = possiblePrefixes
    
    impPrefix(possiblePrefixes) = ""
    
    'sourceVeh(fileName) = impportPrefixStr
    
    sourceVeh = impportPrefixStr        '<<<<<<<<<<<<<<<<< works
    
    possibleVehNum = sourceVehStr
    
    index = index_c
End Sub


Private Property Let sourceVeh(impportPrefixStr As String)         'incorrect order (fileName_m As String, impportPrefixStr As String)  'fileName should 2nd argument, was 1st occurrence in InitializeAttributes
       sourceVehStr = Replace(fileName, impportPrefixStr, "")
End Property
Public Property Get sourceVeh() As String
    Let sourceVeh = sourceVehStr
End Property

Private Property Let impPrefix(impFPcoll As Collection, fileName_m As String)            'impPrefix(fileName_m As String, impFPcoll As Collection)   'incorrect order

    Dim indexCounter As Integer: indexCounter = 1
    Dim isObtained As Boolean: isObtained = False
    
    Do While (indexCounter <= impFPcoll.Count) And Not (isObtained)
        If InStr(fileName, impFPcoll(indexCounter)) Then
        
            impportPrefixStr = impFPcoll(indexCounter)
            
            isObtained = True
        End If
        
        indexCounter = indexCounter + 1
    Loop
End Property
Public Property Get impPrefix(impFPcoll As Collection) As String
    Let impPrefix = impportPrefixStr
End Property

Private Property Let possibleVehNum(sourceVeh_m As String)

    'Dim sourceVeh_PV_local As String: sourceVeh_PV_local = sourceVeh_m      'Failure to do this, changes sourceVehStr globally
        
    vehNumStr = extractVehNum(sourceVeh_m, "")
'    Dim tableSheet As Worksheet: Set tableSheet = Sheets("Sheet2")
'    tableSheet.Range("L1").Value = "Ln 585 is called, VehNumCell.Row: " & VehNumCell.Row & ", value: " & VehNumCell.Value
End Property
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
Public Property Get possibleVehNum() As String
    Let possibleVehNum = vehNumStr
End Property

Private Property Let index(index_m As Integer)
    indexNum = index_m
End Property
Public Property Get index() As Integer
    Let index = indexNum
End Property

Private Property Let isExported(state_m As Boolean)
    isExported_b = state_m
End Property
Private Property Get isExported() As Boolean
    Let isExported = isExported_b
End Property
