Attribute VB_Name = "ReadIniFiles"
'#################################################################################
'*** Public Variablen
'#################################################################################

'Public Const INIPath As String = App.Path & "\config.ini"

Public Masch As String, INIPath As String
' Servername und Benutzerdaten
Public sServer, sUsername, sPassword, sDBName, sTable, sAutor As String
Public cntTyp, cntMasch, cntRes, cntProb, cntFields As Integer
Public ArTyp(), ArMasch(), ArRes(), ArProb(), ArFields() As String
Public NewInsert, Author As String
Public PicBMP As String

'#################################################################################
'### For read the ini-files
'#################################################################################
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias _
                  "GetPrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpDefault As String, _
                  ByVal lpReturnedString As String, _
                  ByVal nSize As Long, _
                  ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias _
                  "WritePrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpString As Any, _
                  ByVal lpFileName As String) As Long
                  
Public Property Get sPath() As String
    sPath = INIPath
End Property
 
Public Property Let sPath(ByVal NewValue As String)
    INIPath = NewValue
End Property
 
Public Function WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)
    WritePrivateProfileString Section, Key, sValue, INIPath
End Function
 
Public Function WriteValue(ByVal Section As String, ByVal Key As String, ByVal vValue As Variant)
    WriteString Section, Key, CStr(vValue)
End Function
 
Public Function GetIniString(ByVal Section As String, ByVal Key As String, _
        Optional ByVal Default As String = "") As String
    
    Dim sTemp As String
 
    sTemp = String(256, 0)
    GetPrivateProfileString Section, Key, "", sTemp, Len(sTemp), INIPath
    If InStr(sTemp, Chr$(0)) Then
        sTemp = Left$(sTemp, InStr(sTemp, vbNullChar$) - 1)
    Else
        sTemp = Default
    End If
    
    GetIniString = sTemp
End Function
 
Public Function GetIniLong(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Long = -1) As Long
Dim sTemp As String
 
    sTemp = GetIniString(Section, Key, CStr(Default))
    If IsNumeric(sTemp) Then
        GetIniLong = CInt(sTemp)
    'Else
        'Evtl. Fehlermeldung ausgeben
    End If
End Function
 
Public Function GetIniBool(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Boolean = False) As Boolean
    GetIniBool = CBool(GetIniLong(Section, Key, CInt(Default)))
End Function
 
Public Sub IniRead()

INIPath = App.Path & "\config.ini"

frmInput.CmbMaschine.Clear
frmInput.CmbTestTyp.Clear
frmInput.CmbProblem.Clear
frmInput.CmbErgebnis.Clear

sServer = GetIniString("dbconnect", "servername", INIPath)
sUsername = GetIniString("dbconnect", "username", INIPath)
sPassword = GetIniString("dbconnect", "password", INIPath)
sDBName = GetIniString("dbconnect", "dbname", INIPath)
sTable = GetIniString("dbconnect", "tablename", INIPath)
sAutor = GetIniString("dbconnect", "autor", INIPath)
PicBMP = GetIniString("dbconnect", "Pic", INIPath)
PicBMP = App.Path & "\" & PicBMP

i = 0
Masch = "Masch" & i
Values0 = GetIniString("Dynos", Masch, INIPath)

ReDim Preserve ArMasch(i)
ArMasch(i) = Values0

frmInput.CmbMaschine.AddItem ""

Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve ArMasch(i)
        ArMasch(i) = Values0
        frmInput.CmbMaschine.AddItem Values0
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "Masch" & i
   Values0 = GetIniString("Dynos", Masch, INIPath)
Loop
 
cntMasch = i
 
i = 0
Values0 = ""
Masch = "Typ" & i
Values0 = GetIniString("TestTyp", Masch, INIPath)

ReDim Preserve ArTyp(i)
ArTyp(i) = Values0

frmInput.CmbTestTyp.AddItem ""

Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve ArTyp(i)
        ArTyp(i) = Values0
        frmInput.CmbTestTyp.AddItem Values0
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "Typ" & i
   Values0 = GetIniString("TestTyp", Masch, INIPath)
Loop
 
cntTyp = i

i = 0
Values0 = ""
Masch = "Prob" & i
Values0 = GetIniString("Problem", Masch, INIPath)

ReDim Preserve ArProb(i)
ArProb(i) = Values0

frmInput.CmbProblem.AddItem ""

Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve ArProb(i)
        ArProb(i) = Values0
        frmInput.CmbProblem.AddItem Values0
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "Prob" & i
   Values0 = GetIniString("Problem", Masch, INIPath)
Loop
 
cntProb = i

i = 0
Values0 = ""
Masch = "Res" & i
Values0 = GetIniString("Result", Masch, INIPath)

ReDim Preserve ArRes(i)
ArRes(i) = Values0

frmInput.CmbErgebnis.AddItem ""

Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve ArRes(i)
        ArRes(i) = Values0
        frmInput.CmbErgebnis.AddItem Values0
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "Res" & i
   Values0 = GetIniString("Result", Masch, INIPath)
Loop

cntRes = i
'cntFields,ArFields()
 
i = 0
Values0 = ""
Masch = "Field" & i
Values0 = GetIniString("Fields", Masch, INIPath)

ReDim Preserve ArFields(i)
ArFields(i) = Values0

Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve ArFields(i)
        ArFields(i) = Values0
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "Field" & i
   Values0 = GetIniString("Fields", Masch, INIPath)
Loop
 
cntFields = i - 1

End Sub

