VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dynamometers Problem report !"
   ClientHeight    =   7275
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10800
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9240
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   9120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdDropTable 
      Caption         =   "Delete Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   8
      Top             =   2655
      Width           =   1590
   End
   Begin VB.CommandButton cmdQueryTable 
      Caption         =   "Query"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   7
      Top             =   3075
      Width           =   1590
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   5
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdCreateTable 
      Caption         =   "Create Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   3
      Top             =   2235
      Width           =   1590
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   2
      Top             =   1500
      Width           =   1590
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   1
      Top             =   1080
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5835
      Left            =   210
      TabIndex        =   0
      Top             =   840
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   10292
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label lblURL 
      Caption         =   "www.little-tools-farm.de"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   7920
      MouseIcon       =   "Form1.frx":0A49
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   11
      Top             =   7020
      Width           =   2385
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright ©2007 by little-tools-farm.de"
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   10
      Top             =   6765
      UseMnemonic     =   0   'False
      Width           =   2760
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Welcome Dynamometers Problem Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   105
      Picture         =   "Form1.frx":0D53
      Stretch         =   -1  'True
      Top             =   105
      Width           =   555
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   210
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   210
      X2              =   5775
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================
Option Explicit

' Die Deklarationen sind fast wie bei ADO
' Connection-Object
Private oConn As New MYSQL_CONNECTION

' Recordset-Object
Private oRs As MYSQL_RS

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Fehlerausgabe bei Verbindungsfehler
Private Function MySQL_Error() As Boolean
  With oConn.Error
    If .Number = 0 Then Exit Function
  
    MsgBox "Error " & .Number & ": " & .Description
    MySQL_Error = True
  End With
End Function

Private Sub cmdAdd_Click()
  ' Datensatz hinzufügen
  Dim nResult As Long
  Dim sVorname, sNachname As String 'Vorsicht hat sich geaendert !
  Dim sTestTyp, sProblem, sMaschine, sResult, sDescription, sTestnumber As String
  Dim sMeetingDate, sComments, sAction As String
   
  Load frmInput
  With frmInput
    .Show vbModal
    If .Tag = True Then
      .txtAutor.Text = sAutor
      sVorname = .txtAutor.Text '.txtVorname.Text
      sNachname = .txtDate.Text  '.txtNachname.Text
      sDescription = .txtDescription.Text
      sTestnumber = .txtTestNo.Text
      sResult = .CmbErgebnis.Text
      sMaschine = .CmbMaschine.Text
      sProblem = .CmbProblem.Text
      sTestTyp = .CmbTestTyp.Text
      sMeetingDate = .txtMeetingdate.Text
      sComments = .txtComments.Text
      sAction = .txtAction.Text
      
      ' Datensatz in Tabelle einfügen
    oConn.Execute "INSERT INTO " & sTable & " (maschine, date, testtyp, testno, autor, problem, result, description, MeetingDate, Comments, Action) " & _
        "VALUES ('" & sMaschine & "', '" & sNachname & "', '" & sTestTyp & "', '" & sTestnumber & "', '" & sVorname & "', '" & sProblem & "', '" & sResult & "', '" & sDescription & "', '" & sMeetingDate & "', '" & sComments & "', '" & sAction & "')"

      If MySQL_Error() = False Then
        ' Datensatz dem Grid hinzufügen
        With MSFlexGrid1
           .AddItem CStr(oConn.LastInsertID) & vbTab & _
            sMaschine & vbTab & sNachname & vbTab & sTestTyp & _
            vbTab & sTestnumber & vbTab & sVorname & vbTab & _
            sProblem & vbTab & sResult & vbTab & sDescription & vbTab & _
            sMeetingDate & vbTab & sComments & vbTab & sAction
            'sVorname & vbTab & sNachname
          .Row = .Rows - 1
          .RowSel = .Row
          .Col = 0
          .ColSel = .Cols - 1
          .Enabled = (.Rows > 1)
        End With
      End If
    End If
  End With
  Unload frmInput
End Sub

Private Sub CmdClose_Click()
  End
End Sub

Private Sub cmdCreateTable_Click()
  Dim sSQL As String
  'CREATE DATABASE `dbproblems`
  ' Tabelle erstellen
  
  sSQL = "CREATE TABLE " & sTable & " (" _
    & "id INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY, " _
    & "maschine CHAR(20) NOT NULL, " _
    & "date CHAR(20) NOT NULL, " _
    & "testtyp CHAR(30) NOT NULL, " _
    & "testno CHAR(20) NOT NULL, " _
    & "autor CHAR(20) NOT NULL, " _
    & "problem CHAR(30) NOT NULL, " _
    & "result CHAR(30) NOT NULL, " _
    & "description LONGTEXT NOT NULL, " _
    & "MeetingDate CHAR(20) NOT NULL, " _
    & "Comments  LONGTEXT NOT NULL, " _
    & "Action LONGTEXT NOT NULL, " _
    & "tstamp TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)"
        
  oConn.Execute sSQL
  
  If MySQL_Error() = False Then
    MsgBox "Tabelle wurde erstellt."
    cmdCreateTable.Enabled = False
    cmdDropTable.Enabled = True
    cmdQueryTable.Enabled = True
  End If
End Sub

' Datensatz löschen
Private Sub cmdDelete_Click()
  Dim sSQL As String
  Dim nId As Long
  
  If MsgBox("Aktuellen Datensatz löschen?", vbYesNo, "Löschen") = vbYes Then
    With MSFlexGrid1
      ' SQL-Abfrage
      nId = .TextMatrix(.Row, 0)
      sSQL = "DELETE FROM " & sTable & " WHERE id = " & CStr(nId)
      
      ' Ausführen
      oConn.Execute sSQL
      If MySQL_Error() = False Then
        ' FlexGrid aktualisieren
        If .Rows = 2 Then
          .Rows = 1
          .Enabled = False
        Else
          .RemoveItem .Row
        End If
      End If
    End With
  End If
End Sub

Private Sub cmdDisconnect_Click()
  ' Verbindung beenden
  If Not oConn Is Nothing Then oConn.CloseConnection
  If MySQL_Error() = False Then
    Set oConn = Nothing
  
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    cmdCreateTable.Enabled = False
    cmdDropTable.Enabled = False
    cmdQueryTable.Enabled = False
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Enabled = False
  End If
End Sub

' Tabelle löschen
Private Sub cmdDropTable_Click()
  Dim sSQL As String
  
  If MsgBox("Tabelle wirklich löschen?", vbYesNo, "Löschen") = vbYes Then
    ' Tabelle löschen
    sSQL = "DROP TABLE IF EXISTS " & sTable
    oConn.Execute sSQL
    If MySQL_Error() = False Then
      cmdCreateTable.Enabled = True
      cmdDropTable.Enabled = False
      cmdQueryTable.Enabled = False
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
    End If
  End If
End Sub

' Datensatz ändern
Private Sub cmdEdit_Click()
  Dim sVorname, sNachname As String 'Vorsicht hat sich geaendert !
  Dim sTestTyp, sProblem, sMaschine, sResult, sDescription, sTestnumber As String
  Dim sMeetingDate, sComments, sAction As String
  Dim nId As Long
  Dim sSQL As String
  
  With MSFlexGrid1
    nId = Val(.TextMatrix(.Row, 0))
    sMaschine = .TextMatrix(.Row, 1)
    sNachname = .TextMatrix(.Row, 2)
    sTestTyp = .TextMatrix(.Row, 3)
    sTestnumber = .TextMatrix(.Row, 4)
    sVorname = .TextMatrix(.Row, 5)
    sProblem = .TextMatrix(.Row, 6)
    sResult = .TextMatrix(.Row, 7)
    sDescription = .TextMatrix(.Row, 8)
    sMeetingDate = .TextMatrix(.Row, 9)
    sComments = .TextMatrix(.Row, 10)
    sAction = .TextMatrix(.Row, 11)
  End With
  
  Load frmInput
  With frmInput
  
         .CmbMaschine = sMaschine
         .txtDate = sNachname   '.txtNachname.Text
         .CmbTestTyp = sTestTyp
         .txtTestNo = sTestnumber
         .txtAutor = sVorname
         .CmbProblem = sProblem
         .CmbErgebnis = sResult
         .txtDescription = sDescription
         .txtMeetingdate = sMeetingDate
         .txtComments = sComments
         .txtAction = sAction
         .Show vbModal
    
    If .Tag Then
      ' Wurden die Daten verändert?
      If .CmbMaschine.Text <> sMaschine Or .txtDate.Text <> sNachname Or _
         .CmbTestTyp.Text <> sTestTyp Or .txtTestNo.Text <> sTestnumber Or _
         .txtAutor.Text <> sVorname Or .CmbProblem.Text <> sProblem Or _
         .CmbErgebnis.Text <> sResult Or .txtDescription.Text <> sDescription Then
         
        ' Ja, also Datensatz aktualisieren
         sMaschine = .CmbMaschine.Text
         sNachname = .txtDate.Text  '.txtNachname.Text
         sTestTyp = .CmbTestTyp.Text
         sTestnumber = .txtTestNo.Text
         sVorname = .txtAutor.Text '.txtVorname.Text
         sProblem = .CmbProblem.Text
         sResult = .CmbErgebnis.Text
         sDescription = .txtDescription.Text
         sMeetingDate = .txtMeetingdate.Text
         sComments = .txtComments.Text
         sAction = .txtAction.Text
        
        ' SQL-Anweisung
        sSQL = "UPDATE " & sTable & " SET " & _
          "maschine = '" & sMaschine & "', " & _
          "date = '" & sNachname & "', " & _
          "testtyp = '" & sTestTyp & "', " & _
          "testno = '" & sTestnumber & "', " & _
          "autor = '" & sVorname & "', " & _
          "problem = '" & sProblem & "', " & _
          "result = '" & sResult & "', " & _
          "description = '" & sDescription & "' " & _
          "MeetingDate = '" & sMeetingDate & "' " & _
          "Comments = '" & sComments & "' " & _
          "Action = '" & sAction & "' " & _
          "WHERE id = " & CStr(nId)

'UPDATE `tbldynoprob` SET `testno` = 'D-2137',
'`description` = 'dritter Testeintrag' WHERE `id` =3 LIMIT 1
        
        ' Ausführen
        oConn.Execute sSQL
        If MySQL_Error() = False Then
          ' Grid aktualisieren
          With MSFlexGrid1
         .TextMatrix(.Row, 1) = sMaschine
         .TextMatrix(.Row, 2) = sNachname
         .TextMatrix(.Row, 3) = sTestTyp
         .TextMatrix(.Row, 4) = sTestnumber
         .TextMatrix(.Row, 5) = sVorname
         .TextMatrix(.Row, 6) = sProblem
         .TextMatrix(.Row, 7) = sResult
         .TextMatrix(.Row, 8) = sDescription
         .TextMatrix(.Row, 9) = sMeetingDate
         .TextMatrix(.Row, 10) = sComments
         .TextMatrix(.Row, 11) = sAction

          End With
        End If
      End If
    End If
  End With
  Unload frmInput
End Sub
Private Sub cmdQueryTable_Click()

  ' RS-Objekt vom Connection-Objekt ableiten und
  ' Status anzeigen.
  '
  ' Nicht wie bei ADO. Bei der MyVbQl.Dll wird das
  ' Recordset direkt von der Connection abgeleitet.
  
  Dim sSQL As String
  Dim bError As Boolean
  
  ' Alle Datensätze selektieren
  sSQL = "SELECT * FROM " & sTable
  Set oRs = oConn.Execute(sSQL)
  bError = MySQL_Error()
  If Not bError Then
    ' Daten im Grid anzeigen
    FillGrid
  Else
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
  End If
  
  ' Recordset schließen
  If Not oRs Is Nothing Then oRs.CloseRecordset
  Set oRs = Nothing
End Sub
' Funktion zum Füllen des FlexGrids
Function FillGrid()
  Dim i As Long
  Dim nRow As Long

  With MSFlexGrid1
    ' Zählvariable für die aktuelle Zeile
    nRow = 0
    .Rows = oRs.RecordCount + 1
    While Not oRs.EOF
      nRow = nRow + 1
      For i = 0 To oRs.FieldCount - 2
        ' Aktuellen Wert ins FlexGrid kopieren
        .TextMatrix(nRow, i) = oRs.Fields(i).Value
        .Col = i 'Spalte zuweisen
      Next i

      ' auch hier gibt es ein "MoveNext" :-)
      oRs.MoveNext
    Wend
    
    If .Row > 0 Then .Row = 1
    .Enabled = (.Rows > 1)
  End With
  
  cmdAdd.Enabled = True
  cmdEdit.Enabled = (MSFlexGrid1.Row > 0)
  cmdDelete.Enabled = (MSFlexGrid1.Row > 0)
End Function
' Ein par Einstellungen fürs FlexGrid

Private Sub Form_Load()

IniRead
  
  '#########
  If PicBMP <> "" Then
     Set Picture1.Picture = LoadPicture(PicBMP)
  End If
  
  
  With Me.MSFlexGrid1
    ' 3 Spalten und zunächst nur 1 Zeile für
    ' die Spaltenüberschriften
    
    .Cols = 12
    .Rows = 1
    
    ' Kein Fokus-Rechteck, sondern FullRowSelect-Modus
    .FocusRect = flexFocusNone
    .Col = 0: .ColSel = .Cols - 1

'sMaschine sNachname sTestTyp sTestnumber sVorname sProblem sResult sDescription
'id maschine date testtyp testno autor problem result description tstamp
    
    ' Spaltenüberschriften
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Maschine"
    .TextMatrix(0, 2) = "Date"
    .TextMatrix(0, 3) = "Test Typ"
    .TextMatrix(0, 4) = "Test No."
    .TextMatrix(0, 5) = "Autor"
    .TextMatrix(0, 6) = "Problem"
    .TextMatrix(0, 7) = "Result"
    .TextMatrix(0, 8) = "Description"
    .TextMatrix(0, 9) = "Meeting Datum"
    .TextMatrix(0, 10) = "Kommentare"
    .TextMatrix(0, 11) = "Aktion"
    
    ' Spaltenbreiten
    .ColWidth(0) = 500
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2000
    .ColWidth(7) = 2000
    .ColWidth(8) = 8000
    .ColWidth(9) = 2000
    .ColWidth(10) = 8000
    .ColWidth(11) = 2000
    
    ' zunächst inaktiv
    .Enabled = False
  End With
  
If sAutor = "" Then
   frmLogin.Show
   Me.Hide
End If
 
End Sub
Private Sub cmdConnect_Click()
  ' Wir öffnen die Verbindung zum MySQL Server
  ' Statt 'Localhost' kann auch die IP verwendet werden. Diese
  ' erfahren Sie im WinMySQLAdmin im Register'Environment',
  ' wenn Sie auf 'Extendet Server Status' klicken.

  oConn.OpenConnection sServer, _
    sUsername, sPassword, sDBName

  ' Statusabfrage
  If (oConn.State = MY_CONN_CLOSED) Then
    ' Falls Verbindung nicht geöffnet, Fehlerangabe!
    MySQL_Error
  Else
    ' Bei erfolgreicher Verbindung, Verbindungsdaten ausgeben
    MsgBox "Connected to Database: " & oConn.DbName, _
      vbInformation, "MySQL-Dyno-Failed-projekt"
      
    ' Prüfen, ob Tabelle existiert
    Set oRs = oConn.Execute("SELECT * FROM " & sTable)
    If oConn.Error.Number = 0 Then
      ' Daten anzeigen
      FillGrid
      
      oRs.CloseRecordset
      cmdCreateTable.Enabled = False
      If sUsername <> "root" Then
         Me.cmdCreateTable.Enabled = False
         Me.cmdDropTable.Enabled = False
      End If
      cmdQueryTable.Enabled = True
    Else
      If sUsername <> "root" Then
         cmdCreateTable.Enabled = True
      End If
    End If
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
          
    ' Recordset sschließen
    If Not oRs Is Nothing Then oRs.CloseRecordset
    Set oRs = Nothing
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Disconnecten
  cmdDisconnect.Value = True
End Sub

Private Sub Image3_Click()
  lblURL_Click
End Sub

' Standard-Browser starten und WWW-Seite aufrufen
Private Sub URLGoTo(ByVal hWnd As Long, ByVal URL As String)
  Screen.MousePointer = 11
  If Left$(URL, 7) <> "http://" Then URL = "http://" & URL
  Call ShellExecute(hWnd, "Open", URL, "", "", 3)
  Screen.MousePointer = 0
End Sub

Private Sub lblURL_Click()
  URLGoTo Me.hWnd, lblURL.Caption
End Sub


Private Sub MSFlexGrid1_RowColChange()
  ' Prüfen
  If MSFlexGrid1.Rows > 0 Then MSFlexGrid1.Col = 0
  cmdEdit.Enabled = (MSFlexGrid1.Row > 0)
  cmdDelete.Enabled = (MSFlexGrid1.Row > 0)
End Sub


Private Sub MSFlexGrid1_SelChange()
  ' Prüfen
  cmdEdit.Enabled = (MSFlexGrid1.Row > 0)
  cmdDelete.Enabled = (MSFlexGrid1.Row > 0)
End Sub

