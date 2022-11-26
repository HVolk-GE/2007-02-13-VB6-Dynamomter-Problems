VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Add/Change"
   ClientHeight    =   7185
   ClientLeft      =   2460
   ClientTop       =   6300
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Caption         =   "Action"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtAction 
      Height          =   1575
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtComments 
      Height          =   2775
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1710
      TabIndex        =   11
      Top             =   720
      Width           =   2265
   End
   Begin VB.TextBox txtTestNo 
      Height          =   285
      Left            =   1710
      TabIndex        =   10
      Top             =   1740
      Width           =   2265
   End
   Begin VB.ComboBox CmbMaschine 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox CmbTestTyp 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtAutor 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox CmbProblem 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox CmbErgebnis 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox txtMeetingdate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   330
      Left            =   3315
      TabIndex        =   1
      Top             =   6480
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   330
      Left            =   1950
      TabIndex        =   0
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   "Kommentar :"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Action :"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Problem Beschreibung:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   225
      Index           =   0
      Left            =   450
      TabIndex        =   19
      Top             =   735
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Test No.:"
      Height          =   225
      Index           =   1
      Left            =   345
      TabIndex        =   18
      Top             =   1770
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "Maschine :"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Test Typ:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Autor:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Problem:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Ergebnis:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Meeting Datum:"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
  If Me.Check1.Value = 1 Then
     Me.Width = 8235
     Else
     Me.Width = 4875
  End If
End Sub

Private Sub cmdAbort_Click()
  ' Abbrechen
  Me.Tag = False
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  ' OK
  Me.Tag = True
  Me.Hide
End Sub

Public Sub Form_Load()
frmInput.Width = 4875
IniRead
cmdOK.Enabled = True
Me.txtDate.Text = Now()
Me.txtAutor.Text = sAutor
Me.txtAutor.Enabled = False
Me.txtDate.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Schlieﬂen
  If UnloadMode <> 1 Then
    Cancel = True
    cmdAbort.Value = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Beenden
 ' For i = 0 To Me.CmbTestTyp.ListCount - 1
 ' Next i
  
  Set frmInput = Nothing

End Sub


