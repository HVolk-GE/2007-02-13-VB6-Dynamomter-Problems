VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anmeldung"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Initialen :"
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'Globale Variable auf False setzen,
    'um eine fehlgeschlagene Anmeldung zu kennzeichnen.
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
Dim newUser As String
    If Me.txtUserName <> "" And Me.txtUserName <> " " And Me.txtUserName <> "  " Then
        LoginSucceeded = True
        Author = Trim(UCase(Me.txtUserName))
        newUser = WriteValue("dbconnect", "autor", Author)
        IniRead
        frmInput.txtAutor = sAutor
        frmMain.Show
        Me.Hide
    End If
End Sub
