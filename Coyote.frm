VERSION 5.00
Begin VB.Form Coyote 
   BackColor       =   &H00C0C000&
   Caption         =   "Create System SQL- DSN"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Coyote.frx":0000
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check Out My New DSN"
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create SQL System DSN"
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "Coyote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'**    Created by Coyote and COYOTE CODE IS COOL           **
'**    http://www.coyotecavern.com/                        **
'**    Use this as you like, no copyright, no restrictions **
'**    I assume NO Responsibility for this code.           **
'**    Copy, Use, Revise, Or even Distribute as your code  **
'**    All I want is to Give, because I have Received      **
'**    So much help from others.  Thank You !              **
'************************************************************
Private Sub Command1_Click()
Dim DriverODBC As String
Dim NameDSN As String

DriverODBC = String(255, Chr(32))
'Edit the next line and change to the DSN name you want to use
NameDSN = "SQLtestDSN"
'Have SQL drivers been installed?


If Not checkSQLDriver(DriverODBC) Then
    MsgBox "You must Install SQL ODBC Drivers before use this program.", vbOKOnly + vbCritical
    MsgBox "Program Being Terminated.", vbOKOnly + vbCritical
    End
End If

'Does the DSN name already exist?

If (SQLDSNWanted(NameDSN)) = True Then



        MsgBox "Program Terminated-DSN Name Already Exist.", vbOKOnly + vbCritical
        End
    Else


        If Not MakeSQLDSN(DriverODBC, NameDSN) Then
            MsgBox "Error Occured-DSN could not be Created", vbOKOnly + vbCritical
        End If
End If
MsgBox "SQL System DSN Has Been Created.", vbOKonlyOnly



End Sub

Private Sub Command2_Click()
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL ODBCCP32.cpl @2, 5")

End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Text1.Text = "When you click the button below, this program will create a new System DSN.  It will be set up to connect to a SQL Server named MyServer and to the Northwind database. It uses NT Authentication, however can be revised to require login and password.(reference code)"
End Sub
