VERSION 5.00
Begin VB.Form Coyote 
   BackColor       =   &H00C0C000&
   Caption         =   "Create System SQL- DSN"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   3180
      TabIndex        =   11
      Top             =   1680
      Width           =   2500
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   3180
      TabIndex        =   9
      Top             =   1140
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   3180
      TabIndex        =   8
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3180
      TabIndex        =   7
      Top             =   120
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   4200
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
      Height          =   720
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   4635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check Out My New DSN"
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create SQL System DSN"
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "OldNews"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF00FF&
      Caption         =   "Database Name:"
      BeginProperty Font 
         Name            =   "OldNews"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Server Name:"
      BeginProperty Font 
         Name            =   "OldNews"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   1140
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "DSN Name:"
      BeginProperty Font 
         Name            =   "OldNews"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   1875
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
' It's dirty but it works for now
' ! However DO NOT REMOVE the test for DSN name ! ! !
' !! DO NOT REMOVE the next 4 lines !!!!
    If Text2.Text = "" Then
        MsgBox "You must fill in a Name for the DSN.", vbOKOnly + vbCritical
        Exit Sub
    End If

    If Text3.Text = "" Then
        MsgBox "You must fill in a Description for the DSN.", vbOKOnly + vbCritical
        Exit Sub
        End If
    If Text4.Text = "" Then
        MsgBox "You must fill in a Server Name for the DSN.", vbOKOnly + vbCritical
        Exit Sub
        End If
    If Text5.Text = "" Then
        MsgBox "You must fill in a Database Name for the DSN.", vbOKOnly + vbCritical
        Exit Sub
            End If
        
'Edit the next line and change to the DSN name you want to use
NameDSN = Text2.Text
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
MsgBox "SQL System DSN Has Been Created.", vbOKOnlyOnly



End Sub

Private Sub Command2_Click()
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL ODBCCP32.cpl @2, 5")

End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Text1.Text = "Enter the above information to Create your SQL System DSN"
End Sub
