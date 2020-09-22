VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2895
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4290
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3360
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BookStore.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BookStore.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Users"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Log In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   1140
   End
   Begin VB.TextBox txtPasswd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   1680
      Width           =   2325
   End
   Begin MSDataListLib.DataCombo dcUser 
      Bindings        =   "frmLogin.frx":0E42
      DataField       =   "User_Name"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ListField       =   "User_Name"
      BoundColumn     =   "User_Name"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your username and enter your password in the space provided bellow."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   720
      TabIndex        =   6
      Top             =   150
      Width           =   3315
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000005&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000005&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   0
      Picture         =   "frmLogin.frx":0E57
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   765
      Left            =   600
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rs As New ADODB.Recordset
Public LoginSucceeded As Boolean
Dim Counter As Integer

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
On Error Resume Next

    If txtPasswd.Text = "" Then
        MsgBox "Enter A Password:", vbCritical
        txtPasswd.SetFocus
        Exit Sub
    End If
    
    'rs.Close  'After the user_id are loaded into dbUserId close the
                  'Recordset for further usage
    rs.Open "Select User_Password from users where User_Name= '" & dcUser.Text & "'", Con, adOpenDynamic, adLockOptimistic
    
    If rs.EOF <> True Then 'If Search is found
         If rs(0) = txtPasswd Then
             
             
             DoEvents
             'Check If user is Admin
             'Checking : To allow Settings Menu Available only to Admin
             
             Unload Me
            
             
             frmMain.Show
             DoEvents
        '     Srchflag = True
             Exit Sub
             rs.Close
         Else
             MsgBox "Invalid Password!!!" & vbCrLf & "Note : Password is same as Username", vbInformation, "Enjoy Freeware"
             txtPasswd.Text = ""
             txtPasswd.SetFocus
             rs.Close
             Exit Sub
         End If
    End If
    
    
    'If Srchflag = False Then 'Display msg when search not found
    '     MsgBox "Invalid Password" & vbCrLf & "No Access!!!", vbCritical, "Invalid User"
    '     End
    'End If
End Sub


