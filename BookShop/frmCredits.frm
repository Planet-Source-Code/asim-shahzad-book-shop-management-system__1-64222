VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Us"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   2
      Left            =   1920
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
      Begin VB.Label L6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "hammar@zaintech.com +920300-5154573"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label L5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Hammar Ahmad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.PictureBox P 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   0
      Left            =   3360
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      Begin VB.Label L3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Sajjad Haider Hani"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label L4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "hani@zaintech.com +920300-5221398"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox P 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      Begin VB.Label L2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "asimshahzad78@hotmail.com +920345-5202322"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Asim Shahzad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Timer T 
      Interval        =   100
      Left            =   3360
      Top             =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Timer tmrScrollTitle 
      Interval        =   100
      Left            =   3840
      Top             =   4200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ZainTech.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmCredits.frx":0E42
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetInitialCaption "Book Shop", 80, Me
End Sub

Private Sub Label1_Click()
ShellExecute Hwnd, "open", "http://www.zaintech.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub T_Timer()
If L1.Top <= -800 Then L1.Top = 1560
If L2.Top <= -800 Then L2.Top = 1560
If L3.Top <= -800 Then L1.Top = 1560
If L4.Top <= -800 Then L2.Top = 1560
If L5.Top <= -800 Then L5.Top = 1560
If L6.Top <= -800 Then L6.Top = 1560

L1.Top = L1.Top - 15
L2.Top = L2.Top - 15
L3.Top = L1.Top - 15
L4.Top = L2.Top - 15
L5.Top = L5.Top - 15
L6.Top = L6.Top - 15
End Sub

Private Sub tmrScrollTitle_Timer()
    ScrollTitle "Book Shop Management System", 80, Me
End Sub
Public Sub SetInitialCaption(Cap As String, Spaces As Integer, FormName As Form)
    FormName.Caption = Space(Spaces)
    FormName.Caption = FormName.Caption + Cap
End Sub
Public Sub ScrollTitle(Cap As String, Spaces As Integer, FormName As Form)
    If Not FormName.Caption = "" Then
        FormName.Caption = Right(FormName.Caption, (Len(FormName.Caption) - 1))
    Else
        Call SetInitialCaption(Cap, Spaces, FormName)
    End If
End Sub
