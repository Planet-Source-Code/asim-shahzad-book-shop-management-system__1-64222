VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "New Capital Book Shop..."
   ClientHeight    =   6240
   ClientLeft      =   2760
   ClientTop       =   1725
   ClientWidth     =   7110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7050
      TabIndex        =   0
      Top             =   5625
      Width           =   7110
      Begin VB.PictureBox picAd 
         Appearance      =   0  'Flat
         BackColor       =   &H00464646&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   735
         ScaleWidth      =   3135
         TabIndex        =   2
         Top             =   0
         Width           =   3135
      End
      Begin SHDocVwCtl.WebBrowser webAdvisory 
         Height          =   975
         Left            =   3120
         TabIndex        =   1
         Top             =   -120
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   1720
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":097D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":230F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5633
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8957
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A2E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D60D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EFA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1055D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11239
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F15
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":138CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":145A9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu backup 
         Caption         =   "Backup &Database"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdNew 
         Caption         =   "&Add New User"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPass 
         Caption         =   "Change Password"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "&LogOut"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuManage 
      Caption         =   "&Magement"
      Begin VB.Menu mnuAdProd 
         Caption         =   "Add New Product"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAdCust 
         Caption         =   "Add New Customer"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAdEmp 
         Caption         =   "Add New Employee"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuAdSup 
         Caption         =   "Add New Supplier"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transactions"
      Begin VB.Menu mnuSale 
         Caption         =   "Sales Invoice"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnnuPurInv 
         Caption         =   "Purchase Invoice"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "&Reports"
      Begin VB.Menu mnuSaleRep 
         Caption         =   "&Sales Report"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPurRep 
         Caption         =   "Purchase Report"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCur 
         Caption         =   "Currency Controler"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "&Calculator"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuNote 
         Caption         =   "&Notepad"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "&Application Skin"
         Begin VB.Menu Default 
            Caption         =   "Default"
         End
         Begin VB.Menu MacGrey 
            Caption         =   "Mac Grey"
         End
         Begin VB.Menu XPBlue 
            Caption         =   "XP Blue"
         End
         Begin VB.Menu CoolGreen 
            Caption         =   "Cool Green"
         End
         Begin VB.Menu LightBrown 
            Caption         =   "Light Brown"
         End
         Begin VB.Menu LightViolet 
            Caption         =   "Light Violet"
         End
         Begin VB.Menu WinClassic 
            Caption         =   "Win Classic"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCont 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredit 
         Caption         =   "Cre&dits"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub backup_Click()
frmBackUp.Show 1
End Sub

Private Sub CoolGreen_Click()
Call select_color_type(3)
sys_color = "3"
End Sub

Private Sub Default_Click()
Call select_color_type(0)
sys_color = "0"
End Sub

Private Sub LightBrown_Click()
Call select_color_type(5)
sys_color = "5"
End Sub

Private Sub LightViolet_Click()
Call select_color_type(4)
sys_color = "4"
End Sub

Private Sub MacGrey_Click()
Call select_color_type(1)
sys_color = "1"
End Sub

Private Sub MDIForm_Load()
 
    Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
    ' Size of Form in Pixels at design resolution
    DesignX = 800
    DesignY = 600
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
    ' Set up the screen values
    Xtwips = Screen.TwipsPerPixelX
    Ytwips = Screen.TwipsPerPixelY
    Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
    Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

    ' Determine scaling factors
    ScaleFactorX = (Xpixels / DesignX)
    ScaleFactorY = (Ypixels / DesignY)
    ScaleMode = 1  ' twips
    'Exit Sub  ' uncomment to see how Form1 looks without resizing
    Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
    '
   '
    MyForm.Height = Me.Height ' Remember the current size
    MyForm.Width = Me.Width
 
    frmContainer.Show
    
    
       'Display the business status
    UpdateInfoMsg
    
        
End Sub

Public Sub UpdateInfoMsg()
    Dim strHTML As String
    Screen.MousePointer = vbHourglass
    ' Header html
    strHTML = "<html><body topmargin=9 leftmargin=0 bgcolor=#" & Hex$(80) & Hex$(80) & Hex$(80) & "><b>"

    ' Body html
    strHTML = strHTML & "<marquee direction=left scrolldelay=75>"

    'For products
    '- For no. of products
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "->&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total products = " & getRecordCount("books") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
    '- For inventory cost
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(238) & Hex$(238) & Hex$(238) & ">" & _
                            "Current inventory value = Rs " & toMoney(getSumOfCost) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
                        "</font>"

    'For Sales
    '- For sales this month
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "->&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total sales in this month = Rs " & toMoney(getSumOfFields) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"

    '- For  sales this year
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(128) & Hex$(191) & Hex$(28) & ">" & _
                            "Total sales in this year = Rs " & toMoney(getSumOfYearly) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
    
    strHTML = strHTML & "</marquee>"

    ' Footer html
    strHTML = strHTML & "</b></body></html>"

    Open Environ$("TMP") & "\SunyeAdvisory.tmp" For Output As #1
        Print #1, strHTML
    Close #1

    strHTML = vbNullString

    'Call SavePicture(ig24x24.ListImages(1).Picture, Environ$("TMP") & "\ar.bmp")
    webAdvisory.Navigate Environ$("TMP") & "\SunyeAdvisory.tmp"
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnsCust_Click()
    FrmCustomers.Show 1
End Sub

Private Sub mnnuPurInv_Click()
FrmPurchases.Show
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAdCust_Click()
FrmCustomers.Show
End Sub

Private Sub mnuAdNew_Click()
frmUserDetails.Show
End Sub

Private Sub mnuAdProd_Click()
FrmProducts.Show
End Sub

Private Sub mnuAdSup_Click()
FrmSuppliers.Show
End Sub

Private Sub mnuCalc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuCredit_Click()
frmCredits.Show 1
End Sub


Private Sub mnuCur_Click()
    frmCurrency.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuLogout_Click()
If MsgBox("Are you sure you want to Log Off the system ?", vbYesNo + vbQuestion, "Log off") = vbYes Then
    
    Unload Me
    frmLogin.Show
End If
End Sub

Private Sub mnuNote_Click()
Shell "Notepad.exe", vbNormalFocus
End Sub

Private Sub mnuPass_Click()
frmChangePassword.Show
End Sub

Private Sub mnuPurRep_Click()
    frmPurchasesReport.Show
End Sub

Private Sub mnuSale_Click()
frmSaleInvoice.Show
End Sub

Private Sub mnuSaleRep_Click()
    frmSalesReport.Show
End Sub

Private Sub WinClassic_Click()
Call select_color_type(6)
sys_color = "6"
End Sub

Private Sub XPBlue_Click()
Call select_color_type(2)
sys_color = "2"
End Sub
