VERSION 5.00
Begin VB.Form frmContainer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   735
   ClientTop       =   0
   ClientWidth     =   3015
   Icon            =   "frmContainer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7710
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image11 
      Height          =   660
      Left            =   480
      Picture         =   "frmContainer.frx":030A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1920
   End
   Begin VB.Shape Shape10 
      Height          =   1095
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Shape Shape9 
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      Height          =   1215
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Shape Shape5 
      Height          =   1215
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   1095
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image9 
      Height          =   975
      Left            =   1800
      MouseIcon       =   "frmContainer.frx":2488
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":2792
      Top             =   2640
      Width           =   960
   End
   Begin VB.Image Image8 
      Height          =   975
      Left            =   1800
      MouseIcon       =   "frmContainer.frx":2BE6
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":2EF0
      Top             =   5280
      Width           =   960
   End
   Begin VB.Image Image7 
      Height          =   1020
      Left            =   1680
      MouseIcon       =   "frmContainer.frx":3794
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":3A9E
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Image Image6 
      Height          =   990
      Left            =   240
      MouseIcon       =   "frmContainer.frx":4010
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":431A
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Image Image5 
      Height          =   1020
      Left            =   240
      MouseIcon       =   "frmContainer.frx":47AD
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":4AB7
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Image Image4 
      Height          =   870
      Left            =   240
      MouseIcon       =   "frmContainer.frx":5075
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":537F
      Top             =   240
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   1680
      MouseIcon       =   "frmContainer.frx":5728
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":5A32
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      MouseIcon       =   "frmContainer.frx":5CF7
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":6001
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   240
      MouseIcon       =   "frmContainer.frx":64BE
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":67C8
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1110
   End
   Begin VB.Image Image10 
      Height          =   1065
      Left            =   1560
      MouseIcon       =   "frmContainer.frx":6B2A
      MousePointer    =   99  'Custom
      Picture         =   "frmContainer.frx":6E34
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer


Private Sub Form_Load()
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
   
    MyForm.Height = Me.Height ' Remember the current size
    MyForm.Width = Me.Width
End Sub

Private Sub Image1_Click()
    FrmCategories.Show
End Sub

Private Sub Image10_Click()
    FrmPurchases.Show
End Sub

Private Sub Image2_Click()
    FrmProducts.Show
End Sub

Private Sub Image3_Click()
    frmBusinessInfo.Show 1
End Sub

Private Sub Image4_Click()
    frmSaleInvoice.Show
End Sub

Private Sub Image5_Click()
    frmPurchasesReport.Show
End Sub

Private Sub Image6_Click()
    FrmCustomers.Show
End Sub

Private Sub Image7_Click()
    frmSalesReport.Show
End Sub

Private Sub Image8_Click()
frmBackUp.Show 1
End Sub

Private Sub Image9_Click()
    FrmSuppliers.Show
End Sub
