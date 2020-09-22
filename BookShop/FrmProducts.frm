VERSION 5.00
Begin VB.Form FrmProducts 
   Caption         =   "Products"
   ClientHeight    =   6765
   ClientLeft      =   1260
   ClientTop       =   615
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      DataField       =   "currency_name"
      Height          =   315
      ItemData        =   "FrmProducts.frx":0000
      Left            =   4560
      List            =   "FrmProducts.frx":0010
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Products Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   4335
      Left            =   1200
      TabIndex        =   26
      Top             =   720
      Width           =   5055
      Begin VB.TextBox txtFields 
         DataField       =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SupplierName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   5
         Tag             =   "Chr"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Category_Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   2
         Tag             =   "Chr"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Book_ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Book_Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Tag             =   "Chr"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Author"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   4
         Tag             =   "num"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "UnitInStock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   8
         Tag             =   "Num"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "UnitPrice"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   9
         Tag             =   "Num"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ISDN"
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cmbCID 
         DataField       =   "CategoryName"
         Height          =   315
         Left            =   2400
         TabIndex        =   38
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbSID 
         DataField       =   "CategoryName"
         Height          =   315
         Left            =   2400
         TabIndex        =   39
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Author Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   36
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Book ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Book Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Subject Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Supplier Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   32
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unit Price:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   31
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Units In Stock:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   30
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Price in Rs:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   29
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "ISBN #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   27
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record Operations"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   5655
      Left            =   6840
      TabIndex        =   22
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   840
         Picture         =   "FrmProducts.frx":0030
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
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
         Height          =   780
         Left            =   2040
         Picture         =   "FrmProducts.frx":0535
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   840
         Picture         =   "FrmProducts.frx":0AA7
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
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
         Height          =   780
         Left            =   2040
         Picture         =   "FrmProducts.frx":1003
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2040
         Picture         =   "FrmProducts.frx":1562
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   840
         Picture         =   "FrmProducts.frx":1AE2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2040
         Picture         =   "FrmProducts.frx":1FE5
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   840
         Picture         =   "FrmProducts.frx":24E8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Record Navigation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1095
      Left            =   1320
      TabIndex        =   23
      Top             =   5160
      Width           =   5055
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   4320
         Picture         =   "FrmProducts.frx":2A0F
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   3600
         Picture         =   "FrmProducts.frx":2FAE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   840
         Picture         =   "FrmProducts.frx":352B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   120
         Picture         =   "FrmProducts.frx":3AB2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Width           =   1920
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT DETAILS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   5
      Left            =   3960
      TabIndex        =   25
      Top             =   0
      Width           =   3795
   End
End
Attribute VB_Name = "FrmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub cmbCID_KeyPress(KeyAscii As Integer)
Dim llngRet As Long
  Dim lstrFind As String
  Dim objcb As New Class1
  
  If KeyAscii >= 33 And KeyAscii <= 126 Then
    If cmbCID.SelLength = 0 Then
       lstrFind = cmbCID.Text & Chr(KeyAscii)
    Else
       lstrFind = Left(cmbCID.Text, cmbCID.SelStart) & Chr(KeyAscii)
    End If
    llngRet = objcb.WinCBFindString(cmbCID.hwnd, lstrFind, False)
    If llngRet <> -1 Then
       cmbCID.ListIndex = llngRet
       cmbCID.SelStart = Len(lstrFind)
       cmbCID.SelLength = Len(cmbCID.Text) - cmbCID.SelStart
       KeyAscii = 0
    End If
End If
End Sub

Private Sub cmbSID_KeyPress(KeyAscii As Integer)
Dim llngRet As Long
  Dim lstrFind As String
  Dim objcb As New Class1
  
  If KeyAscii >= 33 And KeyAscii <= 126 Then
    If cmbSID.SelLength = 0 Then
       lstrFind = cmbSID.Text & Chr(KeyAscii)
    Else
       lstrFind = Left(cmbSID.Text, cmbSID.SelStart) & Chr(KeyAscii)
    End If
    llngRet = objcb.WinCBFindString(cmbSID.hwnd, lstrFind, False)
    If llngRet <> -1 Then
       cmbSID.ListIndex = llngRet
       cmbSID.SelStart = Len(lstrFind)
       cmbSID.SelLength = Len(cmbSID.Text) - cmbSID.SelStart
       KeyAscii = 0
    End If
End If
End Sub

Private Sub cmbSID_LostFocus()
Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim SID As Long
    Dim Sup As String
         
    
        rs.Open "select * from Suppliers where CompanyName = '" & cmbSID.Text & "'", Con, adOpenKeyset, adLockPessimistic
        rs1.Open "select SupplierId from Suppliers", Con, adOpenKeyset, adLockPessimistic
            
        While rs1.EOF = False
            SID = rs1(0) + 1
            rs1.MoveNext
        Wend
        rs1.Close
    
        Sup = cmbSID.Text
        
       'While Not rs.EOF
            
        If rs.EOF = False Then
           Exit Sub
        
        Else
               rs.AddNew
                    rs.Fields(0) = SID
                    rs.Fields(1) = Sup
                rs.Update
                
                rs.Close
                
                cmbSID.Clear
                Dim rs2 As New ADODB.Recordset
        
                    rs2.Open "select CompanyName from Suppliers", Con, adOpenDynamic, adLockOptimistic
        
                    rs2.MoveFirst
        
                While rs2.EOF = False
        
                    cmbSID.AddItem rs2(0)
                    rs2.MoveNext
                Wend
            rs2.Close
        End If
        'Wend
          
                  cmbSID.Text = Sup
End Sub

Private Sub cmdViewAll_Click()
FrmVProducts.Show
End Sub

Private Sub cmbCID_LostFocus()
    'dim rsSearch as New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim CID As Long
    Dim Cat As String
         
    
        rs.Open "select * from categories where categoryname = '" & cmbCID.Text & "'", Con, adOpenKeyset, adLockPessimistic
        rs1.Open "select categoryID from categories", Con, adOpenKeyset, adLockPessimistic
            
        While rs1.EOF = False
            CID = rs1(0) + 1
            rs1.MoveNext
        Wend
        rs1.Close
    
        Cat = cmbCID.Text
        
       'While Not rs.EOF
            
        If rs.EOF = False Then
           Exit Sub
        
        Else
               rs.AddNew
                    rs.Fields(0) = CID
                    rs.Fields(1) = Cat
                rs.Update
                
                rs.Close
                
                cmbCID.Clear
                Dim rs2 As New ADODB.Recordset
        
                    rs2.Open "select CategoryName from Categories", Con, adOpenDynamic, adLockOptimistic
        
                    rs2.MoveFirst
        
                While rs2.EOF = False
        
                    cmbCID.AddItem rs2(0)
                    rs2.MoveNext
                Wend
            rs2.Close
        End If
        'Wend
          
                  cmbCID.Text = Cat
End Sub


Private Sub Combo1_Click()

    If txtFields(3).Text = "" Then
        MsgBox "Please enter money"
        Exit Sub
    End If
    
    Dim rs2 As New ADODB.Recordset
    Dim i As Long
    
    rs2.Open "select cur_value from MoneyChanger where currency_name = '" & Combo1.Text & "' ", Con, adOpenDynamic, adLockPessimistic
    
    i = Val(txtFields(3).Text) * rs2(0)
    
    
    txtFields(8).Text = i
    
    rs2.Close
End Sub

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

Dim rsCategories As Recordset
Dim rsSuppliers As Recordset

Set rsCategories = New ADODB.Recordset
Set rsSuppliers = New ADODB.Recordset

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from Books", Con, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
  On Error Resume Next
  Set oText.DataSource = adoPrimaryRS
     oText.Enabled = True
     oText.Locked = True
  Next

  mbDataChanged = False
    rsSuppliers.Open "select * from Suppliers", Con, adOpenDynamic, adLockOptimistic
If rsSuppliers.EOF = False Then

rsSuppliers.MoveFirst
While rsSuppliers.EOF = False
cmbSID.AddItem rsSuppliers(1)
rsSuppliers.MoveNext
Wend

End If

rsCategories.Open "select * from Categories", Con, adOpenDynamic, adLockOptimistic
Debug.Print rsCategories.RecordCount
Debug.Print rsSuppliers.RecordCount

'If rsCategories.EOF = False Then
    rsCategories.MoveFirst

    While rsCategories.EOF = False

        cmbCID.AddItem rsCategories(1)
        rsCategories.MoveNext
    Wend

'End If

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()

On Error GoTo AddErr

    Dim rsAddProducts As New Recordset
    Dim PID As Long
    Set rsAddProducts = New ADODB.Recordset

    
    'PID = Functions.UID(6, "MedID_")
    rsAddProducts.Open " Select max(Book_ID) from books", Con, adOpenKeyset, adLockPessimistic
    While rsAddProducts.EOF = False
        PID = rsAddProducts(0) + 1
        rsAddProducts.MoveNext
    Wend
        rsAddProducts.Close

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  
  txtFields(1).SetFocus

  txtFields(0).Locked = True
  
  CurrencyManage
  txtFields(3).Locked = False
  Combo1.Locked = False
  txtFields(8).Locked = True
  
  
  cmbCID.ZOrder 0
  cmbSID.ZOrder 0
  cmbCID.TabIndex = 2
  cmbSID.TabIndex = 6
  txtFields(7).TabIndex = 26
  txtFields(9).TabIndex = 27
  
  cmdUpdate.Enabled = True
  cmdCancel.Enabled = True

  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0).Text = PID
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
   If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then
    Exit Sub
  End If

  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

    txtFields(1).SetFocus

  cmdUpdate.Enabled = True
  cmdCancel.Enabled = True
  
     Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  
  txtFields(3).Locked = True

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  
  CurrencyManage
  txtFields(3).Locked = False
  Combo1.Locked = False
  txtFields(8).Locked = True
  
   cmbCID.ZOrder 0
  cmbSID.ZOrder 0
  
 cmbCID.TabIndex = 2
  cmbSID.TabIndex = 6
  txtFields(7).TabIndex = 26
  txtFields(9).TabIndex = 27
  
  Exit Sub
  

EditErr:
  MsgBox err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

    cmdUpdate.Enabled = False
  cmdCancel.Enabled = False
  
  txtFields(3).Locked = True
  Combo1.Locked = True

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  
   cmbCID.ZOrder 1
  cmbSID.ZOrder 1

    cmbCID.TabIndex = 26
  cmbSID.TabIndex = 27
  txtFields(7).TabIndex = 2
  txtFields(9).TabIndex = 6
End Sub

Private Sub cmdUpdate_Click()

    txtFields(7).Text = cmbCID.Text
    txtFields(9).Text = cmbSID.Text

  On Error GoTo UpdateErr

'    txtFields(2).Text = cmbSID.Text
'    txtFields(3).Text = cmbCID.Text

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If
  
   cmbCID.ZOrder 1
  cmbSID.ZOrder 1
  
  cmbCID.TabIndex = 26
  cmbSID.TabIndex = 27
  txtFields(7).TabIndex = 2
  txtFields(9).TabIndex = 6
  
  cmdUpdate.Enabled = False
  cmdCancel.Enabled = False
  
  txtFields(3).Locked = True
  Combo1.Locked = True

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Enabled = bVal
  cmdEdit.Enabled = bVal
  
  cmdDelete.Enabled = bVal
  cmdClose.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  cmdViewAll.Enabled = bVal
End Sub


Private Sub CurrencyManage()
    Dim rs2 As New ADODB.Recordset
    Combo1.Clear
    
    rs2.Open "select currency_name from moneychanger", Con, adOpenDynamic, adLockOptimistic
    
    While rs2.EOF = False
        Combo1.AddItem rs2(0)
        Combo1.Text = rs2(0)
        rs2.MoveNext
    Wend
    
    rs2.Close
        
End Sub
