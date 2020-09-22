VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVProducts 
   Caption         =   "View All Products"
   ClientHeight    =   8490
   ClientLeft      =   960
   ClientTop       =   585
   ClientWidth     =   8880
   Icon            =   "FrmVProducts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
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
      Height          =   1335
      Left            =   2760
      TabIndex        =   9
      Top             =   5520
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   795
         Left            =   4320
         Picture         =   "FrmVProducts.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   795
         Left            =   600
         Picture         =   "FrmVProducts.frx":0DCE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   2400
         Picture         =   "FrmVProducts.frx":1251
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtSearchText 
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   5040
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product ID"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ISDN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Supplier ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit Stock"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Units Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Reorder Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW ALL PRODUCTS"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   4305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Search For"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "FrmVProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Dim strCol As Variant

Private Sub cmbSearch_Click()
cmdFind_Click
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()

Dim rsFind As Recordset
Dim strSQl As String
Dim SQL As String
Dim LItem As ListItem

'if there is nothing to search for then exit
If txtSearchText = "" Then
    Exit Sub
End If

ListView1.ListItems.Clear

Set rsFind = New ADODB.Recordset

       SQL = "SELECT * FROM books"
       SQL = SQL & " WHERE BookID LIKE '*" & txtSearchText & "*'"



'make the search
        strSQl = "SELECT * FROM books WHERE "
        strSQl = strSQl & cmbSearch & " Like " & "'%" & txtSearchText & "%'"

        'SQL = strSQl & " WHERE language LIKE '*" & Text1.Text & "*'"
        'strSQl = strSQl & SQL
        Debug.Print strSQl
        Debug.Print SQL

'show the found records
    rsFind.Open strSQl, Con, adOpenDynamic, adLockPessimistic


    Debug.Print rsFind.RecordCount
    Debug.Print rsFind.Fields.Count

    If Not (rsFind.BOF And rsFind.EOF) Then
        While rsFind.EOF = False
        Set LItem = ListView1.ListItems.Add(, , rsFind(0))

        If rsFind(1) <> "" Then
            LItem.SubItems(1) = rsFind(1)
        End If

        If rsFind(2) <> "" Then
            LItem.SubItems(2) = rsFind(2)
        End If

        If rsFind(3) <> "" Then
            LItem.SubItems(3) = rsFind(3)
        End If

        If rsFind(4) <> "" Then
            LItem.SubItems(4) = rsFind(4)
        End If

        If rsFind(5) <> "" Then
            LItem.SubItems(5) = rsFind(5)
        End If

        If rsFind(6) <> "" Then
            LItem.SubItems(6) = rsFind(6)
        End If


        rsFind.MoveNext
        Wend
    End If


 'show number of records found
    Me.Caption = CStr(rsFind.RecordCount) & " records found"

 'close the recordset
    rsFind.Close


End Sub

Private Sub cmdSearch_Click()
Dim LItem As ListItem

FindItem = InputBox("Enter Product ID", "Find Product Details")

If Not FindItem = "" Then
Set LItem = ListView1.FindItem(FindItem, lvwText, lvwSubItem)
If LItem Is Nothing Then
    NotFound = True
End If
If NotFound Then
    MsgBox "Item not found", vbInformation, "Search Result"
Else
    LItem.EnsureVisible
    LItem.Selected = True
End If
End If


End Sub


Private Sub Command1_Click()
txtSearchText = ""
Form_Load
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
txtSearchText.SetFocus
End Sub

Private Sub Form_Load()

Dim LItem As ListItem
Dim i As Integer

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

cmbSearch.Clear

Dim rsProductID As Recordset
Set rsProductID = New ADODB.Recordset
Dim rsProduct As Recordset
Set rsProduct = New ADODB.Recordset

rsProductID.Open "select * from books", Con, adOpenDynamic, adLockPessimistic

rsProduct.Open "select * from books", Con, adOpenDynamic, adLockPessimistic


For i = 0 To rsProductID.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsProductID(i).Name, i
Next i
rsProductID.Close

ListView1.ListItems.Clear



    If Not (rsProduct.BOF And rsProduct.EOF) Then
        While rsProduct.EOF = False
        Set LItem = ListView1.ListItems.Add(, , rsProduct(0))

        If rsProduct(1) <> "" Then
            LItem.SubItems(1) = rsProduct(1)
        End If

        If rsProduct(2) <> "" Then
            LItem.SubItems(2) = rsProduct(2)
        End If

        If rsProduct(3) <> "" Then
            LItem.SubItems(3) = rsProduct(3)
        End If

        If rsProduct(4) <> "" Then
            LItem.SubItems(4) = rsProduct(4)
        End If

        If rsProduct(5) <> "" Then
            LItem.SubItems(5) = rsProduct(5)
        End If

        If rsProduct(6) <> "" Then
            LItem.SubItems(6) = rsProduct(6)
        End If


rsProduct.MoveNext
Wend
End If
rsProduct.Close
cmbSearch.Text = cmbSearch.List(0)


End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

If strCol <> ColumnHeader Then
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = ColumnHeader.Index - 1
    strCol = ColumnHeader
Else
    ListView1.SortOrder = lvwDescending
    ListView1.SortKey = ColumnHeader.Index - 1
    strCol = ""
End If


End Sub

Private Sub txtSearchText_Change()
cmdFind_Click
End Sub

