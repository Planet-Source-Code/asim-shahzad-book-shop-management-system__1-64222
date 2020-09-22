VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVCustomers 
   Caption         =   "View All Customers"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "FrmVCustomers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtSearchText 
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Top             =   5040
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7011
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer ID"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Billing Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Postal Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Phone Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "NIC #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "EmailAddress"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   2880
      TabIndex        =   8
      Top             =   5520
      Width           =   5775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   2400
         Picture         =   "FrmVCustomers.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   795
         Left            =   600
         Picture         =   "FrmVCustomers.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   795
         Left            =   4320
         Picture         =   "FrmVCustomers.frx":11F3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW ALL CUSTOMERS"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   4545
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
      Left            =   1920
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
      Left            =   6000
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "FrmVCustomers"
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

       SQL = "SELECT * FROM Customers"
       SQL = SQL & " WHERE CustomerID LIKE '*" & txtSearchText & "*'"



'make the search
        strSQl = "SELECT * FROM Customers WHERE "
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

        If rsFind(7) <> "" Then
            LItem.SubItems(7) = rsFind(7)
        End If

        If rsFind(8) <> "" Then
            LItem.SubItems(8) = rsFind(8)
        End If

        If rsFind(9) <> "" Then
            LItem.SubItems(9) = rsFind(9)
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

FindItem = InputBox("Enter Customers ID", "Find Customers Details")

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


Dim rsPatientID As Recordset
Set rsPatientID = New ADODB.Recordset
Dim rsPatient As Recordset
Set rsPatient = New ADODB.Recordset

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

rsPatientID.Open "select * from Customers", Con, adOpenDynamic, adLockPessimistic

rsPatient.Open "select * from Customers", Con, adOpenDynamic, adLockPessimistic


For i = 0 To rsPatientID.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsPatientID(i).Name, i
Next i
rsPatientID.Close

ListView1.ListItems.Clear

    If Not (rsPatient.BOF And rsPatient.EOF) Then
        While rsPatient.EOF = False
        Set LItem = ListView1.ListItems.Add(, , rsPatient(0))

        If rsPatient(1) <> "" Then
            LItem.SubItems(1) = rsPatient(1)
        End If

        If rsPatient(2) <> "" Then
            LItem.SubItems(2) = rsPatient(2)
        End If

        If rsPatient(3) <> "" Then
            LItem.SubItems(3) = rsPatient(3)
        End If

        If rsPatient(4) <> "" Then
            LItem.SubItems(4) = rsPatient(4)
        End If

        If rsPatient(5) <> "" Then
            LItem.SubItems(5) = rsPatient(5)
        End If

        If rsPatient(6) <> "" Then
            LItem.SubItems(6) = rsPatient(6)
        End If

        If rsPatient(7) <> "" Then
            LItem.SubItems(7) = rsPatient(7)
        End If

        If rsPatient(8) <> "" Then
            LItem.SubItems(8) = rsPatient(8)
        End If

        If rsPatient(9) <> "" Then
            LItem.SubItems(9) = rsPatient(9)
        End If

              
rsPatient.MoveNext
Wend
End If
rsPatient.Close
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


