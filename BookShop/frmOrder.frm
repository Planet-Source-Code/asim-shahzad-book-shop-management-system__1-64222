VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSaleInvoice 
   Caption         =   "Invoice"
   ClientHeight    =   8490
   ClientLeft      =   1485
   ClientTop       =   450
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport PInvoice 
      Left            =   3360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4320
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Remove Selected Item"
      Height          =   855
      Left            =   10080
      Picture         =   "frmOrder.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cndew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3360
      Picture         =   "frmOrder.frx":0503
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   1185
   End
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add to the List"
      Default         =   -1  'True
      Height          =   855
      Left            =   10080
      Picture         =   "frmOrder.frx":09B6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4800
      Picture         =   "frmOrder.frx":0EC5
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   6000
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7680
      Picture         =   "frmOrder.frx":1437
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click To Close"
      Top             =   6000
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6240
      Picture         =   "frmOrder.frx":19B7
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdVProducts 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.ComboBox cmbPID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox cmbCID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtpayable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtdisgvn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtBillID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox txtStock 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmbPtID 
         Height          =   315
         ItemData        =   "frmOrder.frx":1E6F
         Left            =   6120
         List            =   "frmOrder.frx":1E71
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtmedname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txttotamt 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtdis 
         Height          =   285
         Left            =   10320
         TabIndex        =   10
         Tag             =   "Amt"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtrpu 
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Tag             =   "Num"
         Top             =   1560
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPIssue 
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   43843585
         CurrentDate     =   38353
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Total"
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
         Height          =   195
         Left            =   4200
         TabIndex        =   27
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Units Available"
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
         Height          =   375
         Left            =   4200
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Height          =   195
         Left            =   4200
         TabIndex        =   37
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   195
         Left            =   8520
         TabIndex        =   29
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Given"
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
         Height          =   195
         Left            =   8520
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Per Unit"
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
         Height          =   195
         Left            =   8520
         TabIndex        =   26
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date"
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
         Height          =   195
         Left            =   4200
         TabIndex        =   24
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name"
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
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
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
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1935
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ForeColorSel    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   9360
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   43843585
      CurrentDate     =   38353
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   11415
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      Height          =   195
      Left            =   480
      TabIndex        =   32
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES INVOICE"
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
      Left            =   4080
      TabIndex        =   39
      Top             =   0
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   11535
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount Payable"
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
      Height          =   195
      Left            =   7800
      TabIndex        =   36
      Top             =   5520
      Width           =   1980
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Given"
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
      Height          =   195
      Left            =   4080
      TabIndex        =   35
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      Height          =   195
      Left            =   840
      TabIndex        =   34
      Top             =   5520
      Width           =   1140
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Height          =   195
      Left            =   8400
      TabIndex        =   33
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "frmSaleInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Dim DesignX As Integer
Dim DesignY As Integer

Dim stock As Integer
Dim OID As String
Dim rsProductName As New ADODB.Recordset

Private Sub cmbCID_Click()

Dim rsProducts As Recordset
Set rsProducts = New ADODB.Recordset
cmbPID.Clear

rsProducts.Open "Select * from books where Category_name = '" & cmbCID & "' ", Con, adOpenDynamic, adLockOptimistic

If rsProducts.EOF = False Then
    rsProducts.MoveFirst

    While rsProducts.EOF = False
        cmbPID.AddItem rsProducts(0)
        cmbPID.Text = rsProducts(0)
        rsProducts.MoveNext
    Wend
End If
If cmbPID.ListCount = 0 Then
txtmedname = ""
txtRPU = "0"
End If

rsProducts.Close

End Sub

Private Sub cmbPID_Click()

rsProductName.Open "Select * from Books where Book_ID = " & cmbPID & "", Con, adOpenStatic, adLockBatchOptimistic


If rsProductName.RecordCount > 1 Then
    MsgBox " Database Error"
    stock = 0
    txtStock = stock
    Exit Sub
ElseIf rsProductName.RecordCount = 0 Then
    txtmedname = ""
    txtRPU = "0.00"
    stock = 0

Else
    txtmedname = rsProductName(1)
    txtRPU = rsProductName(6)
    stock = rsProductName(7)

End If

txtStock = stock

rsProductName.Close

End Sub



Private Sub cmbPtID_Click()
Dim rs3 As New ADODB.Recordset

rs3.Open "Select ContactFirstName,ContactLastName from customers where CustomerID = " & cmbPtID & "", Con, adOpenStatic, adLockBatchOptimistic


If rs3.RecordCount > 1 Then
    MsgBox " Database Error"
  
    Exit Sub
ElseIf rs3.RecordCount = 0 Then
    Text1 = ""
    

Else
    Text1.Text = rs3.Fields(0) & " " & rs3.Fields(1)

End If
rs3.Close
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim selectedRow As Integer
Dim tot As Long
Dim rsProdName As New ADODB.Recordset
Dim i As Integer
selectedRow = MFG.Row


'Add quatity back
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        rsProdName.Open "Select * from Books where Book_ID = " & Val(MFG.TextMatrix(i, 2)) & "", Con, adOpenStatic, adLockOptimistic
        
        tot = Val(rsProdName(7)) + Val(MFG.TextMatrix(i, 4))
        'tot = Val(txtStock.Text) + Val(MFG.TextMatrix(i, 4))
        
        rsProdName(7) = tot
    
        'txtStock.Text = rsProdName(7)
    
        rsProdName.UpdateBatch adAffectCurrent
        rsProdName.Close
    Next i
End If

If selectedRow = MFG.Rows - 1 Then
    MsgBox "Invalid Selection.", vbCritical
    Exit Sub
End If


If Not MFG.TextMatrix(1, 1) = "" Then
    MFG.RemoveItem (selectedRow)
    Call CalFinal
End If



End Sub

Private Sub cmdVProducts_Click()
FrmVProducts.Show
End Sub

Private Sub cndew_Click()
Dim ctl As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    End If
Next
cmbPtID.Enabled = True
cmdSave.Enabled = True
MFG.Clear
MFG.Refresh
MFG.Rows = 2

Call SetData
Call BillID
Call CustDetails
Call MFGVALUES
Command1.Enabled = False
DTPDate.Value = Date
DTPIssue.Value = Date
End Sub

Private Sub Command1_Click()
'On Error Resume Next
Dim strReport As String
Dim strTXT As Integer

strTXT = Val(txtBillID.Text)
'PInvoice.DataFiles(0) = App.Path & "\BookShop.mdb"
'strReport = App.Path & "\invoice.rpt"
PInvoice.ReportFileName = App.Path & "\Sale Invoice.rpt"

PInvoice.DiscardSavedData = True

PInvoice.SelectionFormula = "{qryInvoice.OrderID} = " & strTXT
PInvoice.WindowState = crptMaximized
PInvoice.Action = 1
End Sub

Private Sub Command2_Click()
MsgBox MFG.Row
MsgBox MFG.Rows
End Sub

Private Sub Form_Load()

            
Call SetData
Call BillID
Call CustDetails
Call MFGVALUES
Command1.Enabled = False
DTPDate.Value = Date
DTPIssue.Value = Date


End Sub

Public Sub SetData()

Dim rsCategories As Recordset
Set rsCategories = New ADODB.Recordset

rsCategories.Open "select * from Categories", Con, adOpenDynamic, adLockOptimistic
cmbCID.Clear
While rsCategories.EOF = False
cmbCID.AddItem rsCategories(1)
rsCategories.MoveNext

Wend
rsCategories.Close


End Sub

Public Sub BillID()


   Dim BID As Double
   Dim rsOrderID As Recordset
    Set rsOrderID = New ADODB.Recordset
    ' Generatin Order Details ID
        
    'BID = Functions.UID(6, "MODRID_")
    
    rsOrderID.Open "Select max(orderID) from Orders", Con, adOpenKeyset, adLockOptimistic
    
    
    
    If rsOrderID.EOF = False Then
    
        If IsNull(rsOrderID(0)) Then
                BID = 1
                'rsOrderID.MoveFirst
        Else
        While rsOrderID.EOF = False
            BID = rsOrderID(0) + 1
        rsOrderID.MoveNext
        Wend
        End If
    End If

txtBillID = BID


End Sub

Public Sub MFGVALUES()
MFG.TextMatrix(0, 1) = "ORDER ID"
MFG.TextMatrix(0, 2) = "PRODUCT ID"
MFG.TextMatrix(0, 3) = "PRODUCT NAME"
MFG.TextMatrix(0, 4) = "QUANTITY"
MFG.TextMatrix(0, 5) = "UNIT PRICE"
MFG.TextMatrix(0, 6) = "DISCOUNT"
MFG.TextMatrix(0, 7) = "TOTAL AMOUNT"
'SizeColumnHeaders MFG, Me

End Sub

Public Sub CustDetails()

Dim rsAddCust As Recordset
Set rsAddCust = New ADODB.Recordset

rsAddCust.Open "Select * from Customers", Con, adOpenDynamic, adLockReadOnly

cmbPtID.Clear
If rsAddCust.EOF = False Then
rsAddCust.MoveFirst

'cmbPtID.Text = "0"

While rsAddCust.EOF = False
    cmbPtID.AddItem rsAddCust(0)
    cmbPtID.ListIndex = 0
    rsAddCust.MoveNext
Wend



End If

'cmbPID.Text = Str(rsAddCust(0)) & " "

rsAddCust.Close


End Sub







Private Sub MFG_Click()
      rsProductName.Open "Select * from Books where Book_ID = " & cmbPID & "", Con, adOpenStatic, adLockBatchOptimistic
      
      txtStock = stock

rsProductName.Close
End Sub

Private Sub txtdis_Change()
Dim dis As Long
txtAmount = Val(txtRPU) * Val(txtqty)
dis = (Val(txtAmount) * Val(txtdis)) / 100
txttotamt = Val(txtAmount) - dis
End Sub

Private Sub txtqty_Change()

txtAmount = Val(txtRPU) * Val(txtqty)
txttotamt = Val(txtAmount) - Val(txtdis)
End Sub


Private Sub cmdAddList_Click()
'On Error Resume Next
    Dim rsMed As Recordset
    Dim i As Integer
    Dim tot As Double
    Dim rsProdName As New ADODB.Recordset
   
    
    
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        If MFG.TextMatrix(i, 2) = cmbPID Then
            MsgBox "Medicine Already Exist In The List Cannot Add Same Medicine Again.....", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next i
End If

If txtAmount = "" Or txttotamt = "" Or txtqty = "" Or txtRPU = "" Then
    MsgBox "Please Enter the relevant Fields"
    Exit Sub
End If
If Val(txtqty) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    Exit Sub
End If
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
    Exit Sub
End If
If cmbPtID = "" Then
    'cmbPtID = ""
    Exit Sub
End If

'Remove quantity
rsProdName.Open "Select * from Books where Book_ID = " & cmbPID & "", Con, adOpenStatic, adLockOptimistic

tot = Val(txtStock.Text) - Val(txtqty.Text)
rsProdName(7) = tot


txtStock.Text = rsProdName(7)

rsProdName.UpdateBatch adAffectCurrent
rsProdName.Close
'

Row = MFG.Rows - 1
With MFG

        .Rows = .Rows + 1
                
        MFG.TextMatrix(Row, 1) = txtBillID
        MFG.TextMatrix(Row, 2) = cmbPID
        MFG.TextMatrix(Row, 3) = txtmedname
        MFG.TextMatrix(Row, 4) = txtqty
        MFG.TextMatrix(Row, 5) = txtRPU
        MFG.TextMatrix(Row, 6) = txtdis
        MFG.TextMatrix(Row, 7) = txttotamt
        
  
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     'SizeColumns MFG, Me
     MFGVALUES
     
     Row = Row + 1
     
End With

cmbPtID.Enabled = False

Call CalFinal
Call TextClear

End Sub
Public Sub CalFinal()
Dim amount As Double
Dim Discount As Double
Dim Total As Double

If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        amount = amount + (Val(MFG.TextMatrix(i, 5)) * Val(MFG.TextMatrix(i, 4)))
        Discount = Discount + Val(MFG.TextMatrix(i, 6))
        Total = Total + Val(MFG.TextMatrix(i, 7))
        
    Next i
End If
txtgrndtot = amount
txtdisgvn = Discount
txtpayable = Total
Debug.Print Val(amount) - Val(Discount)

End Sub

Public Sub TextClear()
txtqty = ""
txtdis = ""
txtAmount = ""
txttotamt = ""

cmbPID_Click
End Sub



Private Sub cmdsave_click()
    Dim rsOrderID As Recordset

   If MFG.Rows = 2 Then
    MsgBox "Please Add Items to list before you save", vbCritical, "Error Occured"
    Exit Sub
   End If
   
Dim flag, flag1, flag2 As Boolean
flag = False
flag1 = False
flag2 = False

   
    Set rsOrderID = New ADODB.Recordset
   
    rsOrderID.Open " Select * from Orders", Con, adOpenDynamic, adLockPessimistic
      
    
    rsOrderID.AddNew
        rsOrderID(0) = txtBillID
        rsOrderID(1) = cmbPtID
        rsOrderID(2) = DTPDate
        rsOrderID(3) = Text1.Text
        rsOrderID(4) = txtdisgvn
        rsOrderID(5) = txtpayable
    rsOrderID.Update
    flag2 = True
    
    rsOrderID.Close
    
Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim MID As Long
Dim RQuantity As Integer

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset
Dim rsStock As Recordset
Set rsStock = New ADODB.Recordset

    

rsMed.Open "SELECT * FROM OrderDetails", Con, adOpenDynamic, adLockPessimistic

For i = 1 To MFG.Rows - 2 Step 1

    ' Generatin Order Details ID
    'MID = Functions.UID(6, "ODRDTL_")
    
    rsAddPatient.Open " Select * from OrderDetails", Con, adOpenDynamic, adLockReadOnly
    If rsAddPatient.EOF = False Then
    While rsAddPatient.EOF = False
'        If rsAddPatient(0) = MID Then
'            MID = Functions.UID(6, "ODRDTL_")
'            rsAddPatient.MoveFirst
'        End If
    MID = rsAddPatient(0) + 1
    rsAddPatient.MoveNext
    Wend
    End If
    rsAddPatient.Close

        With rsMed
            .AddNew
                !OrderDetailID = MID
                !OrderID = txtBillID
                !ProductID = MFG.TextMatrix(i, 2)
                !QUANTITY = MFG.TextMatrix(i, 4)
                !UNITPRICE = MFG.TextMatrix(i, 5)
                !amount = Val(MFG.TextMatrix(i, 4)) * Val(MFG.TextMatrix(i, 5))
            .Update
            flag = True
        End With
        
        ' Substract the stock from the products
        rsStock.Open "select * from Books where book_ID= " & MFG.TextMatrix(i, 2) & "", Con, adOpenDynamic, adLockPessimistic
        If rsStock.EOF = False Then
            rsStock(4) = rsStock(4) - Val(MFG.TextMatrix(i, 4))
            rsStock.Update
            flag1 = True
        End If
        rsStock.Close
Next

rsMed.Close

If flag = True And flag1 = True And flag2 = True Then
    MsgBox "Record Saved Succesfully !!"
Else
    MsgBox "Error Updating Record", vbCritical
End If

Command1.Enabled = True
cmdSave.Enabled = False



End Sub

Private Sub txtqty_LostFocus()
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
'    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub


