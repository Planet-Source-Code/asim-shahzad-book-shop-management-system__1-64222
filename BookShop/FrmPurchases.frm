VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FrmPurchases 
   Caption         =   "Purchases"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Remove Selected Item"
      Height          =   855
      Left            =   9960
      Picture         =   "FrmPurchases.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4560
      Width           =   1695
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
      Height          =   795
      Left            =   5640
      Picture         =   "FrmPurchases.frx":0503
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cndew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2760
      Picture         =   "FrmPurchases.frx":09BB
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
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
      Height          =   795
      Left            =   6960
      Picture         =   "FrmPurchases.frx":0E6E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Click To Close"
      Top             =   6360
      Width           =   1185
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
      Height          =   795
      Left            =   4200
      Picture         =   "FrmPurchases.frx":13EE
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   6360
      Width           =   1275
   End
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add to the List"
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
      Height          =   855
      Left            =   9960
      Picture         =   "FrmPurchases.frx":1960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtpayable 
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
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox txtdisgvn 
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1935
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   8
      SelectionMode   =   1
   End
   Begin VB.TextBox txtBillID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   11430
      Begin VB.CommandButton cmdVSuppliers 
         Caption         =   "..."
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdVProducts 
         Caption         =   "..."
         Height          =   255
         Left            =   7800
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtNet 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtdis 
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
         Left            =   9960
         TabIndex        =   10
         Tag             =   "num"
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtRPU 
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
         Left            =   9960
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtUPurchased 
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
         Left            =   6120
         TabIndex        =   8
         Tag             =   "Num"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtunits 
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPName 
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbPID 
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
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSCName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtSName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cmbSID 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Height          =   495
         Left            =   8400
         TabIndex        =   37
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
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
         Height          =   495
         Left            =   8400
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
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
         Left            =   8400
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
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
         Height          =   495
         Left            =   8400
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Units Purchased"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   33
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Units In Stock"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   4440
         TabIndex        =   31
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   9720
      TabIndex        =   27
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
      Format          =   43778049
      CurrentDate     =   38729
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   11295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE ORDER"
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
      Left            =   3960
      TabIndex        =   41
      Top             =   0
      Width           =   3705
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
      Left            =   7200
      TabIndex        =   40
      Top             =   5760
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
      Left            =   3600
      TabIndex        =   39
      Top             =   5760
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
      Left            =   600
      TabIndex        =   38
      Top             =   5760
      Width           =   1140
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
      Left            =   360
      TabIndex        =   29
      Top             =   720
      Width           =   615
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
      TabIndex        =   28
      Top             =   720
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   11430
   End
End
Attribute VB_Name = "FrmPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub cmbPID_Click()

Dim rsProdName As Recordset
Set rsProdName = New ADODB.Recordset


rsProdName.Open "Select * from Books where Book_ID = " & cmbPID & "", Con, adOpenDynamic, adLockReadOnly


If rsProdName.RecordCount > 1 Then
    MsgBox " Database Error"
    Exit Sub
Else
    txtPName = rsProdName(1)
    txtunits = rsProdName(7)
    txtRPU = rsProdName(6)
End If

rsProdName.Close

End Sub

Private Sub cmbSID_Click()

Dim rsSupplierName As Recordset
Set rsSupplierName = New ADODB.Recordset


rsSupplierName.Open "Select * from Suppliers where SupplierID = " & cmbSID & "", Con, adOpenDynamic, adLockReadOnly


If rsSupplierName.RecordCount > 1 Then
    MsgBox " Database Error"
    Exit Sub
Else
    txtSName = rsSupplierName(1)
    txtSCName = rsSupplierName(2)

End If

rsSupplierName.Close


End Sub

Private Sub cmdAddList_Click()

'On Error Resume Next
    Dim rsMed As Recordset
    Dim i As Integer
    Dim rsProdName As New ADODB.Recordset
    Dim tot As Long


If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        If MFG.TextMatrix(i, 2) = cmbPID Then
            MsgBox "Medicine Already Exist In The List Cannot Add Same Medicine Again.....", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next i
End If

If txtAmount = "" Or txtNet = "" Or txtunits = "" Or txtRPU = "" Then
    MsgBox "Please Enter the relevant Fields"
    Exit Sub
End If
If Val(txtUPurchased) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    Exit Sub
End If

'Remove quantity
rsProdName.Open "Select * from Books where Book_ID = " & cmbPID & "", Con, adOpenStatic, adLockOptimistic

tot = Val(txtunits.Text) - Val(txtUPurchased.Text)
rsProdName(7) = tot


txtunits.Text = rsProdName(7)

rsProdName.UpdateBatch adAffectCurrent
rsProdName.Close


Row = MFG.Rows - 1
With MFG

        .Rows = .Rows + 1

        MFG.TextMatrix(Row, 1) = txtBillID
        MFG.TextMatrix(Row, 2) = cmbPID
        MFG.TextMatrix(Row, 3) = txtPName
        MFG.TextMatrix(Row, 4) = txtUPurchased
        MFG.TextMatrix(Row, 5) = txtRPU
        MFG.TextMatrix(Row, 6) = txtdis
        MFG.TextMatrix(Row, 7) = txtNet



    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     SizeColumns MFG, Me
     MFGVALUES

     Row = Row + 1

End With


'cmbSID.Enabled = False
Call CalcFinal
Call TextClear


End Sub
Public Sub CalcFinal()

Dim amount As Double
Dim Discount As Double
Dim Total As Double

If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        amount = amount + (Val(MFG.TextMatrix(i, 5)) * Val(MFG.TextMatrix(i, 4)))
        Discount = Discount + MFG.TextMatrix(i, 6)
        Total = Total + MFG.TextMatrix(i, 7)

    Next i
End If
txtgrndtot = amount
txtdisgvn = Discount
txtpayable = Total
Debug.Print Val(amount) - Val(Discount)

End Sub

Public Sub TextClear()
txtunits = ""
txtUPurchased = ""
txtdis = "0"
txtAmount = ""
txtNet = ""

cmbPID_Click
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim selectedRow As Integer
Dim rsProdName As New ADODB.Recordset
Dim tot As Long
selectedRow = MFG.Row

'Add quatity back
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        rsProdName.Open "Select * from Books where Book_ID = " & Val(MFG.TextMatrix(i, 2)) & "", Con, adOpenStatic, adLockOptimistic

        tot = Val(rsProdName(7)) + Val(MFG.TextMatrix(i, 4))

        rsProdName(7) = tot

        'txtunits.Text = rsProdName(7)

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
    Call CalcFinal
End If
End Sub

Private Sub cmdsave_click()

   If MFG.Rows = 2 Then
    MsgBox "Please Add Items to list before you save", vbCritical, "Error Occured"
    Exit Sub
   End If

Dim flag, flag1, flag2 As Boolean

flag = False
flag1 = False
flag2 = False


    Dim rsOrderID As Recordset
    Dim OID As String
    Set rsOrderID = New ADODB.Recordset

    rsOrderID.Open " Select * from Purchase_Orders", Con, adOpenDynamic, adLockPessimistic



    rsOrderID.AddNew
        rsOrderID(0) = txtBillID
        rsOrderID(1) = cmbSID
        rsOrderID(2) = DTPDate

    rsOrderID.Update
    flag = True

    rsOrderID.Close

Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim MID As Long
Dim RQuantity As Integer

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset
Dim rsStock As Recordset
Set rsStock = New ADODB.Recordset



rsMed.Open "SELECT * FROM Purchase_Orde_Details", Con, adOpenDynamic, adLockPessimistic

For i = 1 To MFG.Rows - 2 Step 1

    ' Generating Purchase Order Details ID
    'MID = Functions.UID(6, "PODRDTL_")
    rsAddPatient.Open " Select * from Purchase_Orde_Details", Con, adOpenDynamic, adLockReadOnly
    If rsAddPatient.EOF = False Then
    While rsAddPatient.EOF = False
'        If rsAddPatient(0) = MID Then
'            MID = Functions.UID(6, "PODRDTL_")
'            rsAddPatient.MoveFirst
'        End If
    MID = rsAddPatient(0) + 1
    rsAddPatient.MoveNext
    Wend
    End If
    rsAddPatient.Close

        With rsMed
            .AddNew
                !PurchaseOrderDetailID = MID
                !PurchaseOrderID = txtBillID
                !PurchaseProductID = MFG.TextMatrix(i, 2)
                !PurchaseQUANTITY = Val(MFG.TextMatrix(i, 4))
                !PurchaseUnitPrice = Val(MFG.TextMatrix(i, 5))
                !PurchaseDiscount = Val(MFG.TextMatrix(i, 6))
                !NetValue = txtpayable
             .Update
             flag1 = True
        End With

        rsStock.Open "select * from books where Book_ID= " & MFG.TextMatrix(i, 2) & "", Con, adOpenDynamic, adLockPessimistic
        If rsStock.EOF = False Then
            rsStock(4) = rsStock(4) + Val(MFG.TextMatrix(i, 4))
            rsStock.Update
            flag2 = True
        End If
        rsStock.Close
Next
rsMed.Close
If flag = True And flag2 = True And flag1 = True Then
    MsgBox "Record Saved Succesfully !!", vbInformation, "Record Added"
    cmdSave.Enabled = False
Else
    MsgBox "An Error Occured while saving to the database", vbCritical
    Exit Sub
End If

Command1.Enabled = True
cmdSave.Enabled = False


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
cmbSID.Enabled = True
cmdSave.Enabled = True
MFG.Clear
MFG.Refresh
MFG.Rows = 2

Call SetData
Call BillID
Call ProdDetails
Call MFGVALUES


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
    


Call SetData
Call BillID
Call ProdDetails
Call MFGVALUES

End Sub

Public Sub SetData()

Dim rsSuppliers As Recordset
Set rsSuppliers = New ADODB.Recordset

  mbDataChanged = False
  rsSuppliers.Open "select * from Suppliers", Con, adOpenDynamic, adLockOptimistic

rsSuppliers.MoveFirst

cmbSID.Clear
While rsSuppliers.EOF = False
cmbSID.AddItem rsSuppliers(0)
rsSuppliers.MoveNext

Wend
rsSuppliers.Close


End Sub

Public Sub BillID()

   Dim BID As Long
   Dim rsOrderID As Recordset
   Set rsOrderID = New ADODB.Recordset
    ' Generatin Order Details ID

    'BID = UID(6, "MODRID_")
    rsOrderID.Open " Select * from Purchase_Orders", Con, adOpenDynamic, adLockPessimistic
    If rsOrderID.EOF = False Then
    While rsOrderID.EOF = False
'        If rsOrderID(0) = BID Then
'            BID = Functions.UID(6, "MODRID_")
'            rsOrderID.MoveFirst
'        End If

    BID = rsOrderID(0) + 1
    rsOrderID.MoveNext
    Wend
    End If
txtBillID = BID
rsOrderID.Close

End Sub

Public Sub MFGVALUES()
MFG.TextMatrix(0, 1) = "ORDER ID"
MFG.TextMatrix(0, 2) = "PRODUCT ID"
MFG.TextMatrix(0, 3) = "PRODUCT NAME"
MFG.TextMatrix(0, 4) = "QUANTITY"
MFG.TextMatrix(0, 5) = "UNIT PRICE"
MFG.TextMatrix(0, 6) = "DISCOUNT"
MFG.TextMatrix(0, 7) = "TOTAL AMOUNT"
SizeColumnHeaders MFG, Me

End Sub

Public Sub ProdDetails()

Dim rsAddProd As Recordset
Set rsAddProd = New ADODB.Recordset

rsAddProd.Open "Select * from Books", Con, adOpenDynamic, adLockReadOnly


cmbPID.Clear
If rsAddProd.EOF = False Then
rsAddProd.MoveFirst

While rsAddProd.EOF = False
    cmbPID.AddItem rsAddProd(0)
    'cmbPID.Text = rsAddProd(0)
    rsAddProd.MoveNext
Wend


End If

rsAddProd.Close


End Sub

Private Sub txtdis_Change()
Dim dis As Long
txtAmount = Val(txtRPU) * Val(txtUPurchased)
dis = (Val(txtAmount) * Val(txtdis)) / 100
txtNet = Val(txtAmount) - dis
End Sub

Private Sub txtUPurchased_Change()

txtAmount = Val(txtRPU) * Val(txtUPurchased)
txtNet = Val(txtAmount) - Val(txtdis)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

