VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPurchasesReport 
   Caption         =   "Pharmacy Purchases Report"
   ClientHeight    =   7785
   ClientLeft      =   1590
   ClientTop       =   705
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   8775
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crPurchase 
      Left            =   840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   855
      Left            =   8520
      Picture         =   "frmPurchasesReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Range"
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   20
      Top             =   600
      Width           =   8175
      Begin VB.CommandButton cmdViewReport1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":0580
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38725
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   3840
         TabIndex        =   22
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   17
      Top             =   2040
      Width           =   8175
      Begin VB.TextBox txtTotalto 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalFrom 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewReport2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":089C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "To"
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
         Left            =   3960
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Total Amount"
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
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criteria"
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   15
      Top             =   3480
      Width           =   8175
      Begin VB.TextBox txtOrder 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         ItemData        =   "frmPurchasesReport.frx":0BB8
         Left            =   600
         List            =   "frmPurchasesReport.frx":0BC8
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdViewReport3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":0C09
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   " = "
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
         Left            =   3240
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   13
      Top             =   4920
      Width           =   8175
      Begin VB.TextBox txtCustomerID 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdViewReport4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":0F25
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE REPORTS"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmPurchasesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdViewReport1_Click()
On Error Resume Next

crPurchase.ReportFileName = App.Path & "\Purchase.rpt"
crPurchase.DiscardSavedData = True
crPurchase.ReplaceSelectionFormula ("{qryPurchases.PurchaseOrderDate}   >=#" & DTPFrom & "#  and {qryPurchases.PurchaseOrderDate}  <=#" & DTPTo & "#  ")

crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport2_Click()
On Error Resume Next
Dim strReport As String
strReport = App.Path & "\Purchase.rpt"

crPurchase.ReportFileName = App.Path & "\Purchase.rpt"
crPurchase.DiscardSavedData = True
crPurchase.ReplaceSelectionFormula ("{qryPurchases.NetValue}   >=" & Val(txtTotalFrom) & "  and {qryPurchases.NetValue}  <=" & Val(txtTotalto) & "")


crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport3_Click()
On Error Resume Next
Dim strReport As String
strReport = App.Path & "\Purchase.rpt"


crPurchase.ReportFileName = App.Path & "\Purchase.rpt"
crPurchase.DiscardSavedData = True

If cmbOrder.ListIndex = 0 Then
crPurchase.ReplaceSelectionFormula ("{qryPurchases." & cmbOrder & "}  ='" & txtOrder & "'")
Else
crPurchase.ReplaceSelectionFormula ("{qryPurchases." & cmbOrder & "}  =" & txtOrder & "")
End If

crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport4_Click()
On Error Resume Next
Dim strReport As String
strReport = App.Path & "\Purchase.rpt"


crPurchase.ReportFileName = App.Path & "\Purchase.rpt"
crPurchase.DiscardSavedData = True

crPurchase.ReplaceSelectionFormula ("{qryPurchases.SupplierID}  =" & txtCustomerID & "")


crPurchase.WindowState = crptMaximized
crPurchase.Action = 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub



