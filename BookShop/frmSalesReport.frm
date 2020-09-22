VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSalesReport 
   Caption         =   "Sales Report"
   ClientHeight    =   7905
   ClientLeft      =   1605
   ClientTop       =   540
   ClientWidth     =   8730
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   7905
   ScaleWidth      =   8730
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crSale 
      Left            =   840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdViewReport1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Report"
      Height          =   975
      Left            =   8520
      Picture         =   "frmSalesReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Customer ID"
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
      Left            =   1680
      TabIndex        =   15
      Top             =   4800
      Width           =   8175
      Begin VB.CommandButton cmdViewReport4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6840
         Picture         =   "frmSalesReport.frx":031C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCustomerID 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   1440
         TabIndex        =   21
         Top             =   480
         Width           =   1215
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
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   8175
      Begin VB.CommandButton cmdViewReport3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6840
         Picture         =   "frmSalesReport.frx":0638
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         ItemData        =   "frmSalesReport.frx":0954
         Left            =   600
         List            =   "frmSalesReport.frx":0964
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtOrder 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   480
         Width           =   2535
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
         TabIndex        =   20
         Top             =   480
         Width           =   375
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
      Left            =   1680
      TabIndex        =   13
      Top             =   1920
      Width           =   8175
      Begin VB.CommandButton cmdViewReport2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   6840
         Picture         =   "frmSalesReport.frx":0994
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalFrom 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalto 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Tag             =   "Amt"
         Top             =   480
         Width           =   1455
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
         TabIndex        =   19
         Top             =   480
         Width           =   1575
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
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
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
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Width           =   8175
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19595265
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
         Format          =   19595265
         CurrentDate     =   38725
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
         TabIndex        =   17
         Top             =   480
         Width           =   510
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
         TabIndex        =   16
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   855
      Left            =   8640
      Picture         =   "frmSalesReport.frx":0CB0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES REPORT"
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
      TabIndex        =   22
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim DesignX As Integer
Dim DesignY As Integer
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdViewReport1_Click()
'On Error Resume Next

    Dim sDateTo As String
    Dim sDateFrom As String
    
    sDateTo = Format(DTPTo.Value, "mmmm dd yyyy")
    sDateFrom = Format(DTPFrom.Value, "mmmm dd yyyy")

crSale.ReportFileName = App.Path & "\Sales.rpt"
crSale.DiscardSavedData = True
'crSale.SelectionFormula = "{qrySales.OrderDate} >=#" & DTPFrom & "# and {qrySales.OrderDate} <=#" & DTPTo & "#"
crSale.SelectionFormula = "{qrySales.OrderDate} >= date('" & sDateFrom & "') And {qrySales.OrderDate} <= date('" & sDateTo & "')"

crSale.WindowState = crptMaximized
crSale.Action = 1

End Sub

Private Sub cmdViewReport3_Click()
crSale.ReportFileName = App.Path & "\Sales.rpt"
crSale.DiscardSavedData = True
If cmbOrder.ListIndex = 0 Or cmbOrder.ListIndex = 1 Then
    crSale.SelectionFormula = ("{qrySales." & cmbOrder & "}  ='" & txtOrder & "'")
Else
    crSale.SelectionFormula = ("{qrySales." & cmbOrder & "}  =" & txtOrder & "")
End If

crSale.WindowState = crptMaximized
crSale.Action = 1

End Sub

Private Sub cmdViewReport4_Click()
Dim strReport As String
strReport = App.Path & "\Sales.rpt"


crSale.ReportFileName = App.Path & "\Sales.rpt"
crSale.DiscardSavedData = True
crSale.SelectionFormula = "{qrySales.CustomerID}  =" & txtCustomerID


crSale.WindowState = crptMaximized
crSale.Action = 1


End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdViewReport2_Click()
On Error Resume Next
Dim strReport As String
strReport = App.Path & "\Sales.rpt"


crSale.ReportFileName = App.Path & "\Sales.rpt"
crSale.DiscardSavedData = True
crSale.ReplaceSelectionFormula ("{qrySales.NetValue}   >=" & Val(txtTotalFrom) & "  and {qrySales.NetValue}  <=" & Val(txtTotalto) & "")


crSale.WindowState = crptMaximized
crSale.Action = 1

End Sub

'Private Sub Form_Resize()
'    Dim ScaleFactorX As Single, ScaleFactorY As Single
'
'    If Not DoResize Then  ' To avoid infinite loop
'       DoResize = True
'       Exit Sub
'    End If
'
'    RePosForm = False
'    ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
'    ScaleFactorY = Me.Height / MyForm.Height
'    Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'    MyForm.Height = Me.Height ' Remember the current size
'    MyForm.Width = Me.Width
'End Sub

