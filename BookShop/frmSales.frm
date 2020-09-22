VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmSales 
   Caption         =   "Sales Report..."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptSales
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
Dim s As String

Report.Database.SetDataSource Con

Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
'CRViewer1.Refresh

RefreshViewer

Screen.MousePointer = vbDefault


 Set fld1 = Report.Field1
'    fld1. SetUnboundFieldSource frmSaleInvoice.cmbPtID


CRViewer1.ViewReport
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

' Filter the data in the viewer to display only the records that we want to see.
'
Private Sub RefreshViewer()
    Dim s As String                         ' Temporary variable of convenience
        
    Screen.MousePointer = vbHourglass
    
    
    If ReportChoice = 1 Then
        s = s & "{ado.OrderDate}   >=#" & frmSalesReport.DTPFrom & "#  and {ado.OrderDate}  <=#" & frmSalesReport.DTPTo & "#"
    ElseIf ReportChoice = 2 Then
        s = s & "{qrySales.NetValue}   >=" & Val(frmSalesReport.txtTotalFrom) & "  and {qrySales.NetValue}  <=" & Val(frmSalesReport.txtTotalto) & ""
    ElseIf ReportChoice = 3 Then
        If frmSalesReport.cmbOrder.ListIndex = 3 Or frmSalesReport.cmbOrder.ListIndex = 4 Or frmSalesReport.cmbOrder.ListIndex = 5 Or frmSalesReport.cmbOrder.ListIndex = 6 Then
            s = s & ("{qrySales." & frmSalesReport.cmbOrder & "}  =" & frmSalesReport.txtOrder & "")
        Else
            s = s & ("{qrySales." & frmSalesReport.cmbOrder & "}  =" & frmSalesReport.txtOrder & "")
        End If
    ElseIf ReportChoice = 4 Then
        s = s & "{qrySales.CustomerID}  =" & frmSalesReport.txtCustomerID & ""
    End If
        
   Report.RecordSelectionFormula = s   ' Apply the new filter-formula to the report
    'RefreshLabels                           ' Refresh the "useful information" labels
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
    
End Sub

