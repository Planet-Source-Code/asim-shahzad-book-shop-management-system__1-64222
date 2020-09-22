Attribute VB_Name = "ModSub"

'Procedure used to highlight text when focus
Public Sub HLText(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cmdViewAll.Visible = bVal
  
End Sub

Public Sub SizeColumns(ByVal flx As MSFlexGrid, frm As Form)

Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
        'wid = TextWidth(flx.TextMatrix(r, c))
        wid = frm.TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + wid
    Next c
 
End Sub
