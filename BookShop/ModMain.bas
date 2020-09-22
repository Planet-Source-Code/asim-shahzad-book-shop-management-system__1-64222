Attribute VB_Name = "ModMain"

Public Con As ADODB.Connection
Public mySQL As String
Public ReportChoice As Integer

Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

Type FRMSIZE
   Height As Long
   Width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean

'The Project is start from here
Public Sub Main()

    If Date = "2/28/2007" Then
        MsgBox "Your License time limit is expired.", vbCritical
        End
    End If
    
    
    
    Set Con = New ADODB.Connection
    Con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; " & _
         "Persist Security Info = False; " & _
        "Data Source=" & App.Path & "\BookStore.mdb"
        
    Con.CursorLocation = adUseClient
    
    frmSplashScreen.Show
    'frmMain.Show
    'frmLogin.Show
    'frmCurrency.Show
End Sub

Sub Resize_For_Resolution(ByVal SFX As Single, _
                          ByVal SFY As Single, MyForm As Form)
    Dim i As Integer
    Dim SFFont As Single

    SFFont = (SFX + SFY) / 2  ' average scale
    ' Size the Controls for the new resolution
    On Error Resume Next  ' for read-only or nonexistent properties
    With MyForm
      For i = 0 To .Count - 1
       If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
         .Controls(i).Left = .Controls(i).Left * SFX
         .Controls(i).Top = .Controls(i).Top * SFY
         .Controls(i).Width = .Controls(i).Width * SFX
       Else
         .Controls(i).Move .Controls(i).Left * SFX, _
          .Controls(i).Top * SFY, _
          .Controls(i).Width * SFX, _
          .Controls(i).Height * SFY
       End If
         ' Be sure to resize and reposition before changing the FontSize
         .Controls(i).FontSize = .Controls(i).FontSize * SFFont
      Next i
      If RePosForm Then
        ' Now size the Form
        .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
      End If
    End With
End Sub

