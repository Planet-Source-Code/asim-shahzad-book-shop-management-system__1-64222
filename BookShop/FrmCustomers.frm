VERSION 5.00
Begin VB.Form FrmCustomers 
   Caption         =   "Pharmacy - Customer Details"
   ClientHeight    =   7905
   ClientLeft      =   1530
   ClientTop       =   450
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   9810
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Navigation"
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
      Height          =   855
      Left            =   2040
      TabIndex        =   22
      Top             =   5880
      Width           =   8535
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   7200
         Picture         =   "FrmCustomers.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6480
         Picture         =   "FrmCustomers.frx":059F
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   2025
         Picture         =   "FrmCustomers.frx":0B1C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1320
         Picture         =   "FrmCustomers.frx":10A3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   3360
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
      Height          =   4935
      Left            =   7560
      TabIndex        =   25
      Top             =   840
      Width           =   3015
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   360
         Picture         =   "FrmCustomers.frx":1637
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1560
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
         Left            =   1560
         Picture         =   "FrmCustomers.frx":1B5E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
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
         Left            =   360
         Picture         =   "FrmCustomers.frx":2061
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
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
         Left            =   1560
         Picture         =   "FrmCustomers.frx":2564
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
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
         Left            =   1560
         Picture         =   "FrmCustomers.frx":2AE4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2520
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
         Left            =   360
         Picture         =   "FrmCustomers.frx":3043
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
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
         Left            =   1560
         Picture         =   "FrmCustomers.frx":359F
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
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
         Left            =   360
         Picture         =   "FrmCustomers.frx":3B11
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Product Details"
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
      Height          =   4935
      Left            =   2040
      TabIndex        =   26
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtFields 
         DataField       =   "ContactTitle"
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   3
         Tag             =   "Chr"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CustomerID"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ContactFirstName"
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   1
         Tag             =   "Chr"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ContactLastName"
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   2
         Tag             =   "Chr"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "EmailAddress"
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   9
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PhoneNumber"
         Height          =   285
         Index           =   6
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Num"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "FaxNumber"
         Height          =   285
         Index           =   7
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Num"
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "BillingAddress"
         Height          =   885
         Index           =   10
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PostalCode"
         Height          =   285
         Index           =   13
         Left            =   2520
         TabIndex        =   5
         Tag             =   "Num"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cmbContactTitle 
         DataField       =   "ContactTitle"
         Height          =   315
         ItemData        =   "FrmCustomers.frx":4016
         Left            =   2520
         List            =   "FrmCustomers.frx":4023
         TabIndex        =   27
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "NIC"
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Num"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerID:"
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
         Left            =   720
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Left            =   720
         TabIndex        =   36
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   720
         TabIndex        =   35
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "ContactTitle:"
         DataField       =   "ContactTitle:"
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
         Left            =   720
         TabIndex        =   34
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "EmailAddress:"
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
         Index           =   5
         Left            =   600
         TabIndex        =   33
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "PhoneNumber:"
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
         Left            =   600
         TabIndex        =   32
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "FaxNumber:"
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
         Left            =   600
         TabIndex        =   31
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "BillingAddress:"
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
         Left            =   720
         TabIndex        =   30
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "PostalCode:"
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
         Index           =   13
         Left            =   600
         TabIndex        =   29
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "NIC #:"
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
         Left            =   600
         TabIndex        =   28
         Top             =   3360
         Width           =   1815
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
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
      Left            =   3840
      TabIndex        =   24
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "FrmCustomers"
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

Private Sub cmdViewAll_Click()
FrmVCustomers.Show

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
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from Customers", Con, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Enabled = True
    oText.Locked = True
  Next

  mbDataChanged = False
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
    
    Dim rsAddCustomers As New Recordset
    Dim CID As Long
    Set rsAddCustomers = New ADODB.Recordset
  
'    CID = Functions.UID(6, "CID_")
    rsAddCustomers.Open " Select max(CustomerID) from Customers", Con, adOpenKeyset, adLockPessimistic
 While rsAddCustomers.EOF = False
        CID = rsAddCustomers(0) + 1
        rsAddCustomers.MoveNext
    Wend
        rsAddCustomers.Close
    
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  
  cmbContactTitle.ZOrder 0
    
  txtFields(0).Locked = True
  
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0).Text = CID
    
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

     Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next

    cmbContactTitle.ZOrder 0

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  
  cmbContactTitle.ZOrder 1
  
    Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = True
  Next

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

End Sub

Private Sub cmdUpdate_Click()
  txtFields(4).Text = cmbContactTitle.Text
  
  cmbContactTitle.ZOrder 1
  
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If
  
    Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = True
  Next

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
  cmbContactTitle.Visible = Not bVal
  cmdViewAll.Enabled = bVal
  
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
'    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
'End Sub


