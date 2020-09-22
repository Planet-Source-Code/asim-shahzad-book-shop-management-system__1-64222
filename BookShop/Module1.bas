Attribute VB_Name = "ModFunctions"


Public DBPath                       As String
Public CN                           As New Connection

'Function used to connect to database
'I created and use this function with SQL Server or Oracle only in client/server application together
'with CloseDB procedure.I use this function and procedure to connect only to
'the server if neccessary to save server resources (ex. When updating or displaying I use
'OpenDB and after the record displayed or update I use the closeDB).
'
'I DID NOT USE THE MAIN PURPOSE OF THIS CODDE WITH CloseDB BECAUSE
'THIS SYSTEM IS A STAND ALONE SYSTEM.
'
'--> This code is also available in VB.NET,J# and C# using ADO.NET. If you want it just e-mail me.

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

'Function used to get the end day number of a cetain month
Public Function getEndDay(ByVal srcDate As Date) As Byte
    Dim h1 As String
    h1 = Format(srcDate, "mm")
    On Error GoTo err
    Select Case h1
        Case Is = "01": getEndDay = 31
        Case Is = "02": getEndDay = Day(h1 & "/29/" & Format(srcDate, "yy"))
        Case Is = "03": getEndDay = 31
        Case Is = "04": getEndDay = 30
        Case Is = "05": getEndDay = 31
        Case Is = "06": getEndDay = 30
        Case Is = "07": getEndDay = 31
        Case Is = "08": getEndDay = 31
        Case Is = "09": getEndDay = 30
        Case Is = "10": getEndDay = 31
        Case Is = "11": getEndDay = 30
        Case Is = "12": getEndDay = 31
    End Select
    h1 = ""
    Exit Function
err:
        If err.Number = 13 Then getEndDay = 28: h1 = "" 'Day if encounter not a left-year
End Function
'Function that return true if the control is empty
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function
Public Function DataEntryValidation(Key As Integer, Param As String) As Integer
    'If BckSpace then allow
    If Key = 8 Then DataEntryValidation = Key: Exit Function
    'Enforce only Digits

    Select Case Param
        Case "Num"
        If Key < Asc("0") Or Key > Asc("9") Then
            DataEntryValidation = 0
        Else
            DataEntryValidation = Key
        End If
        
        Case "Amt"
        If Key < Asc("0") Or Key > Asc("9") Then

            If Key <> Asc(".") Then
                DataEntryValidation = 0
            Else
                DataEntryValidation = Key
            End If
        Else
            DataEntryValidation = Key
        End If
        
        Case "Chr"
        Key = Asc(UCase(Chr(Key)))
        If Key = 8 Or Key = 32 Then DataEntryValidation = Key: Exit Function
        If Key < Asc("A") Or Key > Asc("Z") Then
            DataEntryValidation = 0
        Else
            DataEntryValidation = Asc(UCase(Chr(Key)))
        End If
        
        Case "Cbo"
        DataEntryValidation = 0
        
 
        Case Else
        DataEntryValidation = Asc(UCase(Chr(Key)))
        
    End Select
End Function

' This function is used to create Unique ID's
Public Function UID(mLen As Integer, mPrefix As String) As String
    
    Dim mStr As String, i As Integer, j As Integer, mTable() As String * 1
    ReDim mTable(1 To 61)
    mTable(1) = "1": mTable(2) = "2": mTable(3) = "3": mTable(4) = "4"
    mTable(5) = "5": mTable(6) = "6": mTable(7) = "7": mTable(8) = "8"
    mTable(9) = "9": mTable(10) = "0"
    mTable(11) = "a": mTable(12) = "b": mTable(13) = "c": mTable(14) = "d"
    mTable(15) = "e": mTable(16) = "f": mTable(17) = "g": mTable(18) = "h"
    mTable(19) = "i": mTable(20) = "j": mTable(21) = "k": mTable(22) = "l"
    mTable(23) = "m": mTable(24) = "n": mTable(25) = "o": mTable(26) = "p"
    mTable(27) = "q": mTable(28) = "r": mTable(29) = "s": mTable(30) = "t"
    mTable(31) = "u": mTable(32) = "v": mTable(33) = "w": mTable(34) = "x"
    mTable(35) = "y": mTable(36) = "z"
    mTable(37) = "A": mTable(38) = "B": mTable(39) = "C": mTable(40) = "D"
    mTable(41) = "E": mTable(42) = "F": mTable(43) = "G": mTable(44) = "H":
    mTable(45) = "I": mTable(46) = "J": mTable(47) = "K": mTable(48) = "L"
    mTable(49) = "M": mTable(50) = "N": mTable(51) = "O": mTable(52) = "P"
    mTable(52) = "Q": mTable(53) = "R": mTable(54) = "S": mTable(55) = "T"
    mTable(56) = "U": mTable(57) = "V": mTable(58) = "W": mTable(59) = "X":
    mTable(60) = "Y": mTable(61) = "Z":
    mStr = mPrefix
    For i = 1 To mLen
    
    For j = 0 To 10
        DoEvents
        Next j
        Randomize
        mStr = mStr & mTable(Int((60) * Rnd + 1))
    Next i
    
        UID = mStr
        
End Function

Public Sub SizeColumnHeaders(ByVal flx As MSFlexGrid, frm As Form)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

max_wid = 0
        
    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        'wid = TextWidth(flx.TextMatrix(0, c))
         wid = frm.TextWidth(flx.TextMatrix(r, c))
        If max_wid < wid Then
            max_wid = wid
        End If
        
        flx.ColWidth(c) = max_wid + wid
    Next c
End Sub
'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT COUNT(book_ID) as TCount FROM " & srcTable & srcCondition, Con, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(rs![TCount], "#,##0")
    Else
        getRecordCount = rs![TCount]
    End If
    Set rs = Nothing
End Function
'Function that return the count of the rows in the customer table
Public Function getCustCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT COUNT(CustomerID) as TCount FROM " & srcTable & srcCondition, CN, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getCustCount = Format$(rs![TCount], "#,##0")
    Else
        getCustCount = rs![TCount]
    End If
    Set rs = Nothing
End Function
'Function used to get the sum  of Monthly Sale
Public Function getSumOfFields() As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * from OrderDetails_Query", Con, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfFields = getSumOfFields + rs.Fields(1)
            rs.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If

    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function


Public Function getSumOfCost() As Double
    'On Error GoTo err
    Dim rs As New ADODB.Recordset
    'Dim Total As Long

    'rs.CursorLocation = adUseClient
    rs.Open "SELECT Sum(UnitPrice)as Total FROM books", Con, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfCost = getSumOfCost + rs.Fields("Total")
            rs.MoveNext
        Loop
    Else
        getSumOfCost = 0
    End If
    
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfCost = 0: Resume Next
End Function
'Function used to get the sum  of Yearly Sale
Public Function getSumOfYearly() As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * from YearlySale_Query", Con, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfYearly = getSumOfYearly + rs.Fields(1)
            rs.MoveNext
        Loop
    Else
        getSumOfYearly = 0
    End If

    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfYearly = 0: Resume Next
End Function
