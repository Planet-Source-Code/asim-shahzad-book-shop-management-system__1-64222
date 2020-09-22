VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageStr Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, _
   lparam As String) As Long
   
   Private Const CB_FINDSTRING = &H14C
   Private Const CB_SHOWDROPDOWN = &H14F
   Private Const CB_FINDSTRINGEXACT = &H158
   
   Public Function WinCBFindString(ByVal phwnd As Long, _
          ByVal pSearch As String, ByVal flExact As Boolean) As Long
          Dim llngMess As Long
          If flExact Then
             llngMess = CB_FINDSTRINGEXACT
          Else
             llngMess = CB_FINDSTRING
          End If
          WinCBFindString = SendMessageStr(phwnd, llngMess, -1, ByVal pSearch)
    End Function
   Public Function WinCBShowDropDown(ByVal phwnd As Long) As Boolean
   Dim llngRet As Long
   
   llngRet = SendMessageStr(phwnd, CB_SHOWDROPDOWN, 1, vbNull)
   If llngRet = -1 Then
      WinCBShowDropDown = False
   Else
      WinCBShowDropDown = True
   End If
   End Function

