VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function SetKey(ByVal sSection As String, ByVal sKey As String, ByVal vntValue As Variant)

    WritePrivateProfileString sSection, sKey, CStr(vntValue), App.Path & "\desktop\desktop.ini"

End Function
Public Function GetKey(ByVal sSection As String, ByVal sKey As String, default As String) As String

    Dim xRet As Long
    Dim sReturnStr As String
    Dim nStringLen As Integer
  
    nStringLen = 1024
    sReturnStr = String(nStringLen, Chr$(0))
  
    xRet = GetPrivateProfileString(sSection, sKey, "", sReturnStr, nStringLen, App.Path & "\desktop\desktop.ini")

    If xRet = 0 Then
        sReturnStr = default
        xRet = Len(default)
    End If
    GetKey = Left(sReturnStr, xRet)
  
End Function
Public Sub DelKey(ByVal sSection As String, sKey As String)

    WritePrivateProfileString sSection, CStr(sKey), 0&, App.Path & "\desktop\desktop.ini"

End Sub
Public Sub CreateSection(ByVal sSection As String)

    WritePrivateProfileString sSection, "", "", App.Path & "\desktop\desktop.ini"

End Sub
Public Sub DelSection(ByVal sSection As String)

    WritePrivateProfileString sSection, 0&, 0&, App.Path & "\desktop\desktop.ini"

End Sub
Public Function GetKeys(ByVal sSection As String) As String()

    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    Dim keys() As String
  
    ReDim keys(0)
    
    sBuf = Space$(1024)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(sSection, 0&, "Null", sBuf, iSize, App.Path & "\mmtools.ini")
  
    If (iSize > 0) Then
        sBuf = Left$(sBuf, iRetCode)
        While Len(sBuf) > 0
            ReDim Preserve keys(UBound(keys) + 1)
            keys(UBound(keys) - 1) = Left$(sBuf, InStr(1, sBuf, Chr$(0)) - 1)
            sBuf = Mid(sBuf, InStr(1, sBuf, Chr$(0)) + 1)
        Wend
        GetKeys = keys
    Else
        GetKeys = keys
    End If
    
    ReDim keys(0)
    sBuf = Space$(0)
    
End Function
