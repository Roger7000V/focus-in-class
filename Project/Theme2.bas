Attribute VB_Name = "Module3"
Public Declare Function SkinH_Attach Lib "SkinH_VB6.dll" () As Long
Public Declare Function SkinH_AttachEx Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal strBuffer As String, ByVal lngSize As Long) As Long
Private Const MAX_PATH = 260
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const MF_BYPOSITION = &H400&

Public Function TheSystemDir() As String
    Dim strBuffer As String
    Dim l As Long
    
    strBuffer = Space(255)
    l = GetSystemDirectory(strBuffer, 255)
    TheSystemDir = Left(strBuffer, l)
    
End Function
Public Function windir() As String
    Dim lpBuffer As String
    lpBuffer = Space$(MAX_PATH)
    windir = Left$(lpBuffer, GetWindowsDirectory(lpBuffer, MAX_PATH))
End Function
