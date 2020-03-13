Attribute VB_Name = "Module2"
Sub Main()
    Dim bas As String
    bas = App.Path & "Skin"
    Dim OPEN1() As Byte
    OPEN1 = LoadResData(101, "CUSTOM")
    Open bas For Binary As #1
    Put #1, , OPEN1
    Close #1
    
    Dim dll As String
    dll = App.Path & "SkinH_VB6.dll"
    Dim OPEN3() As Byte
    OPEN3 = LoadResData(102, "CUSTOM")
    Open dll For Binary As #1
    Put #1, , OPEN3
    Close #1
    
    Dim she As String
    she = App.Path & "\Theme\Aero.she"
    Dim OPEN2() As Byte
    OPEN2 = LoadResData(103, "CUSTOM")
    Open she For Binary As #1
    Put #1, , OPEN2
    Close #1
    
    Form1.Show
End Sub

