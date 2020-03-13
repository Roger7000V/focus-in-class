VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Focus-in-Class"
   ClientHeight    =   4050
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7320
   Icon            =   "Interface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Quicknotes 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Interface.frx":25CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Notepad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6240
      Top             =   960
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Actions"
      Begin VB.Menu mnuStart 
         Caption         =   "Start Class"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuStartForce 
         Caption         =   "Start Class (Force)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuRemain 
         Caption         =   "View Remining Time"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Quick Notes"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Import Quick Notes"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuDuration 
         Caption         =   "Set Class Duration"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuForce 
         Caption         =   "Enter Force Mode"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuTheme 
         Caption         =   "Theme"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Focus-in-Class"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim duration As String
Dim n As Integer
Dim FileType, FiType, FileName As String

' Define maximum form size
Private FormOldWidth  As Long ' Initial width
Private FormOldHeight  As Long ' Initial height

Private resize_state As Integer

' Hide taskbar
Private Function Fun_DisplayTaskBar(ByVal bShow As Boolean) As Integer
Dim lTaskBarHWND As Long
Dim lRet As Long
Dim lFlags As Long
On Error GoTo vbErrorHandler
lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
lTaskBarHWND = FindWindow("Shell_TrayWnd", "")
lRet = ShowWindow(lTaskBarHWND, lFlags)
If lRet < 0 Then
      Exit Function
End If
vbErrorHandler:
End Function

' Change the size of each component in the form in scale.
' Call the ReSizeInit function before calling ReSizeForm
Public Sub ResizeInit(FormName As Form)
    Dim Obj  As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
    Obj.Tag = Obj.Left & "    " & Obj.Top & "    " & Obj.Width & "    " & Obj.Height & "    "
    Next Obj
    On Error GoTo 0
End Sub

Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)      As Double
    Dim i      As Long, TempPos        As Long, StartPos        As Long
    Dim Obj      As Control
    Dim ScaleX      As Double, ScaleY        As Double
    ScaleX = FormName.ScaleWidth / FormOldWidth
    ' Save scale for form width
    ScaleY = FormName.ScaleHeight / FormOldHeight
    ' Save scale for form height
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
    ' Read controls'initial position and size
    TempPos = InStr(StartPos, Obj.Tag, "    ", vbTextCompare)
    If TempPos > 0 Then
        Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
    Else
        Pos(i) = 0
    End If
    ' Reposition and resize controls based on their original position and scale of the window resizing
    Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
     Next i
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    Dim mnuAbout_change As Boolean
    mnuAbout_change = False
    If mnuAbout_state = True Then
        mnuAbout_state = False
        mnuAbout_change = True
    End If
    If mnuAbout_change = True Then
        If resize_state = 1 Then
            ' Set Form on top
            Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    Else
        If resize_state = 1 Then
            MsgBox LoadResString(137)
            If FindWindow("Notepad", vbNullString) = 0 Then
                ' Set Form on top
                Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mnuAction.Caption = LoadResString(101)
    mnuStart.Caption = LoadResString(102)
    mnuStartForce.Caption = LoadResString(103)
    mnuSave.Caption = LoadResString(104)
    mnuOpen.Caption = LoadResString(105)
    mnuRemain.Caption = LoadResString(106)
    mnuSettings.Caption = LoadResString(107)
    mnuDuration.Caption = LoadResString(108)
    mnuForce.Caption = LoadResString(109)
    mnuTheme.Caption = LoadResString(134)
    mnuHelp.Caption = LoadResString(138)
    mnuAbout.Caption = LoadResString(139) & "Focus-in-Class"
    
    Label1.Caption = LoadResString(110)
    Label2.Caption = LoadResString(111)
    
    Command1.Caption = LoadResString(112)
    Command2.Caption = LoadResString(113)
    
    Quicknotes.Text = LoadResString(114)
    
    
    Me.AutoRedraw = True
    resize_state = 0
    Call ResizeInit(Me)    ' Make sure controls change as form changes
    Me.Picture = LoadPicture(App.Path & "\Background.jpg")
    
    SkinH_AttachEx App.Path & "\Aero.she", ""
    ' Set Form1 on top
    Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Resize()
    ' 0: focus mode off  1: focus mode on    2: window auto maximized (prevent dialog appearing twice)
    If resize_state = 1 Then
        MsgBox LoadResString(115), vbExclamation + vbOKOnly
        resize_state = 2
        WindowState = vbMaximized
    Else
        If resize_state = 2 Then
            resize_state = 1
        End If
    End If
    ResizeForm Me      ' Control changes
    Me.PaintPicture Me.Picture, 0, 0, Form1.Width, Form1.Height      ' Background image changes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuForce.Checked = True Then
        MsgBox LoadResString(116), vbExclamation + vbOKOnly
        Cancel = vbCancel
    Else
        Cancel = (MsgBox(LoadResString(117), vbQuestion + vbOKCancel) <> vbOK)
    End If
End Sub

Private Sub Command1_Click()
    If resize_state = 1 Then
        ' Set the precious top most form after calculator
        Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        
        Form2.Show 1
        
        Shell "calc.exe", 1
    Else
        Shell "calc.exe", 1
    End If
End Sub

Private Sub Command2_Click()
    If resize_state = 1 Then
        ' Set the precious top most form after notepad
        Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        
        Form2.Show 1
        
        Shell "notepad.exe", 1
        
        Dim note_hWnd As Long
        
        ' Find hwnd of Notepad
        note_hWnd = FindWindow("Notepad", vbNullString)
        
        ' Set Notepad on top
        Call SetWindowPos(note_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Shell "notepad.exe", 1
    End If
End Sub

Private Sub Label3_Click()
    Quicknotes.SelBold = True
End Sub

Private Sub Label4_Click()
    Quicknotes.SelItalic = True
End Sub

Private Sub Label5_Click()
    Quicknotes.SelUnderline = True
End Sub

Private Sub Label6_Click()
    Quicknotes.SelColor = RGB(255, 0, 0)
End Sub

Private Sub Label7_Click()
    Quicknotes.SelBold = False
End Sub

Private Sub Label8_Click()
    Quicknotes.SelItalic = False
End Sub

Private Sub Label9_Click()
    Quicknotes.SelUnderline = False
End Sub

Private Sub Label10_Click()
    Quicknotes.SelColor = RGB(0, 0, 0)
End Sub

Private Sub mnuAbout_Click()
    mnuAbout_state = True
    ' Set the precious top most form after notepad
    Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    frmAbout.Show 1
End Sub

Private Sub mnuTheme_Click()
    ' Set the precious top most form after notepad
    Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    Dialog.Show 1
End Sub

Private Sub mnuDuration_Click()
    duration = InputBox(LoadResString(121), LoadResString(122), 0)
    If StrPtr(duration) <> 0 Then
        If IsNumeric(duration) = False Then
            MsgBox LoadResString(118), vbExclamation + vbOKOnly
        ElseIf Val(duration) > 0 Then
            MsgBox LoadResString(119) & duration & LoadResString(120), vbInformation + vbOKOnly
            n = Val(duration)
        Else
            MsgBox LoadResString(118), vbExclamation + vbOKOnly
        End If
    End If
End Sub

Private Sub mnuForce_Click()
    If mnuForce.Checked = False Then
        Dim response_before As Integer
        response_before = MsgBox(LoadResString(123) & Chr(13) & LoadResString(124), vbExclamation + vbYesNo)
        If response_before = vbYes Then
            mnuTheme.Enabled = False
            Fun_DisplayTaskBar False
            ' Install hook
            lHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallKeyHookProc, App.hInstance, 0)
            
            WindowState = vbMaximized
            mnuForce.Checked = True
            resize_state = 1
        End If
    Else
        Dim response_after As Integer
        response_after = MsgBox(LoadResString(125), vbExclamation + vbYesNo)
        If response_after = vbYes Then
            mnuTheme.Enabled = True
            Fun_DisplayTaskBar True
            ' Uninstall hook
            Call UnhookWindowsHookEx(lHook)
            
            resize_state = 0
            WindowState = vbNormal
            mnuForce.Checked = False
        End If
    End If
End Sub

Private Sub mnuOpen_Click()
    CommonDialog1.Filter = "Text Files(*.txt)|*.txt|RTF Files(*.rtf)|*.rtf|All Files(*.*)|*.*"
    CommonDialog1.ShowOpen
    Quicknotes.Text = ""
    FileName = CommonDialog1.FileName
    Quicknotes.LoadFile FileName
End Sub

Private Sub mnuRemain_click()
    MsgBox n & LoadResString(126), vbInformation + vbOKOnly
End Sub

Private Sub mnuSave_Click()
    CommonDialog1.Filter = "Text Files(*.txt)|*.txt|RTF Files(*.rtf)|*.rtf|All Files(*.*)|*.*"
    CommonDialog1.ShowSave
    FileType = CommonDialog1.FileTitle
    FiType = LCase(Right(FileType, 3))
    FileName = CommonDialog1.FileName
    Select Case FiType
    Case "txt"
        Quicknotes.SaveFile FileName, rtfText
    Case "rtf"
        Quicknotes.SaveFile FileName, rtfRTF
    Case "*.*"
        Quicknotes.SaveFile FileName
    End Select
End Sub

Private Sub mnuStart_Click()
    If Val(duration) = 0 Then
        MsgBox LoadResString(127), vbExclamation + vbOKOnly
    Else
        mnuTheme.Enabled = False
        mnuStart.Enabled = False
        mnuStartForce.Enabled = False
        Timer2.Enabled = True
        mnuRemain.Enabled = True
        Label1.Caption = LoadResString(128)
    End If
End Sub

Private Sub mnuStartForce_Click()
    If Val(duration) = 0 Then
        MsgBox LoadResString(127), vbExclamation + vbOKOnly
    Else
        Dim response_before As Integer
        response_before = MsgBox(LoadResString(129) & Chr(13) & LoadResString(124), vbExclamation + vbYesNo)
        If response_before = vbYes Then
            mnuTheme.Enabled = False
            Fun_DisplayTaskBar False
            ' Install hook
            lHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallKeyHookProc, App.hInstance, 0)
            
            WindowState = vbMaximized
            mnuForce.Checked = True
            mnuForce.Enabled = False
            resize_state = 1
            
            mnuStart.Enabled = False
            mnuStartForce.Enabled = False
            Timer2.Enabled = True
            mnuRemain.Enabled = True
            Label1.Caption = LoadResString(128)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = Format(Now, "yyyy-MM-dd hh:mm")
End Sub

Private Sub Timer2_Timer()
    n = n - 1
    If n = 0 Then
        MsgBox LoadResString(130), vbInformation + vbOKOnly
        mnuRemain.Enabled = False
        mnuStart.Enabled = True
        mnuStartForce.Enabled = True
        n = Val(duration)
        Label1.Caption = LoadResString(110)
        If resize_state = 1 Then
            Fun_DisplayTaskBar True
            ' Uninstall hook
            Call UnhookWindowsHookEx(lHook)
            
            resize_state = 0
            WindowState = vbNormal
            mnuForce.Checked = False
            mnuForce.Enabled = True
        End If
        mnuTheme.Enabled = True
        Timer2.Enabled = False
    End If
End Sub
