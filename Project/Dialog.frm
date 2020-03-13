VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Theme Selection"
   ClientHeight    =   2475
   ClientLeft      =   4440
   ClientTop       =   4410
   ClientWidth     =   5415
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(136)
    OKButton.Caption = LoadResString(132)
    CancelButton.Caption = LoadResString(133)
    
    ' Set Dialog on top
    Call SetWindowPos(Dialog.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    File1.Path = App.Path & "\Theme"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Set Form1 on top
    Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub OKButton_Click()
    SkinH_AttachEx App.Path & "\Theme\" & File1.FileName, ""
    Unload Me
End Sub
