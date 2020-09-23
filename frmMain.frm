VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEL Script Executor"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   5310
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":0442
      Top             =   630
      Width           =   6450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   15
      X2              =   6615
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   6615
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JEL - Justin's Elite Language"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   675
      TabIndex        =   1
      Top             =   180
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmMain.frx":0594
      Top             =   90
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMake 
         Caption         =   "Make EXE..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
      Begin VB.Menu mnuRunRun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpCommands 
         Caption         =   "Quick Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' JEL Script v1.0 : frmMain.frm
' (c) Copyright 2000 By www.oogle.net
'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileMake_Click()
    On Error GoTo ErrorCatch:
    
    If Not FileExists(App.Path & "\jelexe.exe") Then
        MsgBox "Program could not find the second JEL executable, aborting.", vbCritical, "Error"
        Exit Sub
    End If
    
    cd.FileName = ""
    cd.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
    cd.ShowSave
    If cd.FileName <> "" Then
        If FileExists(cd.FileName) Then
            If MsgBox("Overwrite existing file?", vbQuestion + vbYesNo, "JEl") = vbNo Then
                Exit Sub
            Else
                Kill cd.FileName
            End If
        End If
        
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.Path & "\jelexe.exe", cd.FileName
        Open cd.FileName For Output As #nFile
        Print #nFile, "|*JEL*|" & txtScript.Text
        Close #nFile
        MsgBox "File Compiled!", vbInformation, "JEL"
        
        Dim sTemp As String, sTemp2 As String
        
        Open cd.FileName For Output As #1
        Open App.Path & "\jelexe.exe" For Binary As #2
        
        ' Copy data from jelexe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*JEL*|" & txtScript.Text
        
        Close #2
        Close #1
        
        
    End If
    
    Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & Err.Description, vbCritical, "Error"
    Resume Next
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpCommands_Click()
    frmHelp.Show
End Sub

Private Sub mnuOpen_Click()
    cd.FileName = ""
    cd.Filter = "JEL Source Files (*.JEL)|*.jel|All Files (*.*)|*.*"
    cd.ShowOpen
    If cd.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open cd.FileName For Input As nFile
        txtScript.Text = Input(LOF(nFile), nFile)
        Close nFile
    End If
End Sub

Private Sub mnuRunRun_Click()
    Dim myScript As New CJel
    myScript.Script = txtScript.Text
    myScript.ScriptExecute
End Sub

Private Sub mnuSave_Click()
    cd.FileName = ""
    cd.Filter = "JEL Source Files (*.JEL)|*.jel|All Files (*.*)|*.*"
    cd.ShowSave
    If cd.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open cd.FileName For Output As nFile
        Print #nFile, txtScript.Text
        Close nFile
    End If
End Sub

Private Sub txtScript_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
        KeyAscii = Asc(" ")
    End If
End Sub

Private Function FileExists(sFilename As String) As Boolean
    If Len(sFilename) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir(sFilename)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
