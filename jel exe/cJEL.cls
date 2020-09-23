VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' JEL Main Executing Module (with class)
' (c) Copyright 2000 By www.oogle.net
'

Option Explicit

Public Script As String
Private VarNames As New Collection
Private Variables As New Collection
Private ScriptForm As New frmProgram

Public Event LineChange(nLine As Integer, sContents As String)

Public Sub ScriptExecute()
    Dim sCurrentLine As String, nCurrentLine As Integer
    Dim lCrLfSpot As Long, lCrLfSpotOld As Long
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String, sTemp2 As String
    Dim arTemp() As String, arTemp2() As String
    Dim StopLabel As String
    Dim Skip As Boolean
    Dim IfSkip As Boolean
    
    Dim InsideLoop As New Collection
    
    arTemp = Split(Script, vbCrLf)
    
    For nCurrentLine = LBound(arTemp) To UBound(arTemp)
        sCurrentLine = arTemp(nCurrentLine)
        If sCurrentLine = "" Then GoTo SkipCurrentDo:
        RaiseEvent LineChange(nCurrentLine, sCurrentLine)
        
        ' Get rid of comments
        iTemp = InStr(1, sCurrentLine, "#")
        If iTemp <> 0 Then
            sCurrentLine = Mid(sCurrentLine, 1, iTemp - 1)
            If sCurrentLine = "" Then GoTo SkipCurrentDo:
        End If
        
        ' Check for a label
        If StopLabel <> "" Then
            If sCurrentLine = StopLabel Then
                StopLabel = ""
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Check for end of a block
        If Skip = True Then
            If LCase(sCurrentLine) = "end" Then
                Skip = False
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Skipping if stuff...
        If IfSkip = True Then
            Select Case LCase(sCurrentLine)
                Case "end"
                    IfSkip = False
                Case "else"
                    IfSkip = False
            End Select
            GoTo SkipCurrentDo:
        Else
            If LCase(sCurrentLine) = "else" Then
                Skip = True
                GoTo SkipCurrentDo:
            End If
        End If
        
        ' Looping block
        If LCase(sCurrentLine) = "end" And InsideLoop.Count > 0 Then
            If Eval(Mid(InsideLoop(InsideLoop.Count), InStr(1, InsideLoop(InsideLoop.Count), ":") + 1)) Then
                nCurrentLine = CInt(Mid(InsideLoop(InsideLoop.Count), 1, InStr(1, InsideLoop(InsideLoop.Count), ":") - 1)) - 1
            Else
                InsideLoop.Remove InsideLoop.Count
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Chck if its one of those crazy things
        iTemp = InStr(1, sCurrentLine, " ")
        If iTemp <> 0 Then
            Select Case LCase(Mid(sCurrentLine, 1, iTemp - 1))
                Case "if" ' Need I explain?
                    If Not Eval(Mid(sCurrentLine, iTemp + 1)) Then
                        IfSkip = True
                    End If
                    
                    GoTo SkipCurrentDo:
                Case "goto" ' Goto a label
                    StopLabel = Trim(Mid(sCurrentLine, iTemp + 1))
                    
                    GoTo SkipCurrentDo:
                Case "var" ' Create variable
                    arTemp2 = Split(Trim(Mid(sCurrentLine, iTemp + 1)), ",")
                    For i = LBound(arTemp2) To UBound(arTemp2)
                        VarNames.Add Trim(arTemp2(i))
                        Variables.Add ""
                    Next i
                    GoTo SkipCurrentDo:
                Case "while" ' Loop
                    If Eval(Mid(sCurrentLine, iTemp + 1)) Then
                        InsideLoop.Add nCurrentLine + 1 & ":" & Trim(Mid(sCurrentLine, iTemp + 1))
                    Else
                        Skip = True
                    End If
                    GoTo SkipCurrentDo:
            End Select
        End If
        
        For i = 1 To Len(sCurrentLine)
            Select Case Mid(sCurrentLine, i, 1)
                Case "=" ' Assignment
                    ChangeVar Trim(Mid(sCurrentLine, 1, InStr(1, sCurrentLine, "=") - 1)), Equation(Mid(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                    GoTo SkipCurrentDo:
                Case "~" ' Assignment from Function
                    ChangeVar Trim(Mid(sCurrentLine, 1, InStr(1, sCurrentLine, "~") - 1)), DoFunction(Mid(sCurrentLine, InStr(1, sCurrentLine, "~") + 1))
                    GoTo SkipCurrentDo:
                Case "(" ' Procedure Call
                    DoFunction sCurrentLine
                    GoTo SkipCurrentDo:
            End Select
        Next i
        
SkipCurrentDo:
    Next nCurrentLine
    
    ' Its the clean up crew!
    For i = VarNames.Count To 1 Step -1
        VarNames.Remove i
        Variables.Remove i
    Next i
End Sub
    
Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    Dim i As Integer
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            Variables.Remove i
            VarNames.Remove i
            
            Variables.Add NewVal
            VarNames.Add TheVar
            Exit Sub
        End If
    Next i
End Sub

Private Function GetVar(TheVar As String) As Variant
    Dim i As Integer
    For i = 1 To VarNames.Count
        If VarNames(i) = TheVar Then
            GetVar = Variables(i)
            Exit Function
        End If
    Next i
End Function

Private Function Eval(ByVal sFormula As String) As Boolean
    Dim i As Integer, iWait As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    
    ' Get the left value
    For i = 1 To Len(sFormula)
        sTemp = Mid(sFormula, i, 1)
        Select Case sTemp
            Case "("
                iWait = iWait + 1
            Case ")"
                iWait = iWait - 1
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Chr(34)
                i = InStr(i + 1, sFormula, Chr(34))
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Else
                If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
                    If iWait = 0 Then
                        LeftVal = Trim(Mid(sFormula, 1, i - 1))
                        sFormula = Trim(Mid(sFormula, i))
                        Exit For
                    End If
                End If
        End Select
    Next i
    
    ' Get the comparison operator
    sTemp = Mid(sFormula, 2, 1)
    If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
        Operator = Left(sFormula, 2)
        sFormula = Mid(sFormula, 3)
    Else
        Operator = Left(sFormula, 1)
        sFormula = Mid(sFormula, 2)
    End If
    
    ' Set the rest to the right side operator
    RightVal = sFormula
    sFormula = ""
    
    ' Solve these sides
    LeftVal = Equation(LeftVal)
    RightVal = Equation(RightVal)
    
    Select Case Operator
        Case ">"
            If Val(LeftVal) > Val(RightVal) Then Eval = True
        Case "<"
            If Val(LeftVal) < Val(RightVal) Then Eval = True
        Case "<>"
            If Val(LeftVal) <> Val(RightVal) Then Eval = True
        Case ">="
            If Val(LeftVal) >= Val(RightVal) Then Eval = True
        Case "<="
            If Val(LeftVal) <= Val(RightVal) Then Eval = True
        Case "="
            If LeftVal = RightVal Then Eval = True
        Case "=="
            If LCase(LeftVal) = LCase(RightVal) Then Eval = True
    End Select
End Function

Private Function Equation(ByVal sEquation As String) As Variant
    ' For solving full equations like: iTemp + ("3" * ("5" / "4"))
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String
    
    ' First lets change anything inside ()'s to a constant value
    Do
        For i = 1 To Len(sEquation)
            sTemp = Mid(sEquation, i, 1)
            If sTemp = Chr(34) Then
                i = InStr(i + 1, sEquation, Chr(34))
            ElseIf sTemp = "(" Then
                iTemp = i
            ElseIf sTemp = ")" Then
                iTemp2 = i
                Exit For
            End If
        Next i
        If iTemp <> 0 Then
            sEquation = Mid(sEquation, 1, iTemp - 1) & " " & Chr(34) & Solve(Mid(sEquation, iTemp + 1, iTemp2 - (iTemp + 1))) & Chr(34) & " " & Mid(sEquation, iTemp2 + 1)
            iTemp = 0: iTemp2 = 0
        Else
            Exit Do
        End If
    Loop
    
    ' Now that all the ()'s are constant values, lets finish it up
    Equation = Solve(sEquation)
End Function

Private Function Solve(sEquation As String) As Variant
    ' For solving equations without ()'s like: itemp + "3" / "2"
    Dim i As Integer, iTemp As Integer
    Dim sTemp As String, sTemp2 As String
    Dim WaitVal As Variant
    Dim WaitOp As String
    Dim WaitVar
    
    For i = 1 To Len(sEquation)
        sTemp = Mid(sEquation, i, 1)
        If sTemp = Chr(34) Then
            iTemp = InStr(i + 1, sEquation, Chr(34))
            WaitVal = Mid(sEquation, i + 1, iTemp - (i + 1))
            i = iTemp
            If WaitOp <> "" Then
                Select Case WaitOp
                    Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                    Case "-"
                        Solve = Val(Solve) - Val(WaitVal)
                    Case "/"
                        Solve = Val(Solve) / Val(WaitVal)
                    Case "\"
                        Solve = Val(Solve) \ Val(WaitVal)
                    Case "^"
                        Solve = Val(Solve) ^ Val(WaitVal)
                    Case "*"
                        Solve = Val(Solve) * Val(WaitVal)
                    Case "&" ' String Concentration
                        Solve = Solve & WaitVal
                End Select
                WaitOp = ""
            Else
                Solve = WaitVal
            End If
        ElseIf sTemp = " " Then
            ' Make sure it skips this
        ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Or sTemp = "&" Or sTemp = "*" Then
            If WaitVar <> 0 Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i - (WaitVar + 1)))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
                
                WaitVar = 0
            End If
            
            WaitOp = sTemp
        Else
            If WaitVar = 0 Then WaitVar = i
            If i >= Len(sEquation) Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                            Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
            End If
        End If
    Next i
End Function

Private Function DoFunction(daCall As String) As Variant
    Dim sName As String
    Dim argList() As String
    Dim i As Integer
    
    sName = Trim(Mid(daCall, 1, InStr(1, daCall, "(") - 1))
    argList = Split(Mid(daCall, InStr(1, daCall, "(") + 1, InStr(1, daCall, ")") - (InStr(1, daCall, "(") + 1)), ",")
    
    For i = LBound(argList) To UBound(argList)
        argList(i) = Equation(Trim(argList(i)))
    Next i
    
    DoFunction = ExecFunction(sName, argList())
End Function

Private Function ExecFunction(daFunction As String, argList() As String) As Variant
    On Error GoTo ErrorCatch:
    
    Select Case LCase(daFunction)
        Case "msgbox"
            ExecFunction = MsgBox(argList(0), argList(1), argList(2))
        Case "inputbox"
            ExecFunction = InputBox(argList(0), argList(1), argList(2))
        
        ' Window Related Crap
        Case "show"
            Load ScriptForm
            ScriptForm.Caption = argList(0)
            ScriptForm.Show
        Case "hide"
            ScriptForm.Hide
        Case "cls"
            ScriptForm.Cls
        Case "line"
            ScriptForm.Line (argList(0), argList(1))-(argList(2), argList(3))
        Case "point"
            ScriptForm.Point argList(0), argList(1)
        Case "print"
            ScriptForm.Print argList(0)
        Case "setforecolor"
            ScriptForm.ForeColor = RGB(argList(0), argList(1), argList(2))
        Case "setbackcolor"
            ScriptForm.BackColor = RGB(argList(0), argList(1), argList(2))
        Case "setcaption"
            ScriptForm.Caption = argList(0)
        Case "getcaption"
            ExecFunction = ScriptForm.Caption
        Case "setcurrentx"
            ScriptForm.CurrentX = argList(0)
        Case "getcurrentx"
            ExecFunction = ScriptForm.CurrentX
        Case "setcurrenty"
            ScriptForm.CurrentY = argList(0)
        Case "getcurrenty"
            ExecFunction = ScriptForm.CurrentY
        Case "getwidth"
            ExecFunction = ScriptForm.ScaleWidth
        Case "setwidth"
            ScriptForm.Width = argList(0)
        Case "getheight"
            ExecFunction = ScriptForm.ScaleHeight
        Case "setheight"
            ScriptForm.Height = argList(0)
        Case "textwidth"
            ExecFunction = ScriptForm.TextWidth(argList(0))
        Case "textheight"
            ExecFunction = ScriptForm.TextHeight(argList(0))
        Case "setfont"
            ScriptForm.Font = argList(0)
        Case "bold"
            ScriptForm.Font.Bold = Not ScriptForm.Font.Bold
        Case "italic"
            ScriptForm.Font.Italic = Not ScriptForm.Font.ita
        Case "underline"
            ScriptForm.Font.Underline = Not ScriptForm.Font.Underline
        Case "strike"
            ScriptForm.Font.Strikethrough = Not ScriptForm.Font.Strikethrough
        Case "picture"
            ScriptForm.Picture = LoadPicture(argList(0))
        
        ' String Manipulating Junk
        Case "mid"
            ExecFunction = Mid(argList(0), argList(1), argList(2))
        Case "lcase"
            ExecFunction = LCase(argList(0))
        Case "ucase"
            ExecFunction = UCase(argList(0))
        Case "instr"
            ExecFunction = InStr(argList(0), argList(1), argList(2))
        Case "len"
            ExecFunction = Len(argList(0))
        Case "chr"
            ExecFunction = Chr(argList(0))
        Case "asc"
            ExecFunction = Asc(argList(0))
        Case "reverse"
            ExecFunction = StrReverse(argList(0))
        
        ' Math stuff
        Case "cos"
            ExecFunction = Cos(argList(0))
        Case "tan"
            ExecFunction = Tan(argList(0))
        Case "log"
            ExecFunction = Log(argList(0))
        Case "rnd"
            ExecFunction = Rnd(argList(0))
        Case "randomize"
            Randomize
        Case "int"
            ExecFunction = Int(argList(0))
        
        ' Printer stuff
        Case "printer_print"
            Printer.Print argList(0)
        Case "finish"
            Printer.EndDoc
        
        ' Other Stuff
        Case "shell"
            Shell argList(0)
    End Select
    
    Exit Function
    
ErrorCatch:
    Select Case Err.Number
        Case 9
            MsgBox "Argument not optional in call to procedure: " & daFunction, vbCritical, "Error"
    End Select
    Resume Next
End Function
