Attribute VB_Name = "modMain"
Sub Main()
    Dim FileSize As Long, iTemp As Long
    Dim FileData As String, sTemp As String
    
    Open App.Path + "\" + App.EXEName + ".EXE" For Binary As #1
    
    FileSize = LOF(1)
    FileData = Space$(LOF(1))
    
    Get #1, , FileData

    iTemp = InStr(1, FileData, "|*JEL*|")
    If iTemp <> 0 Then
        iTemp = iTemp + 7
        sTemp = String(1000, 0)
        Get #1, iTemp, sTemp
        
        Dim myScript As New CJel
        myScript.Script = sTemp
    End If
    
    Close #1
    
    myScript.ScriptExecute
End Sub
