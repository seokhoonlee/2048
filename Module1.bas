Attribute VB_Name = "Module1"
Sub StartButton()
Attribute StartButton.VB_ProcData.VB_Invoke_Func = "S\n14"
    Sheets(1).Protect Password:="secret", _
    UserInterFaceOnly:=True
    
    Call Module2.ClearGameCell
    Call Module2.GenerateRandomTwo
    Call Module2.ColorTiles
    Call Module2.CalculateScore
End Sub
Sub UpKey()
Attribute UpKey.VB_ProcData.VB_Invoke_Func = "I\n14"
    Dim isShifted
    Dim isMerged
    
    isShifted = Module2.ShiftUp
    isMerged = Module2.MergeUp
    Call Module2.ShiftUp
    
    If isShifted Or isMerged Then
        Call Module2.GenerateRandomTwo
        Call Module2.ColorTiles
        Call Module2.CalculateScore
        isGameover = Module2.CheckGameOver
    End If
    
    If isGameover Then
        Call Module2.HandleGameOver
    End If
End Sub
Sub DownKey()
Attribute DownKey.VB_ProcData.VB_Invoke_Func = "K\n14"
    Dim isShifted
    Dim isMerged

    isShifted = Module2.ShiftDown
    isMerged = Module2.MergeDown
    Call Module2.ShiftDown
    
    If isShifted Or isMerged Then
        Call Module2.GenerateRandomTwo
        Call Module2.ColorTiles
        Call Module2.CalculateScore
        isGameover = Module2.CheckGameOver
    End If
    
    If isGameover Then
        Call Module2.HandleGameOver
    End If
End Sub
Sub LeftKey()
Attribute LeftKey.VB_ProcData.VB_Invoke_Func = "J\n14"
    Dim isShifted
    Dim isMerged
    
    isShifted = Module2.ShiftLeft
    isMerged = Module2.MergeLeft
    Call Module2.ShiftLeft
    
    If isShifted Or isMerged Then
        Call Module2.GenerateRandomTwo
        Call Module2.ColorTiles
        Call Module2.CalculateScore
        isGameover = Module2.CheckGameOver
    End If
    
    If isGameover Then
        Call Module2.HandleGameOver
    End If
End Sub
Sub RightKey()
Attribute RightKey.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim isShifted
    Dim isMerged
    
    isShifted = Module2.ShiftRight
    isMerged = Module2.MergeRight
    Call Module2.ShiftRight
    
    If isShifted Or isMerged Then
        Call Module2.GenerateRandomTwo
        Call Module2.ColorTiles
        Call Module2.CalculateScore
        isGameover = Module2.CheckGameOver
    End If
    
    If isGameover Then
        Call Module2.HandleGameOver
    End If
End Sub
