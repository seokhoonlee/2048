Attribute VB_Name = "Module2"
Public Function GenerateRandomTwo()
    Dim isEmpty As Boolean
    isEmpty = False
    
    Do While isEmpty = False
        isEmpty = True
    
        ' Generate random number
        Dim randRowNum As Integer
        Dim randColNum As Integer
        randRowNum = Int(Rnd() * 4) + 4
        randColNum = Int(Rnd() * 4) + 2
        
        ' Get cell address from random number
        Dim randomAddress
        randomAddress = Cells(randRowNum, randColNum).Address
        
        If Range(randomAddress) <> "" Then
            isEmpty = False
            
            If WorksheetFunction.CountBlank(Range("B4:E7")) = 0 Then
                Exit Do
            End If
        Else
             ' Fill random cell with value 2
            Range(randomAddress) = 2
        End If
    Loop
    
    If isEmpty = False Then
        GenerateRandomTwo = False
    Else
        GenerateRandomTwo = True
    End If
End Function
Public Function CalculateScore()
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    Dim i As Integer
    Dim j As Integer
    Dim score As Integer
    
    score = 0
    
    For i = 1 To 4
        For j = 1 To 4
            If matrix4x4(i, j) <> "" Then
                score = score + (matrix4x4(i, j) * GetLogOfTwo(CInt(matrix4x4(i, j))))
            End If
        Next j
    Next i
    
    Range("I2") = score

End Function
Public Function GetLogOfTwo(score As Integer)
    GetLogOfTwo = Log(score) / Log(2)
End Function
Public Function CheckGameOver()
    CheckGameOver = False

    If WorksheetFunction.CountBlank(Range("B4:E7")) = 0 Then
        If ShiftUp() = False And MergeUp() = False And ShiftDown = False And MergeDown = False And ShiftLeft = False And MergeLeft = False And ShiftRight = False And MergeRight = False Then
            CheckGameOver = True
        End If
    End If
End Function
Public Function HandleGameOver()
    Dim playerRemark As String
    playerRemark = InputBox("Game over! Share remarks, if any.")
    
    Dim playerDivision As String
    Dim playerRank As String
    Dim playerName As String
    Dim playerScore As String
    
    playerDivision = Cells(2, 3)
    playerRank = Cells(2, 5)
    playerName = Cells(2, 7)
    playerScore = Cells(2, 9)
    
    Dim playerDate
    Dim playerTime

    playerDate = Date
    playerTime = Time()
    
    Dim highscore As Variant
    highscore = Range("C11:I1010")
    
    Dim i As Integer
    Dim isRecorded As Boolean
    
    isRecored = False
    
    For i = 1 To 1000
        If highscore(i, 1) = "" Then
            highscore(i, 1) = playerDivision
            highscore(i, 2) = playerRank
            highscore(i, 3) = playerName
            highscore(i, 4) = playerDate
            highscore(i, 5) = playerTime
            highscore(i, 6) = playerScore
            highscore(i, 7) = playerRemark
            
            isRecorded = True
        End If
        
        If isRecorded Then
            Exit For
        End If
    Next i
    
    Dim scoreRange As Range
    Set scoreRange = Range("C11:I1010")
    scoreRange.Value = highscore
    
    scoreRange.Sort key1:=Range("H11:H1010"), order1:=xlDescending, Header:=xlNo
End Function
Sub ClearGameCell()
    ' Remove score
    Range("I2:K2").ClearContents
    
    ' Remove contents from 4x4
    Range("B4:E7").ClearContents
    
    ' Remove color from 4x4
    With Range("B4:E7").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Format starting cell
    With Range("B4:E7")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        
    ' Format style of text in starting cell
    With Range("B4:E7").Font
        .name = "Calibri"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Sub
Public Function ShiftUp()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isShifted As Boolean
    
    ShiftUp = False
    
    'Shift numbers if empty towards direction of the key
    For i = 1 To 4
        isShifted = True
        While isShifted = True
            isShifted = False
            For j = 1 To 3
                If matrix4x4(j, i) = "" And matrix4x4(j + 1, i) <> "" Then
                    matrix4x4(j, i) = matrix4x4(j + 1, i)
                    matrix4x4(j + 1, i) = ""
                    isShifted = True
                    ShiftUp = True
                End If
            Next j
        Wend
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function MergeUp()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isMerged As Boolean
    
    MergeUp = False
    
    'Merge numbers if same and not empty towards direction of the key
    For i = 1 To 4
        For j = 1 To 3
            If matrix4x4(j, i) <> "" And matrix4x4(j, i) = matrix4x4(j + 1, i) Then
                matrix4x4(j, i) = "" & CStr(CInt(matrix4x4(j, i)) * 2)
                matrix4x4(j + 1, i) = ""
                isMerged = True
                MergeUp = True
            End If
        Next j
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function ShiftDown()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isShifted As Boolean
    
    ShiftDown = False
    
    'Shift numbers if empty towards direction of the key
    For i = 1 To 4
        isShifted = True
        While isShifted = True
            isShifted = False
            For j = 1 To 3
                If matrix4x4(5 - j, i) = "" And matrix4x4(5 - j - 1, i) <> "" Then
                    matrix4x4(5 - j, i) = matrix4x4(5 - j - 1, i)
                    matrix4x4(5 - j - 1, i) = ""
                    isShifted = True
                    ShiftDown = True
                End If
            Next j
        Wend
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function MergeDown()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isMerged As Boolean
    
    MergeDown = False
    
    'Merge numbers if same and not empty towards direction of the key
    For i = 1 To 4
        For j = 1 To 3
            If matrix4x4(5 - j, i) <> "" And matrix4x4(5 - j, i) = matrix4x4(5 - j - 1, i) Then
                matrix4x4(5 - j, i) = "" & CStr(CInt(matrix4x4(5 - j, i)) * 2)
                matrix4x4(5 - j - 1, i) = ""
                isMerged = True
                MergeDown = True
            End If
        Next j
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function ShiftLeft()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isShifted As Boolean
    
    ShiftLeft = False
    
    'Shift numbers if empty towards direction of the key
    For i = 1 To 4
        isShifted = True
        While isShifted = True
            isShifted = False
            For j = 1 To 3
                If matrix4x4(i, j) = "" And matrix4x4(i, j + 1) <> "" Then
                    matrix4x4(i, j) = matrix4x4(i, j + 1)
                    matrix4x4(i, j + 1) = ""
                    isShifted = True
                    ShiftLeft = True
                End If
            Next j
        Wend
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function MergeLeft()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isMerged As Boolean
    
    MergeLeft = False
    
    'Merge numbers if same and not empty towards direction of the key
    For i = 1 To 4
        For j = 1 To 3
            If matrix4x4(i, j) <> "" And matrix4x4(i, j) = matrix4x4(i, j + 1) Then
                matrix4x4(i, j) = "" & CStr(CInt(matrix4x4(i, j)) * 2)
                matrix4x4(i, j + 1) = ""
                isMerged = True
                MergeLeft = True
            End If
        Next j
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function ShiftRight()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isShifted As Boolean
    
    ShiftRight = False
    
    'Shift numbers if empty towards direction of the key
    For i = 1 To 4
        isShifted = True
        While isShifted = True
            isShifted = False
            For j = 1 To 3
                If matrix4x4(i, 5 - j) = "" And matrix4x4(i, 5 - j - 1) <> "" Then
                    matrix4x4(i, 5 - j) = matrix4x4(i, 5 - j - 1)
                    matrix4x4(i, 5 - j - 1) = ""
                    isShifted = True
                    ShiftRight = True
                End If
            Next j
        Wend
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Public Function MergeRight()
    'Declare array for computation with values from 4x4 cells
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    'Declare necessary variables for shifting logic
    Dim i As Integer
    Dim j As Integer
    Dim isMerged As Boolean
    
    MergeRight = False
    
    'Merge numbers if same and not empty towards direction of the key
    For i = 1 To 4
        For j = 1 To 3
            If matrix4x4(i, 5 - j) <> "" And matrix4x4(i, 5 - j) = matrix4x4(i, 5 - j - 1) Then
                matrix4x4(i, 5 - j) = "" & CStr(CInt(matrix4x4(i, 5 - j)) * 2)
                matrix4x4(i, 5 - j - 1) = ""
                isMerged = True
                MergeRight = True
            End If
        Next j
    Next i
    
    'Copy back array to cells
    Dim gameRange As Range
    Set gameRange = Range("B4:E7")
    gameRange.Value = matrix4x4
End Function
Sub ColorTiles()
    ' Remove color from 4x4
    With Range("B4:E7").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Color according to content which is number
    Dim matrix4x4() As Variant
    matrix4x4 = Range("B4:E7")
    
    Dim i As Integer
    Dim j As Integer
    
    Dim colorAddress
    
    For i = 1 To 4
        For j = 1 To 4
            If matrix4x4(i, j) <> "" Then
                colorAddress = Cells(i + 3, j + 1).Address
                
                ' Color starting cell
                With Range(colorAddress).Interior
                    .Color = 39423
                    .TintAndShade = 1 - (GetLogOfTwo(CInt(matrix4x4(i, j)) - 1) * 0.05)
                End With
            End If
        Next j
    Next i
End Sub

