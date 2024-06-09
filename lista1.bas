Attribute VB_Name = "Module1"
'1a, 1b
Sub FibonacciNumbersModified()
    Dim i As Integer
    Dim a As Integer
    Dim b As Integer
    Dim Limit As Long
    
    Columns(1).ClearContents
    Columns(2).ClearContents
    
    Limit = InputBox(" Podaj n")
    a = InputBox(" Podaj a")
    b = InputBox(" Podaj b")
    
    i = 1
    Cells(i, 1).Value = a
    Cells(i, 2).Value = a
    i = i + 1
    Cells(i, 1).Value = b
    Cells(i, 2).Value = b ^ (1 / 2)
    
    Do While i < Limit
        i = i + 1
        Cells(i, 1).Value = Cells(i - 1, 1).Value + Cells(i - 2, 1).Value
        
        'limit
        Cells(i, 2).Value = (Cells(i, 1).Value) ^ (1 / i)
    Loop
End Sub
'2a
Sub ArmstrongModified()

    Dim kNum As Byte, Num As Long
    Dim Limit As Long
    Dim Digit As Byte, i As Byte
    Dim DigitPower As Long, Sum As Long
    Dim Counter As Byte
    Dim NumStr As String, Msg As String
    Dim m As Integer
    
    kNum = InputBox(" Wpisz parametr n")
    m = InputBox(" Wpisz parametr m")
    
    Limit = (10 ^ kNum) - 1
    
    For Num = 1 To Limit
        NumStr = CStr(Num)
        
        For i = 1 To Len(NumStr)
            Digit = Mid(NumStr, i, 1)
            DigitPower = Digit ^ m
            Sum = Sum + DigitPower
        Next i
        
        If Num = Sum And Len(CStr(Num)) = kNum Then
            Msg = Msg & Num & ", "
            Counter = Counter + 1
        End If
        
        Sum = 0
        
    Next Num
    
    Select Case Counter
    
        Case 0
            MsgBox " Nie ma " & kNum & "-cyfrowych takich liczby ."
        Case 1
            Msg = Left(Msg, Len(Msg) - 2)
            MsgBox " Jest jedna " & kNum _
            & "-cyfrowa taka liczba : " & Msg & "."
        Case 2, 3, 4
            Msg = Left(Msg, Len(Msg) - 2)
            MsgBox "S a " & Counter & " " & kNum _
            & "-cyfrowe ltakie liczby : " & Msg & "."
        Case Is > 4
            Msg = Left(Msg, Len(Msg) - 2)
            MsgBox " Jest " & Counter & " " & kNum _
            & "-cyfrowych takich liczb : " & Msg & "."
            
    End Select
    
End Sub

'2b
Sub Munchhausen()

    Dim kNum As Long, Num As Long
    Dim Limit As Long
    Dim Digit As Long, i As Byte
    Dim DigitPower As Long, Sum As Long
    Dim Counter As Byte
    Dim NumStr As String, Msg As String
    Dim CzyZera As Integer
    
    kNum = InputBox(" Wpisz maksymalna ilosc cyfr: ")
    CzyZera = MsgBox("Czy 0^0 = 0?", vbYesNo)
    Limit = (10 ^ kNum) - 1
    
    For Num = 0 To Limit
        NumStr = CStr(Num)
    For i = 1 To Len(NumStr)
        Digit = Mid(NumStr, i, 1)
        If Digit > 0 Then
        DigitPower = Digit ^ Digit
        Sum = Sum + DigitPower
        Else
            Sum = Sum + 1
        End If
        Next i
    If Num = Sum And Len(CStr(Num)) = kNum Then
        Msg = Msg & Num & ", "
        Counter = Counter + 1
    End If
    Sum = 0
    Next Num
    
    Select Case Counter
        Case 0
        MsgBox " Nie ma " & kNum & "-cyfrowych takich liczb."
    
    Case Is > 0
        Msg = Left(Msg, Len(Msg) - 2)
        MsgBox " Jest " & Counter & " " & kNum _
        & "-cyfrowych takich liczb : " & Msg & "."
    
    End Select

End Sub



