Attribute VB_Name = "Module1"
Function Pierwsza(Nn As Long) As Boolean

    Dim limit As Long, i As Long
    
    If Nn = 0 Or Nn = 1 Then Exit Function
    'fix for 2 and 3
    If Nn = 2 Or Nn = 3 Then Pierwsza = True
    
    limit = Application.WorksheetFunction.RoundUp(Sqr(Nn), 0)
    
    For i = 2 To limit
        If Nn Mod i = 0 Then Exit Function
        Next i
        Pierwsza = True
    
End Function
Sub Czas()

    Dim Start As Double
    Dim Finish As Double
    Start = Timer
    Dim n As Long
    n = 6
    [A1] = " Czy liczba " & n & " jest pierwsza ?"
    [A2] = Pierwsza(n)
    Finish = Round(Timer - Start, 2)
    [A3] = " Czas noblicze " & Finish & " sekund "
    
End Sub
'1
Sub FirstPrimeTimed()

    Dim n As Long
    Dim i As Long
    Dim Start As Double
    Dim Finish As Double
    
    
    n = InputBox("Wpisz liczbê", "Znajdowanie najmniejszej liczby pierwszej wiêkszej lub równej podanej liczbie")
    i = n
    
    Start = Timer
    
    While Pierwsza(n) = False
        n = n + 1
    Wend
        
    Finish = Round(Timer - Start, 2)
    
    MsgBox ("Najmniejsza liczba pierwsza wiêksza lub równa " & i & " to " & n & ". Czas wykonania wyniós³ " & Finish & " sekund.")

End Sub
'2
Function isSemiPrime(n As Long) As Boolean
    Dim i As Long
    If n <= 1 Then
        isSemiPrime = False
        Exit Function
    End If
    
    For i = 2 To Int(Sqr(n)) + 1
        If n Mod i = 0 Then
            If Pierwsza(i) And Pierwsza(n \ i) Then
                isSemiPrime = True
                Exit Function
            End If
        End If
    Next i
    
    isSemiPrime = False
End Function

Sub CheckIfSemiPrime()

    Dim userInput As String
    Dim number As Long
    Dim isSemiPrimeResult As Boolean
    
    userInput = InputBox("WprowadŸ liczbê:", "Czy liczba jest pó³pierwsza")
    
    If IsNumeric(userInput) Then
        number = CLng(userInput)
        isSemiPrimeResult = isSemiPrime(number)
        
        If isSemiPrimeResult Then
            MsgBox number & " jest pó³pierwsza.", vbInformation
        Else
            MsgBox number & " nie jest pó³pierwsza.", vbInformation
        End If
    Else
        MsgBox "Nieprawid³owe dane wejœciowe. Proszê wprowadziæ prawid³ow¹ liczbê.", vbCritical
    End If
End Sub
'3
Sub NajpopularniejszeImiona()

    Dim userInput As String
    Dim yearInput As Long
    Dim filePath As String
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim lastRow As Long
    Dim maleMaxCount As Long
    Dim femaleMaxCount As Long
    Dim maleName As String
    Dim femaleName As String

    userInput = InputBox("WprowadŸ rok od 2000 do 2019:", "Year Input")
    
    If IsNumeric(userInput) Then
        yearInput = CLng(userInput)
        If yearInput < 2000 Or yearInput > 2019 Then
            MsgBox "WprowadŸ rok od 2000 do 2019!", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "WprowadŸ rok od 2000 do 2019!", vbExclamation
        Exit Sub
    End If

    filePath = ThisWorkbook.Path & "\Imiona_nadane_w_Polsce_w_latach_2000_2019.xlsx"

    If Dir(filePath) = "" Then
        MsgBox "Nie znaleziono pliku!", vbCritical
        Exit Sub
    End If

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets(1)

    maleMaxCount = 0
    femaleMaxCount = 0
    maleName = ""
    femaleName = ""

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = yearInput Then
            If ws.Cells(i, 4).Value = "M" Then
                If ws.Cells(i, 3).Value > maleMaxCount Then
                    maleMaxCount = ws.Cells(i, 3).Value
                    maleName = ws.Cells(i, 2).Value
                End If
            ElseIf ws.Cells(i, 4).Value = "K" Then
                If ws.Cells(i, 3).Value > femaleMaxCount Then
                    femaleMaxCount = ws.Cells(i, 3).Value
                    femaleName = ws.Cells(i, 2).Value
                End If
            End If
        End If
    Next i

    wb.Close SaveChanges:=False

    MsgBox "Najpopularniejsze imiona w " & yearInput & " roku to:" & vbCrLf & _
           "Mêskie: " & maleName & vbCrLf & _
           "¯eñskie: " & femaleName, vbInformation
End Sub














