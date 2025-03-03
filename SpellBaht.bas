Attribute VB_Name = "SpellBath"
Option Explicit

' Main Function
Function SpellBaht(ByVal MyNumber)
    Dim Baht As String, Satang As String, Temp As String, Satangs As String
    Dim DecimalPlace As Integer, Count As Integer
    ReDim Place(9) As String
    
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

    ' Convert number to string and trim spaces
    MyNumber = Trim(Str(MyNumber))

    ' Find decimal place position (0 if none)
    DecimalPlace = InStr(MyNumber, ".")

    ' Convert Satangs and set MyNumber to Baht amount
    If DecimalPlace > 0 Then
        Satangs = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    Else
        Satangs = "Zero"
    End If

    ' Convert Baht part
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Baht = Temp & Place(Count) & Baht
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop

    ' Assign correct Baht wording
    Select Case Baht
        Case ""
            Baht = "Zero Baht"
        Case "One"
            Baht = "One Baht"
        Case Else
            Baht = Baht & " Baht"
    End Select

    ' Assign correct Satang wording
    Select Case Satangs
        Case "", "Zero"
            Satang = " and Zero Satang"
        Case "One"
            Satang = " and One Satang"
        Case Else
            Satang = " and " & Satangs & " Satang"
    End Select

    ' Return final result
    SpellBaht = Baht & Satang
End Function

' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)

    ' Convert the hundreds place
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If

    ' Convert the tens and ones place
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If

    GetHundreds = Result
End Function

' Converts a number from 10 to 99 into text
Function GetTens(TensText)
    Dim Result As String
    Result = ""

    ' If value between 10-19
    If Val(Left(TensText, 1)) = 1 Then
        Select Case Val(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
        End Select
    Else
        ' If value between 20-99
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
        End Select
        Result = Result & GetDigit(Right(TensText, 1)) ' Retrieve ones place
    End If

    GetTens = Result
End Function

' Converts a number from 1 to 9 into text
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

