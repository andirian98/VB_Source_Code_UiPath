 Dim SmallNumberToWord As Func(Of Integer, String, String) = Function(ByVal number As Integer, ByVal words As String)As String
 Dim NumberToWords As Func(Of Double, String) = Function(ByVal doubleNumber As Double)As String
 Dim NumberToWordsInt As Func(Of Integer, String) = Function(ByVal number As Integer)As String
 
 Dim NumberToWords As Func(Of Double, String) = Function(ByVal doubleNumber As Double)As String
    Dim beforeFloatingPoint =CInt(Math.Floor(doubleNumber))
    Dim beforeFloatingPointWord =String.Format("{0} rupees", NumberToWordsInt(beforeFloatingPoint))
    Dim afterFloatingPointWord =String.Format("{0} paisa only.", SmallNumberToWord(CInt(((doubleNumber - beforeFloatingPoint) * 100)),""))
 
    If CInt(((doubleNumber - beforeFloatingPoint) * 100)) > 0Then
        Return String.Format("{0} and {1}", beforeFloatingPointWord, afterFloatingPointWord)
    Else
        Return String.Format("{0} only", beforeFloatingPointWord)
    End If
End Function

Dim NumberToWordsInt As Func(Of Integer, String) = Function(ByVal number As Integer)As String
    If number = 0Then Return "zero"
    If number< 0Then Return "minus " & NumberToWords(Math.Abs(number))
    Dim words =""
 
    If number / 1000000000 > 0Then
        words += NumberToWords(number / 1000000000) &" billion "
        number = number Mod 1000000000
    End If
 
    If number / 1000000 > 0Then
        words += NumberToWords(number / 1000000) &" million "
        number = number Mod 1000000
    End If
 
    If number / 1000 > 0Then
        words += NumberToWords(number / 1000) &" thousand "
        number = number Mod 1000
    End If
 
    If number / 100 > 0Then
        words += NumberToWords(number / 100) &" hundred "
        number = number Mod 100
    End If
 
    words = SmallNumberToWord(number, words)
    Return words
End Function
 
Dim SmallNumberToWord As Func(Of Integer, String, String) = Function(ByVal number As Integer, ByVal words As String)As String
    If number <= 0Then Return words
    If words<>"" Then words +=" "
    Dim unitsMap = {"zero","one","two","three","four","five","six","seven","eight","nine","ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen" }
    Dim tensMap = {"zero","ten","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety" }
 
    If number< 20Then
        words += unitsMap(number)
    Else
        words += tensMap(number / 10)
        If(number Mod 10) > 0Then words +=" " & unitsMap(number Mod 10)
    End If
 
    Return words
End Function

result = NumberToWords(input_num)
final_result = SmallNumberToWord(input_num,result)