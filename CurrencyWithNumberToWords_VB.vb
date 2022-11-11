Dim Amt_Before_Decimal, Amt_After_Decimal, Country_currency, Decimal_Singular, Decimal_Plural As String
Dim  DecimalCount, DecimalCount_1, DecimalCount_10 As String
Dim Before_Amt, After_Amt As String
Dim DecimalPlace As Int32

MyNumber = Trim(Str(MyNumber)) ' String representation of amount.

DecimalPlace = InStr(MyNumber, ".") ' Position of decimal place 0 if none.

If DecimalPlace > 1  Then
Before_Amt = left(MyNumber,DecimalPlace-1)
Else If DecimalPlace = 0 And len(MyNumber) > 0 Then
Before_Amt = MyNumber
Else 
	Before_Amt = ""
End If
Num2Word = Before_Amt




'''''''''''''' VARIABLES FOR NUMBERS '''''''''''''''''''

Dim Digit, GetDigit, Tens, GetTens, Hundreds, GetHundreds As String

Dim Thousands, GetThousands, GetThousands_1, GetThousands_10, GetThousands_100 As String

Dim Millions, GetMillions, GetMillions_1, GetMillions_10, GetMillions_100 As String

Dim Billions, GetBillions, GetBillions_1, GetBillions_10, GetBillions_100 As String

Dim Trillions, GetTrillions, GetTrillions_1, GetTrillions_10, GetTrillions_100 As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''' ONES ''''''''''''''''''''''''''''



If Len(Before_Amt)=1 Or (Left(right(Before_Amt,2),1)= "0" Or Left(right(Before_Amt,2),1)<> "1") Then
Digit = Right(Before_Amt,1)
Select Case Digit
 Case "1" : GetDigit="One"
 Case "2" : GetDigit="Two"
 Case "3" : GetDigit="Three"
 Case "4" : GetDigit="Four"
 Case "5" : GetDigit="Five"
 Case "6" : GetDigit="Six"
 Case "7" : GetDigit="Seven"
 Case "8" : GetDigit="Eight"
 Case "9" : GetDigit="Nine" 
 Case Else: GetDigit = ""
 End Select
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''' TENS ''''''''''''''''''''''''''''



If Len(Before_Amt)>=2 And Left(right(Before_Amt,2),1)= "1" Then

Tens = right(Before_Amt,2)

Select Case Tens
 Case "10" : GetTens="Ten"
 Case "11" : GetTens="Eleven"
 Case "12" : GetTens="Twelve"
 Case "13" : GetTens="Thirteen"
 Case "14" : GetTens="Forteen"
 Case "15" : GetTens="Fifteen"
 Case "16" : GetTens="Sixteen"
 Case "17" : GetTens="Seventeen"
 Case "18" : GetTens="Eighteen"
 Case "19" : GetTens="Nineteen"
 Case Else : GetTens=""
 End Select

ElseIf Len(Before_Amt)>=2 And Left(right(Before_Amt,2),1)<> "1" Then

Tens = Left(right(Before_Amt,2),1)

Select Case Tens
 Case "2" : GetTens="Twenty"
 Case "3" : GetTens="Thirty"
 Case "4" : GetTens="Fourty"
 Case "5" : GetTens="Fifty"
 Case "6" : GetTens="Sixty"
 Case "7" : GetTens="Seventy"
 Case "8" : GetTens="Eighty"
 Case "9" : GetTens="Ninety"
 Case Else : GetTens=""
 End Select
 End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''' HUNDREDS '''''''''''''''''''''''''''


If Len(Before_Amt)>=3 And Left(right(Before_Amt,3),1)<> "0" Then

Hundreds = Left(right(Before_Amt,3),1)

Select Case Hundreds
 Case "1" : GetHundreds="One"
 Case "2" : GetHundreds="Two"
 Case "3" : GetHundreds="Three"
 Case "4" : GetHundreds="Four"
 Case "5" : GetHundreds="Five"
 Case "6" : GetHundreds="Six"
 Case "7" : GetHundreds="Seven"
 Case "8" : GetHundreds="Eight"
 Case "9" : GetHundreds="Nine" 
 Case Else: GetHundreds = ""
 End Select
End If

If GetHundreds<>"" Then

GetHundreds = GetHundreds & " Hundred"

End If




''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''' THOUSANDS ''''''''''''''''''''''''''





If Len(Before_Amt)=4 Then

Thousands = Left(right(Before_Amt,4),1)

Select Case Thousands
 Case "1" : GetThousands_1="One"
 Case "2" : GetThousands_1="Two"
 Case "3" : GetThousands_1="Three"
 Case "4" : GetThousands_1="Four"
 Case "5" : GetThousands_1="Five"
 Case "6" : GetThousands_1="Six"
 Case "7" : GetThousands_1="Seven"
 Case "8" : GetThousands_1="Eight"
 Case "9" : GetThousands_1="Nine"
 Case Else: GetThousands_1 = ""
 End Select

Else If Len(Before_Amt)> 4 And Left(right(Before_Amt,5),1)<> "1" Then

Thousands = Left(right(Before_Amt,4),1)

Select Case Thousands
 Case "1" : GetThousands_1="One"
 Case "2" : GetThousands_1="Two"
 Case "3" : GetThousands_1="Three"
 Case "4" : GetThousands_1="Four"
 Case "5" : GetThousands_1="Five"
 Case "6" : GetThousands_1="Six"
 Case "7" : GetThousands_1="Seven"
 Case "8" : GetThousands_1="Eight"
 Case "9" : GetThousands_1="Nine"
 Case Else: GetThousands_1 = ""
 End Select

 End If



If Len(Before_Amt)>=5 And Left(right(Before_Amt,5),1)<> "0" And Left(right(Before_Amt,5),1)= "1" Then


Thousands = Left(right(Before_Amt,5),2)

Select Case Thousands
 Case "10" : GetThousands_10="Ten"
 Case "11" : GetThousands_10="Eleven"
 Case "12" : GetThousands_10="Twelve"
 Case "13" : GetThousands_10="Thirteen"
 Case "14" : GetThousands_10="Forteen"
 Case "15" : GetThousands_10="Fifteen"
 Case "16" : GetThousands_10="Sixteen"
 Case "17" : GetThousands_10="Seventeen"
 Case "18" : GetThousands_10="Eighteen"
 Case "19" : GetThousands_10="Nineteen"
 Case Else : GetThousands_10=""
 End Select

Else If Len(Before_Amt)>=5 And Left(right(Before_Amt,5),1)<> "0" And Left(right(Before_Amt,5),1)<> "1" Then

Thousands = Left(right(Before_Amt,5),1)

Select Case Thousands
 Case "2" : GetThousands_10="Twenty"
 Case "3" : GetThousands_10="Thirty"
 Case "4" : GetThousands_10="Fourty"
 Case "5" : GetThousands_10="Fifty"
 Case "6" : GetThousands_10="Sixty"
 Case "7" : GetThousands_10="Seventy"
 Case "8" : GetThousands_10="Eighty"
 Case "9" : GetThousands_10="Ninety"
 Case Else : GetThousands_10=""
 End Select
End If



If Len(Before_Amt)>=6 And Left(right(Before_Amt,6),1)<> "0" Then

Thousands = Left(right(Before_Amt,6),1)

Select Case Thousands
 Case "1" : GetThousands_100="One"
 Case "2" : GetThousands_100="Two"
 Case "3" : GetThousands_100="Three"
 Case "4" : GetThousands_100="Four"
 Case "5" : GetThousands_100="Five"
 Case "6" : GetThousands_100="Six"
 Case "7" : GetThousands_100="Seven"
 Case "8" : GetThousands_100="Eight"
 Case "9" : GetThousands_100="Nine" 
 Case Else: GetThousands_100 = ""
 End Select
End If

If GetThousands_100<>"" Then

GetThousands_100 = GetThousands_100 & " Hundred"

End If



If GetThousands_100="" And GetThousands_10="" And GetThousands_1="" Then

GetThousands = ""

Else 
	GetThousands = GetThousands_100 &" " & GetThousands_10 & " " & GetThousands_1 & " Thousand"

End If






''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' MILLIONS ''''''''''''''''''''''''''






If Len(Before_Amt)=7 Then

Millions = Left(right(Before_Amt,7),1)

Select Case Millions
 Case "1" : GetMillions_1="One"
 Case "2" : GetMillions_1="Two"
 Case "3" : GetMillions_1="Three"
 Case "4" : GetMillions_1="Four"
 Case "5" : GetMillions_1="Five"
 Case "6" : GetMillions_1="Six"
 Case "7" : GetMillions_1="Seven"
 Case "8" : GetMillions_1="Eight"
 Case "9" : GetMillions_1="Nine"
 Case Else: GetMillions_1 = ""
 End Select

Else If Len(Before_Amt)>=7 And Left(right(Before_Amt,8),1)<> "1" Then

Millions = Left(right(Before_Amt,7),1)

Select Case Millions
 Case "1" : GetMillions_1="One"
 Case "2" : GetMillions_1="Two"
 Case "3" : GetMillions_1="Three"
 Case "4" : GetMillions_1="Four"
 Case "5" : GetMillions_1="Five"
 Case "6" : GetMillions_1="Six"
 Case "7" : GetMillions_1="Seven"
 Case "8" : GetMillions_1="Eight"
 Case "9" : GetMillions_1="Nine"
 Case Else: GetMillions_1 = ""
 End Select


 End If

If Len(Before_Amt)>=8 And Left(right(Before_Amt,8),1)<> "0" And Left(right(Before_Amt,8),1)= "1" Then


Millions = Left(right(Before_Amt,8),2)

Select Case Millions
 Case "10" : GetMillions_10="Ten"
 Case "11" : GetMillions_10="Eleven"
 Case "12" : GetMillions_10="Twelve"
 Case "13" : GetMillions_10="Thirteen"
 Case "14" : GetMillions_10="Forteen"
 Case "15" : GetMillions_10="Fifteen"
 Case "16" : GetMillions_10="Sixteen"
 Case "17" : GetMillions_10="Seventeen"
 Case "18" : GetMillions_10="Eighteen"
 Case "19" : GetMillions_10="Nineteen"
 Case Else : GetMillions_10=""
 End Select

Else If Len(Before_Amt)>=8 And Left(right(Before_Amt,8),1)<> "0" And Left(right(Before_Amt,8),1)<> "1" Then

Millions = Left(right(Before_Amt,8),1)

Select Case Millions
 Case "2" : GetMillions_10="Twenty"
 Case "3" : GetMillions_10="Thirty"
 Case "4" : GetMillions_10="Fourty"
 Case "5" : GetMillions_10="Fifty"
 Case "6" : GetMillions_10="Sixty"
 Case "7" : GetMillions_10="Seventy"
 Case "8" : GetMillions_10="Eighty"
 Case "9" : GetMillions_10="Ninety"
 Case Else : GetMillions_10=""
 End Select
End If



If Len(Before_Amt)>=9 And Left(right(Before_Amt,9),1)<> "0" Then

Millions = Left(right(Before_Amt,9),1)

Select Case Millions
 Case "1" : GetMillions_100="One"
 Case "2" : GetMillions_100="Two"
 Case "3" : GetMillions_100="Three"
 Case "4" : GetMillions_100="Four"
 Case "5" : GetMillions_100="Five"
 Case "6" : GetMillions_100="Six"
 Case "7" : GetMillions_100="Seven"
 Case "8" : GetMillions_100="Eight"
 Case "9" : GetMillions_100="Nine" 
 Case Else: GetMillions_100 = ""
 End Select
End If

If GetMillions_100<>"" Then

GetMillions_100 = GetMillions_100 & " Hundred"

End If



If GetMillions_100="" And GetMillions_10="" And GetMillions_1="" Then

GetMillions = ""

Else 
	GetMillions = GetMillions_100 &" " & GetMillions_10 & " " & GetMillions_1 & " Million"

End If






''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' BILLIONS ''''''''''''''''''''''''''




If Len(Before_Amt)=10 Then

Billions = Left(right(Before_Amt,10),1)

Select Case Billions
 Case "1" : GetBillions_1="One"
 Case "2" : GetBillions_1="Two"
 Case "3" : GetBillions_1="Three"
 Case "4" : GetBillions_1="Four"
 Case "5" : GetBillions_1="Five"
 Case "6" : GetBillions_1="Six"
 Case "7" : GetBillions_1="Seven"
 Case "8" : GetBillions_1="Eight"
 Case "9" : GetBillions_1="Nine"
 Case Else: GetBillions_1 = ""
 End Select

Else If Len(Before_Amt)>=10 And Left(right(Before_Amt,11),1)<> "1" Then

Billions = Left(right(Before_Amt,10),1)

Select Case Billions
 Case "1" : GetBillions_1="One"
 Case "2" : GetBillions_1="Two"
 Case "3" : GetBillions_1="Three"
 Case "4" : GetBillions_1="Four"
 Case "5" : GetBillions_1="Five"
 Case "6" : GetBillions_1="Six"
 Case "7" : GetBillions_1="Seven"
 Case "8" : GetBillions_1="Eight"
 Case "9" : GetBillions_1="Nine"
 Case Else: GetBillions_1 = ""
 End Select
 End If

If Len(Before_Amt)>=11 And Left(right(Before_Amt,11),1)<> "0" And Left(right(Before_Amt,11),1)= "1" Then


Billions = Left(right(Before_Amt,11),2)

Select Case Billions
 Case "10" : GetBillions_10="Ten"
 Case "11" : GetBillions_10="Eleven"
 Case "12" : GetBillions_10="Twelve"
 Case "13" : GetBillions_10="Thirteen"
 Case "14" : GetBillions_10="Forteen"
 Case "15" : GetBillions_10="Fifteen"
 Case "16" : GetBillions_10="Sixteen"
 Case "17" : GetBillions_10="Seventeen"
 Case "18" : GetBillions_10="Eighteen"
 Case "19" : GetBillions_10="Nineteen"
 Case Else : GetBillions_10=""
 End Select

Else If Len(Before_Amt)>=11 And Left(right(Before_Amt,11),1)<> "0" And Left(right(Before_Amt,11),1)<> "1" Then

Billions = Left(right(Before_Amt,11),1)

Select Case Billions
 Case "2" : GetBillions_10="Twenty"
 Case "3" : GetBillions_10="Thirty"
 Case "4" : GetBillions_10="Fourty"
 Case "5" : GetBillions_10="Fifty"
 Case "6" : GetBillions_10="Sixty"
 Case "7" : GetBillions_10="Seventy"
 Case "8" : GetBillions_10="Eighty"
 Case "9" : GetBillions_10="Ninety"
 Case Else : GetBillions_10=""
 End Select
End If



If Len(Before_Amt)>=12 And Left(right(Before_Amt,12),1)<> "0" Then

Billions = Left(right(Before_Amt,12),1)

Select Case Billions
 Case "1" : GetBillions_100="One"
 Case "2" : GetBillions_100="Two"
 Case "3" : GetBillions_100="Three"
 Case "4" : GetBillions_100="Four"
 Case "5" : GetBillions_100="Five"
 Case "6" : GetBillions_100="Six"
 Case "7" : GetBillions_100="Seven"
 Case "8" : GetBillions_100="Eight"
 Case "9" : GetBillions_100="Nine" 
 Case Else: GetBillions_100 = ""
 End Select
End If

If GetBillions_100<>"" Then

GetBillions_100 = GetBillions_100 & " Hundred"

End If



If GetBillions_100="" And GetBillions_10="" And GetBillions_1="" Then

GetBillions = ""

Else 
	GetBillions = GetBillions_100 &" " & GetBillions_10 & " " & GetBillions_1 & " Billion"

End If






''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' TRILLIONS '''''''''''''''''''''''''





If Len(Before_Amt)=13 Then

Trillions = Left(right(Before_Amt,13),1)

Select Case Trillions
 Case "1" : GetTrillions_1="One"
 Case "2" : GetTrillions_1="Two"
 Case "3" : GetTrillions_1="Three"
 Case "4" : GetTrillions_1="Four"
 Case "5" : GetTrillions_1="Five"
 Case "6" : GetTrillions_1="Six"
 Case "7" : GetTrillions_1="Seven"
 Case "8" : GetTrillions_1="Eight"
 Case "9" : GetTrillions_1="Nine"
 Case Else: GetTrillions_1 = ""
 End Select

Else If Len(Before_Amt)>13 And Left(right(Before_Amt,14),1)<> "1" Then

Trillions = Left(right(Before_Amt,13),1)

Select Case Trillions
 Case "1" : GetTrillions_1="One"
 Case "2" : GetTrillions_1="Two"
 Case "3" : GetTrillions_1="Three"
 Case "4" : GetTrillions_1="Four"
 Case "5" : GetTrillions_1="Five"
 Case "6" : GetTrillions_1="Six"
 Case "7" : GetTrillions_1="Seven"
 Case "8" : GetTrillions_1="Eight"
 Case "9" : GetTrillions_1="Nine"
 Case Else: GetTrillions_1 = ""
 End Select
 End If

If Len(Before_Amt)>=14 And Left(right(Before_Amt,14),1)<> "0" And Left(right(Before_Amt,14),1)= "1" Then


Trillions = Left(right(Before_Amt,14),2)

Select Case Trillions
 Case "10" : GetTrillions_10="Ten"
 Case "11" : GetTrillions_10="Eleven"
 Case "12" : GetTrillions_10="Twelve"
 Case "13" : GetTrillions_10="Thirteen"
 Case "14" : GetTrillions_10="Forteen"
 Case "15" : GetTrillions_10="Fifteen"
 Case "16" : GetTrillions_10="Sixteen"
 Case "17" : GetTrillions_10="Seventeen"
 Case "18" : GetTrillions_10="Eighteen"
 Case "19" : GetTrillions_10="Nineteen"
 Case Else : GetTrillions_10=""
 End Select

Else If Len(Before_Amt)>=14 And Left(right(Before_Amt,14),1)<> "0" And Left(right(Before_Amt,14),1)<> "1" Then

Trillions = Left(right(Before_Amt,14),1)

Select Case Trillions
 Case "2" : GetTrillions_10="Twenty"
 Case "3" : GetTrillions_10="Thirty"
 Case "4" : GetTrillions_10="Fourty"
 Case "5" : GetTrillions_10="Fifty"
 Case "6" : GetTrillions_10="Sixty"
 Case "7" : GetTrillions_10="Seventy"
 Case "8" : GetTrillions_10="Eighty"
 Case "9" : GetTrillions_10="Ninety"
 Case Else : GetTrillions_10=""
 End Select
End If



If Len(Before_Amt)>=15 And Left(right(Before_Amt,15),1)<> "0" Then

Trillions = Left(right(Before_Amt,15),1)

Select Case Trillions
 Case "1" : GetTrillions_100="One"
 Case "2" : GetTrillions_100="Two"
 Case "3" : GetTrillions_100="Three"
 Case "4" : GetTrillions_100="Four"
 Case "5" : GetTrillions_100="Five"
 Case "6" : GetTrillions_100="Six"
 Case "7" : GetTrillions_100="Seven"
 Case "8" : GetTrillions_100="Eight"
 Case "9" : GetTrillions_100="Nine" 
 Case Else: GetTrillions_100 = ""
 End Select
End If

If GetTrillions_100<>"" Then

GetTrillions_100 = GetTrillions_100 & " Hundred"

End If

If GetTrillions_100="" And GetTrillions_10="" And GetTrillions_1="" Then

GetTrillions = ""

Else 
	GetTrillions = GetTrillions_100 &" " & GetTrillions_10 & " " & GetTrillions_1 & " Trillion"

End If



Amt_Before_Decimal = GetTrillions & " " &GetBillions & " " &GetMillions & " " & GetThousands & " " & GetHundreds & " " & GetTens & " " & GetDigit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' BEFORE DECIMAL ENDS HERE '''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''' DECIMALS '''''''''''''''''''''''''



If DecimalPlace <> 0 Then

DecimalCount = Mid(MyNumber.ToString, InStr(MyNumber.ToString, ".")+1, MyNumber.ToString.count - InStr(MyNumber.ToString, "."))


If Len(DecimalCount)=2 And left(DecimalCount,1) <> "1" Then

Select Case right(DecimalCount,1)
 Case "1" : DecimalCount_1 = "One"
 Case "2" : DecimalCount_1 = "Two"
 Case "3" : DecimalCount_1 = "Three"
 Case "4" : DecimalCount_1 = "Four"
 Case "5" : DecimalCount_1 = "Five"
 Case "6" : DecimalCount_1 = "Six"
 Case "7" : DecimalCount_1 = "Seven"
 Case "8" : DecimalCount_1 = "Eight"
 Case "9" : DecimalCount_1 = "Nine" 
 Case Else: DecimalCount_1 = ""
 End Select
End If

If Len(DecimalCount)<>0 And left(DecimalCount,1) <> "0" And left(DecimalCount,1) = "1"

Select Case DecimalCount
 Case "1"  : DecimalCount_10 = "Ten"
 Case "10" : DecimalCount_10 = "Ten"
 Case "11" : DecimalCount_10 = "Eleven"
 Case "12" : DecimalCount_10 = "Twelve"
 Case "13" : DecimalCount_10 = "Thirteen"
 Case "14" : DecimalCount_10 = "Forteen"
 Case "15" : DecimalCount_10 = "Fifteen"
 Case "16" : DecimalCount_10 = "Sixteen"
 Case "17" : DecimalCount_10 = "Seventeen"
 Case "18" : DecimalCount_10 = "Eighteen"
 Case "19" : DecimalCount_10 = "Nineteen"
 Case Else : DecimalCount_10 = ""
 End Select

Else If Len(DecimalCount)<>0 And left(DecimalCount,1) <> "0" And left(DecimalCount,1) <> "1"

Select Case left(DecimalCount,1)
 Case "2" : DecimalCount_10 = "Twenty"
 Case "3" : DecimalCount_10 = "Thirty"
 Case "4" : DecimalCount_10 = "Fourty"
 Case "5" : DecimalCount_10 = "Fifty"
 Case "6" : DecimalCount_10 = "Sixty"
 Case "7" : DecimalCount_10 = "Seventy"
 Case "8" : DecimalCount_10 = "Eighty"
 Case "9" : DecimalCount_10 = "Ninety"
 Case Else : DecimalCount_10 = ""
 End Select
 End If

Amt_After_Decimal = " " & DecimalCount_10 & " " & DecimalCount_1

Else If DecimalPlace = 0 Then

Amt_After_Decimal = ""

End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' AFTER DECIMAL ENDS HERE '''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''' COUNTRY TYPE AND CURRENCY '''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select UCase(Curr_Type)


Case "AFN"                          'Country:  Afghanistan         Currency : Afghan afghani
Country_currency = "Afghani"   
Decimal_Singular = "Pul"
Decimal_Plural = "Pul"


Case "€"                          'Country:  Akrotiri and Dhekelia (UK)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Aland Islands (Finland)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ALL"                          'Country:  Albania         Currency : Albanian lek
Country_currency = "lek"   
Decimal_Singular = "Qindarkë"
Decimal_Plural = "Qindarka"


Case "DZD"                          'Country:  Algeria         Currency : Algerian dinar
Country_currency = "Dinar"   
Decimal_Singular = "Santeem"
Decimal_Plural = "Santeems"


Case "USD"                          'Country:  American Samoa (USA)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Andorra         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AOA"                          'Country:  Angola         Currency : Angolan kwanza
Country_currency = "Kwanza"   
Decimal_Singular = "Cêntimo"
Decimal_Plural = "Cêntimo"


Case "XCD"                          'Country:  Anguilla (UK)         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XCD"                          'Country:  Antigua and Barbuda         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ARS"                          'Country:  Argentina         Currency : Argentine peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "AMD"                          'Country:  Armenia         Currency : Armenian dram
Country_currency = "Dram"   
Decimal_Singular = "Luma"
Decimal_Plural = "Luma"


Case "AWG"                          'Country:  Aruba (Netherlands)         Currency : Aruban florin
Country_currency = "Florin"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SHP"                          'Country:  Ascension Island (UK)         Currency : Saint Helena pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "AUD"                          'Country:  Australia         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Austria         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AZN"                          'Country:  Azerbaijan         Currency : Azerbaijan manat
Country_currency = "Manat"   
Decimal_Singular = "Qəpik"
Decimal_Plural = "Qəpik"


Case "BSD"                          'Country:  Bahamas         Currency : Bahamian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BHD"                          'Country:  Bahrain         Currency : Bahraini dinar
Country_currency = "Dinar"   
Decimal_Singular = "Fils"
Decimal_Plural = "Fils"


Case "BDT"                          'Country:  Bangladesh         Currency : Bangladeshi taka
Country_currency = "Taka"   
Decimal_Singular = "Poisha"
Decimal_Plural = "Poisha"


Case "BBD"                          'Country:  Barbados         Currency : Barbadian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BYN"                          'Country:  Belarus         Currency : Belarusian ruble
Country_currency = "Ruble"   
Decimal_Singular = "Kapiejka"
Decimal_Plural = "Kapiejka"


Case "EUR"                          'Country:  Belgium         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BZD"                          'Country:  Belize         Currency : Belize dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XOF"                          'Country:  Benin         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "BMD"                          'Country:  Bermuda (UK)         Currency : Bermudian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BTN"                          'Country:  Bhutan         Currency : Bhutanese ngultrum
Country_currency = "Ngultrum"   
Decimal_Singular = "Chetrum"
Decimal_Plural = "Chetrums"


Case "BOB"                          'Country:  Bolivia         Currency : Bolivian boliviano
Country_currency = "Boliviano"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "$"                          'Country:  Bonaire (Netherlands)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BAM"                          'Country:  Bosnia and Herzegovina         Currency : Bosnia and Herzegovina convertible mark
Country_currency = "Mark"   
Decimal_Singular = "Fening"
Decimal_Plural = "Fenings"


Case "BWP"                          'Country:  Botswana         Currency : Botswana pula
Country_currency = "Pula"   
Decimal_Singular = "Thebe"
Decimal_Plural = "Thebe"


Case "BRL"                          'Country:  Brazil         Currency : Brazilian real
Country_currency = "Real"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "USD"                          'Country:  British Indian Ocean Territory (UK)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  British Virgin Islands (UK)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BND"                          'Country:  Brunei         Currency : Brunei dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "BGN"                          'Country:  Bulgaria         Currency : Bulgarian lev
Country_currency = "Lev"   
Decimal_Singular = "Stotinka"
Decimal_Plural = "Stotinki"


Case "XOF"                          'Country:  Burkina Faso         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "BIF"                          'Country:  Burundi         Currency : Burundi franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "CVE"                          'Country:  Cabo Verde         Currency : Cape Verdean escudo
Country_currency = "Escudo"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "KHR"                          'Country:  Cambodia         Currency : Cambodian riel
Country_currency = "Riel"   
Decimal_Singular = "Sen"
Decimal_Plural = "Sen"


Case "XAF"                          'Country:  Cameroon         Currency : Central African CFA franc
Country_currency = "CFA Franc BEAC"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "CAD"                          'Country:  Canada         Currency : Canadian dollar
Country_currency = "Canadian Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  Caribbean Netherlands (Netherlands)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "KYD"                          'Country:  Cayman Islands (UK)         Currency : Cayman Islands dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XAF"                          'Country:  Central African Republic         Currency : Central African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "XAF"                          'Country:  Chad         Currency : Central African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "NZD"                          'Country:  Chatham Islands (New Zealand)         Currency : New Zealand dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "CLP"                          'Country:  Chile         Currency : Chilean peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "CNY"                          'Country:  China         Currency : Chinese Yuan Renminbi
Country_currency = "Yuan"   
Decimal_Singular = "Fen"
Decimal_Plural = "Fen"


Case "AUD"                          'Country:  Christmas Island (Australia)         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AUD"                          'Country:  Cocos (Keeling) Islands (Australia)         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "COP"                          'Country:  Colombia         Currency : Colombian peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "KMF"                          'Country:  Comoros         Currency : Comorian franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "CDF"                          'Country:  Congo, Democratic Republic of the         Currency : Congolese franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "XAF"                          'Country:  Congo, Republic of the         Currency : Central African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "none"                          'Country:  Cook Islands (New Zealand)         Currency : Cook Islands dollar
Country_currency = "Dollar"   
Decimal_Singular = "Tene"
Decimal_Plural = "Tene"


Case "CRC"                          'Country:  Costa Rica         Currency : Costa Rican colon
Country_currency = "Colon"   
Decimal_Singular = "Centimo"
Decimal_Plural = "Centimos"


Case "XOF"                          'Country:  Cote d'Ivoire         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "HRK"                          'Country:  Croatia         Currency : Croatian kuna
Country_currency = "Kuna"   
Decimal_Singular = "Lipa"
Decimal_Plural = "Lipas"


Case "CUP"                          'Country:  Cuba         Currency : Cuban peso
Country_currency = "Cuban Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "ANG"                          'Country:  Curacao (Netherlands)         Currency : Netherlands Antillean guilder
Country_currency = "Guilder"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Cyprus         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "CZK"                          'Country:  Czech Republic         Currency : Czech koruna
Country_currency = "Krona"   
Decimal_Singular = "Haler"
Decimal_Plural = "Halers"


Case "DKK"                          'Country:  Denmark         Currency : Danish krone
Country_currency = "Krone"   
Decimal_Singular = "Ore"
Decimal_Plural = "Ore"


Case "DJF"                          'Country:  Djibouti         Currency : Djiboutian franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "XCD"                          'Country:  Dominica         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "DOP"                          'Country:  Dominican Republic         Currency : Dominican peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "USD"                          'Country:  Ecuador         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EGP"                          'Country:  Egypt         Currency : Egyptian pound
Country_currency = "Pound"   
Decimal_Singular = "Piastre"
Decimal_Plural = "Piastres"


Case "$"                          'Country:  El Salvador         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XAF"                          'Country:  Equatorial Guinea         Currency : Central African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "ERN"                          'Country:  Eritrea         Currency : Eritrean nakfa
Country_currency = "Nakfa"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Estonia         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SZL"                          'Country:  Eswatini (formerly Swaziland)         Currency : Swazi lilangeni
Country_currency = "Lilangeni"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ETB"                          'Country:  Ethiopia         Currency : Ethiopian birr
Country_currency = "Birr"   
Decimal_Singular = "Santim"
Decimal_Plural = "Santims"


Case "FKP"                          'Country:  Falkland Islands (UK)         Currency : Falkland Islands pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "none"                          'Country:  Faroe Islands (Denmark)         Currency : Faroese krona
Country_currency = "Krona"   
Decimal_Singular = "Oyra"
Decimal_Plural = "Oyru"


Case "FJD"                          'Country:  Fiji         Currency : Fijian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Finland         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  France         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  French Guiana (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XPF"                          'Country:  French Polynesia (France)         Currency : CFP franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "XAF"                          'Country:  Gabon         Currency : Central African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centimes"


Case "GMD"                          'Country:  Gambia         Currency : Gambian dalasi
Country_currency = "Dalasi"   
Decimal_Singular = "Butut"
Decimal_Plural = "Bututs"


Case "GEL"                          'Country:  Georgia         Currency : Georgian lari
Country_currency = "Lari"   
Decimal_Singular = "Tetri"
Decimal_Plural = "Tetri"


Case "EUR"                          'Country:  Germany         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "GHS"                          'Country:  Ghana         Currency : Ghanaian cedi
Country_currency = "Cedi"   
Decimal_Singular = "Pesewa"
Decimal_Plural = "Pesewas"


Case "GIP"                          'Country:  Gibraltar (UK)         Currency : Gibraltar pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "EUR"                          'Country:  Greece         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "DKK"                          'Country:  Greenland (Denmark)         Currency : Danish krone
Country_currency = "Krone"   
Decimal_Singular = "Ore"
Decimal_Plural = "Ore"


Case "XCD"                          'Country:  Grenada         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Guadeloupe (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  Guam (USA)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "GTQ"                          'Country:  Guatemala         Currency : Guatemalan quetzal
Country_currency = "Quetzal"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "GGP"                          'Country:  Guernsey (UK)         Currency : Guernsey Pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "GNF"                          'Country:  Guinea         Currency : Guinean franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "XOF"                          'Country:  Guinea-Bissau         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "GYD"                          'Country:  Guyana         Currency : Guyanese dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "HTG"                          'Country:  Haiti         Currency : Haitian gourde
Country_currency = "Gourde"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "HNL"                          'Country:  Honduras         Currency : Honduran lempira
Country_currency = "Lempira"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "HKD"                          'Country:  Hong Kong (China)         Currency : Hong Kong dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "HUF"                          'Country:  Hungary         Currency : Hungarian forint
Country_currency = "Forint"   
Decimal_Singular = "Filler"
Decimal_Plural = "Filler"


Case "ISK"                          'Country:  Iceland         Currency : Icelandic krona
Country_currency = "Krona"   
Decimal_Singular = "Eyrir"
Decimal_Plural = "Aurar"


Case "INR"                          'Country:  India         Currency : Indian rupee
Country_currency = "Rupee"   
Decimal_Singular = "Paise"
Decimal_Plural = "Paise"


Case "IDR"                          'Country:  Indonesia         Currency : Indonesian rupiah
Country_currency = "Rupiah"   
Decimal_Singular = "Sen"
Decimal_Plural = "Sen"


Case "IRR"                          'Country:  Iran         Currency : Iranian rial
Country_currency = "Rial"   
Decimal_Singular = "Dinar"
Decimal_Plural = "Dinars"


Case "IQD"                          'Country:  Iraq         Currency : Iraqi dinar
Country_currency = "Dinar"   
Decimal_Singular = "Fils"
Decimal_Plural = "Fils"


Case "€"                          'Country:  Ireland         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "IMP"                          'Country:  Isle of Man (UK)         Currency : Manx pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "ILS"                          'Country:  Israel         Currency : Israeli new shekel
Country_currency = "Scheckel"   
Decimal_Singular = "Agora"
Decimal_Plural = "Agoras"


Case "€"                          'Country:  Italy         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "JMD"                          'Country:  Jamaica         Currency : Jamaican dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "JPY"                          'Country:  Japan         Currency : Japanese yen
Country_currency = "Yen"   
Decimal_Singular = "Sen"
Decimal_Plural = "Sen"


Case "JEP"                          'Country:  Jersey (UK)         Currency : Jersey pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "JOD"                          'Country:  Jordan         Currency : Jordanian dinar
Country_currency = "Dinar"   
Decimal_Singular = "Piastre"
Decimal_Plural = "Piastre"


Case "KZT"                          'Country:  Kazakhstan         Currency : Kazakhstani tenge
Country_currency = "Tenge"   
Decimal_Singular = "Tiin"
Decimal_Plural = "Tiin"


Case "KES"                          'Country:  Kenya         Currency : Kenyan shilling
Country_currency = "Shilling"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AUD"                          'Country:  Kiribati         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Kosovo         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "KWD"                          'Country:  Kuwait         Currency : Kuwaiti dinar
Country_currency = "Dinar"   
Decimal_Singular = "Fils"
Decimal_Plural = "Fils"


Case "KGS"                          'Country:  Kyrgyzstan         Currency : Kyrgyzstani som
Country_currency = "Som"   
Decimal_Singular = "Tyiyn"
Decimal_Plural = "Tyiyns"


Case "LAK"                          'Country:  Laos         Currency : Lao kip
Country_currency = "Kip"   
Decimal_Singular = "Att"
Decimal_Plural = "Att"


Case "EUR"                          'Country:  Latvia         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "LBP"                          'Country:  Lebanon         Currency : Lebanese pound
Country_currency = "Pound"   
Decimal_Singular = "Piastre"
Decimal_Plural = "Piastre"


Case "LSL"                          'Country:  Lesotho         Currency : Lesotho loti
Country_currency = "Loti"   
Decimal_Singular = "Sente"
Decimal_Plural = "Sente"


Case "LRD"                          'Country:  Liberia         Currency : Liberian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "LYD"                          'Country:  Libya         Currency : Libyan dinar
Country_currency = "Dinar"   
Decimal_Singular = "Dirham"
Decimal_Plural = "Dirhams"


Case "CHF"                          'Country:  Liechtenstein         Currency : Swiss franc
Country_currency = "Franc"   
Decimal_Singular = "Rappen"
Decimal_Plural = "Rappen"


Case "EUR"                          'Country:  Lithuania         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Luxembourg         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MOP"                          'Country:  Macau (China)         Currency : Macanese pataca
Country_currency = "Pataca"   
Decimal_Singular = "Avo"
Decimal_Plural = "Avo"


Case "MKD"                          'Country:  Macedonia (FYROM)         Currency : Macedonian denar
Country_currency = "Denar"   
Decimal_Singular = "Deni"
Decimal_Plural = "Denari"


Case "MGA"                          'Country:  Madagascar         Currency : Malagasy ariary
Country_currency = "Ariary"   
Decimal_Singular = "Centimes"
Decimal_Plural = "Centimes"


Case "MWK"                          'Country:  Malawi         Currency : Malawian kwacha
Country_currency = "Kwacha"   
Decimal_Singular = "Tambala"
Decimal_Plural = "Tambala"


Case "RM"                          'Country:  Malaysia         Currency : Malaysian ringgit
Country_currency = "Ringgit"   
Decimal_Singular = "Sen"
Decimal_Plural = "Sen"


Case "MVR"                          'Country:  Maldives         Currency : Maldivian rufiyaa
Country_currency = "Rufiyaa"   
Decimal_Singular = "Laari"
Decimal_Plural = "Laari"


Case "XOF"                          'Country:  Mali         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "EUR"                          'Country:  Malta         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  Marshall Islands         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Martinique (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MRU"                          'Country:  Mauritania         Currency : Mauritanian ouguiya
Country_currency = "Ouguiya"   
Decimal_Singular = ""
Decimal_Plural = ""


Case "MUR"                          'Country:  Mauritius         Currency : Mauritian rupee
Country_currency = "Rupee"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "€"                          'Country:  Mayotte (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MXN"                          'Country:  Mexico         Currency : Mexican peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "USD"                          'Country:  Micronesia         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MDL"                          'Country:  Moldova         Currency : Moldovan leu
Country_currency = "Leu"   
Decimal_Singular = "Ban"
Decimal_Plural = "Bani"


Case "€"                          'Country:  Monaco         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MNT"                          'Country:  Mongolia         Currency : Mongolian tugrik
Country_currency = "Tugrik"   
Decimal_Singular = "Mongo"
Decimal_Plural = "Mongo"


Case "€"                          'Country:  Montenegro         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XCD"                          'Country:  Montserrat (UK)         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "MAD"                          'Country:  Morocco         Currency : Moroccan dirham
Country_currency = "Dirham"   
Decimal_Singular = "Santim"
Decimal_Plural = "Santim"


Case "MZN"                          'Country:  Mozambique         Currency : Mozambican metical
Country_currency = "Metical"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "MMK"                          'Country:  Myanmar (formerly Burma)         Currency : Myanmar kyat
Country_currency = "Kyat"   
Decimal_Singular = "Pya"
Decimal_Plural = "Pya"


Case "NAD"                          'Country:  Namibia         Currency : Namibian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AUD"                          'Country:  Nauru         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "NPR"                          'Country:  Nepal         Currency : Nepalese rupee
Country_currency = "Rupee"   
Decimal_Singular = "Paise"
Decimal_Plural = "Paise"


Case "€"                          'Country:  Netherlands         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XPF"                          'Country:  New Caledonia (France)         Currency : CFP franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "NZD"                          'Country:  New Zealand         Currency : New Zealand dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "NIO"                          'Country:  Nicaragua         Currency : Nicaraguan cordoba
Country_currency = "Oro"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavo"


Case "XOF"                          'Country:  Niger         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "NGN"                          'Country:  Nigeria         Currency : Nigerian naira
Country_currency = "Naira"   
Decimal_Singular = "Kobo"
Decimal_Plural = "Kobo"


Case "NZD"                          'Country:  Niue (New Zealand)         Currency : New Zealand dollar
Country_currency = "N.Zeal.Dollars"   
Decimal_Singular = "Cent"
Decimal_Plural = "Centime"


Case "AUD"                          'Country:  Norfolk Island (Australia)         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  Northern Mariana Islands (USA)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "KPW"                          'Country:  North Korea         Currency : North Korean won
Country_currency = "Won"   
Decimal_Singular = "Chon"
Decimal_Plural = "Chon"


Case "NOK"                          'Country:  Norway         Currency : Norwegian krone
Country_currency = "Krone"   
Decimal_Singular = "Ore"
Decimal_Plural = "Øre"


Case "OMR"                          'Country:  Oman         Currency : Omani rial
Country_currency = "Omani Rial"   
Decimal_Singular = "Baisa"
Decimal_Plural = "Baisa"


Case "PKR"                          'Country:  Pakistan         Currency : Pakistani rupee
Country_currency = "Rupee"   
Decimal_Singular = "Paise"
Decimal_Plural = "Paise"


Case "$"                          'Country:  Palau         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ILS"                          'Country:  Palestine         Currency : Israeli new shekel
Country_currency = "Scheckel"   
Decimal_Singular = "Agora"
Decimal_Plural = "Agoras"


Case "$"                          'Country:  Panama         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "PGK"                          'Country:  Papua New Guinea         Currency : Papua New Guinean kina
Country_currency = "Kina"   
Decimal_Singular = "Toea"
Decimal_Plural = "Toea"


Case "PYG"                          'Country:  Paraguay         Currency : Paraguayan guarani
Country_currency = "Guarani"   
Decimal_Singular = "Centimo"
Decimal_Plural = "Centimo"


Case "PEN"                          'Country:  Peru         Currency : Peruvian sol
Country_currency = "Sol"   
Decimal_Singular = "Centimo"
Decimal_Plural = "Centimo"


Case "PHP"                          'Country:  Philippines         Currency : Philippine peso
Country_currency = "Peso"   
Decimal_Singular = "Centavo"
Decimal_Plural = "Centavos"


Case "NZD"                          'Country:  Pitcairn Islands (UK)         Currency : New Zealand dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "PLN"                          'Country:  Poland         Currency : Polish zloty
Country_currency = "Zloty"   
Decimal_Singular = "Grosz"
Decimal_Plural = "Grosz"


Case "€"                          'Country:  Portugal         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "$"                          'Country:  Puerto Rico (USA)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "QAR"                          'Country:  Qatar         Currency : Qatari riyal
Country_currency = "Rial"   
Decimal_Singular = "Dirham"
Decimal_Plural = "Dirhams"


Case "€"                          'Country:  Reunion (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "RON"                          'Country:  Romania         Currency : Romanian leu
Country_currency = "Leu"   
Decimal_Singular = "Ban"
Decimal_Plural = "Bani"


Case "RUB"                          'Country:  Russia         Currency : Russian ruble
Country_currency = "Ruble"   
Decimal_Singular = "Kopek"
Decimal_Plural = "Kopek"


Case "RWF"                          'Country:  Rwanda         Currency : Rwandan franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "$"                          'Country:  Saba (Netherlands)         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Saint Barthelemy (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SHP"                          'Country:  Saint Helena (UK)         Currency : Saint Helena pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "XCD"                          'Country:  Saint Kitts and Nevis         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XCD"                          'Country:  Saint Lucia         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Saint Martin (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Saint Pierre and Miquelon (France)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XCD"                          'Country:  Saint Vincent and the Grenadines         Currency : East Caribbean dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "WST"                          'Country:  Samoa         Currency : Samoan tala
Country_currency = "Tala"   
Decimal_Singular = "Sene"
Decimal_Plural = "Sene"


Case "EUR"                          'Country:  San Marino         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "STN"                          'Country:  Sao Tome and Principe         Currency : Sao Tome and Principe dobra
Country_currency = "Dobra"   
Decimal_Singular = "Cêntimo"
Decimal_Plural = "Cêntimo"


Case "SAR"                          'Country:  Saudi Arabia         Currency : Saudi Arabian riyal
Country_currency = "Rial"   
Decimal_Singular = "Halalah"
Decimal_Plural = "Halalah"


Case "XOF"                          'Country:  Senegal         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "RSD"                          'Country:  Serbia         Currency : Serbian dinar
Country_currency = "Dinar"   
Decimal_Singular = "Para"
Decimal_Plural = "Paras"


Case "SCR"                          'Country:  Seychelles         Currency : Seychellois rupee
Country_currency = "Rupee"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SLL"                          'Country:  Sierra Leone         Currency : Sierra Leonean leone
Country_currency = "Leone"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SGD"                          'Country:  Singapore         Currency : Singapore dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "USD"                          'Country:  Sint Eustatius (Netherlands)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ANG"                          'Country:  Sint Maarten (Netherlands)         Currency : Netherlands Antillean guilder
Country_currency = "Guilder"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Slovakia         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "EUR"                          'Country:  Slovenia         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SBD"                          'Country:  Solomon Islands         Currency : Solomon Islands dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "SOS"                          'Country:  Somalia         Currency : Somali shilling
Country_currency = "Shilling"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "ZAR"                          'Country:  South Africa         Currency : South African rand
Country_currency = "Rand"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "GBP"                          'Country:  South Georgia Island (UK)         Currency : Pound sterling
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "KRW"                          'Country:  South Korea         Currency : South Korean won
Country_currency = "Won"   
Decimal_Singular = "Jeon"
Decimal_Plural = "Jeon"


Case "SSP"                          'Country:  South Sudan         Currency : South Sudanese pound
Country_currency = "Pound"   
Decimal_Singular = "Piaster"
Decimal_Plural = "Piasters"


Case "EUR"                          'Country:  Spain         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "LKR"                          'Country:  Sri Lanka         Currency : Sri Lankan rupee
Country_currency = "Rupee"   
Decimal_Singular = "Paise"
Decimal_Plural = "Paise"


Case "SDG"                          'Country:  Sudan         Currency : Sudanese pound
Country_currency = "Pound"   
Decimal_Singular = "Qirsh"
Decimal_Plural = "Qirsh"


Case "SRD"                          'Country:  Suriname         Currency : Surinamese dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "NOK"                          'Country:  Svalbard and Jan Mayen (Norway)         Currency : Norwegian krone
Country_currency = "Krone"   
Decimal_Singular = "Ore"
Decimal_Plural = "Ore"


Case "SEK"                          'Country:  Sweden         Currency : Swedish krona
Country_currency = "Krona"   
Decimal_Singular = "Ore"
Decimal_Plural = "Ore"


Case "CHF"                          'Country:  Switzerland         Currency : Swiss franc
Country_currency = "Franc"   
Decimal_Singular = "Rappen"
Decimal_Plural = "Rappen"


Case "SYP"                          'Country:  Syria         Currency : Syrian pound
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "TWD"                          'Country:  Taiwan         Currency : New Taiwan dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "TJS"                          'Country:  Tajikistan         Currency : Tajikistani somoni
Country_currency = "Somoni"   
Decimal_Singular = "Diram"
Decimal_Plural = "Dirams"


Case "TZS"                          'Country:  Tanzania         Currency : Tanzanian shilling
Country_currency = "Shilling"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "THB"                          'Country:  Thailand         Currency : Thai baht
Country_currency = "Baht"   
Decimal_Singular = "Satang"
Decimal_Plural = "Satang"


Case "USD"                          'Country:  Timor-Leste         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XOF"                          'Country:  Togo         Currency : West African CFA franc
Country_currency = "CFA Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "NZD"                          'Country:  Tokelau (New Zealand)         Currency : New Zealand dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "TOP"                          'Country:  Tonga         Currency : Tongan pa’anga
Country_currency = "Pa’Anga"   
Decimal_Singular = "Seniti"
Decimal_Plural = "Seniti"


Case "TTD"                          'Country:  Trinidad and Tobago         Currency : Trinidad and Tobago dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "GBP"                          'Country:  Tristan da Cunha (UK)         Currency : Pound sterling
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "TND"                          'Country:  Tunisia         Currency : Tunisian dinar
Country_currency = "Dinar"   
Decimal_Singular = "Millime"
Decimal_Plural = "Millime"


Case "TRY"                          'Country:  Turkey         Currency : Turkish lira
Country_currency = "Lira"   
Decimal_Singular = "Kurus"
Decimal_Plural = "Kurus"


Case "TMT"                          'Country:  Turkmenistan         Currency : Turkmen manat
Country_currency = "Manat"   
Decimal_Singular = "Tenge"
Decimal_Plural = "Tenge"


Case "USD"                          'Country:  Turks and Caicos Islands (UK)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "AUD"                          'Country:  Tuvalu         Currency : Australian dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "UGX"                          'Country:  Uganda         Currency : Ugandan shilling
Country_currency = "Shilling"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "UAH"                          'Country:  Ukraine         Currency : Ukrainian hryvnia
Country_currency = "Hryvnia"   
Decimal_Singular = "Kopiyka"
Decimal_Plural = "Kopiyky"


Case "AED"                          'Country:  United Arab Emirates         Currency : UAE dirham
Country_currency = "Dirham"   
Decimal_Singular = "Fils"
Decimal_Plural = "Fils"


Case "GBP"                          'Country:  United Kingdom         Currency : Pound sterling
Country_currency = "Pound"   
Decimal_Singular = "Penny"
Decimal_Plural = "Pence"


Case "$"                          'Country:  United States of America         Currency : United States dollar
Country_currency = "U.S. Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "UYU"                          'Country:  Uruguay         Currency : Uruguayan peso
Country_currency = "Peso"   
Decimal_Singular = "Centesimo"
Decimal_Plural = "Centesimos"


Case "USD"                          'Country:  US Virgin Islands (USA)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "UZS"                          'Country:  Uzbekistan         Currency : Uzbekistani som
Country_currency = "Soʻm"   
Decimal_Singular = "Tiyin"
Decimal_Plural = "Tiyin"


Case "VUV"                          'Country:  Vanuatu         Currency : Vanuatu vatu
Country_currency = "Vatu"   
Decimal_Singular = ""
Decimal_Plural = ""


Case "EUR"                          'Country:  Vatican City (Holy See)         Currency : European euro
Country_currency = "Euro"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "VEF"                          'Country:  Venezuela         Currency : Venezuelan bolivar
Country_currency = "Bolivar Hard"   
Decimal_Singular = "Centimo"
Decimal_Plural = "Centimo"


Case "VND"                          'Country:  Vietnam         Currency : Vietnamese dong
Country_currency = "Dong"   
Decimal_Singular = "Hao"
Decimal_Plural = "Hao"


Case "USD"                          'Country:  Wake Island (USA)         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"


Case "XPF"                          'Country:  Wallis and Futuna (France)         Currency : CFP franc
Country_currency = "Franc"   
Decimal_Singular = "Centime"
Decimal_Plural = "Centime"


Case "YER"                          'Country:  Yemen         Currency : Yemeni rial
Country_currency = "Ryal"   
Decimal_Singular = "Fils"
Decimal_Plural = "Fils"


Case "ZMW"                          'Country:  Zambia         Currency : Zambian kwacha
Country_currency = "Kwacha"   
Decimal_Singular = "Ngwee"
Decimal_Plural = "Ngwee"


Case "USD"                          'Country:  Zimbabwe         Currency : United States dollar
Country_currency = "Dollar"   
Decimal_Singular = "Cent"
Decimal_Plural = "Cents"

Case Else : 

Country_currency = ""   
Decimal_Singular = ""
Decimal_Plural = ""

End Select




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''' COUNTRY TYPE AND CURRENCY ENDS HERE '''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' 
' Assign money currency before decimal
Select Case Amt_Before_Decimal.replace(" ","")
        Case ""
            Amt_Before_Decimal = ""
        Case "One"
            Amt_Before_Decimal = "One " + StrConv(Country_currency,vbProperCase)
        Case Else
End Select

If Country_currency.Equals("") Then
          Amt_Before_Decimal = LTrim(Amt_Before_Decimal) + " only"
     Else
		 Amt_Before_Decimal = LTrim(Amt_Before_Decimal) + Country_currency
End If

Console.WriteLine("Amt Before Dec: " + Amt_Before_Decimal)

'Get value for after decimal
If Amt_Before_Decimal.Equals("") Then
	If Amt_After_Decimal.Equals("") Then
		Num2Word = ""
	Else If Amt_After_Decimal.Equals("1") Then
		Num2Word = Amt_After_Decimal + " " + Decimal_Singular +" only"
	Else
		Num2Word = Amt_After_Decimal + " " + Decimal_Plural +" only"
	End If
	Else
	If Amt_After_Decimal.Equals("") Then
		Num2Word = Amt_Before_Decimal
	Else If Amt_After_Decimal.Equals("1") Then
		Num2Word = Amt_Before_Decimal + " and " + Amt_After_Decimal + " " + Decimal_Singular +" only"
	Else
		Num2Word = Amt_Before_Decimal + " and " + Amt_After_Decimal + " " + Decimal_Plural +" only"
	End If
End If
 
Console.WriteLine("Amt After Decimal: " + Num2Word)