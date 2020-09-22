<div align="center">

## ConvertBase\+


</div>

### Description

Converts a number of any base into any other base. For example, it

can convert a binary number into decimal, octal, hexadecimal, etc. and

back again. This is an improvement over the previous version as it can now

handle the 'double' data-type. (Very Large Numbers). It is useful for

converting DWords from the registry, etc.
 
### More Info
 
1. A Number to Convert

2. The base of that Number

3. The base of the desired number

A string containing the number in the specified base


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aidan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aidan.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aidan-convertbase__1-1474/archive/master.zip)





### Source Code

```
' Place the following code in a form...
Private Sub Form_Load()
  Debug.Print ConvertBase("10", 10, 16)
End Sub
Public Function ConvertBase(NumIn As String, BaseIn As Integer,_
BaseOut As Integer) As String
  ' Converts a number from one base to another
    ' E.g. Binary = Base 2
    '    Octal = Base 8
    '    Decimal = Base 10
    '    Hexadecimal = Base 16
  ' NumIn is the number which you wish to convert
	' (A string including characters 0 - 9, A - Z)
  ' BaseIn is the base of NumIn (An integer value in
	' decimal between 1 & 36)
  ' BaseOut is the base of the number the function
	' returns (An integer value in decimal between 1 & 36)
  ' Returns a string in the desired base containing the
	' characters 0 - 9, A - Z)
    ' e.g. Debug.Print ConvertBase ("42", 8, 16) converts the octal number 42 into hexadecimal
        ' Returns the string "22"
    ' Returns the word "Error" if any of the input values
	' are incorrect
  Dim i As Integer, CurrentCharacter As String,_
CharacterValue As Integer, PlaceValue As Integer,_
RunningTotal As Double, Remainder As Double,_
BaseOutDouble As Double, NumInCaps As String
  ' Ensure input data is valid
  If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or_
BaseOut < 1 Or BaseOut > 36 Then
    ConvertBase = "Error"
    Exit Function
  End If
  ' Ensure any letters in the input mumber are capitals
  NumInCaps = UCase$(NumIn)
  ' Convert NumInCaps into Decimal
  PlaceValue = Len(NumInCaps)
  For i = 1 To Len(NumInCaps)
    PlaceValue = PlaceValue - 1
    CurrentCharacter = Mid$(NumInCaps, i, 1)
    CharacterValue = 0
    If Asc(CurrentCharacter) > 64 And _
Asc(CurrentCharacter) < 91 Then _
CharacterValue = Asc(CurrentCharacter) - 55
    If CharacterValue = 0 Then
      ' Ensure NumIn is correct
      If Asc(CurrentCharacter) < 48 Or _
Asc(CurrentCharacter) > 57 Then
        ConvertBase = "Error"
        Exit Function
      Else
        CharacterValue = Val(CurrentCharacter)
      End If
    End If
    If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
      ' Ensure NumIn is correct
      ConvertBase = "Error"
      Exit Function
    End If
    RunningTotal = RunningTotal + CharacterValue *_
(BaseIn ^ PlaceValue)
  Next i
  ' Convert Decimal Number into the desired base using
	' Repeated Division
  Do
    BaseOutDouble = CDbl(BaseOut)
    Remainder = ModDouble(RunningTotal, BaseOutDouble)
    RunningTotal = (RunningTotal - Remainder) / BaseOut
    If Remainder >= 10 Then
      CurrentCharacter = Chr$(Remainder + 55)
    Else
      CurrentCharacter = Right$(Str$(Remainder),_
Len(Str$(Remainder)) - 1)
    End If
    ConvertBase = CurrentCharacter & ConvertBase
  Loop While RunningTotal > 0
End Function
Public Function ModDouble(NumIn As Double, DivNum As Double) As Double
  ' Returns the Remainder when a number is divided by another
  ' (Works for double data-type)
  ModDouble = NumIn - (Int(NumIn / DivNum) * DivNum)
End Function
```

