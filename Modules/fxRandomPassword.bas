Attribute VB_Name = "fxRandomPassword"
Option Compare Database
Option Explicit

Public Function RandomPassword(numchar As Integer) As String:
    ' Purpose:	    Generate a random strong password
    ' Return:	    String
    ' Arguments:	numchar = Desired password length
    
    Dim specialCharacters As String ' Set of special characters to be used
    Dim uCase1 As String            ' Upper case letter
    Dim uCase2 As String            ' Upper case letter
    Dim lCase1 As String            ' Lower case letter
    Dim lCase2 As String            ' Lower case letter
    Dim int1 As Integer             ' Numeral
    Dim int2 As Integer             ' Numberal
    Dim sc1 As String               ' Special character
    Dim sc2 As String               ' Special character
    Dim strCompile As String        ' Concatenated string
    
    specialCharacters = "!@#$%^&*()"
    strCompile = ""
    
    While Len(strCompile) < numchar:
    ' Generate random characters until the required length is exceeded
        uCase1 = UCase(Chr(RandBetween(65, 90)))
        uCase2 = UCase(Chr(RandBetween(65, 90)))
        lCase1 = LCase(Chr(RandBetween(97, 122)))
        lCase2 = LCase(Chr(RandBetween(97, 122)))
        int1 = RandBetween(0, 9)
        int2 = RandBetween(0, 9)
        sc1 = Mid(specialCharacters, RandBetween(1, Len(specialCharacters)), 1)
        sc2 = Mid(specialCharacters, RandBetween(1, Len(specialCharacters)), 1)
        strCompile = strCompile & uCase1 & uCase2 & lCase1 & lCase2 & int1 & int2 & sc1 & sc2
        ' Concatenate the randomly generated characters into a string
    Wend
        
    strCompile = Shuffle(strCompile)        ' Shuffle the string
    strCompile = Left(strCompile, numchar)  ' Trim excess characters from the new password
    RandomPassword = strCompile
End Function

Public Function RandBetween(lowerband As Long, upperband As Long) As Long:
    ' Purpose:	    Generate a random number between two values
    ' Return:	    Long
    ' Arguments:	lowerband = minimum value
    '               upperband = maximum value
        
  RandBetween = Int((upperband - lowerband + 1) * Rnd + lowerband)
End Function

Public Function Shuffle(strInput As String) As String:
  ' Purpose:	    Shuffle a string into random order
  ' Return:	    String
  ' Arguments:	strInput = String to be shuffled
        
  Dim strOutput As String   ' Shuffled string
  Dim i As Integer          ' counter
  
  strOutput = ""
  
  While Len(strInput) > 0:
    i = RandBetween(1, Len(strInput) + 1)
    strOutput = strOutput & Mid(strInput, i, 1)
    strInput = Replace(strInput, Mid(strInput, i, 1), "", 1, 1)
  Wend
            
  Shuffle = strOutput
End Function
