Attribute VB_Name = "fxRandomPassword"
Option Compare Database
Option Explicit

Public Function RandomPassword(numchar As Integer) As String:
    Dim specialCharacters As String
    Dim uCase1 As String
    Dim uCase2 As String
    Dim lCase1 As String
    Dim lCase2 As String
    Dim int1 As Integer
    Dim int2 As Integer
    Dim sc1 As String
    Dim sc2 As String
    Dim strCompile As String
    
    specialCharacters = "!@#$%^&*()"
    strCompile = ""
    
    While Len(strCompile) < numchar:
        uCase1 = UCase(Chr(RandBetween(65, 90)))
        uCase2 = UCase(Chr(RandBetween(65, 90)))
        lCase1 = LCase(Chr(RandBetween(97, 122)))
        lCase2 = LCase(Chr(RandBetween(97, 122)))
        int1 = RandBetween(0, 9)
        int2 = RandBetween(0, 9)
        sc1 = Mid(specialCharacters, RandBetween(1, Len(specialCharacters)), 1)
        sc2 = Mid(specialCharacters, RandBetween(1, Len(specialCharacters)), 1)
        strCompile = strCompile & uCase1 & uCase2 & lCase1 & lCase2 & int1 & int2 & sc1 & sc2
    Wend
    strCompile = Shuffle(strCompile)
    strCompile = Left(strCompile, numchar)
    RandomPassword = strCompile
End Function

Public Function RandBetween(lowerband As Long, upperband As Long) As Long:
  RandBetween = Int((upperband - lowerband + 1) * Rnd + lowerband)
End Function

Public Function Shuffle(strInput As String) As String:
  Dim strOutput As String
  Dim i As Integer
  
  strOutput = ""
  
  While Len(strInput) > 0:
    i = RandBetween(1, Len(strInput) + 1)
    strOutput = strOutput & Mid(strInput, i, 1)
    strInput = Replace(strInput, Mid(strInput, i, 1), "", 1, 1)
  Wend
  Shuffle = strOutput
End Function
