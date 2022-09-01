Attribute VB_Name = "fxEncrypt"
Option Explicit

' Encryption key to be used
Public Const Key As String = "Secret Phrase"

Public Function Encrypt(strText As String)
' Purpose:	Encrypt input string
' Return:		String
' Arguments:	strText = Text string to be encrypted

  Dim bText() As Byte   ' Bytes for storing input text
  Dim bKey() As Byte    ' Bytes for storing encryption key
  Dim lText As Long     ' Long value of input text
  Dim lKey As Long      ' Long value of encryption key
  Dim lTextPos As Long  ' Input text position counter
  Dim lKeyPos As Long   ' Encryption key position counter
  
  ' Convert to bytes
  bText = StrConv(strText, vbFromUnicode)
  bKey = StrConv(Key, vbFromUnicode)
  
  ' Convert to long
  lText = UBound(bText)
  lKey = UBound(bKey)
  
  ' XOR Comparison of bits
  For lTextPos = 0 To lText
    bText(lTextPos) = bText(lTextPos) Xor bKey(lKeyPos)
    If lKeyPos < lKey Then
      lKeyPos = lKeyPos + 1
    Else
      lKeyPos = 0
    End If
  Next lTextPos

' Convert byte output to Unicode
  Encrypt = StrConv(bText, vbUnicode)
End Function
