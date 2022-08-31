Attribute VB_Name = "fxEncrypt"
Option Explicit

Public Const Key As String = "Secret Phrase"

Public Function Encrypt(strText As String)
  Dim bText() As Byte
  Dim bKey() As Byte
  Dim lText As Long
  Dim lKey As Long
  Dim lTextPos As Long
  Dim lKeyPos As Long
  
  bText = StrConv(strText, vbFromUnicode)
  bKey = StrConv(Key, vbFromUnicode)
  lText = UBound(bText)
  lKey = UBound(bKey)
  For lTextPos = 0 To lText
    bText(lTextPos) = bText(lTextPos) Xor bKey(lKeyPos)
    If lKeyPos < lKey Then
      lKeyPos = lKeyPos + 1
    Else
      lKeyPos = 0
    End If
  Next lTextPos
  Encrypt = StrConv(bText, vbUnicode)
End Function
