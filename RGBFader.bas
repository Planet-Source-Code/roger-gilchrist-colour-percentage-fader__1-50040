Attribute VB_Name = "RGBFader2Mod"
Option Explicit
'Type used by ToRGB to break long colours int RGB
Type RGB_Type
  R As Long
  G As Long
  B As Long
End Type

Public Function Percent(a As Variant, _
                        ByVal B As Variant) As Double

  'calculate percentage given two numbers(using Variant allows you st send any type
  'Provided for this demo
  'you probably have your own percent function already
  Dim c As Single

  If a > B Then
    a = B
  End If
  If B = 0 Then
    Percent = 0
    Exit Function
  End If
  c = Int(a / B * 100)
  If c > 100 Then
    c = 100
  End If
  Percent = c

End Function

Public Function RGBPercentage(lngStartColor As Long, _
                              lngEndColor As Long, _
                              ByVal phase As Single) As Long

  'generate a colour that is a certain percentage between the 2 input colours
  
  Dim Scols As RGB_Type
  Dim Ecols As RGB_Type
'break the colours up
  Scols = ToRGB(lngStartColor)
  Ecols = ToRGB(lngEndColor)
'generate the new colour
  With Scols
    RGBPercentage = RGB(.R + ((Ecols.R - .R) * phase / 100), .G + ((Ecols.G - .G) * phase / 100), .B + ((Ecols.B - .B) * phase / 100))
  End With 'Scols

End Function

Private Function ToRGB(ByVal LngColor As Long) As RGB_Type
'break a long colour value into its RGB parts
'there are lots of ways to do this (many faster than this)
'this is just the one I had stored in my templates folder
  Dim ColorStr As String

  ColorStr = Right$("000000" & Hex$(LngColor), 6)
  With ToRGB
    .R = Val("&h" & Right$(ColorStr, 2))
    .G = Val("&h" & Mid$(ColorStr, 3, 2))
    .B = Val("&h" & Left$(ColorStr, 2))
  End With

End Function

':)Roja's VB Code Fixer V1.1.59 (22/11/2003 1:05:02 PM) 7 + 52 = 59 Lines Thanks Ulli for inspiration and lots of code.

