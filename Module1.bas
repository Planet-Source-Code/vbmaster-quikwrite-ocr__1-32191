Attribute VB_Name = "Module1"
Public Type DWORD
    low As Integer
    high As Integer
End Type
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Enum StretchBltModes
  BLACKONWHITE = 1
  WHITEONBLACK = 2
  COLORONCOLOR = 3
  HALFTONE = 4
  MAXSTRETCHBLTMODE = 4
End Enum
Public Const SRCCOPY = &HCC0020
Const CharacterWidth As Long = 30
Const CharacterHeight As Long = 20




Function GetTrueExtents(TextExtents As DWORD) As RECT
Dim i As Long
Dim j As Long
Dim Colour As Long
Dim RealTextExtent As RECT

  With RealTextExtent
    .Left = -1
    .Top = -1
    .Right = -1
    .Bottom = -1
  End With
With Form1
  ' Top
  For i = 0 To TextExtents.high
    For j = 0 To TextExtents.low
      Colour = GetPixel(.pict.hdc, j, i)
      If Colour <> &HFFFFFF And Colour <> -1 Then
        RealTextExtent.Top = i
        Exit For
      End If
    Next j
    If RealTextExtent.Top <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Top = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If
  ' Left
  For i = 0 To TextExtents.low
    For j = RealTextExtent.Top To TextExtents.high
      Colour = GetPixel(.pict.hdc, i, j)
      If Colour <> &HFFFFFF And Colour <> -1 Then
        RealTextExtent.Left = i
        Exit For
      End If
    Next j
    If RealTextExtent.Left <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Left = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  ' Right
  For i = TextExtents.low To RealTextExtent.Left Step -1
    For j = RealTextExtent.Top To TextExtents.high
      Colour = GetPixel(.pict.hdc, i, j)
      If Colour <> -1 And Colour <> &HFFFFFF Then
        RealTextExtent.Right = i
        Exit For
      End If
    Next j
    If RealTextExtent.Right <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Right = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  ' Bottom
  For i = TextExtents.high To RealTextExtent.Top Step -1
    For j = RealTextExtent.Left To TextExtents.low
      Colour = GetPixel(.pict.hdc, j, i)
      If Colour <> -1 And Colour <> &HFFFFFF Then
        RealTextExtent.Bottom = i
        Exit For
      End If
    Next j
    If RealTextExtent.Bottom <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Bottom = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  GetTrueExtents = RealTextExtent
End With
End Function

