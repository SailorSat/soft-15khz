Attribute VB_Name = "Module1"
Option Explicit

Function LeadZero(Data As String, Length As Integer) As String
  Dim Total As Integer
  Total = Length - Len(Data)
  LeadZero = String(Total, "0") & Data
End Function

Function HexToBinary(HexLine As String) As String
  Dim BinaryString As String
  Dim Index As Integer
  
  BinaryString = ""
  Index = 1
  
  While Index < Len(HexLine)
    BinaryString = BinaryString & Chr("&H" & Mid(HexLine, Index, 2))
    Index = Index + 2
  Wend
  
  HexToBinary = BinaryString
End Function

Function BinaryToHex(Binaryline As String) As String
  Dim HexString As String
  Dim Index As Integer
  Dim Dummy As String
  
  HexString = ""
  Index = 1
  
  While Index <= Len(Binaryline)
    Dummy = Hex(Asc(Mid(Binaryline, Index, 1)))
    While Len(Dummy) < 2
      Dummy = "0" & Dummy
    Wend
    
    HexString = HexString & Dummy
    Index = Index + 1
  Wend
  
  BinaryToHex = HexString
End Function

Public Function HexToBits(Hex As String)
  Dim Index As Integer
  Dim Bits As String
  Dim Dummy As String
  For Index = 1 To Len(Hex)
    Dummy = Mid(Hex, Index, 1)
    Select Case Dummy
      Case "0"
        Dummy = "0000"
      Case "1"
        Dummy = "0001"
      Case "2"
        Dummy = "0010"
      Case "3"
        Dummy = "0011"
      Case "4"
        Dummy = "0100"
      Case "5"
        Dummy = "0101"
      Case "6"
        Dummy = "0110"
      Case "7"
        Dummy = "0111"
      Case "8"
        Dummy = "1000"
      Case "9"
        Dummy = "1001"
      Case "A"
        Dummy = "1010"
      Case "B"
        Dummy = "1011"
      Case "C"
        Dummy = "1100"
      Case "D"
        Dummy = "1101"
      Case "E"
        Dummy = "1110"
      Case "F"
        Dummy = "1111"
    End Select
    Bits = Bits & Dummy
  Next
  HexToBits = Bits
End Function

Public Function BitsToHex(Bits As String)
  Dim Index As Integer
  Dim Hex As String
  Dim Dummy As String
  While Not Len(Bits) Mod 4 = 0
    Bits = "0" & Bits
  Wend
  For Index = 1 To Len(Bits) Step 4
    Dummy = Mid(Bits, Index, 4)
    Select Case Dummy
      Case "0000"
        Dummy = "0"
      Case "0001"
        Dummy = "1"
      Case "0010"
        Dummy = "2"
      Case "0011"
        Dummy = "3"
      Case "0100"
        Dummy = "4"
      Case "0101"
        Dummy = "5"
      Case "0110"
        Dummy = "6"
      Case "0111"
        Dummy = "7"
      Case "1000"
        Dummy = "8"
      Case "1001"
        Dummy = "9"
      Case "1010"
        Dummy = "A"
      Case "1011"
        Dummy = "B"
      Case "1100"
        Dummy = "C"
      Case "1101"
        Dummy = "D"
      Case "1110"
        Dummy = "E"
      Case "1111"
        Dummy = "F"
    End Select
    Hex = Hex & Dummy
  Next
  BitsToHex = Hex
End Function

Public Function Intel_GMA_ModelineToBinary(ResName As String, Modeline As String) As String
  Dim LineBinary As String
  Dim LineParam() As String

  Dim Dummy As String
  Dim DummyX As String
  Dim DummyY As String
  Dim Index As Integer

  Dim P_FREQ As Single
  Dim H_ACTIVE As Integer
  Dim H_FIRST As Integer
  Dim H_LAST As Integer
  Dim H_TOTAL As Integer
  Dim V_ACTIVE As Integer
  Dim V_FIRST As Integer
  Dim V_LAST As Integer
  Dim V_TOTAL As Integer
  Dim M_OPTIONS As Integer

  LineBinary = ""
  LineParam = Split(Modeline, " ", 12)

  P_FREQ = CSng(LineParam(2))
  If InStr(1, LCase(LineParam(11)), "interlace") Then
    M_OPTIONS = M_OPTIONS + 2
  End If
  If InStr(1, LCase(LineParam(11)), "-hsync") Then
    M_OPTIONS = M_OPTIONS + 4
  End If
  If InStr(1, LCase(LineParam(11)), "-vsync") Then
    M_OPTIONS = M_OPTIONS + 8
  End If
  H_ACTIVE = LineParam(3)
  H_FIRST = LineParam(4)
  H_LAST = LineParam(5)
  H_TOTAL = LineParam(6)
  V_ACTIVE = LineParam(7)
  V_FIRST = LineParam(8)
  V_LAST = LineParam(9)
  V_TOTAL = LineParam(10)

  If UCase(LineParam(0)) = "MODELINE" Then
    LineBinary = ""
    'P_FREQ
    Dummy = LeadZero(Hex(CLng(P_FREQ * 100)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'H_ACTIVE
    '12 Bit
    Dummy = LeadZero(Hex(H_ACTIVE), 3)
    DummyX = Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    LineBinary = LineBinary & Dummy

    'H_TOTAL - H_ACTIVE
    '12 Bit
    Dummy = LeadZero(Hex(H_TOTAL - H_ACTIVE), 3)
    DummyX = DummyX & Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    LineBinary = LineBinary & Dummy

    'DummyX
    LineBinary = LineBinary & DummyX

    'V_ACTIVE
    '12 Bit
    If M_OPTIONS And 2 Then
      Dummy = LeadZero(Hex(V_ACTIVE / 2), 3)
    Else
      Dummy = LeadZero(Hex(V_ACTIVE), 3)
    End If
    DummyX = Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    LineBinary = LineBinary & Dummy

    'V_TOTAL - V_ACTIVE
    '12 Bit
    If M_OPTIONS And 2 Then
      Dummy = LeadZero(Hex((V_TOTAL - V_ACTIVE) / 2), 3)
    Else
      Dummy = LeadZero(Hex(V_TOTAL - V_ACTIVE), 3)
    End If
    DummyX = DummyX & Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    LineBinary = LineBinary & Dummy

    'DummyX
    LineBinary = LineBinary & DummyX

    'H_FIRST - H_ACTIVE
    '10 Bit
    Dummy = LeadZero(Hex(H_FIRST - H_ACTIVE), 3)
    DummyX = Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    DummyY = Right(HexToBits(DummyX), 2)
    LineBinary = LineBinary & Dummy

    'H_LAST - H_FIRST
    '10 Bit
    Dummy = LeadZero(Hex(H_LAST - H_FIRST), 3)
    DummyX = Left(Dummy, 1)
    Dummy = Mid(Dummy, 2)
    DummyY = DummyY & Right(HexToBits(DummyX), 2)
    LineBinary = LineBinary & Dummy

    'V_FIRST - V_ACTIVE
    '6 Bit
    If M_OPTIONS And 2 Then
      Dummy = LeadZero(Hex(CInt((V_FIRST - V_ACTIVE) / 2)), 3)
    Else
      Dummy = LeadZero(Hex(V_FIRST - V_ACTIVE), 3)
    End If
    DummyX = Right(HexToBits(Dummy), 6)
    DummyY = DummyY & Left(DummyX, 2)
    LineBinary = LineBinary & BitsToHex(Right(DummyX, 4))

    'V_LAST - V_FIRST
    '6 Bit
    If M_OPTIONS And 2 Then
      Dummy = LeadZero(Hex((V_LAST - V_FIRST) / 2), 3)
    Else
      Dummy = LeadZero(Hex(V_LAST - V_FIRST), 3)
    End If
    DummyX = Right(HexToBits(Dummy), 6)
    DummyY = DummyY & Left(DummyX, 2)
    LineBinary = LineBinary & BitsToHex(Right(DummyX, 4))

    'DummyY
    LineBinary = LineBinary & BitsToHex(DummyY)

    'NOT_USED
    LineBinary = LineBinary & String(10, "0")

    'M_OPTIONS
    If M_OPTIONS And 2 Then
      Dummy = "1"
    Else
      Dummy = "0"
    End If
    Dummy = Dummy & "0011"
    If M_OPTIONS And 4 Then
      Dummy = Dummy & "0"
    Else
      Dummy = Dummy & "1"
    End If
    If M_OPTIONS And 8 Then
      Dummy = Dummy & "0"
    Else
      Dummy = Dummy & "1"
    End If
    Dummy = Dummy & "0"
    LineBinary = LineBinary & BitsToHex(Dummy)
  End If
  Intel_GMA_ModelineToBinary = LineBinary
End Function

Public Function Intel_GMA_BinaryToModeline(ResName As String, Intel As String) As String
  Dim Dummy As String

  Dim P_FREQ As Single
  Dim H_ACTIVE As Integer
  Dim H_FIRST As Integer
  Dim H_LAST As Integer
  Dim H_TOTAL As Integer
  Dim V_ACTIVE As Integer
  Dim V_FIRST As Integer
  Dim V_LAST As Integer
  Dim V_TOTAL As Integer
  Dim M_OPTIONS As Integer

  Dim V_FREQ As Single

  P_FREQ = CSng("&H" & Mid(Intel, 3, 2) & Mid(Intel, 1, 2)) / 100

  H_ACTIVE = CInt("&H" & Mid(Intel, 9, 1) & Mid(Intel, 5, 2))
  H_TOTAL = H_ACTIVE + CInt("&H" & Mid(Intel, 10, 1) & Mid(Intel, 7, 2))

  V_ACTIVE = CInt("&H" & Mid(Intel, 15, 1) & Mid(Intel, 11, 2))
  V_TOTAL = V_ACTIVE + CInt("&H" & Mid(Intel, 16, 1) & Mid(Intel, 13, 2))

  H_FIRST = H_ACTIVE + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 1, 2) & HexToBits(Mid(Intel, 17, 2))))
  H_LAST = H_FIRST + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 3, 2) & HexToBits(Mid(Intel, 19, 2))))

  V_FIRST = V_ACTIVE + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 5, 2) & Mid(HexToBits(Mid(Intel, 21, 2)), 1, 4)))
  V_LAST = V_FIRST + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 7, 2) & Mid(HexToBits(Mid(Intel, 21, 2)), 5, 4)))

  Dummy = HexToBits(Mid(Intel, 35, 2))
  If Mid(Dummy, 1, 1) = "1" Then
    M_OPTIONS = M_OPTIONS + 2
  End If
  If Mid(Dummy, 6, 1) = "0" Then
    M_OPTIONS = M_OPTIONS + 4
  End If
  If Mid(Dummy, 7, 1) = "0" Then
    M_OPTIONS = M_OPTIONS + 8
  End If

  If M_OPTIONS And 2 Then
    V_ACTIVE = V_ACTIVE * 2
    V_TOTAL = V_TOTAL * 2
    V_FIRST = V_FIRST * 2
    V_LAST = V_LAST * 2
    V_FREQ = Round((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * (CSng(V_TOTAL) / 2)), 3)
  Else
    V_FREQ = Round((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL)), 3)
  End If

  Dummy = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "@" & V_FREQ & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
  If M_OPTIONS And 1 Then
    Dummy = Dummy & " doublescan"
  End If
  If M_OPTIONS And 2 Then
    Dummy = Dummy & " interlace"
  End If
  If M_OPTIONS And 4 Then
    Dummy = Dummy & " -hsync"
  End If
  If M_OPTIONS And 8 Then
    Dummy = Dummy & " -vsync"
  End If
  Intel_GMA_BinaryToModeline = Dummy
End Function

