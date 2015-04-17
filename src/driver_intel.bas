Attribute VB_Name = "driver_intel"
' -= Soft-15kHz - DRIVER - Intel GMA / EGD
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


' -= Intel GMA =-
Sub Intel_GMA_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  Dim Key() As String
  Dim Index As Integer
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  'Modes
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "TotalDTDCount", 0
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "TotalStaticModes", 0
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "___Soft-15KHz_Modes___"
  'Definitions
  Key = GetAllValues(HKEY_LOCAL_MACHINE, CurObjectPath)
  For Index = LBound(Key) To UBound(Key)
    If Left(LCase(Key(Index)), 4) = "dtd_" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Key(Index)
    End If
    If Left(LCase(Key(Index)), 12) = "static_mode_" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Key(Index)
    End If
  Next
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub Intel_GMA_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim TotalDTDCount As Long
  Dim Index As Integer
  Dim Res() As String
  Dim ResIndex As Integer
  Dim ResCount As Integer
  Dim Soft15KHz_Modes As String
  Dim Dummy As String
  Dim Dummy2 As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Soft15KHz_Modes = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "___Soft-15KHz_Modes___", ""), Chr(0), "")
  TotalDTDCount = GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "TotalDTDCount", 0)
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        If Not Intel_GMA_ModelineToBinary(.ModeName, .Modeline) = "" Then
          ResIndex = 0
          If InStr(1, Soft15KHz_Modes, .ModeName, vbBinaryCompare) Then
            Res = Split(Soft15KHz_Modes, ";")
            For ResCount = 0 To UBound(Res)
              If Res(ResCount) = .ModeName Then
                ResIndex = ResCount
                Exit For
              End If
            Next
          Else
            Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
            If Adapter(AdapterIndex).ModeCount = 6 Then If Not Host.RuntimeFlags And 2 Then MsgBox "Intel drivers don't allow more than 5 resolution defintions.", vbInformation + vbOKOnly, "Soft-15kHz"
          End If
          If ResIndex = 0 Then
            TotalDTDCount = TotalDTDCount + 1
            ResIndex = TotalDTDCount
            Soft15KHz_Modes = Soft15KHz_Modes & ";" & .ModeName
          End If
          SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DTD_" & ResIndex, HexToBinary(Intel_GMA_ModelineToBinary(.ModeName, .Modeline))
          'Static Modes Test...
          Res = Split(.ModeName, ",")
          'X
          Dummy = Hex(Res(0))
          While Len(Dummy) < 4
            Dummy = "0" & Dummy
          Wend
          Dummy2 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
          'Y
          Dummy = Hex(Res(1))
          While Len(Dummy) < 4
            Dummy = "0" & Dummy
          Wend
          Dummy2 = Dummy2 & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
          'Refresh etc.
          Dummy2 = Dummy2 & "0100" '60hz only
          Dummy2 = Dummy2 & "07" ' 8 + 16 + 32 Bit
          Dummy2 = Dummy2 & "0F" ' don't know
          SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "STATIC_MODE_" & ResIndex, HexToBinary(Dummy2)
        End If
      End If
    End With
  Next
  SaveSettingString HKEY_LOCAL_MACHINE, CurObjectPath, "___Soft-15KHz_Modes___", Soft15KHz_Modes
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "TotalDTDCount", TotalDTDCount
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "TotalStaticModes", TotalDTDCount
End Sub

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

    'FLAGS (don't know!)
    LineBinary = LineBinary & "3701"
  End If
  Intel_GMA_ModelineToBinary = LineBinary
End Function

''Public Function Intel_GMA_BinaryToModeline(ResName As String, Intel As String) As String
''  Dim Dummy As String
''
''  Dim P_FREQ As Single
''  Dim H_ACTIVE As Integer
''  Dim H_FIRST As Integer
''  Dim H_LAST As Integer
''  Dim H_TOTAL As Integer
''  Dim V_ACTIVE As Integer
''  Dim V_FIRST As Integer
''  Dim V_LAST As Integer
''  Dim V_TOTAL As Integer
''  Dim M_OPTIONS As Integer
''
''  Dim V_FREQ As Single
''
''  P_FREQ = CSng("&H" & Mid(Intel, 3, 2) & Mid(Intel, 1, 2)) / 100
''
''  H_ACTIVE = CInt("&H" & Mid(Intel, 9, 1) & Mid(Intel, 5, 2))
''  H_TOTAL = H_ACTIVE + CInt("&H" & Mid(Intel, 10, 1) & Mid(Intel, 7, 2))
''
''  V_ACTIVE = CInt("&H" & Mid(Intel, 15, 1) & Mid(Intel, 11, 2))
''  V_TOTAL = V_ACTIVE + CInt("&H" & Mid(Intel, 16, 1) & Mid(Intel, 13, 2))
''
''  H_FIRST = H_ACTIVE + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 1, 2) & HexToBits(Mid(Intel, 17, 2))))
''  H_LAST = H_FIRST + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 3, 2) & HexToBits(Mid(Intel, 19, 2))))
''
''  V_FIRST = V_ACTIVE + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 5, 2) & Mid(HexToBits(Mid(Intel, 21, 2)), 1, 4)))
''  V_LAST = V_FIRST + CInt("&H" & BitsToHex(Mid(HexToBits(Mid(Intel, 23, 1)), 7, 2) & Mid(HexToBits(Mid(Intel, 21, 2)), 5, 4)))
''
''  Dummy = HexToBits(Mid(Intel, 35, 2))
''  If Mid(Dummy, 1, 1) = "1" Then
''    M_OPTIONS = M_OPTIONS + 2
''  End If
''  If Mid(Dummy, 6, 1) = "0" Then
''    M_OPTIONS = M_OPTIONS + 4
''  End If
''  If Mid(Dummy, 7, 1) = "0" Then
''    M_OPTIONS = M_OPTIONS + 8
''  End If
''
''  If M_OPTIONS And 2 Then
''    V_ACTIVE = V_ACTIVE * 2
''    V_TOTAL = V_TOTAL * 2
''    V_FIRST = V_FIRST * 2
''    V_LAST = V_LAST * 2
''    V_FREQ = Round((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * (CSng(V_TOTAL) / 2)), 3)
''  Else
''    V_FREQ = Round((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL)), 3)
''  End If
''
''  Dummy = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "@" & V_FREQ & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
''  If M_OPTIONS And 1 Then
''    Dummy = Dummy & " doublescan"
''  End If
''  If M_OPTIONS And 2 Then
''    Dummy = Dummy & " interlace"
''  End If
''  If M_OPTIONS And 4 Then
''    Dummy = Dummy & " -hsync"
''  End If
''  If M_OPTIONS And 8 Then
''    Dummy = Dummy & " -vsync"
''  End If
''  Intel_GMA_BinaryToModeline = Dummy
''End Function


' -= Intel EGD =-
Sub Intel_EGD_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  Dim Dummy As String
  Dim ConfigID As Long
  Dim PortID As Long
  Dim PortOrder As String
  Dim IndexA As Integer
  Dim IndexB As Integer
  Dim IndexC As Integer
  Dim Key() As String
  Dim Value() As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  ConfigID = GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "ConfigId", 0)
  PortOrder = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\General", "PortOrder", ""), Chr(0), "")
  For IndexA = 1 To Len(PortOrder)
    PortID = CLng(Mid(PortOrder, IndexA, 1))
    If Not PortID = 0 Then
      Key = GetAllKeys(HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD")
      If Not Key(0) = "" Then
        For IndexB = LBound(Key) To UBound(Key)
          Value = GetAllValues(HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD\" & Key(IndexB))
          For IndexC = LBound(Value) To UBound(Value)
            DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD\" & Key(IndexB), Value(IndexC)
          Next IndexC
          DeleteKey HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD\" & Key(IndexB)
          DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD", "___Soft-15KHz_Modes___"
          DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\General", "Edid"
          DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\General", "EdidAvail"
          DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\General", "EdidNotAvail"
        Next
      End If
    End If
  Next
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub Intel_EGD_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim DummyA As String
  Dim DummyB As String
  Dim ConfigID As Long
  Dim PortID As Long
  Dim PortOrder As String
  Dim IndexA As Integer
  Dim IndexB As Integer
  Dim IndexC As Integer
  
  Dim Res() As String
  Dim ResIndex As Integer
  Dim ResCount As Integer
  Dim Soft15KHz_Modes As String
  
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  ConfigID = GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "ConfigId", 0)
  PortOrder = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\General", "PortOrder", ""), Chr(0), "")
  For IndexA = 1 To Len(PortOrder)
    PortID = CLng(Mid(PortOrder, IndexA, 1))
    If Not PortID = 0 Then
      CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\DTD"
      Soft15KHz_Modes = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath & "\ALL\" & ConfigID & "\Port\" & PortID & "\DTD", "___Soft-15KHz_Modes___", ""), Chr(0), "")
      Res = Split(Soft15KHz_Modes, ";")
      IndexC = UBound(Res)
      If IndexC = -1 Then IndexC = 0
      For IndexB = 1 To ModeCount
        With Mode(IndexB)
          If .ModeFreq = Freq Then
            If Not Intel_EGD_ModelineToStrings(.ModeName, .Modeline) = "" Then
              ResIndex = 0
              If InStr(1, Soft15KHz_Modes, .ModeName, vbBinaryCompare) Then
                Res = Split(Soft15KHz_Modes, ";")
                For ResCount = 0 To UBound(Res)
                  If Res(ResCount) = .ModeName Then
                    ResIndex = ResCount
                    Exit For
                  End If
                Next
              Else
                Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
                If Adapter(AdapterIndex).ModeCount = 6 Then If Not Host.RuntimeFlags And 2 Then MsgBox "Intel drivers don't allow more than 5 resolution defintions.", vbInformation + vbOKOnly, "Soft-15kHz"
              End If
              If ResIndex = 0 Then
                IndexC = IndexC + 1
                ResIndex = IndexC
                Soft15KHz_Modes = Soft15KHz_Modes & ";" & .ModeName
              End If
              CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\DTD\" & ResIndex
              Res = Split(Intel_EGD_ModelineToStrings(.ModeName, .Modeline), ";")
              For ResCount = 0 To UBound(Res)
                DummyA = Left(Res(ResCount), InStr(1, Res(ResCount), "=", vbBinaryCompare) - 1)
                DummyB = Mid(Res(ResCount), InStr(1, Res(ResCount), "=", vbBinaryCompare) + 1)
                SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\DTD\" & ResIndex, DummyA, CLng(DummyB)
              Next
            End If
          End If
        End With
      Next
      SaveSettingString HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\DTD", "___Soft-15KHz_Modes___", Soft15KHz_Modes
      SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\General", "Edid", 0
      SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\General", "EdidAvail", 0
      SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath & "\Config\" & ConfigID & "\Port\" & PortID & "\General", "EdidNotAvail", 4
    End If
  Next
End Sub

Function Intel_EGD_ModelineToStrings(ResName As String, Modeline As String) As String
  Dim LineBinary As String
  Dim LineParam() As String

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
  
  Dim FLAGS As Long
  
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
    'P_FREQ
    LineBinary = "PixelClock=" & CLng(P_FREQ * 1000)
  
    'H_ACTIVE
    LineBinary = LineBinary & ";HorzActive=" & CInt(H_ACTIVE)
  
    'H_TOTAL - H_ACTIVE
    LineBinary = LineBinary & ";HorzBlank=" & CInt(H_TOTAL - H_ACTIVE)
  
    'H_FIRST - H_ACTIVE
    LineBinary = LineBinary & ";HorzSync=" & CInt(H_FIRST - H_ACTIVE)

    'H_TOTAL - H_LAST
    LineBinary = LineBinary & ";HorzSyncPulse=" & CInt(H_TOTAL - H_LAST)
  
    'V_ACTIVE
    LineBinary = LineBinary & ";VertActive=" & CInt(V_ACTIVE)
  
    'V_TOTAL - V_ACTIVE
    LineBinary = LineBinary & ";VertBlank=" & CInt(V_TOTAL - V_ACTIVE)
  
    'V_FIRST - V_ACTIVE
    LineBinary = LineBinary & ";VertSync=" & CInt(V_FIRST - V_ACTIVE)

    'V_TOTAL - V_LAST
    LineBinary = LineBinary & ";VertSyncPulse=" & CInt(V_TOTAL - V_LAST)

    'M_OPTIONS
    FLAGS = 0
    If M_OPTIONS And 2 Then
      FLAGS = FLAGS Or -2147483648#
    End If
    If Not M_OPTIONS And 4 Then
      FLAGS = FLAGS Or 67108864
    End If
    If Not M_OPTIONS And 8 Then
      FLAGS = FLAGS Or 134217728
    End If
    LineBinary = LineBinary & ";Flags=" & FLAGS
    Intel_EGD_ModelineToStrings = LineBinary
  End If
End Function
