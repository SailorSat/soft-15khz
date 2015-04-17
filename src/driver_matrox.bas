Attribute VB_Name = "driver_matrox"
' -= Soft-15kHz - DRIVER - Matrox PowerDesk
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


Sub Matrox_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  Dim Key() As String
  Dim Index As Integer
  Dim Dummy As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "Mga.SingleResolutions", ""
  Key = GetAllValues(HKEY_LOCAL_MACHINE, CurObjectPath)
  For Index = LBound(Key) To UBound(Key)
    If Left(LCase(Key(Index)), 8) = "graphic." Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Key(Index)
    End If
  Next
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub Matrox_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim Mga_SingleResolutions As String
  Dim Dummy As String
  Dim Dummy2 As String
  Dim Index As Integer
  Dim Res() As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Mga.SingleResolutions", "")
  Mga_SingleResolutions = BinaryToHex(Dummy)
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        Res = Split(.ModeName, ",")
        If Not Res(0) Mod 8 = 0 Then
          Res(0) = (Res(0) + 8) - (Res(0) Mod 8)
        End If
        Dummy = Matrox_ResolutionToBinary(CInt(Res(0)), CInt(Res(1)))
        If InStr(1, Mga_SingleResolutions, Dummy, vbBinaryCompare) = 0 Then
          Mga_SingleResolutions = Mga_SingleResolutions & Dummy
          Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
          Dummy2 = ""
        Else
          Dummy2 = BinaryToHex(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Graphic." & Res(0) & "." & Res(1), ""))
        End If
        Dummy = Matrox_ModelineToBinary(.ModeName, .Modeline)
        SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "Graphic." & Res(0) & "." & Res(1), HexToBinary(Dummy2 & Dummy)
      End If
    End With
  Next
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "Mga.SingleResolutions", HexToBinary(Mga_SingleResolutions)
End Sub

Function Matrox_ResolutionToBinary(X As Integer, Y As Integer) As String
  Dim Dummy As String
  Dummy = Hex(X)
  While Len(Dummy) < 4
    Dummy = "0" & Dummy
  Wend
  Matrox_ResolutionToBinary = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
  Dummy = Hex(Y)
  While Len(Dummy) < 4
    Dummy = "0" & Dummy
  Wend
  Matrox_ResolutionToBinary = Matrox_ResolutionToBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
End Function

Function Matrox_ModelineToBinary(ResName As String, Modeline As String) As String
  Dim LineBinary As String
  Dim LineParam() As String
  Dim Dummy As String
  
  Dim P_FREQ As Single
  Dim M_OPTIONS As Integer
  Dim H_ACTIVE As Integer
  Dim H_FIRST As Integer
  Dim H_LAST As Integer
  Dim H_TOTAL As Integer
  Dim V_ACTIVE As Integer
  Dim V_FIRST As Integer
  Dim V_LAST As Integer
  Dim V_TOTAL As Integer
  
  Dim V_FREQ As Long
  Dim H_FREQ As Long
  
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
  
  Dim Res() As String
  Res = Split(ResName, ",")
  V_ACTIVE = Res(1)
  
  '-V_TOTAL must be multiple of 4-
  V_TOTAL = V_TOTAL - (V_TOTAL Mod 4)
  
  V_FREQ = Round(((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL))) * 1000, 3)
  H_FREQ = Round((1000000 * CSng(P_FREQ)) / CSng(H_TOTAL), 2)
  
  '-interlace workaround-
  If M_OPTIONS And 2 Then
    V_ACTIVE = V_ACTIVE / 2
    V_FIRST = V_FIRST / 2
    V_LAST = V_LAST / 2
    V_TOTAL = V_TOTAL / 2
  End If
  
    
  If UCase(LineParam(0)) = "MODELINE" Then
    LineBinary = ""
    'V_FREQ
    Dummy = LeadZero(Hex(CInt(V_FREQ / 1000)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_FREQ
    Dummy = LeadZero(Hex(CLng(H_FREQ / 1000)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'P_FREQ
    Dummy = LeadZero(Hex(CLng(P_FREQ * 1000)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'NOT_USED
    LineBinary = LineBinary & String(4, 48)
    'H_FIRST - H_ACTIVE
    Dummy = LeadZero(Hex(H_FIRST - H_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_LAST - H_FIRST (evtl. verdreht? ->)
    Dummy = LeadZero(Hex(H_LAST - H_FIRST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_TOTAL - H_LAST (<- evtl. verdreht?)
    Dummy = LeadZero(Hex(H_TOTAL - H_LAST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_FIRST - V_ACTIVE
    Dummy = LeadZero(Hex(V_FIRST - V_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_LAST - V_FIRST (evtl. verdreht? ->)
    Dummy = LeadZero(Hex(V_LAST - V_FIRST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_TOTAL - V_LAST (<- evtl. verdreht?)
    Dummy = LeadZero(Hex(V_TOTAL - V_LAST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'M_OPTIONS
    If M_OPTIONS And 2 Then
      Dummy = 1
    Else
      Dummy = 0
    End If
    If Not M_OPTIONS And 4 Then
      Dummy = Dummy + 4
    End If
    If Not M_OPTIONS And 8 Then
      Dummy = Dummy + 8
    End If
    Dummy = LeadZero(Hex(Dummy), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'NOT_USED
    LineBinary = LineBinary & String(4, 48)
  End If
  Matrox_ModelineToBinary = LineBinary
End Function

''Public Function Matrox_BinaryToModeline(ResName As String, Matrox As String)
''  Dim Modeline As String
''  Dim HexLine As String
''  Dim Dummy As String
''
''  Dim P_FREQ As Single
''  Dim M_OPTIONS As Integer
''  Dim H_ACTIVE As Integer
''  Dim H_FIRST As Integer
''  Dim H_LAST As Integer
''  Dim H_TOTAL As Integer
''  Dim V_ACTIVE As Integer
''  Dim V_FIRST As Integer
''  Dim V_LAST As Integer
''  Dim V_TOTAL As Integer
''  Dim CHECKSUM As Long
''
''  Dim V_FREQ As Single
''
''  Dim Res() As String
''  HexLine = Matrox
''
''  Res = Split(ResName, ",")
''  H_ACTIVE = Res(0)
''  V_ACTIVE = Res(1)
''  V_FREQ = Res(2)
''
''  'M_OPTIONS
''  Dummy = Mid(HexLine, 43, 2) & Mid(HexLine, 41, 2)
''  Dummy = CInt("&H" & Dummy)
''  If Dummy And 1 Then
''    M_OPTIONS = M_OPTIONS + 2
''  End If
''  If Not Dummy And 4 Then
''    M_OPTIONS = M_OPTIONS + 4
''  End If
''  If Not Dummy And 8 Then
''    M_OPTIONS = M_OPTIONS + 8
''  End If
''  'P_FREQ
''  Dummy = Mid(HexLine, 11, 2) & Mid(HexLine, 9, 2)
''  P_FREQ = ("&H" & Dummy) / 1000
''  'H_FIRST
''  Dummy = Mid(HexLine, 19, 2) & Mid(HexLine, 17, 2)
''  H_FIRST = ("&H" & Dummy) + H_ACTIVE
''  'H_LAST
''  Dummy = Mid(HexLine, 23, 2) & Mid(HexLine, 21, 2)
''  H_LAST = ("&H" & Dummy) + H_FIRST
''  'H_TOTAL
''  Dummy = Mid(HexLine, 27, 2) & Mid(HexLine, 25, 2)
''  H_TOTAL = ("&H" & Dummy) + H_LAST
''  'V_FIRST
''  Dummy = Mid(HexLine, 31, 2) & Mid(HexLine, 29, 2)
''  If M_OPTIONS And 2 Then
''    V_FIRST = (("&H" & Dummy) * 2) + V_ACTIVE
''  Else
''    V_FIRST = ("&H" & Dummy) + V_ACTIVE
''  End If
''  'V_LAST
''  Dummy = Mid(HexLine, 35, 2) & Mid(HexLine, 33, 2)
''  If M_OPTIONS And 2 Then
''    V_LAST = (("&H" & Dummy) * 2) + V_FIRST
''  Else
''    V_LAST = ("&H" & Dummy) + V_FIRST
''  End If
''  'V_TOTAL
''  Dummy = Mid(HexLine, 39, 2) & Mid(HexLine, 37, 2)
''  If M_OPTIONS And 2 Then
''    V_TOTAL = (("&H" & Dummy) * 2) + V_LAST
''  Else
''    V_TOTAL = ("&H" & Dummy) + V_LAST
''  End If
''
''  Modeline = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "@" & V_FREQ & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
''  If M_OPTIONS And 1 Then
''    Modeline = Modeline & " doublescan"
''  End If
''  If M_OPTIONS And 2 Then
''    Modeline = Modeline & " interlace"
''  End If
''  If M_OPTIONS And 4 Then
''    Modeline = Modeline & " -hsync"
''  End If
''  If M_OPTIONS And 8 Then
''    Modeline = Modeline & " -vsync"
''  End If
''  Matrox_BinaryToModeline = Modeline
''End Function

