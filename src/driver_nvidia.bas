Attribute VB_Name = "driver_nvidia"
' -= Soft-15kHz - DRIVER - NVidia ForceWare
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


' -= Windows NT5 - 2000 / XP =-
Sub NVidia_NT5_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  SaveSettingMultiString HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", ""
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CUST_MODE"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "NV_CustModes"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "DevicesConnected"
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub NVidia_NT5_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim NV_Modes(0 To 255) As String
  Dim CUST_MODE As String
  Dim Res() As String
  Dim Dummy As String
  Dim Index As Integer
  Dim Count As Integer
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", "")
  If Right(Dummy, 1) = Chr(0) Then Dummy = Mid(Dummy, 1, Len(Dummy) - 1)
  If Len(Dummy) > 0 Then
    Res = Split(Dummy, ";")
    For Index = LBound(Res) To UBound(Res)
      If Not Res(Index) = "" Then
        Count = "&H" & Mid(Res(Index), InStr(1, Res(Index), "=") + 2)
        NV_Modes(Count) = Mid(Res(Index), 5, InStr(1, Res(Index), "=") - 5)
      End If
    Next
  End If
  Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "CUST_MODE", "")
  CUST_MODE = BinaryToHex(Dummy)
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        Res = Split(.ModeName, ",")
        Dummy = " " & Res(0) & "x" & Res(1)
        If InStr(1, NV_Modes(Res(2)), Dummy) = 0 Then
          NV_Modes(Res(2)) = NV_Modes(Res(2)) & Dummy
          Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
          If Adapter(AdapterIndex).ModeCount = 33 Then If Not Host.RuntimeFlags And 2 Then MsgBox "NVidia drivers don't allow more than 32 resolution defintions.", vbInformation + vbOKOnly, "Soft-15kHz"
        End If
        Dummy = NVidia_NT5_ModelineToBinary(.ModeName, .Modeline)
        Count = 1
        While Count < Len(CUST_MODE)
          If Left(Dummy, 40) = Mid(CUST_MODE, Count, 40) Then
            CUST_MODE = Left(CUST_MODE, Count - 1) & Mid(CUST_MODE, Count + 184)
          End If
          Count = Count + 184
        Wend
        CUST_MODE = CUST_MODE & Dummy
      End If
    End With
  Next
  Dummy = ""
  For Index = 0 To 255
    If Not NV_Modes(Index) = "" Then
      Dummy = Dummy & "{*}S" & NV_Modes(Index) & "=" & Hex(CInt(-32768 + Index)) & ";"
    End If
  Next
  SaveSettingMultiString HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", Dummy
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "CUST_MODE", HexToBinary(CUST_MODE)
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DevicesConnected", HexToBinary("03000000")
End Sub

Function NVidia_NT5_ModelineToBinary(ResName As String, Modeline As String) As String
  Dim LineBinary As String
  Dim LineParam() As String
  Dim Dummy As String
  
  Dim Header As String
  Dim Data1 As String
  Dim Data2 As String
  Dim Foot As String
  
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
  
  Dim Res() As String
  
  LineBinary = ""
  LineParam = Split(Modeline, " ", 12)
  
  P_FREQ = CSng(LineParam(2))
  If InStr(1, LCase(LineParam(11)), "doublescan") Then
    M_OPTIONS = M_OPTIONS + 1
  End If
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
  
  ' -= interlace workaround =-
  If M_OPTIONS And 2 Then
    If V_TOTAL Mod 2 = 1 Then
      V_TOTAL = V_TOTAL + 1
    End If
    V_ACTIVE = V_ACTIVE / 2
    V_FIRST = V_FIRST / 2
    V_LAST = V_LAST / 2
    V_TOTAL = V_TOTAL / 2
  End If
  ' -= doublescan workaround =-
  If M_OPTIONS And 1 Then
    V_FIRST = V_FIRST * 2
    V_LAST = V_LAST * 2
    V_TOTAL = V_TOTAL * 2
  End If
  
  V_FREQ = Round(((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL))) * 1000, 3)
  
  Res = Split(ResName, ",")
    
  If UCase(LineParam(0)) = "MODELINE" Then
    '--Header--
    'NOT_USED
    LineBinary = String(16, 48)
    'DEVICE_MASK
    LineBinary = LineBinary & "03000000"
    'Res(0)   'X
    Dummy = LeadZero(Hex(Res(0)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'Res(1)   'Y
    Dummy = LeadZero(Hex(Res(1)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'NOT_USED
    LineBinary = LineBinary & String(4, 48)
    'Res(2)   'Hz
    Dummy = LeadZero(Hex(Res(2)), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Header = LineBinary
    
    '--Data1--
    'P_FREQ
    Dummy = LeadZero(Hex(P_FREQ * 100), 4)
    LineBinary = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'NOT_USED
    LineBinary = LineBinary & String(4, 48)
    'H_ACTIVE
    Dummy = LeadZero(Hex(H_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_ACTIVE
    Dummy = LeadZero(Hex(V_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_TOTAL
    Dummy = LeadZero(Hex(H_TOTAL), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_FIRST - H_ACTIVE
    Dummy = LeadZero(Hex(H_FIRST - H_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_LAST - H_FIRST
    Dummy = LeadZero(Hex(H_LAST - H_FIRST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_TOTAL
    Dummy = LeadZero(Hex(V_TOTAL), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_FIRST - V_ACTIVE
    Dummy = LeadZero(Hex(V_FIRST - V_ACTIVE), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_LAST - V_FIRST
    Dummy = LeadZero(Hex(V_LAST - V_FIRST), 4)
    LineBinary = LineBinary & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'NOT_USED
    LineBinary = LineBinary & String(8, 48)
    'M_OPTION
    '-hsync
    If M_OPTIONS And 4 Then
      LineBinary = LineBinary & "01"
    Else
      LineBinary = LineBinary & "00"
    End If
    '-vsync
    If M_OPTIONS And 8 Then
      LineBinary = LineBinary & "01"
    Else
      LineBinary = LineBinary & "00"
    End If
    'interlace
    If M_OPTIONS And 2 Then
      LineBinary = LineBinary & "01"
    Else
      LineBinary = LineBinary & "00"
    End If
    'doublescan
    If M_OPTIONS And 1 Then
      LineBinary = LineBinary & "01"
    Else
      LineBinary = LineBinary & "00"
    End If
    'V_FREQ
    Dummy = LeadZero(Hex(V_FREQ), 8)
    LineBinary = LineBinary & Mid(Dummy, 7, 2) & Mid(Dummy, 5, 2) & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Data1 = LineBinary
    
    '--Data2--
    Data2 = Data1
    
    '--Foot--
    Foot = "FF20" & String(12, 48) 'String(16, 48)
    
    LineBinary = Header & Data1 & Data2 & Foot
  End If
  NVidia_NT5_ModelineToBinary = LineBinary
End Function

''Function NVidia_NT5_BinaryToModeline(ResName As String, NVidia As String) As String
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
''  Dim Header As String
''  Dim Data1 As String
''  Dim Data2 As String
''  Dim Foot As String
''
''  HexLine = NVidia
''
''  Header = Mid(HexLine, 1, 40)
''  Data1 = Mid(HexLine, 41, 64)
''  Data2 = Mid(HexLine, 105, 64)
''  Foot = Mid(HexLine, 169, 16)
''
''  'H_ACTIVE
''  Dummy = Mid(Header, 27, 2) & Mid(Header, 25, 2)
''  H_ACTIVE = "&H" & Dummy
''  'V_ACTIVE
''  Dummy = Mid(Header, 31, 2) & Mid(Header, 29, 2)
''  V_ACTIVE = "&H" & Dummy
''  'V_FREQ
''  Dummy = Mid(Header, 39, 2) & Mid(Header, 37, 2)
''  V_FREQ = "&H" & Dummy
''
''  'P_FREQ (sometimes)
''  Dummy = Mid(Data1, 3, 2) & Mid(Data1, 1, 2)
''  P_FREQ = ("&H" & Dummy) / 100
''  'H_ACTIVE
''  Dummy = Mid(Data1, 11, 2) & Mid(Data1, 9, 2)
''  H_ACTIVE = "&H" & Dummy
''  'V_ACTIVE
''  Dummy = Mid(Data1, 15, 2) & Mid(Data1, 13, 2)
''  V_ACTIVE = "&H" & Dummy
''  'H_TOTAL
''  Dummy = Mid(Data1, 19, 2) & Mid(Data1, 17, 2)
''  H_TOTAL = "&H" & Dummy
''  'H_FIRST
''  Dummy = Mid(Data1, 23, 2) & Mid(Data1, 21, 2)
''  H_FIRST = ("&H" & Dummy) + H_ACTIVE
''  'H_LAST
''  Dummy = Mid(Data1, 27, 2) & Mid(Data1, 25, 2)
''  H_LAST = ("&H" & Dummy) + H_FIRST
''  'V_TOTAL
''  Dummy = Mid(Data1, 31, 2) & Mid(Data1, 29, 2)
''  V_TOTAL = "&H" & Dummy
''  'V_FIRST
''  Dummy = Mid(Data1, 35, 2) & Mid(Data1, 33, 2)
''  V_FIRST = ("&H" & Dummy) + V_ACTIVE
''  'V_LAST
''  Dummy = Mid(Data1, 39, 2) & Mid(Data1, 37, 2)
''  V_LAST = ("&H" & Dummy) + V_FIRST
''  'M_OPTIONS
''  If Mid(Data1, 49, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 4
''  End If
''  If Mid(Data1, 51, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 8
''  End If
''  If Mid(Data1, 53, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 2
''  End If
''  'V_FREQ
''  Dummy = Mid(Data1, 59, 2) & Mid(Data1, 57, 2)
''  V_FREQ = ("&H" & Dummy) / 1000
''
''  If P_FREQ = 0 Then
''    P_FREQ = Round(((CSng(H_TOTAL) * CSng(V_TOTAL)) * V_FREQ) / 1000000, 3)
''  End If
''
''  If M_OPTIONS And 2 Then
''    Modeline = "modeline '" & H_ACTIVE & "x" & (V_ACTIVE * 2) & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & (V_ACTIVE * 2) & " " & (V_FIRST * 2) & " " & (V_LAST * 2) & " " & (V_TOTAL * 2)
''  Else
''    Modeline = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
''  End If
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
''  NVidia_NT5_BinaryToModeline = Modeline
''End Function


' -= Windows NT6 - Vista/Seven =-
Sub NVidia_NT6_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  SaveSettingMultiString HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", ""
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CustomDisplay"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "DevicesConnected"
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub NVidia_NT6_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim NV_Modes(0 To 255) As String
  Dim CustomDisplay As String
  Dim Res() As String
  Dim Dummy As String
  Dim Index As Integer
  Dim Count As Integer
  Dim Splice As Integer
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Select Case Adapter(AdapterIndex).DriverVersion
    Case 6
      Splice = 176 * 2  '0x00B0
    Case 7
      Splice = 272 * 2  '0x0110
    Case 8
      Splice = 280 * 2  '0x0118
  End Select
  Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", "")
  If Right(Dummy, 1) = Chr(0) Then Dummy = Mid(Dummy, 1, Len(Dummy) - 1)
  If Len(Dummy) > 0 Then
    Res = Split(Dummy, ";")
    For Index = LBound(Res) To UBound(Res)
      If Not Res(Index) = "" Then
        Count = "&H" & Mid(Res(Index), InStr(1, Res(Index), "=") + 2)
        NV_Modes(Count) = Mid(Res(Index), 5, InStr(1, Res(Index), "=") - 5)
      End If
    Next
  End If
  Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "CustomDisplay", "")
  CustomDisplay = BinaryToHex(Dummy)
  While Right(CustomDisplay, Splice) = String(Splice, 48)
    CustomDisplay = Left(CustomDisplay, Len(CustomDisplay) - Splice)
  Wend
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        Res = Split(.ModeName, ",")
        Dummy = " " & Res(0) & "x" & Res(1)
        If InStr(1, NV_Modes(Res(2)), Dummy) = 0 Then
          NV_Modes(Res(2)) = NV_Modes(Res(2)) & Dummy
          Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
          If Adapter(AdapterIndex).ModeCount = 33 Then If Not Host.RuntimeFlags And 2 Then MsgBox "NVidia drivers don't allow more than 32 resolution defintions.", vbInformation + vbOKOnly, "Soft-15kHz"
        End If
        Dummy = NVidia_NT6_ModelineToBinary(.ModeName, .Modeline, Splice)
        Count = 1
        While Count < Len(CustomDisplay)
          If Left(Dummy, 32) = Mid(CustomDisplay, Count, 32) Then
            CustomDisplay = Left(CustomDisplay, Count - 1) & Mid(CustomDisplay, Count + Splice)
          End If
          Count = Count + Splice
        Wend
        CustomDisplay = CustomDisplay & Dummy
      End If
    End With
  Next
  Dummy = ""
  For Index = 0 To 255
    If Not NV_Modes(Index) = "" Then
      Dummy = Dummy & "{*}S" & NV_Modes(Index) & "=" & Hex(CInt(-32768 + Index)) & ";"
    End If
  Next
  Index = Splice * 32
  If Len(CustomDisplay) > Index Then CustomDisplay = Left(CustomDisplay, Index)
  CustomDisplay = CustomDisplay & String(Index - Len(CustomDisplay), 48)
  SaveSettingMultiString HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", Dummy
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "CustomDisplay", HexToBinary(CustomDisplay)
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DevicesConnected", HexToBinary("03000000")
End Sub

'? NVidia_NT6_ModelineToBinary("640,240,60","modeline '640x240' 13,22 640 672 736 832 240 243 246 265 -hsync -vsync",544)
Function NVidia_NT6_ModelineToBinary(ResName As String, Modeline As String, Splice As Integer) As String
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

  Dim Res() As String
  
  LineBinary = ""
  LineParam = Split(Modeline, " ", 12)

  P_FREQ = CSng(LineParam(2))
  If InStr(1, LCase(LineParam(11)), "doublescan") Then
    M_OPTIONS = M_OPTIONS + 1
  End If
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

  Res = Split(ResName, ",")
  
  ' -= interlace workaround =-
  If M_OPTIONS And 2 Then
    If V_TOTAL Mod 2 = 1 Then
      V_TOTAL = V_TOTAL + 1
    End If
    V_ACTIVE = V_ACTIVE / 2
    V_FIRST = V_FIRST / 2
    V_LAST = V_LAST / 2
    V_TOTAL = V_TOTAL / 2
    Res(2) = Res(2) / 2
  End If
  ' -= doublescan workaround =-
  If M_OPTIONS And 1 Then
    V_FIRST = V_FIRST * 2
    V_LAST = V_LAST * 2
    V_TOTAL = V_TOTAL * 2
  End If

  V_FREQ = Round(((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL))) * 1000, 3)
  H_FREQ = V_TOTAL * (V_FREQ / 1000)

  If UCase(LineParam(0)) = "MODELINE" Then
    Dim LineBinary2 As String
    Dim LineBinary3 As String
    
    ' Header
    Dim LineBinary0 As String
    
    Dim LB0_01 As String
    Dim LB0_02 As String
    Dim LB0_03 As String
    Dim LB0_04 As String
    Dim LB0_05 As String
    Dim LB0_06 As String
    Dim LB0_07 As String
    
    'DEVICEMASK
    '"01000000" on NT6
    '"01010002" on NT7 ' doesn't work!
    '"00011000" on NT7 300+
    Select Case Splice
      Case 352
        LB0_01 = "01000000"
      Case 544
        LB0_01 = "00010002"
      Case 560
        LB0_01 = "00012000" '"00011000"
    End Select
    
    'HARDWAREID
    LB0_02 = "ffffffff" ' Monitor ID = "PNP0815" = "150841d0"
    
    'H_ACTIVE (Windows)
    Dummy = LeadZero(Hex(Res(0)), 4)
    LB0_03 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"
    
    'V_ACTIVE (Windows)
    Dummy = LeadZero(Hex(Res(1)), 4)
    LB0_04 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"
    
    'UNKNOWN
    LB0_05 = "15000000"
    
    'COLOR DEPTH
    LB0_06 = "20000000"
    
    'ZERO
    LB0_07 = "00000000"
    
    LineBinary0 = LB0_01 & LB0_02 & LB0_03 & LB0_04 & LB0_05 & LB0_06 & LB0_07 & _
                  LB0_07 & LB0_07 & LB0_03 & LB0_04 & _
                  LB0_07 & LB0_07 & LB0_03 & LB0_04 & _
                  LB0_07 & LB0_07 & LB0_03 & LB0_04
                  
                  
    ' Details
    Dim LineBinary1 As String
    Dim LB1_01 As String
    Dim LB1_02 As String
    Dim LB1_03 As String
    Dim LB1_04 As String
    Dim LB1_05 As String
    Dim LB1_06 As String
    Dim LB1_07 As String
    Dim LB1_08 As String
    Dim LB1_09 As String
    Dim LB1_10 As String
    Dim LB1_11 As String
    Dim LB1_12 As String
    Dim LB1_13 As String
    Dim LB1_14 As String
    Dim LB1_15 As String
    Dim LB1_16 As String
    Dim LB1_17 As String
    Dim LB1_18 As String
    Dim LB1_19 As String
    Dim LB1_20 As String
                  
    'H_ACTIVE (Resolution)
    Dummy = LeadZero(Hex(H_ACTIVE), 4)
    LB1_01 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"
    
    'H_FIRST - H_ACTIVE
    Dummy = LeadZero(Hex(H_FIRST - H_ACTIVE), 4)
    LB1_02 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    
    'H_LAST - H_FIRST
    Dummy = LeadZero(Hex(H_LAST - H_FIRST), 4)
    LB1_03 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    
    'H_TOTAL
    Dummy = LeadZero(Hex(H_TOTAL), 4)
    LB1_04 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    
    'H_SYNC_POL
    LB1_05 = IIf(M_OPTIONS And 4, "01", "00") & "00"
    
    'V_ACTIVE (Resolution)
    Dummy = LeadZero(Hex(V_ACTIVE), 4)
    LB1_06 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"
    
    'V_FIRST - V_ACTIVE
    Dummy = LeadZero(Hex(V_FIRST - V_ACTIVE), 4)
    LB1_07 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    
    'V_LAST - V_LAST
    Dummy = LeadZero(Hex(V_LAST - V_FIRST), 4)
    LB1_08 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    
    'V_TOTAL
    Dummy = LeadZero(Hex(V_TOTAL), 4)
    LB1_09 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'V_SYNC_POL
    LB1_10 = IIf(M_OPTIONS And 8, "01", "00") & "00"

    'INTERLACE
    LB1_11 = IIf(M_OPTIONS And 2, "01", "00") & "000000"

    'P_FREQ
    Dummy = LeadZero(Hex(P_FREQ * 100), 4)
    LB1_12 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"

    'ZERO
    LB1_13 = "00000000"
    
    'V_FREQ (OS)
    Dummy = LeadZero(Hex(Res(2)), 4)
    LB1_14 = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2) & "0000"

    'V_FREQ
    Dummy = LeadZero(Hex(V_FREQ), 8)
    LB1_15 = Mid(Dummy, 7, 2) & Mid(Dummy, 5, 2) & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'ZERO
    LB1_16 = "00000000"

    '?
    Select Case Splice
      Case 352, 544
        LB1_17 = "01000000"
        LB1_18 = ""
      Case 560
        LB1_17 = "01000200"
        LB1_18 = "0000463A"
    End Select

    'FORMAT
    '01 - ?
    '03 - DMT
    '05 - CVT
    '07 - automatic?/GTF
    '0A/1D - manual
    Select Case Splice
      Case 352, 544
        LB1_19 = "0A00"
      Case 560
        LB1_19 = "001D"
    End Select
    
    'STRING
    Dummy = "@CUST:" & Res(0) & "x" & Res(1) & "x" & Replace(Format((CSng(V_FREQ) / 1000), "0.000"), ",", ".") & "Hz"
    If M_OPTIONS And 2 Then
      Dummy = Dummy & "/i"
    End If
    Dummy = "00" & BinaryToHex(Dummy)
    LB1_20 = Dummy & String(84 - Len(Dummy), 48)
    
    LineBinary1 = LB1_01 & LB1_02 & LB1_03 & LB1_04 & LB1_05 & _
                  LB1_06 & LB1_07 & LB1_08 & LB1_09 & LB1_10 & _
                  LB1_11 & LB1_12 & LB1_13 & LB1_14 & LB1_15 & _
                  LB1_16 & LB1_17 & LB1_18 & LB1_19 & LB1_20
    
    
    ' Details Copy?
    Select Case Splice
      Case 352
        LineBinary2 = ""
      Case 544, 560
        LineBinary2 = LineBinary1
    End Select
    
    ' Footer?
    LineBinary3 = "00000000"
    
    LineBinary = LineBinary0 & LineBinary1 & LineBinary2 & LineBinary3
  End If
  NVidia_NT6_ModelineToBinary = LineBinary
End Function

''Function NVidia_NT6_BinaryToModeline(ResName As String, HexLine As String) As String
''  Dim Modeline As String
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
''
''  Dim V_FREQ As Single
''
''  'H_ACTIVE
''  Dummy = Mid(HexLine, 19, 2) & Mid(HexLine, 17, 2)
''  H_ACTIVE = "&H" & Dummy
''  'V_ACTIVE
''  Dummy = Mid(HexLine, 27, 2) & Mid(HexLine, 25, 2)
''  V_ACTIVE = "&H" & Dummy
''  'V_FREQ
''  Dummy = Mid(HexLine, 227, 2) & Mid(HexLine, 225, 2)
''  V_FREQ = "&H" & Dummy
''
''  'P_FREQ (sometimes)
''  Dummy = Mid(HexLine, 211, 2) & Mid(HexLine, 209, 2)
''  P_FREQ = ("&H" & Dummy) / 100
''  'H_ACTIVE
''  Dummy = Mid(HexLine, 155, 2) & Mid(HexLine, 153, 2)
''  H_ACTIVE = "&H" & Dummy
''  'V_ACTIVE
''  Dummy = Mid(HexLine, 179, 2) & Mid(HexLine, 177, 2)
''  V_ACTIVE = "&H" & Dummy
''  'H_TOTAL
''  Dummy = Mid(HexLine, 171, 2) & Mid(HexLine, 169, 2)
''  H_TOTAL = "&H" & Dummy
''  'H_FIRST
''  Dummy = Mid(HexLine, 163, 2) & Mid(HexLine, 161, 2)
''  H_FIRST = ("&H" & Dummy) + H_ACTIVE
''  'H_LAST
''  Dummy = Mid(HexLine, 167, 2) & Mid(HexLine, 165, 2)
''  H_LAST = ("&H" & Dummy) + H_FIRST
''  'V_TOTAL
''  Dummy = Mid(HexLine, 195, 2) & Mid(HexLine, 193, 2)
''  V_TOTAL = "&H" & Dummy
''  'V_FIRST
''  Dummy = Mid(HexLine, 187, 2) & Mid(HexLine, 185, 2)
''  V_FIRST = ("&H" & Dummy) + V_ACTIVE
''  'V_LAST
''  Dummy = Mid(HexLine, 191, 2) & Mid(HexLine, 189, 2)
''  V_LAST = ("&H" & Dummy) + V_FIRST
''  'M_OPTIONS
''  If Mid(HexLine, 173, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 4
''  End If
''  If Mid(HexLine, 199, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 8
''  End If
''  If Mid(HexLine, 201, 2) = "01" Then
''    M_OPTIONS = M_OPTIONS + 2
''  End If
''  'V_FREQ
''  Dummy = Mid(HexLine, 239, 2) & Mid(HexLine, 237, 2) & Mid(HexLine, 235, 2) & Mid(HexLine, 233, 2)
''  V_FREQ = ("&H" & Dummy) / 1000
''
''  If P_FREQ = 0 Then
''    P_FREQ = Round(((CSng(H_TOTAL) * CSng(V_TOTAL)) * V_FREQ) / 1000000, 3)
''  End If
''
''  If M_OPTIONS And 2 Then
''    Modeline = "modeline '" & H_ACTIVE & "x" & (V_ACTIVE * 2) & "@" & V_FREQ & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & (V_ACTIVE * 2) & " " & (V_FIRST * 2) & " " & (V_LAST * 2) & " " & (V_TOTAL * 2)
''  Else
''    Modeline = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "@" & V_FREQ & "' " & P_FREQ & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
''  End If
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
''  NVidia_NT6_BinaryToModeline = Modeline
''End Function

