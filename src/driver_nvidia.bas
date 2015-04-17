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
      Splice = 272 * 2  '0x0110"
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
    'DUMMY
    LineBinary = String(Splice, 48)
    'DEVICEMASK
    '"01000000" on NT6
    '"01010002" on NT7 ' doesn't work!
    Mid(LineBinary, 1, 8) = "00010002"
    
    'HARDWAREID
    Mid(LineBinary, 9, 8) = "150841d0" ' Monitor ID = "PNP0815"

''    'CHECKSUM?
''    Mid(LineBinary, 343, 8) = "18340c00"
    
    'HEADER
    Mid(LineBinary, 33, 2) = "15"
    'COLOR DEPTH
    Mid(LineBinary, 41, 2) = "20"

    'H_ACTIVE (Windows)
    Dummy = LeadZero(Hex(Res(0)), 4)
    Mid(LineBinary, 17, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 73, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 105, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 137, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_ACTIVE (Windows)
    Dummy = LeadZero(Hex(Res(1)), 4)
    Mid(LineBinary, 25, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 81, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 113, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    Mid(LineBinary, 145, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'H_ACTIVE (Resolution)
    Dummy = LeadZero(Hex(H_ACTIVE), 4)
    Mid(LineBinary, 153, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_FIRST - H_ACTIVE
    Dummy = LeadZero(Hex(H_FIRST - H_ACTIVE), 4)
    Mid(LineBinary, 161, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_LAST - H_FIRST
    Dummy = LeadZero(Hex(H_LAST - H_FIRST), 4)
    Mid(LineBinary, 165, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'H_TOTAL
    Dummy = LeadZero(Hex(H_TOTAL), 4)
    Mid(LineBinary, 169, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'V_ACTIVE (Resolution)
    Dummy = LeadZero(Hex(V_ACTIVE), 4)
    Mid(LineBinary, 177, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_FIRST - V_ACTIVE
    Dummy = LeadZero(Hex(V_FIRST - V_ACTIVE), 4)
    Mid(LineBinary, 185, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_LAST - V_LAST
    Dummy = LeadZero(Hex(V_LAST - V_FIRST), 4)
    Mid(LineBinary, 189, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)
    'V_TOTAL
    Dummy = LeadZero(Hex(V_TOTAL), 4)
    Mid(LineBinary, 193, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'P_FREQ
    Dummy = LeadZero(Hex(P_FREQ * 100), 4)
    Mid(LineBinary, 209, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'V_FREQ (OS)
    Dummy = LeadZero(Hex(Res(2)), 4)
    'Dummy = LeadZero(Hex(CInt(V_FREQ / 1000)), 4)
    Mid(LineBinary, 225, 4) = Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    'V_FREQ
    Dummy = LeadZero(Hex(V_FREQ), 8)
    Mid(LineBinary, 233, 8) = Mid(Dummy, 7, 2) & Mid(Dummy, 5, 2) & Mid(Dummy, 3, 2) & Mid(Dummy, 1, 2)

    '?
    Mid(LineBinary, 249, 2) = "01"

    'FORMAT
    '01 - ?
    '03 - DMT
    '05 - CVT
    '07 - automatic?/GTF
    '0A - manual
    Mid(LineBinary, 257, 2) = "0A"

    'M_OPTIONS
    If M_OPTIONS And 2 Then
      Mid(LineBinary, 201, 2) = "01"
    End If
    If M_OPTIONS And 4 Then
      Mid(LineBinary, 173, 2) = "01"
    End If
    If M_OPTIONS And 8 Then
      Mid(LineBinary, 197, 2) = "01"
    End If

    'STRING
    Dummy = "@CUST:" & Res(0) & "x" & Res(1) & "x" & Replace(Format((CSng(V_FREQ) / 1000), "0.000"), ",", ".") & "Hz"
    If M_OPTIONS And 2 Then
      Dummy = Dummy & "/i"
    End If
    Dummy = BinaryToHex(Dummy)

    Mid(LineBinary, 263, Len(Dummy)) = Dummy
    
    'Splice
    If Splice = 544 Then
      'CopyCat because I'm lazy ;D
      Mid(LineBinary, 345, 108) = Mid(LineBinary, 153, 108)
      Dummy = Mid(Dummy, 3)
      Mid(LineBinary, 457, Len(Dummy)) = Dummy
      'Mid(LineBinary, 537, 8) = "76007300"
      'Mid(LineBinary, 537, 8) = "64006300"
      Mid(LineBinary, 537, 8) = "EE3FFDF7"
    End If
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

