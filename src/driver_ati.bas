Attribute VB_Name = "driver_ati"
' -= Soft-15kHz - DRIVER - ATI Catalyst
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


Sub ATI_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Dim Key() As String
  Dim Index As Integer
  Dim SubKey() As String
  Dim SubIndex As Integer
  Dim Dummy As String
  
  Key = GetAllValues(HKEY_LOCAL_MACHINE, CurObjectPath)
  For Index = LBound(Key) To UBound(Key)
    Dummy = Key(Index)
    If Left(Dummy, 23) = "DALLargeDesktopModesBCD" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
    End If
    If Left(Dummy, 22) = "DALNonStandardModesBCD" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
    End If
    If Left(Dummy, 21) = "DALRestrictedModesBCD" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
    End If
    If Left(Dummy, 12) = "DALDTMCRTBCD" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
    End If
    If Left(Dummy, 12) = "DALDTMDFPBCD" Then
      DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
    End If
    If Left(Dummy, 16) = "DALRULE_RESTRICT" Then
      If Right(Dummy, 4) = "MODE" Then
        DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
      End If
    End If
    If Left(Dummy, 9) = "DALR6 CRT" Then
      If Dummy = "DALR6 CRT_MaxModeInfo" Then
      ElseIf Dummy = "DALR6 CRT" Then
      ElseIf Dummy = "DALR6 CRT2_MaxModeInfo" Then
      ElseIf Dummy = "DALR6 CRT2" Then
      Else
        DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, Dummy
      End If
    End If
  Next
  Dummy = ""
  Dummy = Dummy & ATI_ResolutionToBinary(320, 200, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(320, 240, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(400, 300, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(512, 384, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(640, 400, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(640, 480, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(800, 600, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1024, 768, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1152, 864, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1280, 1024, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1600, 1200, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1792, 1344, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1800, 1440, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1920, 1080, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1920, 1200, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(1920, 1440, 0)
  Dummy = Dummy & ATI_ResolutionToBinary(2048, 1536, 0)
  If Adapter(AdapterIndex).DriverVersion > 6 Then
    Dummy = Dummy & ATI_ResolutionToBinary(1280, 720, 0)
    Dummy = Dummy & ATI_ResolutionToBinary(1280, 800, 0)
  End If
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALRestrictedModesBCD", HexToBinary(Dummy)
  Dummy = String(8, "0")
  Dummy = Dummy & "00040000"
  Dummy = Dummy & "00030000"
  Dummy = Dummy & String(8, "0")
  Dummy = Dummy & "3C000000"
  SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALR6 CRT_MaxModeInfo", HexToBinary(Dummy)
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "DALR6 CRT_Info"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_FORCECRTDAC1DETECTED"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_FORCECRTDAC2DETECTED"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_R520FORCECRTDAC1DETECTED"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_R520FORCECRTDAC2DETECTED"
  DeleteValue HKEY_LOCAL_MACHINE, CurObjectPath, "DALRULE_AUTOGENERATELARGEDESKTOPMODES"
  If Host.Windows = 1 Then
    CurObjectPath = Adapter(AdapterIndex).ObjectPath2
    Key() = GetAllKeys(HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES")
    For Index = LBound(Key) To UBound(Key)
      SubKey = GetAllKeys(HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\" & Key(Index))
      For SubIndex = LBound(SubKey) To UBound(SubKey)
        DeleteKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\" & Key(Index) & "\" & SubKey(SubIndex)
      Next
      DeleteKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\" & Key(Index)
    Next
  End If
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub ATI_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim DALRestrictedModesBCD As String
  Dim DALNonStandardModesBCD As String
  Dim Dummy As String
  Dim Index As Integer
  Dim Res() As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  DALRestrictedModesBCD = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DALRestrictedModesBCD", "")
  DALRestrictedModesBCD = BinaryToHex(DALRestrictedModesBCD)
  DALNonStandardModesBCD = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD", "")
  For Index = 1 To 8
    DALNonStandardModesBCD = DALNonStandardModesBCD & GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD" & Index, "")
  Next
  DALNonStandardModesBCD = BinaryToHex(DALNonStandardModesBCD)
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        Res = Split(.ModeName, ",")
        If Adapter(AdapterIndex).ChipFlags And 1 Then
          If Res(0) = 320 Or Res(0) = 400 Then Res(0) = Res(0) + 1
        End If
        Dummy = ATI_ResolutionToBinary(CInt(Res(0)), CInt(Res(1)), CInt(Res(2)))
        DALRestrictedModesBCD = Replace(DALRestrictedModesBCD, Dummy, "")
        If Not InStr(1, DALNonStandardModesBCD, Dummy, vbBinaryCompare) Then
          DALNonStandardModesBCD = DALNonStandardModesBCD & Dummy
          Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
        End If
        Dummy = ATI_ModelineToBinary(Join(Res, ","), .Modeline)
        SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALDTMCRTBCD" & Res(0) & "x" & Res(1) & "x0x" & Res(2), HexToBinary(Dummy)
        SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALDTMDFPBCD" & Res(0) & "x" & Res(1) & "x0x" & Res(2), HexToBinary(Dummy)
        SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "DALRULE_RESTRICT" & Res(0) & "x" & Res(1) & "MODE", 0
        If Adapter(AdapterIndex).DriverVersion >= 5 Then
          Dummy = String(20, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(6, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(14, "0")
          SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALR6 CRT" & Res(0) & "x" & Res(1) & "x0x" & Res(2), HexToBinary(Dummy)
        ElseIf Adapter(AdapterIndex).DriverVersion >= 3 Then
          Dummy = String(20, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(6, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(6, "0")
          SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALR6 CRT" & Res(0) & "x" & Res(1) & "x0x" & Res(2), HexToBinary(Dummy)
        Else
          Dummy = String(16, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(6, "0")
          Dummy = Dummy & "01"
          Dummy = Dummy & String(6, "0")
          SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALR6 CRT" & Res(0) & "x" & Res(1) & "x0x" & Res(2), HexToBinary(Dummy)
        End If
      End If
    End With
  Next
  Index = 0
  While Len(DALNonStandardModesBCD) > 128
    Dummy = Left(DALNonStandardModesBCD, 128)
    DALNonStandardModesBCD = Mid(DALNonStandardModesBCD, 129)
    If Index = 0 Then
      SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD", HexToBinary(Dummy)
    Else
      SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD" & Index, HexToBinary(Dummy)
    End If
    Index = Index + 1
  Wend
  If Index = 0 Then
    SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD", HexToBinary(DALNonStandardModesBCD)
  Else
    SaveSettingBinary HKEY_LOCAL_MACHINE, CurObjectPath, "DALNonStandardModesBCD" & Index, HexToBinary(DALNonStandardModesBCD)
  End If
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_FORCECRTDAC1DETECTED", 1
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_FORCECRTDAC2DETECTED", 1
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_R520FORCECRTDAC1DETECTED", 1
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "CRTRULE_R520FORCECRTDAC2DETECTED", 1
  SaveSettingLong HKEY_LOCAL_MACHINE, CurObjectPath, "DALRULE_AUTOGENERATELARGEDESKTOPMODES", 0
  If Host.Windows = 1 Then
    CurObjectPath = Adapter(AdapterIndex).ObjectPath2
    For Index = 1 To ModeCount
      With Mode(Index)
        If .ModeFreq = Freq Then
          Res = Split(.ModeName, ",")
          If Adapter(AdapterIndex).ChipFlags And 1 Then
            If Res(0) = 320 Or Res(0) = 400 Then Res(0) = Res(0) + 1
          End If
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\8"
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\16"
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\32"
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\8\" & Res(0) & "," & Res(1)
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\16\" & Res(0) & "," & Res(1)
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\32\" & Res(0) & "," & Res(1)
          SaveSettingString HKEY_LOCAL_MACHINE, CurObjectPath & "\MODES\8\" & Res(0) & "," & Res(1), "", Res(2)
        End If
      End With
    Next
  End If
End Sub

Function ATI_ResolutionToBinary(X As Integer, Y As Integer, Hz As Integer) As String
  Dim Dummy As String
  Dummy = X
  While Len(Dummy) < 4
    Dummy = "0" & Dummy
  Wend
  ATI_ResolutionToBinary = Dummy
  Dummy = Y
  While Len(Dummy) < 4
    Dummy = "0" & Dummy
  Wend
  ATI_ResolutionToBinary = ATI_ResolutionToBinary & Dummy
  ATI_ResolutionToBinary = ATI_ResolutionToBinary & "0000"
  Dummy = Hz
  While Len(Dummy) < 4
    Dummy = "0" & Dummy
  Wend
  ATI_ResolutionToBinary = ATI_ResolutionToBinary & Dummy
End Function

Public Function ATI_ModelineToBinary(ResName As String, Modeline As String) As String
  Dim ResParam() As String
  Dim LineBinary As String
  Dim LineParam() As String
  Dim Dummy As String
  
  Dim P_FREQ As Integer
  Dim M_OPTIONS As Integer
  Dim H_ACTIVE As Integer
  Dim H_FIRST As Integer
  Dim H_LAST As Integer
  Dim H_TOTAL As Integer
  Dim V_ACTIVE As Integer
  Dim V_FIRST As Integer
  Dim V_LAST As Integer
  Dim V_TOTAL As Integer
  Dim CHECKSUM As Long
  
  LineBinary = ""
  LineParam = Split(Modeline, " ", 12)
  
  P_FREQ = CSng(LineParam(2) * 100)
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
  
  '-doublescan workaround-
  If M_OPTIONS And 1 Then
    'V_ACTIVE = V_ACTIVE * 2
    V_FIRST = V_FIRST * 2
    V_LAST = V_LAST * 2
    V_TOTAL = V_TOTAL * 2
  End If
  
  '-boing-
  ResParam = Split(ResName, ",")
  If ResParam(0) > H_ACTIVE Then H_ACTIVE = ResParam(0)
  '-boing-
  
  CHECKSUM = M_OPTIONS + H_ACTIVE + H_FIRST + (H_LAST - H_FIRST) + H_TOTAL + V_ACTIVE + V_FIRST + (V_LAST - V_FIRST) + V_TOTAL + P_FREQ
  CHECKSUM = 65535 - CHECKSUM
  
  If UCase(LineParam(0)) = "MODELINE" Then
    'M_OPTIONS
    LineBinary = LineBinary & LeadZero(Hex(M_OPTIONS), 8)
    
    'H_TOTAL
    LineBinary = LineBinary & LeadZero(CStr(H_TOTAL), 8)
    
    'H_ACTIVE
    LineBinary = LineBinary & LeadZero(CStr(H_ACTIVE), 8)
    
    'H_FIRST
    LineBinary = LineBinary & LeadZero(CStr(H_FIRST), 8)
    
    'H_LAST - H_FIRST
    LineBinary = LineBinary & LeadZero(CStr(H_LAST - H_FIRST), 8)
    
    'V_TOTAL
    LineBinary = LineBinary & LeadZero(CStr(V_TOTAL), 8)
    
    'V_ACTIVE
    LineBinary = LineBinary & LeadZero(CStr(V_ACTIVE), 8)
    
    'V_FIRST
    LineBinary = LineBinary & LeadZero(CStr(V_FIRST), 8)
    
    'V_LAST - V_FIRST
    LineBinary = LineBinary & LeadZero(CStr(V_LAST - V_FIRST), 8)
    
    'P_FREQ
    LineBinary = LineBinary & LeadZero(CStr(P_FREQ), 8)
    
    'NOT_USED
    LineBinary = LineBinary & String(16, 48)
    LineBinary = LineBinary & String(16, 48)
    LineBinary = LineBinary & String(16, 48)
    
    'CHECKSUM
    LineBinary = LineBinary & LeadZero(Hex(CHECKSUM), 8)
  End If
  
  ATI_ModelineToBinary = LineBinary
End Function

''Public Function ATI_BinaryToModeline(ResName As String, ATI As String) As String
''  Dim ModeLine As String
''  Dim HexLine As String
''
''  Dim P_FREQ As Integer
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
''  HexLine = ATI
''
''  M_OPTIONS = "&H" & Mid(HexLine, 5, 4)
''  H_TOTAL = Mid(HexLine, 13, 4)
''  H_ACTIVE = Mid(HexLine, 21, 4)
''  H_FIRST = Mid(HexLine, 29, 4)
''  H_LAST = H_FIRST + (Mid(HexLine, 37, 4))
''  V_TOTAL = Mid(HexLine, 45, 4)
''  V_ACTIVE = Mid(HexLine, 53, 4)
''  V_FIRST = Mid(HexLine, 61, 4)
''  V_LAST = V_FIRST + (Mid(HexLine, 69, 4))
''  P_FREQ = Mid(HexLine, 77, 4)
''
''  If M_OPTIONS And 2 Then
''    V_FREQ = Round((1000000 * (CSng(P_FREQ) / 100)) / (CSng(H_TOTAL) * (CSng(V_TOTAL) / 2)), 3)
''  Else
''    V_FREQ = Round((1000000 * (CSng(P_FREQ) / 100)) / (CSng(H_TOTAL) * CSng(V_TOTAL)), 3)
''  End If
''
''  ModeLine = "modeline '" & H_ACTIVE & "x" & V_ACTIVE & "' " & (CSng(P_FREQ) / 100) & " " & H_ACTIVE & " " & H_FIRST & " " & H_LAST & " " & H_TOTAL & " " & V_ACTIVE & " " & V_FIRST & " " & V_LAST & " " & V_TOTAL
''  If M_OPTIONS And 1 Then
''    ModeLine = ModeLine & " doublescan"
''  End If
''  If M_OPTIONS And 2 Then
''    ModeLine = ModeLine & " interlace"
''  End If
''  If M_OPTIONS And 4 Then
''    ModeLine = ModeLine & " -hsync"
''  End If
''  If M_OPTIONS And 8 Then
''    ModeLine = ModeLine & " -vsync"
''  End If
''  ATI_BinaryToModeline = ModeLine
''End Function

