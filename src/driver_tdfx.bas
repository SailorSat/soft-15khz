Attribute VB_Name = "driver_tdfx"
' -= Soft-15kHz - DRIVER - 3Dfx
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


Sub TDfx_RemoveAllModes(AdapterIndex As Byte)
  Dim CurObjectPath As String
  Dim Dummy As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  Dim Key() As String
  Dim SubKey() As String
  Dim Index As Integer
  Dim SubIndex As Integer
  Key = GetAllKeys(HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS")
  On Error Resume Next
  For Index = LBound(Key) To UBound(Key)
    SubKey = GetAllKeys(HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Key(Index))
    For SubIndex = LBound(SubKey) To UBound(SubKey)
      DeleteKey HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Key(Index) & "\" & SubKey(SubIndex)
    Next
    DeleteKey HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Key(Index)
  Next
  If Err Then Err.Clear
  On Error GoTo 0
  Adapter(AdapterIndex).ModeCount = 0
End Sub

Sub TDfx_AddAllModes(AdapterIndex As Byte, Freq As Byte)
  Dim CurObjectPath As String
  Dim Index As Integer
  Dim Dummy As String
  Dim Res() As String
  CurObjectPath = Adapter(AdapterIndex).ObjectPath1
  For Index = 1 To ModeCount
    With Mode(Index)
      If .ModeFreq = Freq Then
        Res = Split(.ModeName, ",")
        If Not Res(0) Mod 8 = 0 Then
          Res(0) = (Res(0) + 8) - (Res(0) Mod 8)
        End If
        Dummy = TDfx_ModelineToBinary(.ModeName, .Modeline)
        If Not Dummy = "" Then
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Res(0) & "," & Res(1)
          CreateKey HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Res(0) & "," & Res(1) & "\" & Res(2) & "Hz"
          SaveSettingString HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Res(0) & "," & Res(1) & "\" & Res(2) & "Hz", "", Dummy
          SaveSettingString HKEY_LOCAL_MACHINE, CurObjectPath & "\TIMINGS\" & Res(0) & "," & Res(1) & "\" & Res(2) & "Hz", "Supported", "BPP+8+16+32,DDRAW"
          Adapter(AdapterIndex).ModeCount = Adapter(AdapterIndex).ModeCount + 1
        End If
      End If
    End With
  Next
End Sub

Function TDfx_ModelineToBinary(ResName As String, Modeline As String) As String
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
  Dim CHKSUM As Long
  
  Dim Res() As String
  
  LineBinary = ""
  LineParam = Split(Modeline, " ", 12)
  
  P_FREQ = CSng(LineParam(2))
  If InStr(1, LCase(LineParam(11)), "doublescan") Then
    M_OPTIONS = M_OPTIONS + 1
  End If
  If InStr(1, LCase(LineParam(11)), "interlace") Then
    '3Dfx + Interlace = No Go!
    Exit Function
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
  
  '-interlace workaround-
  If M_OPTIONS And 2 Then
    If V_TOTAL Mod 2 = 1 Then
      V_TOTAL = V_TOTAL + 1
    End If
    V_ACTIVE = V_ACTIVE / 2
    V_FIRST = V_FIRST / 2
    V_LAST = V_LAST / 2
    V_TOTAL = V_TOTAL / 2
  End If
  '-doublescan workaround-
  If M_OPTIONS And 1 Then
    V_FIRST = V_FIRST * 2
    V_LAST = V_LAST * 2
    V_TOTAL = V_TOTAL * 2
  End If

  If M_OPTIONS And 2 Then
    V_FREQ = Round(((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * (CSng(V_TOTAL) / 2))) * 1000, 3)
  Else
    V_FREQ = Round(((1000000 * (CSng(P_FREQ))) / (CSng(H_TOTAL) * CSng(V_TOTAL))) * 1000, 3)
  End If

  '3Dfx uses String Values :D
  LineBinary = ""
  'H_TOTAL
  CHKSUM = CHKSUM + H_TOTAL
  LineBinary = H_TOTAL
  'H_FIRST
  CHKSUM = CHKSUM + H_FIRST
  LineBinary = LineBinary & "," & H_FIRST
  'H_LAST
  CHKSUM = CHKSUM + H_LAST
  LineBinary = LineBinary & "," & H_LAST
  'V_TOTAL
  CHKSUM = CHKSUM + V_TOTAL
  LineBinary = LineBinary & "," & V_TOTAL
  'V_FIRST
  CHKSUM = CHKSUM + V_FIRST
  LineBinary = LineBinary & "," & V_FIRST
  'V_LAST
  CHKSUM = CHKSUM + V_LAST
  LineBinary = LineBinary & "," & V_LAST
  'M_OPTIONS
  CHKSUM = CHKSUM + M_OPTIONS
  LineBinary = LineBinary & "," & M_OPTIONS
  'P_FREQ
  CHKSUM = CHKSUM + CLng(CLng(P_FREQ * 1000000) / 10000)
  LineBinary = LineBinary & "," & CLng(P_FREQ * 1000000)
  'V_FREQ
  CHKSUM = CHKSUM + CLng(V_FREQ / 10)
  LineBinary = LineBinary & "," & CLng(V_FREQ / 10)
  'NOT_USED
  CHKSUM = CHKSUM + 8
  LineBinary = LineBinary & ",8"
  'CHKSUM
  LineBinary = LineBinary & "," & CHKSUM
  TDfx_ModelineToBinary = LineBinary
End Function

