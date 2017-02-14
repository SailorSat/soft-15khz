Attribute VB_Name = "base_main"
' -= Soft-15kHz - BASE - Main
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


' -= Types =-
Type tHost
  Windows As Byte                   ' Windows Version (1 = 9x/ME, 5 = 2000/XP, 6 = Vista/Seven)
  DotSeparator As Boolean           ' TRUE = 3.1, FALSE = 3,1
  RuntimeFlags As Byte              ' Runtime Flags (0 = Default, 1 = Skip Backups, 2 = No Warnings)
  CommandLine As String             ' Commands to execute
End Type

Type tMode
  ModeFreq As Byte                  ' Targeted Frequency (15 / 25 / 31 / 255 = USER)
  ModeName As String                ' Resolution "Name" = X,Y,V [X Size, Y Size, V Refresh]
  Modeline As String                ' Modeline itself
End Type

Type tAdapter
  Soft15KHz As Byte                 ' Soft-15kHz Flag (1 = 15, 2 = 25, 4 = 31, 128 = USER)
  AdapterName As String             ' Driver Name
  ModeCount As Integer              ' Counter for installed Modes
  ObjectPath1 As String             ' Primary Registry Path
  ObjectPath2 As String             ' Secondary Registry Path (if available)
  DriverType As Byte                ' Driver Type (0 = Unsupported, 1 = NVidia, 2 = ATI Catalyst, 3 = Matrox PowerDesk, 4 = 3Dfx, 5 = Intel)
  DriverVersion As Byte             ' Driver SubType / Version (0 = Unsupported, [3Dfx] 1 = Supported, [ATI / Matrox / NVidia] # = Version, [Intel] 1 = GMA, 2 = EGD)
  DriverString As String            ' Driver "Name", usualy contains the version as string
  ChipType As String                ' Chipset used
  ChipFamily As Integer             ' Chipset ID if known and needed
  ChipFlags As Byte                 ' Flags (0 = Default, 1 = No 320x Modes, 2 = Multiple of 8 Modes, 4 = No Interlace)
End Type


' -= Objects =-
Public Host As tHost

Public Mode() As tMode
Public ModeCount As Integer

Public Adapter() As tAdapter
Public AdapterCount As Byte


' -= Main Functions =-
Sub Main()
  ' -= Step 1 =-
  CheckWindowsVersion
  CheckDecimalSeperator
  CheckCommandLine
  
  ' -= Step 2 =-
  GenerateModeTable
  ReadCustomModes 15, "custom15khz.txt"
  ReadCustomModes 25, "custom25khz.txt"
  ReadCustomModes 31, "custom31khz.txt"
  ReadCustomModes 255, "usermodes.txt"

  ' -= Step 3 =-
  GenerateAdapterTable
  
  ' -= Step 4 =-
  ShowGUI
End Sub


' -= Service Functions =-
Private Sub CheckWindowsVersion()
  Dim Result As Integer
  Dim OSInfo As OSVERSIONINFO
  Host.Windows = 0
  OSInfo.dwOSVersionInfoSize = 148
  OSInfo.szCSDVersion = Space(128)
  Result = GetVersionExA(OSInfo)
  With OSInfo
    Select Case .dwPlatformId
      Case 1
        ' -= Win 4.x = 9x/Me =-
        If .dwMajorVersion = 4 Then Host.Windows = 1
      Case 2
        ' -= NT 5.x = 2000/XP =-
        If .dwMajorVersion = 5 Then Host.Windows = 5
        ' -= NT 6.x = Vista/Seven =-
        If .dwMajorVersion = 6 Then Host.Windows = 6
    End Select
  End With
#If DEBUGMODE = 2 Then
  Host.Windows = 1
#ElseIf DEBUGMODE = 3 Then
  Host.Windows = 6
#End If
End Sub

Private Sub CheckDecimalSeperator()
  If Mid$(1 / 2, 2, 1) = "." Then
    Host.DotSeparator = True
  Else
    Host.DotSeparator = False
  End If
End Sub

Private Sub CheckCommandLine()
  Host.RuntimeFlags = 0
  Host.CommandLine = Command
  '-- Placeholder for now --
  If Host.CommandLine <> "" Then Host.RuntimeFlags = 2
End Sub

Private Sub ShowGUI()
  base_gui.Show
End Sub

Private Sub GenerateModeTable()
  ReDim Mode(0)
  ModeCount = 0
  
  ' -= 15KHz Progressive =-
  AddModeToTable 15, "240,240,60", "modeline '240x240' 4,83 240 252 276 310 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "256,240,60", "modeline '256x240' 5,30 256 272 296 336 240 244 247 261 -hsync -vsync"
  AddModeToTable 15, "256,256,60", "modeline '256x256' 5,36 256 268 292 330 256 257 260 273 -hsync -vsync"
  AddModeToTable 15, "256,264,60", "modeline '256x264' 5,35 256 268 292 330 264 265 268 278 -hsync -vsync"
  AddModeToTable 15, "288,240,60", "modeline '288x240' 5,84 288 296 328 368 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "296,240,60", "modeline '296x240' 5,95 296 304 336 376 240 243 246 264 -hsync -vsync"
  AddModeToTable 15, "304,240,60", "modeline '304x240' 6,20 304 320 352 396 240 243 246 264 -hsync -vsync"
  AddModeToTable 15, "320,240,60", "modeline '320x240' 6,45 320 336 368 414 240 242 245 264 -hsync -vsync"
  AddModeToTable 15, "320,256,60", "modeline '320x256' 6,68 320 340 372 416 256 257 260 268 -hsync -vsync"
  AddModeToTable 15, "336,240,60", "modeline '336x240' 6,83 336 352 384 433 240 243 246 264 -hsync -vsync"
  AddModeToTable 15, "352,256,60", "modeline '352x256' 7,28 352 368 400 450 256 257 260 271 -hsync -vsync"
  AddModeToTable 15, "352,264,60", "modeline '352x264' 7,35 352 365 405 452 264 265 268 284 -hsync -vsync"
  AddModeToTable 15, "352,288,60", "modeline '352x288' 7,40 352 368 408 464 288 289 292 312 -hsync -vsync"
  AddModeToTable 15, "368,240,60", "modeline '368x240' 7,47 368 384 424 478 240 243 246 264 -hsync -vsync"
  AddModeToTable 15, "384,288,60", "modeline '384x288' 7,85 384 400 440 496 288 289 292 309 -hsync -vsync"
  AddModeToTable 15, "392,240,60", "modeline '392x240' 8,00 392 408 448 504 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "400,256,60", "modeline '400x256' 8,08 400 416 456 519 256 268 271 297 -hsync -vsync"
  AddModeToTable 15, "448,240,60", "modeline '448x240' 9,16 448 464 512 576 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "512,240,60", "modeline '512x240' 10,68 512 544 600 672 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "512,288,60", "modeline '512x288' 10,68 512 544 600 672 288 289 292 312 -hsync -vsync"
  AddModeToTable 15, "632,264,60", "modeline '632x264' 13,00 632 664 728 824 264 265 268 278 -hsync -vsync"
  AddModeToTable 15, "640,240,60", "modeline '640x240' 13,22 640 672 736 832 240 243 246 265 -hsync -vsync"
  AddModeToTable 15, "640,288,60", "modeline '640x288' 13,10 640 672 736 832 288 289 292 309 -hsync -vsync"
  
  ' -= 15KHz Interlace =-
  AddModeToTable 15, "512,448,60", "modeline '512x448' 10,60 512 542 598 672 448 469 472 527 interlace -hsync -vsync"
  AddModeToTable 15, "512,512,60", "modeline '512x512' 10,60 512 538 594 668 512 513 516 545 interlace -hsync -vsync"
  AddModeToTable 15, "640,480,60", "modeline '640x480' 13,09 640 672 736 836 480 486 489 525 interlace -hsync -vsync"
  AddModeToTable 15, "720,480,60", "modeline '720x480' 14,60 720 752 824 928 480 486 489 525 interlace -hsync -vsync"
  AddModeToTable 15, "800,600,60", "modeline '800x600' 16,48 800 840 920 1040 600 602 605 627 interlace -hsync -vsync"
  AddModeToTable 15, "1024,768,60", "modeline '1024x600' 20,90 1024 1072 1176 1328 600 607 610 627 interlace -hsync -vsync"
  
  ' -= 25KHz Progressive =-
  AddModeToTable 25, "496,384,60", "modeline '496x384' 15,4752 496 508 570 620 384 388 391 416 -hsync -vsync"
  AddModeToTable 25, "512,384,60", "modeline '512x384' 14,75 512 520 568 600 384 388 391 410 -hsync -vsync"

  ' -= 31KHz Progressive =-
  AddModeToTable 31, "512,448,60", "modeline '512x448' 21,21 512 542 598 672 448 469 472 527 -hsync -vsync"
  AddModeToTable 31, "512,512,60", "modeline '512x512' 21,21 512 538 594 668 512 513 516 545 -hsync -vsync"
  AddModeToTable 31, "640,480,60", "modeline '640x480' 26,18 640 672 736 836 480 486 489 525 -hsync -vsync"
  AddModeToTable 31, "720,480,60", "modeline '720x480' 29,25 720 752 824 928 480 486 489 526 -hsync -vsync"
  AddModeToTable 31, "800,600,60", "modeline '800x600' 32,96 800 840 920 1040 600 602 605 627 -hsync -vsync"
  AddModeToTable 31, "1024,768,60", "modeline '1024x600' 41,80 1024 1072 1176 1328 600 607 610 627 -hsync -vsync"
End Sub

Private Sub AddModeToTable(ModeFreq As Byte, ModeName As String, Modeline As String)
  Dim Param() As String
  Param = Split(Modeline, " ", 12)
  If Host.DotSeparator Then
    Param(2) = Replace(Param(2), ",", ".")
  Else
    Param(2) = Replace(Param(2), ".", ",")
  End If
  Modeline = Join(Param, " ")
  
  ModeCount = ModeCount + 1
  If ModeCount > UBound(Mode) Then
    ReDim Preserve Mode(ModeCount + 4)
  End If
  With Mode(ModeCount)
    .ModeFreq = ModeFreq
    .ModeName = ModeName
    .Modeline = Modeline
  End With
End Sub

Private Sub RemoveModeFromTable(ModeFreq As Byte, ModeName As String)
  Dim Index1 As Integer
  Dim Index2 As Integer
  For Index1 = 1 To ModeCount
    If Mode(Index1).ModeFreq = ModeFreq And Mode(Index1).ModeName = ModeName Then
      For Index2 = Index1 + 1 To ModeCount
        Mode(Index2 - 1) = Mode(Index2)
      Next
      ModeCount = ModeCount - 1
      Exit Sub
    End If
  Next
End Sub

Private Sub ReadCustomModes(ModeFreq As Byte, FileName As String)
  Dim Line As String
  Dim Res(0 To 2) As String
  Dim Param() As String
  Dim Index1 As Integer
  Dim Index2 As Integer
  Dim Long1 As Long
  Dim Long2 As Long
  If Dir(App.Path & "\" & FileName) = "" Then Exit Sub
  On Error Resume Next
  Open App.Path & "\" & FileName For Input As #1
  If Err Then
    Err.Clear
    '-- ErrorHandler here! --
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  While Not EOF(1)
    Line Input #1, Line
    Line = LCase(Line)
    Line = Replace(Line, Chr(34), "'")
    Index1 = InStr(1, Line, "'")
    If Index1 > 0 Then
      Index2 = InStr(Index1 + 1, Line, "'")
      Line = Left(Line, Index1) & Replace(Mid(Line, Index1 + 1, Index2 - Index1 - 1), " ", "_") & Mid(Line, Index2)
    End If
    While InStr(1, Line, "  ")
      Line = Replace(Line, "  ", " ")
    Wend
    Param = Split(Line, " ", 12)
    If Not Line = "" Then
      If Param(0) = "modeline" Then
        If UBound(Param) >= 10 Then
          Res(0) = Param(3)
          Res(1) = Param(7)
          If FileName = "usermodes.txt" Then
            If Host.DotSeparator Then
              Param(2) = Replace(Param(2), ",", ".")
            Else
              Param(2) = Replace(Param(2), ".", ",")
            End If
            Long1 = Param(2) * 1000000
            Long2 = Param(6) * Param(10)
            Res(2) = CInt(Long1 / Long2)
            If InStr(1, Param(11), "interlace") Then Res(2) = Res(2) * 2
            If InStr(1, Param(11), "doublescan") Then Res(2) = Res(2) / 2
          Else
            Res(2) = "60"
          End If
          Line = Join(Param, " ")
          AddModeToTable ModeFreq, Join(Res, ","), Line
        End If
      ElseIf Param(0) = "remove" Then
        If UBound(Param) >= 1 Then
          Param(1) = Replace(Param(1), ".", ",")
          Param(1) = Replace(Param(1), "x", ",")
          Param(1) = Replace(Param(1), "*", ",")
          RemoveModeFromTable ModeFreq, Param(1) & ",60"
        End If
      End If
    End If
  Wend
  Close #1
End Sub

Private Sub GenerateAdapterTable()
  If Host.Windows = 1 Then
    GenerateAdapterTable_Win9x
  Else
    GenerateAdapterTable_WinNT
  End If
End Sub

Private Sub GenerateAdapterTable_Win9x()
  Dim MaxObjectNumber As Integer
  Dim CurObjectNumber As Integer
  Dim CurObjectPath As String
  Dim Key() As String
  
  Dim Index As Integer
  Dim Dummy As String
  
  
  Key = GetAllKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Display")
  MaxObjectNumber = UBound(Key)
  ReDim Adapter(MaxObjectNumber + 1)
  
  For CurObjectNumber = 0 To MaxObjectNumber
    CurObjectPath = "System\CurrentControlSet\Services\Class\Display\" & LeadZero(CStr(CurObjectNumber), 4)
    AdapterCount = AdapterCount + 1
    With Adapter(AdapterCount)
      .ObjectPath1 = CurObjectPath
      .ObjectPath2 = ""
    
      .Soft15KHz = GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "___Soft-15KHz___", 0)
    
      If GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Catalyst_Version", "") <> "" Then
        ' -= ATI Catalyst =-
        .DriverType = 2
        .ObjectPath2 = .ObjectPath1
        .ObjectPath1 = "Software\ATI Technologies\Driver\" & LeadZero(CStr(CurObjectNumber), 4) & "\DAL"
        
        .Soft15KHz = GetSettingLong(HKEY_LOCAL_MACHINE, .ObjectPath1, "___Soft-15KHz___", 0)
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "ReleaseVersion", "")
        Index = InStr(1, Dummy, "-")
        If Index > 0 Then Dummy = Left(Dummy, Index - 1)
        .DriverString = Dummy
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Catalyst_Version", "1.0")
        If Dummy = "0" Then Dummy = "1.0"
        While InStr(1, Dummy, " ")
          Dummy = Mid(Dummy, InStr(1, Dummy, " ") + 1)
        Wend
        Index = InStr(1, Dummy, ".")
        If Index > 0 Then
          If Not Dummy = "1.0" Then .DriverString = .DriverString & " (Catalyst " & CInt(Left(Dummy, Index - 1)) & "." & CInt(Mid(Dummy, Index + 1)) & ")"
          Dummy = Left(Dummy, Index - 1)
        End If
        .DriverVersion = CInt(Dummy)
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        .ChipType = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath & "\INFO", "ChipType", "")
        
        .ChipFlags = 1
      Else
        'Sonstiges
        .DriverType = 0
        .DriverVersion = 0
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        If Dummy = "" Then Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", "")
        .AdapterName = Replace(Dummy, Chr(0), " ")
      End If
    End With
  Next
End Sub

Private Sub GenerateAdapterTable_WinNT()
  Dim MaxObjectNumber As Integer
  Dim CurObjectNumber As Integer
  Dim CurObjectPath As String
  
  Dim Dummy As String
  Dim Temp As String
  Dim Index As Integer
  
  ' -= Read Count =-
  MaxObjectNumber = GetSettingLong(HKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\VIDEO", "MaxObjectNumber", 3)
  ReDim Adapter(MaxObjectNumber + 1)
  
  ' -= Read Adapters =-
  For CurObjectNumber = 0 To MaxObjectNumber
    CurObjectPath = GetSettingString(HKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\VIDEO", "\Device\Video" & CurObjectNumber, "")
    CurObjectPath = Replace(LCase(CurObjectPath), "\registry\machine\", "")
    CurObjectPath = Replace(CurObjectPath, Chr(0), "")
    AdapterCount = AdapterCount + 1
    
#If DEBUGMODE = 1 Then
    Select Case AdapterCount
      Case 1
        ' SYSTEM\CurrentControlSet\Control\Video\{7A741D44-029A-49A2-9318-3CF233EF2299}\0000
        CurObjectPath = LCase("SYSTEM\CurrentControlSet\Control\Video\{7A741D44-029A-49A2-9318-3CF233EF2299}\0000")
    End Select
#End If
    
    With Adapter(AdapterCount)
      .ObjectPath1 = CurObjectPath
      .ObjectPath2 = ""
      .Soft15KHz = GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "___Soft-15KHz___", 0)
    
      If GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "NV_Modes", "!") <> "!" Then
        ' -= NVidia ForceWare =-
        .DriverType = 1
        
        Dummy = Replace(Replace(Right(GetFileVersionInformation(GetSpecialFolder(41) & "\" & Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", ""), Chr(0), "") & ".dll"), 6), ".", ""), ",", "")
        If Dummy = "" Then
          Dummy = Replace(Replace(Right(GetFileVersionInformation(GetSpecialFolder(41) & "\nvapi.dll"), 6), ".", ""), ",", "")
          If Dummy = "" Then Dummy = "99999"
          ' -- Add Message later --
        End If
        If CLng(Dummy) < 6693 Then
          .DriverVersion = 0
          If Not Host.RuntimeFlags And 2 Then MsgBox "Your driver appears to be too old." & vbCrLf & "NVidia ForceWare 66.93 or newer required.", vbCritical + vbOKOnly, "Soft-15kHz"
        Else
          .DriverVersion = 1
        End If
        .DriverString = CLng(Left(Dummy, 3)) & "." & Right(Dummy, 2)
        
        .AdapterName = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", ""), Chr(0), "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
        
        If Host.Windows = 6 Then
          Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "CustomDisplay", "")
          Select Case Len(Dummy)
            Case 5632
              '175.19
              .DriverVersion = 6
            Case 8704
              .DriverVersion = 7
            Case 8960
              .DriverVersion = 8
            Case Else
              ' -- Most likely not correct! --
              .DriverVersion = 8
          End Select
        End If
        If InStr(1, .ChipType, "GeForce") Then
          If InStr(1, .ChipType, "MX") > 0 Then
            ' -- Add Message later --
            .ChipFlags = 4
          End If
          ' -- Check later --
          Dummy = Replace(Replace(Replace(Replace(Replace(.ChipType, "FX", ""), "GTX ", "10"), "GeForce", ""), "GTS ", "10"), "GT ", "10")
          If InStr(2, Dummy, " ") > 0 Then
            Dummy = Mid(Dummy, 2, InStr(2, Dummy, " ") - 2)
          End If
          If IsNumeric(Dummy) Then
            Index = CInt(Dummy)
            If Index > 7999 Then
              If Not Host.RuntimeFlags And 2 Then MsgBox "Please note that GeForce 8 series, GeForce 9 series and GeForce GTX series cards most likely don´t work without the EDID dongle.", vbInformation + vbOKOnly, "Soft-15kHz"
            End If
            .ChipFamily = Index
          End If
        End If
      ElseIf GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Catalyst_Version", "") <> "" Then
        ' -= ATI Catalyst =-
        .DriverType = 2
        
        Dummy = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "ReleaseVersion", ""), Chr(0), "")
        Index = InStr(1, Dummy, "-")
        If Index > 0 Then Dummy = Left(Dummy, Index - 1)
        .DriverString = Dummy
        
        Dummy = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Catalyst_Version", "1.0"), Chr(0), "")
        If Dummy = "0" Then Dummy = "1.0"
        While InStr(1, Dummy, " ")
          Dummy = Mid(Dummy, InStr(1, Dummy, " ") + 1)
        Wend
        Index = InStr(1, Dummy, ".")
        If Index > 0 Then
          If Not Dummy = "1.0" Then .DriverString = .DriverString & " (Catalyst " & CInt(Left(Dummy, Index - 1)) & "." & CInt(Mid(Dummy, Index + 1)) & ")"
          Temp = Mid(Dummy, Index + 1)
          Dummy = Left(Dummy, Index - 1)
        Else
          Temp = "0"
        End If
        .DriverVersion = CInt(Dummy)
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        Dummy = LCase(Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.AdapterString", ""), Chr(0), ""))
        Dummy = Replace(Dummy, "mobility ", "")
        Dummy = Replace(Dummy, " series", "")
        Dummy = Replace(Dummy, "ati ", "")
        Dummy = Replace(Dummy, "amd ", "")
        Dummy = Replace(Dummy, "radeon ", "")
        Dummy = Replace(Dummy, "x1", "11")
        Dummy = Replace(Dummy, "x", "10")
        Dummy = Replace(Dummy, "hd ", "1")
        Index = InStr(1, Dummy, "/")
        If Index > 0 Then Dummy = Left(Dummy, Index - 1)
        Index = InStr(1, Dummy, " ")
        If Index > 0 Then Dummy = Left(Dummy, Index - 1)
        If IsNumeric(Dummy) Then .ChipFamily = CInt(Dummy) Else .ChipFamily = 7000
        If .ChipFamily > 14000 Then
          ' -= HD4000 and newer =-
          .ChipFlags = 0
        Else
          If .ChipFamily < 11000 Then
            If .DriverVersion > 6 Or (.DriverVersion = 6 And CInt(Temp) > 5) Then
              If Not Host.RuntimeFlags And 2 Then MsgBox "For pre-X1000 series Radeon cards Catalyst 6.5 is recommended.", vbOKOnly + vbInformation, "Soft-15kHz"
            End If
          End If
          .ChipFlags = 1
        End If
        
        Dummy = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
        .ChipType = Dummy
        
        If InStr(1, .AdapterName, "ArcadeVGA") Then
          If Not Host.RuntimeFlags And 2 Then MsgBox "ArcadeVGA cards are not supported.", vbOKOnly + vbExclamation, "Soft-15kHz"
          .DriverVersion = 0
        End If
      ElseIf GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Mga.SingleResolutions", "!") <> "!" Then
        ' -= Matrox PowerDesk =-
        .DriverType = 3
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "PackageVer", "1.00.00")
        .DriverString = Dummy
        Index = InStr(1, Dummy, ".")
        If Index > 0 Then Dummy = Left(Dummy, Index - 1)
        If IsNumeric(Dummy) Then .DriverVersion = CInt(Dummy)
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        .AdapterName = Replace(Dummy, Chr(0), "")
        
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
      ElseIf GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "3dfxToolsAPI", "!") <> "!" Then
        ' -= 3Dfx =-
        .DriverType = 4
        .DriverVersion = 1
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", "")
        Dummy = Left(Dummy, InStr(1, Dummy, Chr(0)) - 1)
        .DriverString = GetFileVersionInformation(GetSpecialFolder(41) & "\" & Dummy & ".dll")
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
        .ChipFlags = 2 + 4
      ElseIf GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "TotalDTDCount", -1) <> -1 Then
        ' -= Intel GMA =-
        .DriverType = 5
        .DriverVersion = 1
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        Dummy = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
        .ChipType = Replace(Replace(Replace(Replace(Dummy, " Chipset", ""), " Family", ""), " Express", ""), "(R)", "")
      ElseIf Left(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", ""), 4) = "iegd" Then
        ' -= Intel EGD =-
        .DriverType = 5
        .DriverVersion = 2
        
        Dummy = Hex(GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "PcfVersion", 256))
        .DriverString = Left(Dummy, Len(Dummy) - 2) & "." & Right(Dummy, 2)
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
      ElseIf Left(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", ""), 4) = "igdd" Then
        ' -= Intel EGD 5.1 =-
        .DriverType = 5
        .DriverVersion = 2
        
        Dummy = Hex(GetSettingLong(HKEY_LOCAL_MACHINE, CurObjectPath, "PcfVersion", 256))
        .DriverString = Left(Dummy, Len(Dummy) - 2) & "." & Right(Dummy, 2)
        
        .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If .AdapterName = "" Then .AdapterName = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
      Else
        'Sonstiges
        .DriverType = 0
        .DriverVersion = 0
        
        Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "Device Description", "")
        If Dummy = "" Then Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "DriverDesc", "")
        If Dummy = "" Then Dummy = GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "InstalledDisplayDrivers", "")
        .AdapterName = Replace(Dummy, Chr(0), " ")
        .ChipType = Replace(GetSettingString(HKEY_LOCAL_MACHINE, CurObjectPath, "HardwareInformation.ChipType", ""), Chr(0), "")
      End If
    End With
  Next
End Sub

Sub SaveRegistryBackup(CurObjectNumber As Byte)
  Dim CurObjectPath As String
  Dim Dummy As String
  CurObjectPath = Adapter(CurObjectNumber).ObjectPath1
  If InStr(1, CurObjectPath, "{") Then
    Dummy = Mid(CurObjectPath, InStr(1, CurObjectPath, "{") + 1)
    Dummy = Replace(Dummy, "}\", "--")
  Else
    Dummy = Mid(CurObjectPath, InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\") - 1) - 1) + 1)
    Dummy = Replace(Dummy, "\", "-")
  End If
  If Not Dir(App.Path & "\backup_" & Dummy & ".reg") = "" Then
    Kill App.Path & "\backup_" & Dummy & ".reg"
    While Not Dir(App.Path & "\backup_" & Dummy & ".reg") = ""
      Sleep 250
    Wend
  End If
  ShellExecuteA 0, "open", "regedit.exe", "/e " & Chr(34) & App.Path & "\backup_" & Dummy & ".reg" & Chr(34) & " " & Chr(34) & "HKEY_LOCAL_MACHINE\" & CurObjectPath & Chr(34), vbNullString, 1
  While Dir(App.Path & "\backup_" & Dummy & ".reg") = ""
    Sleep 250
  Wend
  If Host.Windows = 1 Then
    CurObjectPath = Adapter(CurObjectNumber).ObjectPath2
    If InStr(1, CurObjectPath, "{") Then
      Dummy = Mid(CurObjectPath, InStr(1, CurObjectPath, "{") + 1)
      Dummy = Replace(Dummy, "}\", "--")
    Else
      Dummy = Mid(CurObjectPath, InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\") - 1) - 1) + 1)
      Dummy = Replace(Dummy, "\", "-")
    End If
    If Not Dir(App.Path & "\backup_" & Dummy & ".reg") = "" Then
      Kill App.Path & "\backup_" & Dummy & ".reg"
      While Not Dir(App.Path & "\backup_" & Dummy & ".reg") = ""
        Sleep 250
      Wend
    End If
    ShellExecuteA 0, "open", "regedit.exe", "/e " & Chr(34) & App.Path & "\backup_" & Dummy & ".reg" & Chr(34) & " " & Chr(34) & "HKEY_LOCAL_MACHINE\" & CurObjectPath & Chr(34), vbNullString, 1
    While Dir(App.Path & "\backup_" & Dummy & ".reg") = ""
      Sleep 250
    Wend
  End If
End Sub

Sub ReadRegistryBackup(CurObjectNumber As Byte)
  Dim CurObjectPath As String
  Dim Dummy As String
  CurObjectPath = Adapter(CurObjectNumber).ObjectPath1
  If InStr(1, CurObjectPath, "{") Then
    Dummy = Mid(CurObjectPath, InStr(1, CurObjectPath, "{") + 1)
    Dummy = Replace(Dummy, "}\", "--")
  Else
    Dummy = Mid(CurObjectPath, InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\") - 1) - 1) + 1)
    Dummy = Replace(Dummy, "\", "-")
  End If
  ShellExecuteA 0, "open", "regedit.exe", "/s " & Chr(34) & App.Path & "\backup_" & Dummy & ".reg" & Chr(34), vbNullString, 1
  Sleep 500
  If Host.Windows = 1 Then
    CurObjectPath = Adapter(CurObjectNumber).ObjectPath2
    If InStr(1, CurObjectPath, "{") Then
      Dummy = Mid(CurObjectPath, InStr(1, CurObjectPath, "{") + 1)
      Dummy = Replace(Dummy, "}\", "--")
    Else
      Dummy = Mid(CurObjectPath, InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\", InStrRev(CurObjectPath, "\") - 1) - 1) + 1)
      Dummy = Replace(Dummy, "\", "-")
    End If
    ShellExecuteA 0, "open", "regedit.exe", "/s " & Chr(34) & App.Path & "\backup_" & Dummy & ".reg" & Chr(34), vbNullString, 1
    Sleep 500
  End If
End Sub

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

Function HexToBits(Hex As String)
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

Function BitsToHex(Bits As String)
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

Function LeadZero(Data As String, Length As Integer) As String
  Dim Total As Integer
  Total = Length - Len(Data)
  LeadZero = String(Total, "0") & Data
End Function

