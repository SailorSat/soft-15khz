Attribute VB_Name = "Core"
Option Explicit

Sub Main()
  Open App.Path & "\custom15khz.txt" For Output As #1
    Print #1, "# -= 15KHz Progressive =-"
    Print #1, GenerateModeline("", 240, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 256, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 288, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 296, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 304, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 320, 240, 60, 15720, False, True)
    Print #1, GenerateModeline("", 336, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 368, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 392, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 448, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 512, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 640, 240, 59.885, 15720, False, True)
    Print #1, GenerateModeline("", 256, 256, 55.456, 15720, False, True)
    Print #1, GenerateModeline("", 320, 256, 55.456, 15720, False, True)
    Print #1, GenerateModeline("", 352, 256, 55.456, 15720, False, True)
    Print #1, GenerateModeline("", 400, 256, 54.686, 15720, False, True)
    Print #1, GenerateModeline("", 256, 264, 54.497, 15720, False, True)
    Print #1, GenerateModeline("", 352, 264, 54.497, 15720, False, True)
    Print #1, GenerateModeline("", 632, 264, 54.497, 15720, False, True)
    Print #1, GenerateModeline("", 352, 288, 50, 15720, False, True)
    Print #1, GenerateModeline("", 384, 288, 50.481, 15720, False, True)
    Print #1, GenerateModeline("", 512, 288, 50.481, 15720, False, True)
    Print #1, GenerateModeline("", 640, 288, 50.481, 15720, False, True)
    Print #1, GenerateModeline("", 768, 224, 60.606, 15720, False, True)
    Print #1, ""
    Print #1, "# -= 15KHz Interlace =-"
    Print #1, GenerateModeline("", 512, 448, 60, 15720, True, True)
    Print #1, GenerateModeline("", 640, 480, 60, 15720, True, True)
    Print #1, GenerateModeline("", 720, 480, 60, 15720, True, True)
    Print #1, GenerateModeline("", 512, 512, 54.497, 15720, True, True)
    Print #1, GenerateModeline("", 800, 600, 50, 15720, True, True)
    Print #1, GenerateModeline("", 1024, 600, 50, 15720, True, True)
    Print #1, ""
    Print #1, "# -= remove fakes =-"
    Print #1, "remove 1024,768"
  Close #1

  Open App.Path & "\custom25khz.txt" For Output As #1
    Print #1, "# -= 25KHz Progressive =-"
    Print #1, GenerateModeline("", 496, 384, 60, 24960, False, True)
    Print #1, GenerateModeline("", 512, 384, 60, 24960, False, True)
    Print #1, GenerateModeline("", 512, 400, 57.349, 24960, False, True)
    Print #1, GenerateModeline("", 640, 480, 50, 24960, False, True)
    Print #1, GenerateModeline("", 720, 480, 50, 24960, False, True)
    Print #1, ""
    Print #1, "# -= 25KHz Interlace =-"
    Print #1, GenerateModeline("", 1024, 768, 60, 24960, True, True)
    Print #1, GenerateModeline("", 1280, 720, 60, 24960, True, True)
  Close #1

  Open App.Path & "\custom31khz.txt" For Output As #1
    Print #1, "# -= 31KHz Progressive =-"
    Print #1, GenerateModeline("", 512, 448, 60, 31500, False, True)
    Print #1, GenerateModeline("", 640, 480, 60, 31500, False, True)
    Print #1, GenerateModeline("", 720, 480, 60, 31500, False, True)
    Print #1, GenerateModeline("", 512, 512, 54.497, 31500, False, True)
    Print #1, GenerateModeline("", 800, 600, 50, 31500, False, True)
    Print #1, GenerateModeline("", 1024, 600, 50, 31500, False, True)
  Close #1

  Open App.Path & "\usermodes.txt" For Output As #1
    Print #1, GenerateModeline("a2600-ntsc", 160, 192, 60, 15720, False, True)
    Print #1, GenerateModeline("a2600-pal", 160, 228, 50, 15720, False, True)
    Print #1, GenerateModeline("neogeo", 320, 224, 59.186, 15720, False, True)
    Print #1, GenerateModeline("cps1", 384, 224, 59.61, 15720, False, True)
    Print #1, GenerateModeline("rtype", 384, 256, 55.018, 15720, False, True)
    Print #1, GenerateModeline("mkombat", 400, 254, 54.707, 15720, False, True)
    Print #1, GenerateModeline("amiga-ntsc", 720, 240, 60, 15720, False, True)
    Print #1, GenerateModeline("amiga-pal", 720, 288, 50, 15720, False, True)
    Print #1, GenerateModeline("model1", 496, 384, 57.524, 24960, False, True)
    Print #1, GenerateModeline("pal", 720, 576, 50, 15720, True, True)
    Print #1, GenerateModeline("720p-25khz", 1280, 720, 50, 24960, True, True)
    
    Print #1, GenerateModeline("2560x224-15khz", 2560, 224, 60, 15720, False, True)
    Print #1, GenerateModeline("2560x240-15khz", 2560, 240, 60, 15720, False, True)
    Print #1, GenerateModeline("2560x256-15khz", 2560, 256, 60, 15720, False, True)
    Print #1, GenerateModeline("2560x288-15khz", 2560, 288, 50, 15720, False, True)
    Print #1, GenerateModeline("2560x300-15khz", 2560, 300, 50, 15720, False, True)
    
  Close #1
End Sub

Function GenerateModeline(Name As String, Width As Long, Height As Long, Refresh As Double, Frequency As Long, Interlace As Boolean, MultipleOf8 As Boolean) As String
  Dim Second As Double
  Dim Line As Double
  Dim Pixel As Double
  
  Dim Index As Long
  
  Dim H_Active As Long
  Dim H_FrontPorch As Long
  Dim H_SyncWidth As Long
  Dim H_BackPorch As Long
  Dim H_Total As Long
  
  Dim V_Active As Long
  Dim V_FrontPorch As Long
  Dim V_SyncWidth As Long
  Dim V_BackPorch As Long
  Dim V_Total As Long
  
  Dim V_Frequency As Double
  Dim H_Frequency As Long
  Dim P_Frequency As Long
  
  Dim pHActive As Double
  Dim pHSyncWidth As Double
  
  Dim pVSyncWidth As Double
  
  Select Case Frequency
    Case 15720
      ' PALi = 625 * 25 = 15625
      ' NTSCi = 525 * 30 = 15750
      ' Arcade (50p) = 312 * 50 = 15600
      ' Arcade (60p) = 262 * 60 = 15720
      
      ' Standard Resolution (based on neogeo!)
      pHActive = 0.833
      pHSyncWidth = 0.073
            
      pVSyncWidth = 0.03
  
    Case 24960
      ' Medium Resolution (based on model1!)
      pHActive = 0.756
      pHSyncWidth = 0.073
      
      pVSyncWidth = 0.03
    
    Case 31500
      ' High Resolution
      pHActive = 0.833
      pHSyncWidth = 0.073
      
      pVSyncWidth = 0.03
      
    Case Else
      ' Something other?
      pHActive = 0.833
      pHSyncWidth = 0.073
      
      pVSyncWidth = 0.03
  End Select
  
  H_Active = Width
  H_Total = H_Active / pHActive
  H_SyncWidth = H_Total * pHSyncWidth
  H_BackPorch = H_SyncWidth
  H_FrontPorch = H_Total - H_BackPorch - H_SyncWidth - H_Active
  
  V_Active = Height
  V_Total = Frequency / IIf(Interlace, Refresh / 2, Refresh)
  V_SyncWidth = V_Total * pVSyncWidth
  V_BackPorch = (V_Total - V_Active - V_SyncWidth) / 2 + 0.5
  V_FrontPorch = V_Total - V_Active - V_SyncWidth - V_BackPorch

  If MultipleOf8 Then
    H_Total = 8 + H_Total - (H_Total Mod 8)
    H_FrontPorch = 8 + H_FrontPorch - (H_FrontPorch Mod 8)
    H_SyncWidth = 8 + H_SyncWidth - (H_SyncWidth Mod 8)
    H_BackPorch = 8 + H_BackPorch - (H_BackPorch Mod 8)
  End If
  
  V_Frequency = Frequency / V_Total
  P_Frequency = H_Total * V_Total * V_Frequency
  
  Second = 1000000
  Line = Second / Frequency
  Pixel = Line / H_Total
  If Interlace Then Line = Line / 2
  
  If H_FrontPorch > H_BackPorch Then
    Index = H_FrontPorch
    H_FrontPorch = H_BackPorch
    H_BackPorch = Index
  End If
  
  If V_FrontPorch > V_BackPorch Then
    Index = V_FrontPorch
    V_FrontPorch = V_BackPorch
    V_BackPorch = Index
  End If

  Debug.Print "--- " & IIf(Name <> "", Name & "-", "") & Width & "x" & Height & " " & Refresh & "Hz ---"
  Debug.Print "Horizontal Front Porch: " & Format(Pixel * H_FrontPorch, "0.0#") & " 탎, " & H_FrontPorch & " pix"
  Debug.Print "Horizontal Sync Width : " & Format(Pixel * H_SyncWidth, "0.0#") & " 탎, " & H_SyncWidth & " pix"
  Debug.Print "Horizontal Back Porch : " & Format(Pixel * H_BackPorch, "0.0#") & " 탎, " & H_BackPorch & " pix"
  Debug.Print ""
  Debug.Print "Vertical Front Porch  : " & Format(Line * V_FrontPorch, "0.0#") & " 탎, " & V_FrontPorch & " lin"
  Debug.Print "Vertical Sync Width   : " & Format(Line * V_SyncWidth, "0.0#") & " 탎, " & V_SyncWidth & " lin"
  Debug.Print "Vertical Back Porch   : " & Format(Line * V_BackPorch, "0.0#") & " 탎, " & V_BackPorch & " lin"
  Debug.Print ""
  Debug.Print "modeline '" & Width & "x" & Height & "-" & Format(Frequency / 1000, "#.0##") & "kHz-" & Format((IIf(Interlace, V_Frequency * 2, V_Frequency)), "#.0##") & "Hz' " & Format(P_Frequency / 1000000, "#.0##") & " " & H_Active & " " & H_Active + H_FrontPorch & " " & H_Active + H_FrontPorch + H_SyncWidth & " " & H_Total & " " & V_Active & " " & V_Active + V_FrontPorch & " " & V_Active + V_FrontPorch + V_SyncWidth & " " & V_Total & " -hsync -vsync" & IIf(Interlace, " interlace", "")
  Debug.Print ""
  
  GenerateModeline = "modeline '" & IIf(Name <> "", Name & "-", "") & Width & "x" & Height & "-" & Format(Frequency / 1000, "#.0##") & "kHz-" & Format((IIf(Interlace, V_Frequency * 2, V_Frequency)), "#.0##") & "Hz' " & Format(P_Frequency / 1000000, "#.0##") & " " & H_Active & " " & H_Active + H_FrontPorch & " " & H_Active + H_FrontPorch + H_SyncWidth & " " & H_Total & " " & V_Active & " " & V_Active + V_FrontPorch & " " & V_Active + V_FrontPorch + V_SyncWidth & " " & V_Total & " -hsync -vsync" & IIf(Interlace, " interlace", "")
End Function
