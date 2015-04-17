VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Binary As String

Dim Binary_Main As String
Dim Binary_Ext1 As String


Private Sub Command1_Click()
  LoadDAT
  If Len(Binary) = 256 Then
    Binary_Main = Binary
    Binary_Ext1 = ""
  ElseIf Len(Binary) = 512 Then
    Binary_Main = Left(Binary, 256)
    Binary_Ext1 = Mid(Binary, 257)
  End If
End Sub

Private Sub Command2_Click()
  SaveDAT
End Sub

Private Sub Command3_Click()
  CreateDAT_Main
End Sub


Sub LoadDAT()
  Dim Line As String
  Open App.Path & "\soft15khz_old.dat" For Input As #1
  ' Header
  Binary = ""
  Line Input #1, Line
  Line Input #1, Line
  Line Input #1, Line
  ' Load Data
  While Not EOF(1)
    Line Input #1, Line
    Binary = Binary & Replace(Mid(Line, 6), " ", "")
  Wend
  Close #1
End Sub

Sub SaveDAT()
  Dim Offset As Integer, Line As String, Index As Integer
  Open App.Path & "\soft15khz_new.dat" For Output As #1
  ' Header
  Print #1, "EDID BYTES:"
  Print #1, "0x   00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F"
  Print #1, "    ------------------------------------------------"
  ' Save Data
  Offset = 0
  Line = ""
  Binary = Binary_Main & Binary_Ext1
  While Offset < Len(Binary)
    If Offset Mod 32 = 0 Then
      Line = Hex(Offset / 2)
      If Line = "0" Then Line = "00"
      Line = Line & " |"
      For Index = 1 To 32 Step 2
        Line = Line & " " & Mid(Binary, Offset + Index, 2)
      Next
      Print #1, Line
    End If
    Offset = Offset + 2
  Wend
  Close #1
  
  Open App.Path & "\soft15khz_new.bin" For Binary As #1
  Offset = 0
  While Offset < Len(Binary)
    Put #1, , CByte("&H" & Mid(Binary, Offset + 1, 2))
    Offset = Offset + 2
  Wend
  Close #1
  
  Open App.Path & "\soft15khz_new.reg" For Output As #1
  Print #1, "Windows Registry Editor Version 5.00"
  Print #1, ""
  Print #1, "[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\DISPLAY\DEL3000\5&18647256&2&11335578&06&00\Device Parameters]"
  Line = Chr(34) & "EDID" & Chr(34) & "=hex:"
  Offset = 0
  While Offset < Len(Binary)
    Line = Line & Mid(Binary, Offset + 1, 2) & ","
    Offset = Offset + 2
  Wend
  Print #1, Left(Line, Len(Line) - 1)
  Close #1
End Sub

Function CalcChecksum(RAW) As Integer
  Dim Data As String, Checksum As Integer
  Data = Left(RAW, 254)
  While Data <> ""
    Checksum = Checksum + CInt("&H" & Left(Data, 2))
    Data = Mid(Data, 3)
  Wend
  Checksum = Checksum Mod 256
  Checksum = 256 - Checksum
  CalcChecksum = Checksum
End Function

Sub CreateDAT_Main()
  Dim RAW As String, Checksum As Integer
  RAW = ""
  
  ' - Header - (8 Bytes)
  RAW = RAW & "00 FF FF FF FF FF FF 00"
  
  ' - Product Information - (10 Bytes)
  ' Manufactor ID
  RAW = RAW & "41 D0"
  ' Product ID
  RAW = RAW & "15 08"
  ' Serial ID
  RAW = RAW & "D4 4A 0C 66"
  ' Week of Manufact
  RAW = RAW & "2E"
  ' Year of Manufact
  RAW = RAW & "12"
  
  ' - EDID Format - (2 Bytes)
  ' Version
  RAW = RAW & "01"
  ' Revision
  RAW = RAW & "03"
  
  ' - Display Parameters - (5 Bytes)
  ' Video Input
  RAW = RAW & "0C"
  ' Max Horizontal Size
  RAW = RAW & "34"
  ' Max Vertical Size
  RAW = RAW & "27"
  ' Gamma
  RAW = RAW & "78"
  ' Features
  RAW = RAW & "0A"
  
  ' - Colors - (10 Bytes)
  RAW = RAW & "20 79 A0 56 48 9A 26 12 48 4C"
  
  ' - Established Timings - (3 Bytes)
  'RAW = RAW & "21 08 00" ' ; 640x480, 800x600, 1024x768
  'RAW = RAW & "20 00 00" ' ; 640x480
  RAW = RAW & "00 00 00" ' ; none
  
  ' - Standard Timings - (16 Bytes)
  RAW = RAW & "01 01 01 01 01 01 01 01 01 01 01 01 01 01 01 01"
  
  ' - Detailed Timings - (72 Bytes)
  ' Timing #1 {Native Modeline}
  RAW = RAW & "28 05 80 C4 20 F0 17 00 20 40 12 00 00 00 00 00 00 98" ' modeline '640x480@60,036' 13,2 640 672 736 836 480 482 486 526 interlace -hsync -vsync
  
  '' MonDef #1 {Limits}
  'RAW = RAW & "00 00 00 FD 00 19 41 0F 30 15 00 0A 20 20 20 20 20 20" ' H: 15-48 kHz, V: 25-65 Hz, P: -210 MHz
  'RAW = RAW & "00 00 00 FD 00 19 41 0F 23 15 00 0A 20 20 20 20 20 20" ' H: 15-35 kHz, V: 25-65 Hz, P: -210 MHz
  RAW = RAW & "00 00 00 FD 00 19 41 0F 0F 15 00 0A 20 20 20 20 20 20" ' H: 15-15 kHz, V: 25-65 Hz, P: -210 MHz
  
  ' MonDef #2 {Name}
  RAW = RAW & "00 00 00 FC 00 53 31 35 4B 20 2D 20 45 44 49 44 0A 0A" ' "S15K - EDID"
  
  ' Timing #2 {Modeline}
  RAW = RAW & "70 06 20 F0 30 2C 0E 10 28 50 12 00 00 00 00 00 00 98" ' modeline '800x600@50,465' 16,48 800 840 920 1040 600 602 606 628 interlace -hsync -vsync
  
  ' - Extension Flag - (1 Byte)
  RAW = RAW & "00"
  
  ' - Checksum - (1 Byte)
  RAW = Replace(RAW, " ", "")
  Checksum = CalcChecksum(RAW)
  If Checksum < 16 Then
    RAW = RAW & "0" & Hex(Checksum)
  Else
    RAW = RAW & Hex(Checksum)
  End If
  
  Binary_Main = RAW
End Sub

'Sub CreateDAT_Ext1()
'  ' VTB_EXT v1
'  Dim RAW As String, Checksum As Integer
'  RAW = ""
'
'  ' - Extension Header - (5 Bytes)
'  ' Extension Tag
'  RAW = RAW & "10"
'  ' Version
'  RAW = RAW & "01"
'  ' Number of DTBs
'  RAW = RAW & "06"
'  ' Number of CVTs
'  RAW = RAW & "00"
'  ' Number of STs
'  RAW = RAW & "00"
'
'  ' - Detailed Timings - (108 Bytes)
'  ' Timing #2 {modeline '800x600@50,465' 16,48 800 840 920 1040 600 602 606 628 interlace -hsync -vsync}
'  RAW = RAW & "70 06 20 F0 30 2C 0E 10 28 50 12 00 00 00 00 00 00 98"
'
'  ' Timing #3 {modeline '320x240@59,014' 6,45 320 336 368 414 240 242 245 264 -hsync -vsync}
'  RAW = RAW & "85 02 40 5E 10 F0 18 00 10 20 23 00 00 00 00 00 00 18"
'
'  ' Timing #4 {modeline '392x240@59,898' 8 392 408 448 504 240 243 246 265 -hsync -vsync}
'  RAW = RAW & "20 03 88 70 10 F0 19 00 10 28 33 00 00 00 00 00 00 18"
'
'  ' Timing #5 {modeline '512x240@59,973' 10,68 512 544 600 672 240 243 246 265 -hsync -vsync}
'  RAW = RAW & "2C 04 00 A0 20 F0 19 00 20 38 33 00 00 00 00 00 00 18"
'
'  ' Timing #6 {modeline '352x288@51,116' 7,4 352 368 408 464 288 289 292 312 -hsync -vsync}
'  RAW = RAW & "E4 02 60 70 10 20 18 10 10 28 13 00 00 00 00 00 00 18"
'
'  ' Timing #7 {modeline '640x288@50,955' 13,1 640 672 736 832 288 289 292 309 -hsync -vsync}
'  RAW = RAW & "1E 05 80 C0 20 20 15 10 20 40 13 00 00 00 00 00 00 18"
'
'  ' Unused
'  RAW = RAW & String(28, "0")
'
'  ' - Checksum - (1 Byte)
'  RAW = Replace(RAW, " ", "")
'  Debug.Print Len(RAW)
'  Checksum = CalcChecksum(RAW)
'  If Checksum < 16 Then
'    RAW = RAW & "0" & Hex(Checksum)
'  Else
'    RAW = RAW & Hex(Checksum)
'  End If
'  Binary_Ext1 = RAW
'End Sub


Sub CreateDAT_Ext1()
  ' EIA/CEA-861B
  Dim RAW As String, Checksum As Integer
  RAW = ""

  ' - Extension Header - (4 Bytes)
  ' Extension Tag
  RAW = RAW & "02"
  ' Version
  RAW = RAW & "03"
  ' Data Type
  RAW = RAW & "05"
  ' Number of DTDs
  RAW = RAW & "06"
  
  RAW = RAW & "00"

  ' - Detailed Timings - (108 Bytes)
  ' Timing #2 {modeline '800x600@50,465' 16,48 800 840 920 1040 600 602 606 628 interlace -hsync -vsync}
  RAW = RAW & "70 06 20 F0 30 2C 0E 10 28 50 12 00 00 00 00 00 00 98"

  ' Timing #3 {modeline '320x240@59,014' 6,45 320 336 368 414 240 242 245 264 -hsync -vsync}
  RAW = RAW & "85 02 40 5E 10 F0 18 00 10 20 23 00 00 00 00 00 00 18"

  ' Timing #4 {modeline '392x240@59,898' 8 392 408 448 504 240 243 246 265 -hsync -vsync}
  RAW = RAW & "20 03 88 70 10 F0 19 00 10 28 33 00 00 00 00 00 00 18"

  ' Timing #5 {modeline '512x240@59,973' 10,68 512 544 600 672 240 243 246 265 -hsync -vsync}
  RAW = RAW & "2C 04 00 A0 20 F0 19 00 20 38 33 00 00 00 00 00 00 18"

  ' Timing #6 {modeline '352x288@51,116' 7,4 352 368 408 464 288 289 292 312 -hsync -vsync}
  RAW = RAW & "E4 02 60 70 10 20 18 10 10 28 13 00 00 00 00 00 00 18"

  ' Timing #7 {modeline '640x288@50,955' 13,1 640 672 736 832 288 289 292 309 -hsync -vsync}
  RAW = RAW & "1E 05 80 C0 20 20 15 10 20 40 13 00 00 00 00 00 00 18"

  ' Unused
  RAW = RAW & String(28, "0")

  ' - Checksum - (1 Byte)
  RAW = Replace(RAW, " ", "")
  Checksum = CalcChecksum(RAW)
  If Checksum < 16 Then
    RAW = RAW & "0" & Hex(Checksum)
  Else
    RAW = RAW & Hex(Checksum)
  End If
  Binary_Ext1 = RAW
End Sub


Private Sub Command4_Click()
CreateDAT_Ext1
End Sub

