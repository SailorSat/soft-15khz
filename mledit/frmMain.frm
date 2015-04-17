VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Modeline Editor"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClockM 
      Caption         =   "Clock-"
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdClockP 
      Caption         =   "Clock+"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdHoriM 
      Caption         =   ">-<"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdVertM 
      Caption         =   "\/"
      Height          =   615
      Left            =   7680
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdVertP 
      Caption         =   "/\"
      Height          =   615
      Left            =   7680
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtHFreq 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "15,72 kHz"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtVFreq 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "60 Hz"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "\/"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "/\"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtModeline 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "modeline '320x240' 6.45 320 336 368 414 240 242 245 264 -hsync -vsync"
      Top             =   5640
      Width           =   7815
   End
   Begin VB.CommandButton cmdHoriP 
      Caption         =   "<->"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape shpTotal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   4500
      Left            =   600
      Top             =   600
      Width           =   6000
   End
   Begin VB.Shape shpActive 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      FillColor       =   &H00808000&
      FillStyle       =   0  'Ausgefüllt
      Height          =   3975
      Left            =   840
      Top             =   840
      Width           =   5415
   End
   Begin VB.Shape shpHorizontal 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      FillColor       =   &H00008000&
      FillStyle       =   0  'Ausgefüllt
      Height          =   4500
      Left            =   840
      Top             =   600
      Width           =   5415
   End
   Begin VB.Shape shpVertical 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H00008080&
      FillStyle       =   0  'Ausgefüllt
      Height          =   3975
      Left            =   600
      Top             =   840
      Width           =   6000
   End
   Begin VB.Shape shpBackground 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Ausgefüllt
      Height          =   4500
      Left            =   600
      Top             =   600
      Width           =   6000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private ModeParam() As String

Private DotSeparator As Boolean
Private IsInterlace As Boolean


Private Sub cmdUp_Click()
  If ModeParam(9) = ModeParam(10) Then Exit Sub
  ModeParam(9) = ModeParam(9) + 1
  ModeParam(8) = ModeParam(8) + 1
  
  UpdateModeline
End Sub


Private Sub cmdDown_Click()
  If ModeParam(8) = ModeParam(7) Then Exit Sub
  ModeParam(9) = ModeParam(9) - 1
  ModeParam(8) = ModeParam(8) - 1
  
  UpdateModeline
End Sub


Private Sub cmdLeft_Click()
  If ModeParam(5) = ModeParam(6) Then Exit Sub
  ModeParam(5) = ModeParam(5) + 1
  ModeParam(4) = ModeParam(4) + 1
  
  UpdateModeline
End Sub


Private Sub cmdRight_Click()
  If ModeParam(4) = ModeParam(3) Then Exit Sub
  ModeParam(5) = ModeParam(5) - 1
  ModeParam(4) = ModeParam(4) - 1
  
  UpdateModeline
End Sub


Private Sub Form_Load()
  cmdVertP.Caption = "/\" & vbCrLf & "|" & vbCrLf & "\/"
  cmdVertM.Caption = "\/" & vbCrLf & "|" & vbCrLf & "/\"
  
  If Mid$(1 / 2, 2, 1) = "." Then
    DotSeparator = True
  Else
    DotSeparator = False
  End If
  
  Dim Modeline As String
  Modeline = InputBox("Modeline", "Modeline Editor", "modeline '320x240' 6.45 320 336 368 414 240 242 245 264 -hsync -vsync")
  
  ModeParam = Split(Modeline, " ", 12)
  
  If DotSeparator Then
    ModeParam(2) = Replace(ModeParam(2), ",", ".")
  Else
    ModeParam(2) = Replace(ModeParam(2), ".", ",")
  End If
  IsInterlace = InStr(1, ModeParam(11), "interlace", vbTextCompare)
  
  UpdateModeline
End Sub


Sub UpdateModeline()
  UpdateImage
  ModeParam(1) = "'" & ModeParam(3) & "x" & ModeParam(7) & "-" & Replace(txtHFreq.Text, " ", "") & "-" & Replace(txtVFreq.Text, " ", "") & "'"
  txtModeline.Text = Join(ModeParam, " ")
End Sub


Sub UpdateImage()
  Dim HFactor As Double
  Dim VFactor As Double
  Dim FrontPorch As Double
  Dim BackPorch As Double
  Dim SyncPulse As Double
  Dim Active As Double
  Dim HTotal As Double
  Dim VTotal As Double
  Dim PTotal As Double
  Dim PixelClock As Double
  Dim VFreq As Double
  Dim HFreq As Double
  
  'Modeline 'Name' FF.FFF HHH PHH LHH THH VVV PVV LVV TVV <options>
  '0        1      2      3   4   5   6   7   8   9   10  11
  
  ' Horizontal
  Active = ModeParam(3)
  FrontPorch = ModeParam(4) - ModeParam(3)
  SyncPulse = ModeParam(5) - ModeParam(4)
  BackPorch = ModeParam(6) - ModeParam(5)
  HFactor = 400 / (ModeParam(6) - SyncPulse)
  HTotal = ModeParam(6)
  shpHorizontal.Move 40 + CInt(BackPorch * HFactor), 40, CInt(Active * HFactor), 300
  
  ' Vertical
  Active = ModeParam(7)
  FrontPorch = ModeParam(8) - ModeParam(7)
  SyncPulse = ModeParam(9) - ModeParam(8)
  BackPorch = ModeParam(10) - ModeParam(9)
  VFactor = 300 / (ModeParam(10) - SyncPulse)
  VTotal = ModeParam(10)
  shpVertical.Move 40, 40 + CInt(BackPorch * VFactor), 400, CInt(Active * VFactor)
  
  ' Active
  shpActive.Move shpHorizontal.Left, shpVertical.Top, shpHorizontal.Width, shpVertical.Height
  
  ' VFreq / HFreq
  PixelClock = ModeParam(2) * 1000000
  PTotal = HTotal * VTotal
  VFreq = PixelClock / PTotal
  HFreq = VFreq * VTotal
  If IsInterlace Then
    txtVFreq.Text = Format(VFreq * 2, "#0.00") & " iHz"
  Else
    txtVFreq.Text = Format(VFreq, "#0.00") & " Hz"
  End If
  txtHFreq.Text = Format(HFreq / 1000, "#0.00") & " kHz"
End Sub


Private Sub cmdClockP_Click()
  Dim PixelClock As Double
  
  PixelClock = ModeParam(2) * 1000
  
  PixelClock = PixelClock + 2
  
  ModeParam(2) = Format(PixelClock / 1000, "#0.0000")
  UpdateModeline
End Sub


Private Sub cmdClockM_Click()
  Dim PixelClock As Double
  
  PixelClock = ModeParam(2) * 1000
  
  PixelClock = PixelClock - 2
  
  ModeParam(2) = Format(PixelClock / 1000, "#0.0000")
  UpdateModeline
End Sub


Private Sub cmdVertP_Click()
  If ModeParam(8) = ModeParam(7) Then Exit Sub
  
  Dim HTotal As Double
  Dim VTotal As Double
  Dim PTotal As Double
  Dim PixelClock As Double
  Dim VFreq As Double
  Dim HFreq As Double

  ' VFreq / HFreq
  HTotal = ModeParam(6)
  VTotal = ModeParam(10)
  PixelClock = ModeParam(2) * 1000000
  PTotal = HTotal * VTotal
  VFreq = PixelClock / PTotal
  HFreq = VFreq * VTotal

  ModeParam(10) = ModeParam(10) - 1
  ModeParam(9) = ModeParam(9) - 1
  ModeParam(8) = ModeParam(8) - 1
  VFreq = VFreq + 0.2
  
  VTotal = ModeParam(10)
  PTotal = HTotal * VTotal
  PixelClock = VFreq * PTotal
  
  ModeParam(2) = Format(PixelClock / 1000000, "#0.0000")
  UpdateModeline
End Sub


Private Sub cmdVertM_Click()
  If ModeParam(9) = ModeParam(10) Then Exit Sub
  
  Dim HTotal As Double
  Dim VTotal As Double
  Dim PTotal As Double
  Dim PixelClock As Double
  Dim VFreq As Double
  Dim HFreq As Double

  ' VFreq / HFreq
  HTotal = ModeParam(6)
  VTotal = ModeParam(10)
  PixelClock = ModeParam(2) * 1000000
  PTotal = HTotal * VTotal
  VFreq = PixelClock / PTotal
  HFreq = VFreq * VTotal

  ModeParam(10) = ModeParam(10) + 1
  ModeParam(9) = ModeParam(9) + 1
  ModeParam(8) = ModeParam(8) + 1
  VFreq = VFreq - 0.2
  
  VTotal = ModeParam(10)
  PTotal = HTotal * VTotal
  PixelClock = VFreq * PTotal
  
  ModeParam(2) = Format(PixelClock / 1000000, "#0.0000")
  UpdateModeline
End Sub


Private Sub cmdHoriP_Click()
  If ModeParam(4) = ModeParam(3) Then Exit Sub
  
  Dim HTotal As Double
  Dim VTotal As Double
  Dim PTotal As Double
  Dim PixelClock As Double
  Dim VFreq As Double
  Dim HFreq As Double

  ' VFreq / HFreq
  HTotal = ModeParam(6)
  VTotal = ModeParam(10)
  PixelClock = ModeParam(2) * 1000000
  PTotal = HTotal * VTotal
  VFreq = PixelClock / PTotal
  HFreq = VFreq * VTotal

  ModeParam(6) = ModeParam(6) - 1
  ModeParam(5) = ModeParam(5) - 1
  ModeParam(4) = ModeParam(4) - 1
  
  HTotal = ModeParam(6)
  PTotal = HTotal * VTotal
  PixelClock = VFreq * PTotal
  
  ModeParam(2) = Format(PixelClock / 1000000, "#0.0000")
  UpdateModeline
End Sub


Private Sub cmdHoriM_Click()
  If ModeParam(5) = ModeParam(6) Then Exit Sub
  
  Dim HTotal As Double
  Dim VTotal As Double
  Dim PTotal As Double
  Dim PixelClock As Double
  Dim VFreq As Double
  Dim HFreq As Double

  ' VFreq / HFreq
  HTotal = ModeParam(6)
  VTotal = ModeParam(10)
  PixelClock = ModeParam(2) * 1000000
  PTotal = HTotal * VTotal
  VFreq = PixelClock / PTotal
  HFreq = VFreq * VTotal

  ModeParam(6) = ModeParam(6) + 1
  ModeParam(5) = ModeParam(5) + 1
  ModeParam(4) = ModeParam(4) + 1
  
  HTotal = ModeParam(6)
  PTotal = HTotal * VTotal
  PixelClock = VFreq * PTotal
  
  ModeParam(2) = Format(PixelClock / 1000000, "#0.0000")
  UpdateModeline
End Sub

