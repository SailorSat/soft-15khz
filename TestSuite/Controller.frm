VERSION 5.00
Begin VB.Form Controller 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'Kein
   Caption         =   "Test Suite"
   ClientHeight    =   9690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   646
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1019
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Testbild"
      ForeColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   9120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Farbraster"
         Height          =   855
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   3975
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Farbbalken"
         Height          =   855
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   3975
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Kreise"
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Gitter"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdResolution 
      Height          =   855
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton cmdResolution 
      Height          =   855
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmdResolution 
      Height          =   855
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmdResolution 
      Height          =   855
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdResolution 
      BackColor       =   &H00C0FFC0&
      Caption         =   "15,720kHz @ 60Hz - 240p"
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Frequenz"
      ForeColor       =   &H00E0E0E0&
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdResolution 
         Height          =   855
         Index           =   7
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   3975
      End
      Begin VB.CommandButton cmdResolution 
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   4680
         Width           =   3975
      End
      Begin VB.CommandButton cmdResolution 
         Height          =   855
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdResolution_Click(Index As Integer)
  Select Case Index
    Case 0
      SetSecondaryResolution 336, 240
    Case 1
      SetSecondaryResolution 720, 480
    Case 2
      SetSecondaryResolution 352, 288
    Case 3
      SetSecondaryResolution 720, 576
    Case 4
      SetSecondaryResolution 512, 384
    Case 5
      SetSecondaryResolution 640, 480
    Case 6
      SetSecondaryResolution 800, 600
    Case 7
      SetSecondaryResolution 1024, 768
  End Select
End Sub

Private Sub cmdTest_Click(Index As Integer)
  Select Case Index
    Case 0
      Grid.GridTest
    Case 1
      Grid.CrossTest
    Case 2
      Grid.ColorTest
    Case 3
      Grid.ContrastTest
  End Select
End Sub

Private Sub Form_Load()
  cmdResolution(0).Caption = "15,750 kHz" & vbCrLf & "240p - 60Hz"
  cmdResolution(1).Caption = "15,750 kHz" & vbCrLf & "480i - 60Hz"
  cmdResolution(2).Caption = "15,625 kHz" & vbCrLf & "288p - 50Hz"
  cmdResolution(3).Caption = "15,625 kHz" & vbCrLf & "576i - 50Hz"
  cmdResolution(4).Caption = "24,960kHz" & vbCrLf & "384p - 60Hz"
  cmdResolution(5).Caption = "31,500kHz" & vbCrLf & "480p - 60Hz"
  cmdResolution(6).Caption = "37,500kHz" & vbCrLf & "600p - 60Hz"
  cmdResolution(7).Caption = "48,360kHz" & vbCrLf & "768p - 60Hz"
  
  cmdTest(0).Caption = "Gitter" & vbCrLf & "(Konvergenz)"
  cmdTest(1).Caption = "Kreise" & vbCrLf & "(Geometrie)"
  cmdTest(2).Caption = "Verläufe" & vbCrLf & "(Farbtemperatur)"
  cmdTest(3).Caption = "Raster" & vbCrLf & "(Kontrast)"
  
  GetPrimaryResolution
  
  Grid.Show
  Grid.Move Controller.Width, 0
  HookWindowProc
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FreeWindowProc
  End
End Sub

Public Sub Redraw(Width As Integer, Height As Integer)
  Me.Move 0, 0, Width * 15, Height * 15
End Sub

