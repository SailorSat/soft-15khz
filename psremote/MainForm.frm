VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'Kein
   Caption         =   "PSremote"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_USER = 1024


Private Const UM_SETCUSTOMTIMING = WM_USER + 200
Private Const UM_SETREFRESHRATE = WM_USER + 201
Private Const UM_SETPOLARITY = WM_USER + 202
Private Const UM_REMOTECONTROL = WM_USER + 210
Private Const UM_SETGAMMARAMP = WM_USER + 203
Private Const UM_CREATERESOLUTION = WM_USER + 204
Private Const UM_GETTIMING = WM_USER + 205
Private Const UM_GETSETCLOCKS = WM_USER + 206
Private Const UM_SETCUSTOMTIMINGFAST = WM_USER + 211


Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Long) As Long
Private Declare Function GlobalGetAtomName Lib "kernel32.dll" Alias "GlobalGetAtomNameA" (ByVal nAtom As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long


Private hPSWnd As Long

Private sParams() As String
Private sBackup() As String
Private lRefresh As Byte


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim Result As Long
  Dim Changed As Boolean
  Dim Refresh As Double
  Dim PTotal As Long
  Dim PClock As Long
  Changed = False
  Select Case KeyCode
    Case 37 'Left
      If sParams(3) >= 8 Then
        sParams(1) = sParams(1) + 8
        sParams(3) = sParams(3) - 8
        Changed = True
      End If
    Case 38 'Up
      If sParams(7) >= 1 Then
        sParams(5) = sParams(5) + 1
        sParams(7) = sParams(7) - 1
        Changed = True
      End If
    Case 39 'Right
      If sParams(1) >= 8 Then
        sParams(1) = sParams(1) - 8
        sParams(3) = sParams(3) + 8
        Changed = True
      End If
    Case 40 'Down
      If sParams(5) >= 1 Then
        sParams(5) = sParams(5) - 1
        sParams(7) = sParams(7) + 1
        Changed = True
      End If
      
    Case 33 'PageUp
      If sParams(5) >= 2 And sParams(6) >= 1 And sParams(7) >= 2 Then
        sParams(5) = sParams(5) - 2
        sParams(6) = sParams(6) - 1
        sParams(7) = sParams(7) - 2
        Changed = True
      End If
    Case 34 'PageDown
      sParams(5) = sParams(5) + 2
      sParams(6) = sParams(6) + 1
      sParams(7) = sParams(7) + 2
      Changed = True
    
    Case 46 'Delete
      PTotal = (CLng(sParams(0)) + CLng(sParams(1)) + CLng(sParams(2)) + CLng(sParams(3))) * (CLng(sParams(4)) + CLng(sParams(5)) + CLng(sParams(6)) + CLng(sParams(7)))
      PClock = CLng(sParams(8)) * 1000
      Refresh = PClock / PTotal
      
      sParams(1) = sParams(1) + 8
      sParams(2) = sParams(2) + 2
      sParams(3) = sParams(3) + 8
      
      PTotal = (CLng(sParams(0)) + CLng(sParams(1)) + CLng(sParams(2)) + CLng(sParams(3))) * (CLng(sParams(4)) + CLng(sParams(5)) + CLng(sParams(6)) + CLng(sParams(7)))
      PClock = (Refresh * CDbl(PTotal)) / 1000
      
      sParams(8) = CStr(PClock)
      Changed = True
    Case 35 'End
      If sParams(1) >= 8 And sParams(2) >= 4 And sParams(3) >= 8 Then
        PTotal = (CLng(sParams(0)) + CLng(sParams(1)) + CLng(sParams(2)) + CLng(sParams(3))) * (CLng(sParams(4)) + CLng(sParams(5)) + CLng(sParams(6)) + CLng(sParams(7)))
        PClock = CLng(sParams(8)) * 1000
        Refresh = PClock / PTotal
        
        sParams(1) = sParams(1) - 8
        sParams(2) = sParams(2) - 2
        sParams(3) = sParams(3) - 8
        
        PTotal = (CLng(sParams(0)) + CLng(sParams(1)) + CLng(sParams(2)) + CLng(sParams(3))) * (CLng(sParams(4)) + CLng(sParams(5)) + CLng(sParams(6)) + CLng(sParams(7)))
        PClock = (Refresh * CDbl(PTotal)) / 1000
        
        sParams(8) = CStr(PClock)
        Changed = True
      End If
    Case 27 'ESC
      sParams = sBackup
      Changed = True
    Case 13 'Return
      End
  End Select
  If Changed Then
    Dim nAtom As Long
    Dim sData As String
    sData = Join(sParams, ",")
    nAtom = GlobalAddAtom(sData)
    Result = SendMessage(hPSWnd, UM_SETCUSTOMTIMING, Nothing, nAtom)
    GlobalDeleteAtom nAtom
    DrawOSD
  Else
    Debug.Print KeyCode
  End If
End Sub


Private Sub Form_Load()
  Dim nAtom As Long
  Dim nSize As Long
  Dim sData As String
  
  hPSWnd = FindWindow("TPShidden", vbNullString)
  nAtom = SendMessage(hPSWnd, UM_GETTIMING, Nothing, Nothing)
  sData = String(255, 0)
  nSize = GlobalGetAtomName(nAtom, sData, 256)
  sData = Left(sData, nSize)
  GlobalDeleteAtom nAtom
  Debug.Print sData
  sParams = Split(sData, ",")
  sBackup = sParams
  
  Dim Refresh As Double
  Dim PTotal As Long
  Dim PClock As Long
  PTotal = (CLng(sParams(0)) + CLng(sParams(1)) + CLng(sParams(2)) + CLng(sParams(3))) * (CLng(sParams(4)) + CLng(sParams(5)) + CLng(sParams(6)) + CLng(sParams(7)))
  PClock = CLng(sParams(8)) * 1000
  Refresh = PClock / PTotal
  lRefresh = CByte(Refresh)
  
  Me.Show
  DoEvents
  DrawOSD
End Sub

Sub DrawOSD()
  Me.Cls
  Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
  Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(0, Me.ScaleHeight)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(0, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Label1.Caption = sParams(0) & "x" & sParams(4)
  Label2.Caption = sParams(1) & vbCrLf & sParams(2) & vbCrLf & sParams(3)
  Label3.Caption = sParams(5) & vbCrLf & sParams(6) & vbCrLf & sParams(7)
  Label4.Caption = sParams(8) / 1000 & "MHz (" & sParams(0) & ")"
  Label5.Caption = lRefresh & "Hz"
  Label6.Caption = GetModeline
End Sub

Function GetModeline() As String
  Dim Dummy As String
  GetModeline = "modeline '" & sParams(0) & "x" & sParams(4) & "' " & sParams(8) / 1000 & vbCrLf
  GetModeline = GetModeline & "   " & sParams(0) & " " & CStr(CInt(sParams(0)) + CInt(sParams(1))) & " " & CStr(CInt(sParams(0)) + CInt(sParams(1)) + CInt(sParams(2))) & " " & CStr(CInt(sParams(0)) + CInt(sParams(1)) + CInt(sParams(2)) + CInt(sParams(3))) & vbCrLf
  GetModeline = GetModeline & "   " & sParams(4) & " " & CStr(CInt(sParams(4)) + CInt(sParams(5))) & " " & CStr(CInt(sParams(4)) + CInt(sParams(5)) + CInt(sParams(6))) & " " & CStr(CInt(sParams(4)) + CInt(sParams(5)) + CInt(sParams(6)) + CInt(sParams(7))) & vbCrLf
  Dummy = ""
  If sParams(9) And 2 Then
    Dummy = Dummy & "-hsync "
  Else
    Dummy = Dummy & "+hsync "
  End If
  If sParams(9) And 4 Then
    Dummy = Dummy & "-vsync "
  Else
    Dummy = Dummy & "+vsync "
  End If
  If sParams(9) And 32 Then
    If ((sParams(9) And 2) And (sParams(9) And 4)) Then
      Dummy = "-csync "
    ElseIf Not ((sParams(9) And 2) And (sParams(9) And 4)) Then
      Dummy = "+csync "
    End If
  End If
  If sParams(9) And 8 Then Dummy = Dummy & "interlace "
  If sParams(9) And 128 Then Dummy = Dummy & "sync-on-green "
  GetModeline = GetModeline & Dummy
End Function
