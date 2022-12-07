VERSION 5.00
Begin VB.Form Window 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   765
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1380
   Icon            =   "Window.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuResolutions 
      Caption         =   "Resolutions"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuRes 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display Settings"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DisplayName As String = "\\.\DISPLAY1" 'vbNullString '
Private Const DisplayMinBit As Integer = 32

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONUP = &H202
Private Const WM_MOUSEMOVE = &H200

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32


Private Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function RegisterWindowMessageA Lib "user32.dll" (ByVal lpString As String) As Long


Private Declare Function EnumDisplaySettingsA Lib "user32" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettingsExA Lib "user32" (ByVal lpszDeviceName As String, lpDevMode As Any, ByVal hWnd As Long, ByVal dwFlags As Long, ByVal lParam As Long) As Long


Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private Type tResolution
  X As Long
  Y As Long
  Bit As Integer
  Hz As Long
  Num As Long
  Menu As Long
End Type

Private TrayIcon As NOTIFYICONDATA
Private MSG_TaskbarCreated As Long


Private Resolution() As tResolution
Private ResolutionCount As Integer
Private CurrentResolution As Integer
Private MenuCount As Integer


' -- Resolutions
Private Sub AddResolution(X As Long, Y As Long, Bit As Integer, Hz As Long, Num As Long)
  If Bit < DisplayMinBit Then Exit Sub
  Dim Index As Integer
  Dim Index2 As Integer
  Dim Index3 As Integer
  Index2 = -1
  Index3 = -1
  For Index = 1 To ResolutionCount
    If Resolution(Index).Bit >= Bit Then
      If Resolution(Index).Bit > Bit Then
        Index2 = Index
        Exit For
      End If
      If Resolution(Index).X >= X And Resolution(Index).Y > Y Then
        Index2 = Index
        Exit For
      End If
    End If
  Next
  For Index = 1 To ResolutionCount
    If Resolution(Index).X = X And Resolution(Index).Y = Y And Resolution(Index).Bit = Bit Then
      If Hz = 60 Then
        Resolution(Index).Hz = Hz
        Resolution(Index).Num = Num
      Else
        Exit Sub
      End If
    End If
  Next
  ResolutionCount = ResolutionCount + 1
  If ResolutionCount > UBound(Resolution) Then ReDim Preserve Resolution(ResolutionCount + 5)
  If Index2 <> -1 Then
    For Index = ResolutionCount To Index2 + 1 Step -1
      Resolution(Index) = Resolution(Index - 1)
    Next
  Else
    Index2 = ResolutionCount
  End If
  Resolution(Index2).X = X
  Resolution(Index2).Y = Y
  Resolution(Index2).Bit = Bit
  Resolution(Index2).Hz = Hz
  Resolution(Index2).Num = Num
End Sub


Private Sub EnumResolutions()
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(DisplayName, Index, DevM)
  While Not Result = 0
    AddResolution DevM.dmPelsWidth, DevM.dmPelsHeight, DevM.dmBitsPerPel, DevM.dmDisplayFrequency, Index
    Index = Index + 1
    Result = EnumDisplaySettingsA(DisplayName, Index, DevM)
  Wend
End Sub


Private Sub ChangeResolution(iMode As Long)
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(DisplayName, iMode, DevM)
  Result = ChangeDisplaySettingsExA(DisplayName, DevM, 0, 1, 0)
  If MsgBox("Keep current resolution?", vbQuestion + vbOKCancel, "QuickRes") = vbCancel Then
    Result = EnumDisplaySettingsA(DisplayName, mnuRes(CurrentResolution).Tag, DevM)
    Result = ChangeDisplaySettingsExA(DisplayName, DevM, 0, 1, 0)
  End If
End Sub

Private Sub GetCurrentResolution()
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(DisplayName, -1, DevM)
  
  For Index = 1 To ResolutionCount
    If Resolution(Index).X = DevM.dmPelsWidth And Resolution(Index).Y = DevM.dmPelsHeight And Resolution(Index).Bit = DevM.dmBitsPerPel Then
      If CurrentResolution > -1 Then
        mnuRes(CurrentResolution).Checked = False
      End If
      CurrentResolution = Resolution(Index).Menu
      mnuRes(CurrentResolution).Checked = True
      Exit Sub
    End If
  Next Index
End Sub

Private Sub ResetResolutions()
  ReDim Resolution(0)
  ResolutionCount = 0
End Sub


Private Sub Form_Load()
  Me.Visible = False
  AddTrayIcon
End Sub

Private Sub Form_Terminate()
  RemTrayIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Form_Terminate
End Sub


' -- Menu
Private Sub ResetMenu()
  Dim Index As Integer
  For Index = MenuCount To 1 Step -1
    Unload mnuRes(Index)
  Next
End Sub

Private Sub CreateMenu()
  Dim Index As Integer
  Dim Index2 As Integer
  Dim LastBit As Integer
  LastBit = DisplayMinBit
  Index2 = 0
  For Index = 1 To ResolutionCount
    Index2 = Index2 + 1
    Load mnuRes(Index2)
    If LastBit < Resolution(Index).Bit Then
      LastBit = Resolution(Index).Bit
      mnuRes(Index2).Caption = "-"
      Index2 = Index2 + 1
      Load mnuRes(Index2)
    End If
    mnuRes(Index2).Caption = Resolution(Index).X & " x " & Resolution(Index).Y & ", " & Resolution(Index).Bit & "Bit"
    mnuRes(Index2).Tag = Resolution(Index).Num
    Resolution(Index).Menu = Index2
  Next
  Index2 = Index2 + 1
  Load mnuRes(Index2)
  mnuRes(Index2).Caption = "-"
  MenuCount = Index2
End Sub

Private Sub mnuDisplay_Click()
  Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", vbHide)
End Sub


Private Sub mnuRes_Click(Index As Integer)
  ChangeResolution CLng(mnuRes(Index).Tag)
End Sub


' -- TrayIcon
Public Sub AddTrayIcon()
  Dim Result As Long
  TrayIcon.cbSize = Len(TrayIcon)
  TrayIcon.hWnd = Window.hWnd
  TrayIcon.uId = vbNull
  TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  TrayIcon.szTip = "QuickRes" & Chr(0)
  TrayIcon.ucallbackMessage = WM_MOUSEMOVE
  TrayIcon.hIcon = Window.Icon
  Result = Shell_NotifyIconA(NIM_ADD, TrayIcon)
End Sub

Public Sub RemTrayIcon()
  Dim Result As Long
  Result = Shell_NotifyIconA(NIM_DELETE, TrayIcon)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static Message As Long
  Static RR As Boolean
  Dim TPPX As Long
  TPPX = Screen.TwipsPerPixelX
  
  Message = X / TPPX
  
  If RR = False Then
    RR = True
    Select Case Message
      'Right button up (brings up a context menu)
      Case WM_LBUTTONUP
          ResetMenu
          ResetResolutions
          EnumResolutions
          CreateMenu
          GetCurrentResolution
          Me.PopupMenu mnuResolutions
    End Select
    RR = False
  End If
End Sub
