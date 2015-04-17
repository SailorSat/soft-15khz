Attribute VB_Name = "Api"
Option Explicit

Private Const PrimaryDisplayName As String = "\\.\DISPLAY1"
Private Const SecondaryDisplayName As String = "\\.\DISPLAY2"

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32

Private Const GWL_WNDPROC As Long = -4

Private Const WM_DISPLAYCHANGE As Long = &H7E&

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

Private Type RES
  X As Long
  Y As Long
End Type

Public lpPrevWndFunc As Long

Private Declare Function EnumDisplaySettingsA Lib "user32" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettingsExA Lib "user32" (ByVal lpszDeviceName As String, lpDevMode As Any, ByVal hWnd As Long, ByVal dwFlags As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub HookWindowProc()
  ' Hook WindowProc
  lpPrevWndFunc = SetWindowLongA(Grid.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub FreeWindowProc()
  ' Free WindowProc
  Call SetWindowLongA(Grid.hWnd, GWL_WNDPROC, lpPrevWndFunc)
End Sub

Public Sub GetPrimaryResolution()
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(PrimaryDisplayName, -1, DevM)
  
  Controller.Redraw CInt(DevM.dmPelsWidth), CInt(DevM.dmPelsHeight)
End Sub

Public Sub GetSecondaryResolution()
  Dim Result As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(SecondaryDisplayName, -1, DevM)
  Grid.Redraw CInt(DevM.dmPelsWidth), CInt(DevM.dmPelsHeight)
End Sub

Public Sub SetSecondaryResolution(Width As Integer, Height As Integer)
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(SecondaryDisplayName, 0, DevM)
  While Not Result = 0
    If DevM.dmPelsWidth = Width And DevM.dmPelsHeight = Height And DevM.dmBitsPerPel = 32 Then
      Result = ChangeDisplaySettingsExA(SecondaryDisplayName, DevM, 0, 1, 0)
      Exit Sub
    End If
    Index = Index + 1
    Result = EnumDisplaySettingsA(SecondaryDisplayName, Index, DevM)
  Wend
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case iMsg
    Case WM_DISPLAYCHANGE
      ' Display Resolution changed
      GetSecondaryResolution
    
    Case Else
      ' Whatever happend isn't our concern
      WindowProc = CallWindowProcA(lpPrevWndFunc, hWnd, iMsg, wParam, lParam)
  End Select
End Function
