Attribute VB_Name = "Api"
Option Explicit

Private Const DisplayName As String = "\\.\DISPLAY2"

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
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub HookWindowProc()
  ' Hook WindowProc
  lpPrevWndFunc = SetWindowLongA(GridForm.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub GetCurrentResolution()
  Dim Result As Long
  Dim Index As Long
  Dim DevM As DEVMODE
  Result = EnumDisplaySettingsA(DisplayName, -1, DevM)
  
  CheckDisplayResolution CInt(DevM.dmPelsWidth), CInt(DevM.dmPelsHeight)
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case iMsg
    Case WM_DISPLAYCHANGE
      ' Display Resolution changed
      GetCurrentResolution
    
    Case Else
      ' Whatever happend isn't our concern
      WindowProc = CallWindowProcA(lpPrevWndFunc, hWnd, iMsg, wParam, lParam)
  End Select
End Function


Public Sub CheckDisplayResolution(Width As Integer, Height As Integer)
  GridForm.Redraw Width, Height
End Sub

