VERSION 5.00
Begin VB.Form GridForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GridForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "GridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScreenWidth As Integer
Private ScreenHeight As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then End
  
  If Me.Tag = 2 Then
    ColorTest
  ElseIf Me.Tag = 1 Then
    CrossTest
  ElseIf Me.Tag = 3 Then
    ContrastTest
  Else
    GridTest
  End If
End Sub

Private Sub Form_Load()
  Me.Tag = 1
  
  HookWindowProc
  GetCurrentResolution
End Sub

Public Sub Redraw(Width As Integer, Height As Integer)
  ScreenWidth = Width - (Width Mod 2)
  ScreenHeight = Height - (Height Mod 2)
  
  Me.Move 1024 * 15, 0, ScreenWidth * 15, ScreenHeight * 15
  
  Select Case Me.Tag
    Case 1
      GridTest
    Case 2
      CrossTest
    Case 3
      ColorTest
    Case 4
      ContrastTest
  End Select
End Sub
Private Sub CrossTest()
  Me.Tag = 2
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls

  Dim TileSize As Integer
  TileSize = 7
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, ScreenHeight - 1)
  Me.Line (ScreenWidth - 1, 0)-(ScreenWidth - 1, ScreenHeight - 1)
  Me.Line (0, 0)-(ScreenWidth - 1, 0)
  Me.Line (0, ScreenHeight - 1)-(ScreenWidth, ScreenHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)

  Me.Line (ScreenWidth / 2, 1)-(ScreenWidth / 2, ScreenHeight - 1)
  Me.Line ((ScreenWidth / 2) - 1, 1)-((ScreenWidth / 2) - 1, ScreenHeight - 1)
  
  Me.Line ((ScreenHeight / TileSize) + 1, 1)-((ScreenHeight / TileSize) + 1, ScreenHeight - 1)
  
  Me.Line (ScreenWidth - (ScreenHeight / TileSize) - 2, 1)-(ScreenWidth - (ScreenHeight / TileSize) - 2, ScreenHeight - 1)
  
  Me.Line (1, ScreenHeight / 2)-(ScreenWidth - 1, ScreenHeight / 2)
  Me.Line (1, (ScreenHeight / 2) - 1)-(ScreenWidth - 1, (ScreenHeight / 2) - 1)
  
  Me.Line (1, (ScreenHeight / TileSize) + 1)-(ScreenWidth - 1, (ScreenHeight / TileSize) + 1)
  
  Me.Line (1, ScreenHeight - (ScreenHeight / TileSize) - 2)-(ScreenWidth - 1, ScreenHeight - (ScreenHeight / TileSize) - 2)
  
  Me.Circle ((ScreenWidth / 2) - 0, (ScreenHeight / 2) - 0), ScreenHeight / (TileSize - 2)
  Me.Circle ((ScreenWidth / 2) - 1, (ScreenHeight / 2) - 1), ScreenHeight / (TileSize - 2)
  Me.Circle ((ScreenWidth / 2) - 0, (ScreenHeight / 2) - 1), ScreenHeight / (TileSize - 2)
  Me.Circle ((ScreenWidth / 2) - 1, (ScreenHeight / 2) - 0), ScreenHeight / (TileSize - 2)

  Me.Circle ((ScreenHeight / TileSize) + 1, (ScreenHeight / TileSize) + 1), ScreenHeight / TileSize
  Me.Circle ((ScreenHeight / TileSize) + 1, ScreenHeight - (ScreenHeight / TileSize) - 2), ScreenHeight / TileSize
  Me.Circle (ScreenWidth - (ScreenHeight / TileSize) - 2, (ScreenHeight / TileSize) + 1), ScreenHeight / TileSize
  Me.Circle (ScreenWidth - (ScreenHeight / TileSize) - 2, ScreenHeight - (ScreenHeight / TileSize) - 2), ScreenHeight / TileSize
End Sub

Private Sub GridTest()
  Me.Tag = 1
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((ScreenWidth / TileSize) > 20) And ((ScreenHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(ScreenWidth / TileSize) = CSng(CInt(ScreenWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  
  Me.ForeColor = RGB(63, 63, 63)
  Dim LineCount As Integer
  Dim RowCount As Integer
  For RowCount = 0 To ((ScreenWidth / TileSize) - 1)
    Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, ScreenHeight)
    Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight)
  Next
  For LineCount = 0 To ((ScreenHeight / TileSize) - 1)
    Me.Line (0, (LineCount * TileSize))-(ScreenWidth, (LineCount * TileSize) + 0)
    Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth, (LineCount * TileSize) + (TileSize - 1))
  Next

  Me.ForeColor = RGB(127, 127, 127)
  For RowCount = 1 To ((ScreenWidth / TileSize) - 2)
    Me.Line ((RowCount * TileSize) + 0, TileSize)-((RowCount * TileSize) + 0, ScreenHeight - TileSize)
    Me.Line ((RowCount * TileSize) + (TileSize - 1), TileSize)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight - TileSize)
  Next
  For LineCount = 1 To ((ScreenHeight / TileSize) - 2)
    Me.Line (TileSize, (LineCount * TileSize))-(ScreenWidth - TileSize, (LineCount * TileSize) + 0)
    Me.Line (TileSize, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth - TileSize, (LineCount * TileSize) + (TileSize - 1))
  Next

  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, ScreenHeight - 1)
  Me.Line (ScreenWidth - 1, 0)-(ScreenWidth - 1, ScreenHeight - 1)
  Me.Line (0, 0)-(ScreenWidth - 1, 0)
  Me.Line (0, ScreenHeight - 1)-(ScreenWidth, ScreenHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, ScreenHeight - TileSize + 1)
  Me.Line (ScreenWidth - TileSize, TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(ScreenWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, ScreenHeight - TileSize)-(ScreenWidth - TileSize + 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, ScreenHeight - TileSize)
  Me.Line (ScreenWidth - TileSize - 1, TileSize)-(ScreenWidth - TileSize - 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(ScreenWidth - TileSize, TileSize)
  Me.Line (TileSize, ScreenHeight - TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize - 1)
End Sub

Private Sub ColorTest()
  Me.Tag = 3
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((ScreenWidth / TileSize) > 20) And ((ScreenHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(ScreenWidth / TileSize) = CSng(CInt(ScreenWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  Dim PaletteSize As Integer
  Dim PaletteStep As Integer
  While ScreenWidth - PaletteSize > 256
    PaletteSize = PaletteSize + 256
    PaletteStep = PaletteStep + 1
  Wend
  If PaletteSize = 0 Then
    PaletteSize = 256
    PaletteStep = 1
  End If
  
  Me.ForeColor = RGB(63, 63, 63)
  Dim LineCount As Integer
  Dim RowCount As Integer
  RowCount = 0
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, ScreenHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight)
  RowCount = (ScreenWidth / TileSize) - 1
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, ScreenHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight)
  
  LineCount = 0
  Me.Line (0, (LineCount * TileSize) + 0)-(ScreenWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth, (LineCount * TileSize) + (TileSize - 1))
  LineCount = (ScreenHeight / TileSize) - 1
  Me.Line (0, (LineCount * TileSize) + 0)-(ScreenWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth, (LineCount * TileSize) + (TileSize - 1))
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, ScreenHeight - 1)
  Me.Line (ScreenWidth - 1, 0)-(ScreenWidth - 1, ScreenHeight - 1)
  Me.Line (0, 0)-(ScreenWidth - 1, 0)
  Me.Line (0, ScreenHeight - 1)-(ScreenWidth, ScreenHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, ScreenHeight - TileSize + 1)
  Me.Line (ScreenWidth - TileSize, TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(ScreenWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, ScreenHeight - TileSize)-(ScreenWidth - TileSize + 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, ScreenHeight - TileSize)
  Me.Line (ScreenWidth - TileSize - 1, TileSize)-(ScreenWidth - TileSize - 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(ScreenWidth - TileSize, TileSize)
  Me.Line (TileSize, ScreenHeight - TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize - 1)
  
  Dim ColorIndex As Integer
  Dim SubPixel As Integer
  Dim WorkIndex As Integer
  
  '-rot-
  LineCount = 2
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(ColorIndex, 0, 0)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next

  '-gelb-
  LineCount = 3
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(ColorIndex, ColorIndex, 0)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next

  '-grün-'
  LineCount = 4
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, ColorIndex, 0)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next
  
  '-cyan-'
  LineCount = 5
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, ColorIndex, ColorIndex)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next

  '-blau-'
  LineCount = 6
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, 0, ColorIndex)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next
  
  '-magenta-'
  LineCount = 7
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(ColorIndex, 0, ColorIndex)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next

  '-weiss-'
  LineCount = 8
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(ColorIndex, ColorIndex, ColorIndex)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next

  'horizontal
  LineCount = 10
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 0 To TileSize * 3
    If ColorIndex Mod 2 = 0 Then
      Me.ForeColor = RGB(255, 255, 255)
    Else
      Me.ForeColor = RGB(0, 0, 0)
    End If
    Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + (TileSize * 3) + 1)
    SubPixel = SubPixel + 1
    If SubPixel = TileSize Then
      SubPixel = 0
      RowCount = RowCount + 1
    End If
  Next

  'vertikal
  LineCount = 10
  RowCount = ((ScreenWidth / TileSize)) - (((ScreenWidth - PaletteSize) / 2) / TileSize) - 3
  SubPixel = 0
  For ColorIndex = 0 To TileSize * 3
    If ColorIndex Mod 2 = 0 Then
      Me.ForeColor = RGB(255, 255, 255)
    Else
      Me.ForeColor = RGB(0, 0, 0)
    End If
    Me.Line ((RowCount * TileSize) + 0, (LineCount * TileSize) + SubPixel)-((RowCount * TileSize) + (TileSize * 3) + 1, (LineCount * TileSize) + SubPixel)
    SubPixel = SubPixel + 1
    If SubPixel = TileSize Then
      SubPixel = 0
      LineCount = LineCount + 1
    End If
  Next
  
  'cross
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize * ((((ScreenWidth - PaletteSize) / 2) / TileSize) + 4), 10 * TileSize)-((((ScreenWidth / TileSize)) - (((ScreenWidth - PaletteSize) / 2) / TileSize) - 4) * TileSize, 13 * TileSize)
  Me.Line (TileSize * ((((ScreenWidth - PaletteSize) / 2) / TileSize) + 4), 13 * TileSize)-((((ScreenWidth / TileSize)) - (((ScreenWidth - PaletteSize) / 2) / TileSize) - 4) * TileSize, 10 * TileSize)
  Me.Circle (ScreenWidth / 2, 11 * TileSize + (TileSize / 2)), TileSize
End Sub
 
Private Sub ContrastTest()
  Me.Tag = 4
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((ScreenWidth / TileSize) > 20) And ((ScreenHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(ScreenWidth / TileSize) = CSng(CInt(ScreenWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  Dim PaletteSize As Integer
  Dim PaletteStep As Integer
  While ScreenWidth - PaletteSize > 256
    PaletteSize = PaletteSize + 256
    PaletteStep = PaletteStep + 1
  Wend
  If PaletteSize = 0 Then
    PaletteSize = 256
    PaletteStep = 1
  End If
  
  Me.ForeColor = RGB(63, 63, 63)
  Dim LineCount As Integer
  Dim RowCount As Integer
  RowCount = 0
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, ScreenHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight)
  RowCount = (ScreenWidth / TileSize) - 1
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, ScreenHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), ScreenHeight)
  
  LineCount = 0
  Me.Line (0, (LineCount * TileSize) + 0)-(ScreenWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth, (LineCount * TileSize) + (TileSize - 1))
  LineCount = (ScreenHeight / TileSize) - 1
  Me.Line (0, (LineCount * TileSize) + 0)-(ScreenWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(ScreenWidth, (LineCount * TileSize) + (TileSize - 1))
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, ScreenHeight - 1)
  Me.Line (ScreenWidth - 1, 0)-(ScreenWidth - 1, ScreenHeight - 1)
  Me.Line (0, 0)-(ScreenWidth - 1, 0)
  Me.Line (0, ScreenHeight - 1)-(ScreenWidth, ScreenHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, ScreenHeight - TileSize + 1)
  Me.Line (ScreenWidth - TileSize, TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(ScreenWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, ScreenHeight - TileSize)-(ScreenWidth - TileSize + 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, ScreenHeight - TileSize)
  Me.Line (ScreenWidth - TileSize - 1, TileSize)-(ScreenWidth - TileSize - 1, ScreenHeight - TileSize)
  Me.Line (TileSize, TileSize)-(ScreenWidth - TileSize, TileSize)
  Me.Line (TileSize, ScreenHeight - TileSize - 1)-(ScreenWidth - TileSize, ScreenHeight - TileSize - 1)
  
  Dim ColorIndex As Integer
  Dim SubPixel As Integer
  Dim WorkIndex As Integer
  
  '-rot-
  LineCount = 2
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(((ColorIndex \ 25) + 1) * 25, 0, 0)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next
  
  '-grün-
  LineCount = 4
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, ((ColorIndex \ 25) + 1) * 25, 0)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next
  
  '-blau-
  LineCount = 6
  RowCount = ((ScreenWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, 0, ((ColorIndex \ 25) + 1) * 25)
    For WorkIndex = 1 To PaletteStep
      Me.Line ((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + 0)-((RowCount * TileSize) + SubPixel, (LineCount * TileSize) + TileSize)
      SubPixel = SubPixel + 1
      If SubPixel = TileSize Then
        SubPixel = 0
        RowCount = RowCount + 1
      End If
    Next
  Next
End Sub
 

