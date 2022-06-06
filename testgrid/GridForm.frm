VERSION 5.00
Begin VB.Form GridForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
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
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "GridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
  End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If Me.Tag = 2 Then
    ColorTest
  ElseIf Me.Tag = 1 Then
    CrossTest
  ElseIf Me.Tag = 3 Then
    ContrastTest
  Else
    End
  End If
End Sub

Private Sub Form_Load()
  Me.Move 0, 0, Screen.Width, Screen.Height
'  Me.Move 0, 0, 240 * 15, 240 * 15
  GridTest
End Sub

Sub CrossTest()
  Me.Tag = 2
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls

  Dim TileSize As Integer
  TileSize = 7
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
  Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)

  Me.Line (Me.ScaleWidth / 2, 1)-(Me.ScaleWidth / 2, Me.ScaleHeight - 1)
  Me.Line ((Me.ScaleWidth / 2) - 1, 1)-((Me.ScaleWidth / 2) - 1, Me.ScaleHeight - 1)
  
  Me.Line ((Me.ScaleHeight / TileSize) + 1, 1)-((Me.ScaleHeight / TileSize) + 1, Me.ScaleHeight - 1)
  
  Me.Line (Me.ScaleWidth - (Me.ScaleHeight / TileSize) - 2, 1)-(Me.ScaleWidth - (Me.ScaleHeight / TileSize) - 2, Me.ScaleHeight - 1)
  
  Me.Line (1, Me.ScaleHeight / 2)-(Me.ScaleWidth - 1, Me.ScaleHeight / 2)
  Me.Line (1, (Me.ScaleHeight / 2) - 1)-(Me.ScaleWidth - 1, (Me.ScaleHeight / 2) - 1)
  
  Me.Line (1, (Me.ScaleHeight / TileSize) + 1)-(Me.ScaleWidth - 1, (Me.ScaleHeight / TileSize) + 1)
  
  Me.Line (1, Me.ScaleHeight - (Me.ScaleHeight / TileSize) - 2)-(Me.ScaleWidth - 1, Me.ScaleHeight - (Me.ScaleHeight / TileSize) - 2)
  
  Me.Circle ((Me.ScaleWidth / 2) - 0, (Me.ScaleHeight / 2) - 0), Me.ScaleHeight / (TileSize - 2)
  Me.Circle ((Me.ScaleWidth / 2) - 1, (Me.ScaleHeight / 2) - 1), Me.ScaleHeight / (TileSize - 2)
  Me.Circle ((Me.ScaleWidth / 2) - 0, (Me.ScaleHeight / 2) - 1), Me.ScaleHeight / (TileSize - 2)
  Me.Circle ((Me.ScaleWidth / 2) - 1, (Me.ScaleHeight / 2) - 0), Me.ScaleHeight / (TileSize - 2)

  Me.Circle ((Me.ScaleHeight / TileSize) + 1, (Me.ScaleHeight / TileSize) + 1), Me.ScaleHeight / TileSize
  Me.Circle ((Me.ScaleHeight / TileSize) + 1, Me.ScaleHeight - (Me.ScaleHeight / TileSize) - 2), Me.ScaleHeight / TileSize
  Me.Circle (Me.ScaleWidth - (Me.ScaleHeight / TileSize) - 2, (Me.ScaleHeight / TileSize) + 1), Me.ScaleHeight / TileSize
  Me.Circle (Me.ScaleWidth - (Me.ScaleHeight / TileSize) - 2, Me.ScaleHeight - (Me.ScaleHeight / TileSize) - 2), Me.ScaleHeight / TileSize
End Sub

Sub GridTest()
  Me.Tag = 1
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((Me.ScaleWidth / TileSize) > 20) And ((Me.ScaleHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(Me.ScaleWidth / TileSize) = CSng(CInt(Me.ScaleWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  
  Me.ForeColor = RGB(63, 63, 63)
  Dim LineCount As Integer
  Dim RowCount As Integer
  For RowCount = 0 To ((Me.ScaleWidth / TileSize) - 1)
    Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, Me.ScaleHeight)
    Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight)
  Next
  For LineCount = 0 To ((Me.ScaleHeight / TileSize) - 1)
    Me.Line (0, (LineCount * TileSize))-(Me.ScaleWidth, (LineCount * TileSize) + 0)
    Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth, (LineCount * TileSize) + (TileSize - 1))
  Next

  Me.ForeColor = RGB(127, 127, 127)
  For RowCount = 1 To ((Me.ScaleWidth / TileSize) - 2)
    Me.Line ((RowCount * TileSize) + 0, TileSize)-((RowCount * TileSize) + 0, Me.ScaleHeight - TileSize)
    Me.Line ((RowCount * TileSize) + (TileSize - 1), TileSize)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight - TileSize)
  Next
  For LineCount = 1 To ((Me.ScaleHeight / TileSize) - 2)
    Me.Line (TileSize, (LineCount * TileSize))-(Me.ScaleWidth - TileSize, (LineCount * TileSize) + 0)
    Me.Line (TileSize, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth - TileSize, (LineCount * TileSize) + (TileSize - 1))
  Next

  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
  Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, Me.ScaleHeight - TileSize + 1)
  Me.Line (Me.ScaleWidth - TileSize, TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(Me.ScaleWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, Me.ScaleHeight - TileSize)-(Me.ScaleWidth - TileSize + 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, Me.ScaleHeight - TileSize)
  Me.Line (Me.ScaleWidth - TileSize - 1, TileSize)-(Me.ScaleWidth - TileSize - 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(Me.ScaleWidth - TileSize, TileSize)
  Me.Line (TileSize, Me.ScaleHeight - TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize - 1)
End Sub

Sub ColorTest()
  Me.Tag = 3
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((Me.ScaleWidth / TileSize) > 20) And ((Me.ScaleHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(Me.ScaleWidth / TileSize) = CSng(CInt(Me.ScaleWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  Dim PaletteSize As Integer
  Dim PaletteStep As Integer
  While Me.ScaleWidth - PaletteSize > 256
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
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, Me.ScaleHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight)
  RowCount = (Me.ScaleWidth / TileSize) - 1
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, Me.ScaleHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight)
  
  LineCount = 0
  Me.Line (0, (LineCount * TileSize) + 0)-(Me.ScaleWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth, (LineCount * TileSize) + (TileSize - 1))
  LineCount = (Me.ScaleHeight / TileSize) - 1
  Me.Line (0, (LineCount * TileSize) + 0)-(Me.ScaleWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth, (LineCount * TileSize) + (TileSize - 1))
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
  Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, Me.ScaleHeight - TileSize + 1)
  Me.Line (Me.ScaleWidth - TileSize, TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(Me.ScaleWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, Me.ScaleHeight - TileSize)-(Me.ScaleWidth - TileSize + 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, Me.ScaleHeight - TileSize)
  Me.Line (Me.ScaleWidth - TileSize - 1, TileSize)-(Me.ScaleWidth - TileSize - 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(Me.ScaleWidth - TileSize, TileSize)
  Me.Line (TileSize, Me.ScaleHeight - TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize - 1)
  
  Dim ColorIndex As Integer
  Dim SubPixel As Integer
  Dim WorkIndex As Integer
  
  '-rot-
  LineCount = 2
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth / TileSize)) - (((Me.ScaleWidth - PaletteSize) / 2) / TileSize) - 3
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
  Me.Line (TileSize * ((((Me.ScaleWidth - PaletteSize) / 2) / TileSize) + 4), 10 * TileSize)-((((Me.ScaleWidth / TileSize)) - (((Me.ScaleWidth - PaletteSize) / 2) / TileSize) - 4) * TileSize, 13 * TileSize)
  Me.Line (TileSize * ((((Me.ScaleWidth - PaletteSize) / 2) / TileSize) + 4), 13 * TileSize)-((((Me.ScaleWidth / TileSize)) - (((Me.ScaleWidth - PaletteSize) / 2) / TileSize) - 4) * TileSize, 10 * TileSize)
  Me.Circle (Me.ScaleWidth / 2, 11 * TileSize + (TileSize / 2)), TileSize
  
    '-rot-
  LineCount = 14
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(((ColorIndex \ 16) + 1) * 16, 0, 0)
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
  LineCount = 15
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, ((ColorIndex \ 16) + 1) * 16, 0)
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
  LineCount = 16
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
  SubPixel = 0
  For ColorIndex = 255 To 0 Step -1
    Me.ForeColor = RGB(0, 0, ((ColorIndex \ 16) + 1) * 16)
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
 
Sub ContrastTest()
  Me.Tag = 4
  Me.BackColor = RGB(0, 0, 0)
  Me.Cls
  
  Dim TileSize As Integer
  Dim ValidSize As Integer
  TileSize = 2
  ValidSize = 2
  While ((Me.ScaleWidth / TileSize) > 20) And ((Me.ScaleHeight / TileSize) > 15)
    TileSize = TileSize + 2
    If CSng(Me.ScaleWidth / TileSize) = CSng(CInt(Me.ScaleWidth / TileSize)) Then
      ValidSize = TileSize
    End If
  Wend
  TileSize = ValidSize
  Dim PaletteSize As Integer
  Dim PaletteStep As Integer
  While Me.ScaleWidth - PaletteSize > 256
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
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, Me.ScaleHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight)
  RowCount = (Me.ScaleWidth / TileSize) - 1
  Me.Line ((RowCount * TileSize) + 0, 0)-((RowCount * TileSize) + 0, Me.ScaleHeight)
  Me.Line ((RowCount * TileSize) + (TileSize - 1), 0)-((RowCount * TileSize) + (TileSize - 1), Me.ScaleHeight)
  
  LineCount = 0
  Me.Line (0, (LineCount * TileSize) + 0)-(Me.ScaleWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth, (LineCount * TileSize) + (TileSize - 1))
  LineCount = (Me.ScaleHeight / TileSize) - 1
  Me.Line (0, (LineCount * TileSize) + 0)-(Me.ScaleWidth, (LineCount * TileSize) + 0)
  Me.Line (0, (LineCount * TileSize) + (TileSize - 1))-(Me.ScaleWidth, (LineCount * TileSize) + (TileSize - 1))
  
  Me.ForeColor = RGB(255, 0, 0)
  Me.Line (0, 0)-(0, Me.ScaleHeight - 1)
  Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
  Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
  Me.Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1)
  Me.ForeColor = RGB(255, 255, 255)
  Me.Line (TileSize - 1, TileSize - 1)-(TileSize - 1, Me.ScaleHeight - TileSize + 1)
  Me.Line (Me.ScaleWidth - TileSize, TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize + 1)
  Me.Line (TileSize - 1, TileSize - 1)-(Me.ScaleWidth - TileSize + 1, TileSize - 1)
  Me.Line (TileSize - 1, Me.ScaleHeight - TileSize)-(Me.ScaleWidth - TileSize + 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(TileSize, Me.ScaleHeight - TileSize)
  Me.Line (Me.ScaleWidth - TileSize - 1, TileSize)-(Me.ScaleWidth - TileSize - 1, Me.ScaleHeight - TileSize)
  Me.Line (TileSize, TileSize)-(Me.ScaleWidth - TileSize, TileSize)
  Me.Line (TileSize, Me.ScaleHeight - TileSize - 1)-(Me.ScaleWidth - TileSize, Me.ScaleHeight - TileSize - 1)
  
  Dim ColorIndex As Integer
  Dim SubPixel As Integer
  Dim WorkIndex As Integer
  
  '-rot-
  LineCount = 2
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
  RowCount = ((Me.ScaleWidth - PaletteSize) / 2) / TileSize
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
 

