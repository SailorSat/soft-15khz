VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   11595
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOut 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "00000000"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton btnParse 
      Caption         =   "Parse"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtRaw 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblOut 
      Alignment       =   2  'Center
      Caption         =   "DM------"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Label() As String
Private VType() As String

Private Sub btnParse_Click()
  Dim Line As String
  Dim Output As String
  Dim Index As Long
  Dim Offset As Long
  Dim Length As Long
  
  Line = Replace(txtRaw.Text, " ", "")
  Debug.Print Line
  Offset = 1
  Output = ""
  For Index = 0 To UBound(Label)
    Length = Len(Label(Index))
    Output = Output & Mid(Line, Offset, Length) & " "
    txtOut(Index).Text = Mid(Line, Offset, Length)
    lblVal(Index).Caption = Decode(Index, txtOut(Index).Text)
    Offset = Offset + Length
  Next
  Debug.Print Output
  Line = ""
  For Index = 0 To UBound(Label)
    Length = Len(Label(Index))
    Offset = Len(lblVal(Index).Caption)
    If Offset > Length Then Length = Offset
    Line = Line & Space(Length - Offset) & lblVal(Index).Caption & " "
  Next
  Debug.Print Line
  Clipboard.Clear
  Clipboard.SetText Output & vbCrLf & Line & vbCrLf
End Sub

Private Sub Form_Load()
  Dim Labels As String
  Dim VTypes As String
  Labels = "DM------ HWID---- H-ACTIVE V-ACTIVE 15------ COLORDEP 0------- 0------- 0------- H-ACTIVE V-ACTIVE 0------- 0------- H-ACTIVE V-ACTIVE 0------- 0------- H-ACTIVE V-ACTIVE H-ACTIVE H-FP H-SW H-TT H-SP V-ACTIVE V-FP V-SW V-TT V-SP FLAGS--- PCLOCK-- 0------- V-Hz H-Hz V-FREQ-- 0------- 1------- TIMM -- String-- -------- -------- -------- -------- -------- -------- -------- 00 0------- 0------- H-ACTIVE H-FP H-SW H-TT H-SP V-ACTIVE V-FP V-SW V-TT V-SP FLAGS--- PCLOCK-- 0------- V-Hz H-Hz V-FREQ-- 0------- 1------- TIMM 0- String-- -------- -------- -------- -------- -------- -------- -------- -------- -------- 0- ?-------"
  VTypes = "R        R        L        L        L        L        L        L        L        L        L        L        L        L        L        L        L        L        L        L        L    L    L    L    L        L    L    L    L    L        L        L        L    L    L        L        L        L    R  S        S        S        S        S        S        S        S        R  L        L        L        L    L    L    L    L        L    L    L    L    L        L        L        L    L    L        L        L        L    L  S        S        S        S        S        S        S        S        S        S        R  R       "
  While InStr(1, VTypes, "  ")
    VTypes = Replace(VTypes, "  ", " ")
  Wend
  Label = Split(Labels, " ")
  VType = Split(VTypes, " ")
  txtRaw.Text = "0001000200000000a0050000380400001500000020000000000000000000000000000000a0050000380400000000000000000000a0050000380400000000000000000000a005000038040000a0050000600098009007010038040000010003005e04000000000000bb320000000000003c00000061ea0000000000000100000007000040435553543a3134343078313038307836302e303031487a0000000000000000000000000000000000a0050000600098009007010038040000010003005e04000000000000bb320000000000003c00000061ea0000000000000100000007000000435553543a3134343078313038307836302e303031487a0000000000000000000000000000000000ee3ffdf7"
  LoadGrid
End Sub

Private Sub LoadGrid()
  Dim Index As Long
  Dim Row As Long
  
  Dim ColsPerRow As Long
  ColsPerRow = 12
  
  Dim Offset As Long
  
  For Index = 1 To UBound(Label)
    Row = (Index - (Index Mod ColsPerRow)) \ ColsPerRow
    Load lblOut(Index)
    Load txtOut(Index)
    Load lblVal(Index)
    lblOut(Index).Move 120 + ((Index Mod ColsPerRow) * 1200), 600 + (Row * 940)
    txtOut(Index).Move 120 + ((Index Mod ColsPerRow) * 1200), 840 + (Row * 940)
    lblVal(Index).Move 120 + ((Index Mod ColsPerRow) * 1200), 1200 + (Row * 940)
    lblOut(Index).Caption = Label(Index)
    txtOut(Index).Text = String(Len(Label(Index)), "0")
    lblVal(Index).Caption = CLng("&H" & txtOut(Index).Text)
    lblOut(Index).Visible = True
    txtOut(Index).Visible = True
    lblVal(Index).Visible = True
  Next
End Sub

Private Function Decode(Index As Long, Data As String) As String
  Dim Line As String
  Dim Dummy As String
  Select Case VType(Index)
    Case "R"
      Decode = Data
    Case "L"
      Dummy = Data
      While Len(Dummy) >= 2
        Line = Line & Right(Dummy, 2)
        Dummy = Left(Dummy, Len(Dummy) - 2)
      Wend
      Decode = CLng("&H" & Line)
    Case "S"
      Dummy = Data
      While Len(Dummy) >= 2
        Line = Line & Chr(CLng("&H" & Left(Dummy, 2)))
        Dummy = Mid(Dummy, 3)
      Wend
      Decode = Replace(Line, Chr(0), " ")
    Case Else
      Decode = VType(Index)
  End Select
End Function
