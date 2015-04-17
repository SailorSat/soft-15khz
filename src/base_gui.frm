VERSION 5.00
Begin VB.Form base_gui 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Soft-15kHz"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "base_gui.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton CommandUSER 
      Caption         =   "Install USER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "Status:"
      Height          =   2415
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      Begin VB.ListBox StatusList 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2040
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame AdapterFrame 
      Caption         =   "Adapter:"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton AdapterOption 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame CommandFrame 
      Caption         =   "Command:"
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   6735
      Begin VB.CommandButton Command31kHz 
         Caption         =   "Install 31kHz"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CommandUninstall 
         Caption         =   "Uninstall"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command25kHz 
         Caption         =   "Install 25kHz"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command15kHz 
         Caption         =   "Install 15kHz"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblLink2 
      Caption         =   "donate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3960
      MousePointer    =   10  'Aufwärtspfeil
      TabIndex        =   21
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblDonate 
      Alignment       =   1  'Rechts
      Caption         =   "feel free to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblSponsor 
      Alignment       =   1  'Rechts
      Caption         =   "sponsored by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblLink 
      Caption         =   "arcadeshop.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8040
      MousePointer    =   10  'Aufwärtspfeil
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "base_gui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -= Soft-15kHz - GUI - Main
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


Private SelectedAdapter As Byte


Private Sub AdapterOption_Click(Index As Integer)
  SelectAdapter Index + 1
End Sub


Private Sub CommandInstall(Bit As Byte)
  Command15kHz.Enabled = False
  Command25kHz.Enabled = False
  Command31kHz.Enabled = False
  CommandUSER.Enabled = False
  Select Case Bit
    Case 1
      AddStatusMessage "*** Installing 15kHz..."
    Case 2
      AddStatusMessage "*** Installing 25kHz..."
    Case 4
      AddStatusMessage "*** Installing 31kHz..."
    Case 128
      AddStatusMessage "*** Installing USER..."
  End Select
  With Adapter(SelectedAdapter)
    If .Soft15KHz = 0 Then
      If Not Host.RuntimeFlags And 1 Then
        AddStatusMessage "  creating backup..."
        SaveRegistryBackup SelectedAdapter
      End If
    End If
    AddStatusMessage "  clearing modes..."
    Select Case .DriverType
      Case 1
        ' -= NVidia ForceWare =-
        If Host.Windows = 5 Then
          NVidia_NT5_RemoveAllModes SelectedAdapter
        ElseIf Host.Windows = 6 Then
          NVidia_NT6_RemoveAllModes SelectedAdapter
        End If
      Case 2
        ' -= ATI Catalyst =-
        ATI_RemoveAllModes SelectedAdapter
      Case 3
        ' -= Matrox PowerDesk =-
        Matrox_RemoveAllModes SelectedAdapter
      Case 4
        ' -= 3Dfx =-
        TDfx_RemoveAllModes SelectedAdapter
      Case 5
        ' -= Intel GMA / EGD =-
        If .DriverVersion = 1 Then
          Intel_GMA_RemoveAllModes SelectedAdapter
        ElseIf .DriverVersion = 2 Then
          Intel_EGD_RemoveAllModes SelectedAdapter
        End If
    End Select
    .Soft15KHz = .Soft15KHz Or Bit
    If .Soft15KHz And 1 Then
      AddStatusMessage "  adding 15kHz modes..."
      Select Case .DriverType
        Case 1
          ' -= NVidia ForceWare =-
          If Host.Windows = 5 Then
            NVidia_NT5_AddAllModes SelectedAdapter, 15
          Else
            NVidia_NT6_AddAllModes SelectedAdapter, 15
          End If
        Case 2
          ' -= ATI Catalyst =-
          ATI_AddAllModes SelectedAdapter, 15
        Case 3
          ' -= Matrox PowerDesk =-
          Matrox_AddAllModes SelectedAdapter, 15
        Case 4
          ' -= 3Dfx =-
          TDfx_AddAllModes SelectedAdapter, 15
        Case 5
          ' -= Intel GMA / EGD =-
          If .DriverVersion = 1 Then
            Intel_GMA_AddAllModes SelectedAdapter, 15
          ElseIf .DriverVersion = 2 Then
            Intel_EGD_AddAllModes SelectedAdapter, 15
          End If
      End Select
    End If
    If .Soft15KHz And 2 Then
      AddStatusMessage "  adding 25kHz modes..."
      Select Case .DriverType
        Case 1
          ' -= NVidia ForceWare =-
          If Host.Windows = 5 Then
            NVidia_NT5_AddAllModes SelectedAdapter, 25
          Else
            NVidia_NT6_AddAllModes SelectedAdapter, 25
          End If
        Case 2
          ' -= ATI Catalyst =-
          ATI_AddAllModes SelectedAdapter, 25
        Case 3
          ' -= Matrox PowerDesk =-
          Matrox_AddAllModes SelectedAdapter, 25
        Case 4
          ' -= 3Dfx =-
          TDfx_AddAllModes SelectedAdapter, 25
        Case 5
          ' -= Intel GMA / EGD =-
          If .DriverVersion = 1 Then
            Intel_GMA_AddAllModes SelectedAdapter, 25
          ElseIf .DriverVersion = 2 Then
            Intel_EGD_AddAllModes SelectedAdapter, 25
          End If
      End Select
    End If
    If .Soft15KHz And 4 Then
      AddStatusMessage "  adding 31kHz modes..."
      Select Case .DriverType
        Case 1
          ' -= NVidia ForceWare =-
          If Host.Windows = 5 Then
            NVidia_NT5_AddAllModes SelectedAdapter, 31
          Else
            NVidia_NT6_AddAllModes SelectedAdapter, 31
          End If
        Case 2
          ' -= ATI Catalyst =-
          ATI_AddAllModes SelectedAdapter, 31
        Case 3
          ' -= Matrox PowerDesk =-
          Matrox_AddAllModes SelectedAdapter, 31
        Case 4
          ' -= 3Dfx =-
          TDfx_AddAllModes SelectedAdapter, 31
        Case 5
          ' -= Intel GMA / EGD =-
          If .DriverVersion = 1 Then
            Intel_GMA_AddAllModes SelectedAdapter, 31
          ElseIf .DriverVersion = 2 Then
            Intel_EGD_AddAllModes SelectedAdapter, 31
          End If
      End Select
    End If
    If .Soft15KHz And 128 Then
      AddStatusMessage "  adding USER modes..."
      Select Case .DriverType
        Case 1
          ' -= NVidia ForceWare =-
          If Host.Windows = 5 Then
            NVidia_NT5_AddAllModes SelectedAdapter, 255
          Else
            NVidia_NT6_AddAllModes SelectedAdapter, 255
          End If
        Case 2
          ' -= ATI Catalyst =-
          ATI_AddAllModes SelectedAdapter, 255
        Case 3
          ' -= Matrox PowerDesk =-
          Matrox_AddAllModes SelectedAdapter, 255
        Case 4
          ' -= 3Dfx =-
          TDfx_AddAllModes SelectedAdapter, 255
        Case 5
          ' -= Intel GMA / EGD =-
          If .DriverVersion = 1 Then
            Intel_GMA_AddAllModes SelectedAdapter, 255
          ElseIf .DriverVersion = 2 Then
            Intel_EGD_AddAllModes SelectedAdapter, 255
          End If
      End Select
    End If
    AddStatusMessage "  done."
    SaveSettingLong HKEY_LOCAL_MACHINE, .ObjectPath1, "___Soft-15kHz___", .Soft15KHz
  End With
  AddStatusMessage ""
  SelectAdapter CInt(SelectedAdapter)
End Sub


Private Sub Command15kHz_Click()
  CommandInstall 1
End Sub

Private Sub Command25kHz_Click()
  CommandInstall 2
End Sub

Private Sub Command31kHz_Click()
  CommandInstall 4
End Sub

Private Sub CommandUSER_Click()
  CommandInstall 128
End Sub


Private Sub CommandUninstall_Click()
  CommandUninstall.Enabled = False
  AddStatusMessage "*** Uninstalling..."
  AddStatusMessage "  clearing modes..."
  With Adapter(SelectedAdapter)
    Select Case .DriverType
      Case 1
        ' -= NVidia ForceWare =-
        If Host.Windows = 5 Then
          NVidia_NT5_RemoveAllModes SelectedAdapter
        ElseIf Host.Windows = 6 Then
          NVidia_NT6_RemoveAllModes SelectedAdapter
        End If
      Case 2
        ' -= ATI Catalyst =-
        ATI_RemoveAllModes SelectedAdapter
      Case 3
        ' -= Matrox PowerDesk =-
        Matrox_RemoveAllModes SelectedAdapter
      Case 4
        ' -= 3Dfx =-
        TDfx_RemoveAllModes SelectedAdapter
      Case 5
        ' -= Intel GMA / EGD =-
        If .DriverVersion = 1 Then
          Intel_GMA_RemoveAllModes SelectedAdapter
        ElseIf .DriverVersion = 2 Then
          Intel_EGD_RemoveAllModes SelectedAdapter
        End If
    End Select
    If Not Host.RuntimeFlags And 1 Then
      AddStatusMessage "  reading backup..."
      ReadRegistryBackup SelectedAdapter
    End If
    AddStatusMessage "  done."
    .Soft15KHz = 0
    DeleteValue HKEY_LOCAL_MACHINE, .ObjectPath1, "___Soft-15kHz___"
  End With
  AddStatusMessage ""
  SelectAdapter CInt(SelectedAdapter)
End Sub


Private Sub Form_Load()
  Me.Caption = "Soft-15kHz (Build " & App.Revision & ")"
  AddStatusMessage "-/|\- Soft-15kHz (Build " & App.Revision & ") -/|\-"
  AddStatusMessage ""
  ReadAdapterTable
  
  If Not Command = "" Then
    Dim Dummy As String
    Dim Line As String
    Dummy = LCase(Command) & " "
    While InStr(1, Dummy, " ")
      Line = Left(Dummy, InStr(1, Dummy, " ") - 1)
      Dummy = Mid(Dummy, InStr(1, Dummy, " ") + 1)
      If Left(Line, 2) = "-s" Then
        Line = Mid(Line, 3)
        If IsNumeric(Line) Then
          If CInt(Line) <= AdapterCount Then
            If CInt(Line) >= 0 Then
              If AdapterOption(CInt(Line)).Enabled = True Then
                AdapterOption(CInt(Line)).Value = True
                AdapterOption_Click CInt(Line)
              End If
            End If
          End If
        End If
      End If
      If Left(Line, 3) = "-nb" Then
        Host.RuntimeFlags = Host.RuntimeFlags Or 1
      End If
      If Left(Line, 4) = "-i15" Then
        If Command15kHz.Enabled = True Then
          Command15kHz_Click
        End If
      End If
      If Left(Line, 4) = "-i25" Then
        If Command25kHz.Enabled = True Then
          Command25kHz_Click
        End If
      End If
      If Left(Line, 4) = "-i31" Then
        If Command31kHz.Enabled = True Then
          Command31kHz_Click
        End If
      End If
      If Left(Line, 4) = "-ius" Then
        If CommandUSER.Enabled = True Then
          CommandUSER_Click
        End If
      End If
      If Left(Line, 2) = "-u" Then
        If CommandUninstall.Enabled = True Then
          CommandUninstall_Click
        End If
      End If
      If Left(Line, 2) = "-q" Then
        End
      End If
    Wend
  End If
End Sub

Sub ReadAdapterTable()
  Dim Index As Integer
  For Index = 1 To AdapterCount
    If Index = 10 Then Exit For
    With AdapterOption(Index - 1)
      .Visible = True
      .Caption = Adapter(Index).AdapterName
      If Len(.Caption) > 25 Then .Caption = Left(.Caption, 22) & "..."
      If Adapter(Index).DriverType = 0 Then
        .Enabled = False
      Else
        .Enabled = True
      End If
    End With
  Next
  For Index = AdapterCount + 1 To 9
    If Index = 10 Then Exit For
    With AdapterOption(Index - 1)
      .Visible = False
      .Enabled = False
    End With
  Next
  For Index = 1 To AdapterCount
    If AdapterOption(Index - 1).Enabled = True Then
      AdapterOption(Index - 1).Value = True
      Exit Sub
    End If
  Next
End Sub

Sub SelectAdapter(AdapterIndex As Integer)
  Dim Dummy As String
  SelectedAdapter = AdapterIndex
  AddStatusMessage "*** Adapter #" & AdapterIndex & " selected."
  AddStatusMessage "  Name       : " & Adapter(AdapterIndex).AdapterName
  Select Case Adapter(AdapterIndex).DriverType
    Case 1
      ' -= NVidia ForceWare =-
      AddStatusMessage "  Driver     : NVIDIA ForceWare"
    Case 2
      ' -= ATI Catalyst =-
      AddStatusMessage "  Driver     : ATI Catalyst"
    Case 3
      ' -= Matrox PowerDesk =-
      AddStatusMessage "  Driver     : Matrox Powerdesk"
    Case 4
      ' -= 3Dfx =-
      AddStatusMessage "  Driver     : 3Dfx"
    Case 5
      ' -= Intel GMA / EGD =-
      If Adapter(AdapterIndex).DriverVersion = 1 Then
        AddStatusMessage "  Driver     : Intel GMA"
      Else
        AddStatusMessage "  Driver     : Intel IEGD"
      End If
  End Select
  If Not Adapter(AdapterIndex).DriverString = "" Then
    AddStatusMessage "  Version    : " & Adapter(AdapterIndex).DriverString
  End If
  If Not Adapter(AdapterIndex).ChipType = "" Then
    AddStatusMessage "  Chipset    : " & Adapter(AdapterIndex).ChipType
  End If
  If Not Adapter(AdapterIndex).DriverVersion = 0 Then
    If Adapter(AdapterIndex).Soft15KHz = 0 Then
      AddStatusMessage "  Soft-15kHz : not active"
      CommandUninstall.Enabled = False
      Command15kHz.Enabled = True
      Command25kHz.Enabled = True
      Command31kHz.Enabled = True
      CommandUSER.Enabled = True
    Else
      CommandUninstall.Enabled = True
      Command15kHz.Enabled = True
      Command25kHz.Enabled = True
      Command31kHz.Enabled = True
      CommandUSER.Enabled = True
      Dummy = ""
      If Adapter(AdapterIndex).Soft15KHz And 1 Then
        Command15kHz.Enabled = False
        Dummy = Dummy & ",15"
      End If
      If Adapter(AdapterIndex).Soft15KHz And 2 Then
        Command25kHz.Enabled = False
        Dummy = Dummy & ",25"
      End If
      If Adapter(AdapterIndex).Soft15KHz And 4 Then
        Command31kHz.Enabled = False
        Dummy = Dummy & ",31"
      End If
      Dummy = Mid(Dummy, 2)
      If Adapter(AdapterIndex).Soft15KHz = 128 Then
        CommandUSER.Enabled = False
        AddStatusMessage "  Soft-15kHz : active ( USER )"
      ElseIf Adapter(AdapterIndex).Soft15KHz And 128 Then
        CommandUSER.Enabled = False
        AddStatusMessage "  Soft-15kHz : active ( " & Dummy & "kHz + USER )"
      Else
        AddStatusMessage "  Soft-15kHz : active ( " & Dummy & "kHz )"
      End If
    End If
  Else
    CommandUninstall.Enabled = False
    Command15kHz.Enabled = False
    Command25kHz.Enabled = False
    Command31kHz.Enabled = False
    CommandUSER.Enabled = False
  End If
  AddStatusMessage ""
End Sub

Public Sub AddStatusMessage(Line As String)
  If StatusList.ListCount > 20 Then
    StatusList.RemoveItem 0
  End If
  StatusList.AddItem Line
  StatusList.ListIndex = StatusList.ListCount - 1
  StatusList.ListIndex = -1
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblLink.ForeColor = &H8000000D
  lblLink2.ForeColor = &H8000000D
End Sub

Private Sub lblLink_Click()
  ShellExecuteA hWnd, "open", "http://www.arcadeshop.de", vbNullString, vbNullString, 1
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblLink.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub lblLink2_Click()
  ShellExecuteA hWnd, "open", "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=sailorsat%40animeger%2ede&item_name=Soft%2d15kHz&no_shipping=0&no_note=1&tax=0&currency_code=EUR&lc=EN&bn=PP%2dDonationsBF&charset=UTF%2d8", vbNullString, vbNullString, 1
End Sub

Private Sub lblLink2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblLink2.ForeColor = RGB(255, 0, 0)
End Sub
