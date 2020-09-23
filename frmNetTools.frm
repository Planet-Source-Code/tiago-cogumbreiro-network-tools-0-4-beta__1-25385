VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Tools"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmNetTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   6360
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   50
      Top             =   10740
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      Caption         =   "fraMain"
      Height          =   4215
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   6855
      Begin VB.Frame fraDesc 
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   4335
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            Height          =   1275
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   4065
         End
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Single Port Scan"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Single\Multi Ping"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3495
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Multi Ping and for each pinged machine make a Port Scan"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "PortScan"
      Height          =   6615
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      Begin VB.CommandButton cmdSkip 
         Caption         =   "Skip"
         Height          =   375
         Left            =   4080
         TabIndex        =   56
         Top             =   5760
         Width           =   855
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   5040
         TabIndex        =   55
         Top             =   5760
         Width           =   855
      End
      Begin VB.Frame fraPingStatus 
         Caption         =   "Status"
         Height          =   1335
         Left            =   4080
         TabIndex        =   44
         Top             =   1800
         Width           =   3975
         Begin MSComDlg.CommonDialog cdl 
            Left            =   360
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ProgressBar ProgressPing 
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   960
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Min             =   1e-4
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            Height          =   255
            Left            =   3480
            TabIndex        =   48
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "0%"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblPingStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Percentage done"
            Height          =   675
            Left            =   480
            TabIndex        =   46
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame fraPingHost 
         Caption         =   "Hosts Settings:"
         Height          =   3495
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   3975
         Begin VB.CheckBox chkHostAdv 
            Caption         =   "Advanced Options:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox chkTo 
            Caption         =   "To:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdPingClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            ToolTipText     =   "Clear List"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdDel 
            Appearance      =   0  'Flat
            Caption         =   "Rem"
            Height          =   375
            Left            =   120
            Picture         =   "frmNetTools.frx":0442
            TabIndex        =   37
            ToolTipText     =   "Remove Entry"
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   2880
            TabIndex        =   36
            ToolTipText     =   "Add Entry"
            Top             =   600
            Width           =   495
         End
         Begin VB.ListBox lstAdr 
            Height          =   1425
            Left            =   840
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txtIP 
            Height          =   285
            Left            =   840
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   34
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtLast 
            Height          =   285
            Left            =   840
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   33
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblNote 
            BackStyle       =   0  'Transparent
            Caption         =   "Note: Remember that from computer 0.0.0.1 to computer 0.0.1.1 it has go through 255 computers"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   41
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label lblFrom 
            Caption         =   "From:"
            Height          =   255
            Left            =   360
            TabIndex        =   40
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Timer tmrStats 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   4560
         Top             =   1920
      End
      Begin MSWinsockLib.Winsock sckScan 
         Left            =   4560
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrCheckStatus 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   4560
         Top             =   240
      End
      Begin VB.Frame fraPortSettings 
         Caption         =   "Port Settings:"
         Height          =   2535
         Left            =   0
         TabIndex        =   15
         Top             =   3480
         Width           =   3975
         Begin VB.CommandButton cmdPortClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton cmdPortDel 
            Caption         =   "Rem"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   615
         End
         Begin VB.ListBox lstPortInterval 
            Height          =   840
            Left            =   840
            TabIndex        =   31
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkPortInterval 
            Caption         =   "To:"
            Height          =   255
            Left            =   1080
            TabIndex        =   30
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdPortAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   2520
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.CheckBox chkPortAdv 
            Caption         =   "Advanced Options:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEndPort 
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Text            =   "32767"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtStartPort 
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Text            =   "1"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblSelPort 
            Caption         =   "Selected Ports:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblStart 
            Caption         =   "Starting Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblEnd 
            Caption         =   "Ending Port:"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame fraLog 
         Caption         =   "Scan Status:"
         Height          =   2415
         Left            =   4080
         TabIndex        =   14
         Top             =   3240
         Width           =   3015
         Begin VB.CommandButton cmdLogClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtLog 
            Height          =   1335
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblHost 
            Caption         =   "Host:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraScanStatus 
         Caption         =   "Status:"
         Height          =   1815
         Left            =   4080
         TabIndex        =   1
         Top             =   0
         Width           =   3975
         Begin MSComctlLib.ProgressBar ProgressScan 
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Percent 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            Height          =   255
            Left            =   3600
            TabIndex        =   13
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Metrics 
            BackStyle       =   0  'Transparent
            Caption         =   " ports/sec"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label RemPort 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label EstRem 
            Caption         =   "#"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Progress:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblRemPort 
            Caption         =   "Remaining Ports:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblRem 
            Caption         =   "Estimated Remaining Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Port 
            Caption         =   "#"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Rate 
            AutoSize        =   -1  'True
            Caption         =   "#"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1080
            TabIndex        =   5
            Top             =   720
            Width           =   105
         End
         Begin VB.Label Label5 
            Caption         =   "Rate Speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status - Idle"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Port Scan"
            Object.ToolTipText     =   "Perform a simple Port Scan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ping"
            Object.ToolTipText     =   "Perform a simple Multi/Single Ping"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Multi Port Scan"
            Object.ToolTipText     =   "Perform PorScan on multiple computers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About the autor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RememberOptions
    LastPingHostOpt As Byte
    LastScanPortOpt As Byte
    LastMultiScanPortOpt As Byte
End Type
Dim Remb As RememberOptions
Dim CurPort As Long 'This specifies the next port that needs to be scanned
Dim LastPort As Long
Dim StartPort As Integer, EndPort As Integer
Dim LastFrame As Integer
Dim Working As Boolean
Dim LastChk(2, 2) As Byte
Dim SimpPing As Boolean, SimpScan As Boolean, MultiScan As Boolean
Dim Scanning As Boolean
Dim PortInt() As Integer 'Selected ports -> Advanced Portscan
Dim Action As Integer
Dim LastOpenPort As Integer
Dim HostPorts() As String 'An array of host and ports to scan
Dim bolStop As Boolean

Private Sub chkHostAdv_Click()
    'Visual fx:
    '(sorry but i didn't had the time to coment each and every line) =(
    If chkHostAdv.Value = 1 Then
        lblFrom = "From:"
        cmdAdd.Visible = True
        txtLast.Visible = True
        chkTo.Visible = True
        cmdDel.Visible = True
        cmdPingClear.Visible = True
        lstAdr.Visible = True
        lblNote.Visible = True
        fraPingHost.Height = 3495
        If Action = 3 Then
            chkHostAdv.Visible = False
            fraPortSettings.Left = fraPingHost.Left + fraPingHost.Width + 105
            fraPortSettings.Width = fraLog.Width
            fraPortSettings.Top = fraPingHost.Top
            fraScanStatus.Top = fraPingHost.Height + fraPingHost.Top
            fraLog.Top = fraPortSettings.Top + fraPortSettings.Height
            fraLog.Height = 4830 - fraPortSettings.Height
            txtLog.Height = fraLog.Height - 1080
            cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
        Else
            fraPingStatus.Top = fraPingHost.Top + fraPingHost.Height
            fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
            fraPortSettings.Left = fraPingHost.Left
            fraPingStatus.Left = fraPingHost.Left
            fraLog.Top = fraPingHost.Top
            fraLog.Height = 4830
            txtLog.Height = fraLog.Height - 1080
            cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
            fraPortSettings.Width = fraPingHost.Width
        End If
    Else
        lblFrom = "Host:"
        cmdAdd.Visible = False
        txtLast.Visible = False
        chkTo.Visible = False
        cmdDel.Visible = False
        cmdPingClear.Visible = False
        lstAdr.Visible = False
        lblNote.Visible = False
        fraPingHost.Height = 975
        fraLog.Top = fraPingHost.Top
        fraLog.Height = 4830
        txtLog.Height = fraLog.Height - 1080
        cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
        If Action = 3 Then
        'the following lines are unusefull since we'll always need
        'multiple hosts
'            fraPortSettings.Left = fraPingHost.Left
'            fraPortSettings.Top = fraPingHost.Top + fraPingHost.Height
'            fraPortSettings.Width = fraPingHost.Width
'            If chkPortAdv.Value = 0 Then
'                fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
'                fraScanStatus.Left = fraPortSettings.Left
'            End If
        Else
            fraPortSettings.Top = fraPingHost.Top + fraPingHost.Height
            fraPingStatus.Top = fraPingHost.Top + fraPingHost.Height
            fraPingStatus.Left = fraPingHost.Left
            fraPortSettings.Left = fraPingHost.Left
            fraScanStatus.Left = fraPingHost.Left
            fraPortSettings.Width = fraPingHost.Width
        End If
    End If
    If SimpPing Then Remb.LastPingHostOpt = chkHostAdv.Value
End Sub

Private Sub chkPortAdv_Click()
'Visual fx:
'Sorry again (same as above)
    If chkPortAdv.Value = 1 Then
        chkPortInterval.Visible = True
        lblSelPort.Visible = True
        lstPortInterval.Visible = True
        cmdPortAdd.Visible = True
        fraPortSettings.Height = 2415
        fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
        If Action = 3 Then
            If chkHostAdv = 1 Then
                If fraScanStatus.Top <> fraPingHost.Top + fraPingHost.Height Then fraScanStatus.Top = fraPingHost.Top + fraPingHost.Height
                fraLog.Top = fraPortSettings.Top + fraPortSettings.Height
                fraLog.Height = 4830 - fraPortSettings.Height
                txtLog.Height = fraLog.Height - 1080
                cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
            Else
            'following lines are unusefull:
'                If fraScanStatus.Top <> fraPortSettings.Top + fraPortSettings.Height Then fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
'                fraLog.Top = fraPingHost.Top
'                fraLog.Height = 4830
'                txtLog.Height = 4830 - 1080
'                cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
            End If
        End If
    Else
        chkPortInterval.Visible = False
        lblSelPort.Visible = False
        lstPortInterval.Visible = False
        cmdPortAdd.Visible = False
        fraPortSettings.Height = 1455
        fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
        If Action = 3 Then
            If chkHostAdv = 1 Then
                fraScanStatus.Top = fraPingHost.Top + fraPingHost.Height
                fraLog.Top = fraPortSettings.Top + fraPortSettings.Height
                fraLog.Height = 4830 - fraPortSettings.Height
                txtLog.Height = fraLog.Height - 1080
                cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
'            Else
'                fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
'                fraLog.Top = fraPingHost.Top
'                fraLog.Height = 4830
'                txtLog.Height = 4830 - 1080
'                cmdLogClear.Top = txtLog.Top + txtLog.Height + 108
            End If
        End If
    End If
    If SimpScan Then
        Remb.LastScanPortOpt = chkPortAdv.Value
    ElseIf MultiScan Then
        Remb.LastMultiScanPortOpt = chkPortAdv.Value
    End If
End Sub

Private Sub cmdLogClear_Click()
    'Clear the log
    txtLog = ""
End Sub

Private Sub cmdNext_Click()
    'Here we have a next and back button:
    If cmdNext.Caption = "Next >>" Then
        'If it's in general menu:
        Dim i As Byte
        'Use a for loop to get the position selected
        For i = 0 To optAction.UBound
            'When you get it it's of no use proceding with the loop
            'and we would lost the position
            If optAction(i).Value Then Exit For
        Next
        
        'Treat user option:
        Select Case i
            Case 0 'Single\Multi Ping
                Action = 1
                TabStrip.Tabs(3).Selected = True
            Case 1 'Multi Ping
                Action = 3
                TabStrip.Tabs(4).Selected = True
            Case 2 'Single Port Scan
                TabStrip.Tabs(2).Selected = True
                Action = 2
        End Select
    Else
        TabStrip.Tabs(1).Selected = True
    End If
End Sub

Private Sub cmdPingClear_Click()
    'Clear inserted addresses
    lstAdr.Clear
    'Check for delete button (there won't be nothin to delete)
    If cmdDel.Enabled = True Then cmdDel.Enabled = False
End Sub

Private Sub cmdPortAdd_Click()
Dim StartPort As Integer, EndPort As Integer
    If Not Working Then
        'Initialize flag for reconizing when user presses cancel
        Working = True
        'Change caption on button to let know that it can be stoped
        cmdPortAdd.Caption = "Stop"
        If IsNumeric(txtStartPort) Then
            If (Int(txtStartPort) - Val(txtStartPort)) = 0 Then
                'Verify values:
                StartPort = CInt(txtStartPort.Text)
                If StartPort < 1 Then StartPort = 1
                If StartPort >= 32767 Then StartPort = 32766
                
                'check if we're adding an interval of values
                If chkPortInterval.Value = 1 And IsNumeric(txtEndPort) Then
                'Both values must be integer!
                    If (Int(txtEndPort) - Val(txtEndPort)) <> 0 Then
                        MsgBox "Ending port has a non integer value!", vbCritical, "Initialize Error"
                        Exit Sub
                    Else
                        EndPort = CInt(txtEndPort.Text)
                        If EndPort > 32767 Then EndPort = 32767
                        If EndPort < 2 Then EndPort = 2
                        If StartPort >= txtEndPort Then
                            If MsgBox("Starting Port is bigger than ending Port. Ending Port will be set to a bigger value", vbCritical Or vbOKCancel, "Initialize Error") = vbOK Then
                                EndPort = StartPort + 1
                            Else
                                Exit Sub
                            End If
                        End If
                    End If
                    Dim i As Integer
                    For i = StartPort To EndPort
                        If Not EntryExists(CStr(i), lstPortInterval) Then lstPortInterval.AddItem i
                        DoEvents
                        If Not Working Then GoTo CleanUp
                    Next i
                    Exit Sub
                End If
                If Not EntryExists(CStr(StartPort), lstPortInterval) Then lstPortInterval.AddItem StartPort
            Else
                MsgBox "Starting port has a non integer value!", vbCritical, "Initialize Error"
                Exit Sub
            End If
        Else
            MsgBox "Starting port has a non numerical value!", vbCritical, "Initialize Error"
            Exit Sub
        End If
    Else
        Working = False
        Exit Sub
    End If
CleanUp:
    cmdPortAdd.Caption = "Add"
End Sub

Private Sub cmdPortClear_Click()
    'Clear inserted addresses
    lstPortInterval.Clear
    'Check for delete button (there won't be nothin to delete)
    If cmdPortDel.Enabled = True Then cmdPortDel.Enabled = False
End Sub

Private Sub cmdPortDel_Click()
    With lstPortInterval
        'Assure error free:
        If .ListCount = 0 Then cmdDel.Enabled = False: Exit Sub
        'Assure there is somethin selected
        If .ListIndex > -1 Then
            If .Selected(.ListIndex) Then
                .RemoveItem .ListIndex
            End If
        End If
        'If for some reason it's still enabled and
        'there's nothin to delete, disable button:
        If .ListCount = 0 Then cmdPortDel.Enabled = False
    End With
End Sub

Private Sub cmdSkip_Click()
    StartScan , True
    Working = False
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Caption = "Cancel" Then
        bolStop = True
        Working = False
    Else
        bolStop = False
        Working = False
    End If
    Select Case Action
        Case 1
            'Simple ping
            StartPing
        Case 2
            'simple scan
            StartScan
        Case 3
            'ping->scan
            Dim i As Integer
            For i = 0 To lstAdr.ListCount - 1
                StartScan lstAdr.List(i)
                Do
                    If Not Scanning Then Exit Do
                    If bolStop Then
                        StartScan , True
                        StatusBar.SimpleText = "Operation aborted by user!"
                        Exit Sub
                    End If
                    DoEvents
                Loop
            Next i
    End Select
End Sub

Private Sub Form_Load()
'Initialize
    Me.Height = 6555
    Me.Width = 7425
    Dim i As Integer
    For i = 0 To fraMain.UBound
        fraMain(i).Top = TabStrip.Height
        fraMain(i).BorderStyle = 0
        fraMain(i).Left = 120
        fraMain(i).Visible = False
    Next i
    StartScan , True
    fraScanStatus.Left = fraPortSettings.Left
    fraPingStatus.Left = fraPortSettings.Left
    LastFrame = 1
    'TabStrip_Click
    TabStrip.Tabs(2).Selected = True
    TabStrip.Tabs(1).Selected = True
    lblDesc = "Welcome to Network Tools"
    'Refresh check buttons bindings
    'buttons are affected by the following procedure thus they must be place under the calls
    chkPortAdv_Click
    chkHostAdv_Click
    fraScanStatus.Top = fraPortSettings.Top + fraPortSettings.Height
    cmdStart.Top = fraLog.Top + fraLog.Height + 105
    cmdSkip.Top = cmdStart.Top
'Multi-Ping:
    'Check if Del button is going to be enabled:
    If lstAdr.ListCount = 0 Then cmdDel.Enabled = False
    'Clean status label:
    lblPingStatus = vbLf & vbLf & "Percentage Done"
    'Check if last txtbox is going to be enabled:
    If chkTo.Value = 0 Then txtLast.Enabled = False
    'Refresh tooltip text:
    If lstAdr.ListCount = 0 Then lstAdr.ToolTipText = "No adresses to ping"
End Sub

Private Sub Form_Resize()
    Dim i As Byte
    For i = 0 To fraMain.UBound
        fraMain(i).Width = Me.Width - 120
        fraMain(i).Height = Me.Height - TabStrip.Height
    Next i
    cmdNext.Top = cmdStart.Top + TabStrip.Height
    cmdNext.Left = cmdStart.Left + cmdStart.Width + 210
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "If you liked it or something got buggy or i doesn't work or you hate it please leave a comment.", vbExclamation, "Sorry about the nag screen =)"
End Sub

Private Sub optAction_Click(Index As Integer)
    cmdNext.Enabled = True
    Const PING_SIMP = 0
    Const PING_ADV = 1
    Const SCANPORT_SIMP = 2
    Const SCANPORT_ADV = 3
    Select Case Index
        Case PING_SIMP
            lblDesc.Caption = "Perform a simple Ping on one or more machines of your network."
        Case PING_ADV
            lblDesc.Caption = "Perform a Ping on more then one machine and a PortScan on each machine found."
        Case SCANPORT_SIMP
            lblDesc.Caption = "Perform a simple Port Scan in one machine. Note: the machine will always pinged first."
    End Select
End Sub

Private Sub sckScan_Connect()
    'Adds to the list current open port
    LastOpenPort = sckScan.RemotePort
    txtLog = txtLog & "Found open port: " & LastOpenPort & vbNewLine
    sckScan.Close
End Sub

Private Sub TabStrip_Click()
    If LastFrame <> TabStrip.SelectedItem.Index Then
        'Something has changed:
        If TabStrip.SelectedItem.Index = 1 Then
            'User clicked General
            fraMain(0).Visible = True
            fraMain(1).Visible = False
        Else
            'user clicked other 2
            fraMain(0).Visible = False
            fraMain(1).Visible = True
        End If
        
        LastFrame = TabStrip.SelectedItem.Index
        
        'Reset flag for remembering last Ping checked option
        'because it is reseted when user makes a port scan
        SimpPing = False: SimpScan = False: MultiScan = False
        Select Case LastFrame
            Case 2 'Simple portscan:
                Action = 2
                SimpScan = True
                'Set to simple mode:
                chkHostAdv.Value = 0
                'Remember option
                chkPortAdv.Value = Remb.LastScanPortOpt
                'Refresh frame:
                chkHostAdv_Click
                'Disable advanced options
                chkHostAdv.Visible = False
                'Show Port Settings:
                If Not fraPortSettings.Visible Then fraPortSettings.Visible = True
                fraPingStatus.Visible = False
                fraScanStatus.Visible = True
                'There is no use of skiping since skiping a port is unnecessary because of it's speed
                cmdSkip.Visible = False
            Case 3 'Simple Ping
                Action = 1
                SimpPing = True
                'Make sure adv options is visible:
                If Not chkHostAdv.Visible Then chkHostAdv.Visible = True
                'hide port scan:
                If fraPortSettings.Visible Then fraPortSettings.Visible = False
                'Remember last option
                chkHostAdv.Value = Remb.LastPingHostOpt
                
                fraPingStatus.Visible = True
                fraScanStatus.Visible = False
                'There is no use of skiping since we can't intercept the ping fucntion
                cmdSkip.Visible = False
            Case 4 'ping->scan:
                Action = 3
                MultiScan = True
                chkHostAdv.Value = 1
                chkHostAdv.Visible = False
                fraPingHost.Visible = True
                fraPortSettings.Visible = True
                fraPingStatus.Visible = False
                fraScanStatus.Visible = True
                chkHostAdv.Visible = True
                fraPingStatus.Visible = False
                cmdSkip.Visible = True
                'Remember last opt:
                chkPortAdv.Value = Remb.LastMultiScanPortOpt
            Case 5 'about:
                frmAbout.Show
                TabStrip.Tabs(5).Selected = False
                TabStrip.Tabs(1).Selected = True
        End Select
        chkHostAdv_Click
        chkPortAdv_Click
        If LastFrame > 1 Then
            cmdNext.Caption = "Back <<"
            cmdNext.Cancel = True
            cmdNext.Default = False
        Else
            cmdNext.Caption = "Next >>"
            cmdNext.Cancel = False
            cmdNext.Default = True
        End If
    End If
End Sub

Private Sub tmrCheckStatus_Timer()
    If sckScan.State <> sckConnected Or sckScan.State <> sckConnecting Or sckScan.State <> sckHostResolved Or sckScan.State <> sckResolvingHost Then      'If the control isn't connected, connecting, resolving it's host, or already resolved it's host then...(in other words, it's hanging)
        sckScan.Close    'Stop the control from doing whatever it's doing and close it
        If chkPortAdv.Value = 0 Then
            'Set the remote port for the next port needed to be scanned
            sckScan.RemotePort = CurPort
        Else
            sckScan.RemotePort = PortInt(CurPort)
        End If
        sckScan.Connect  'Have it connect
        lblStatus.Caption = "Status - Scanning Port:"
        Port = sckScan.RemotePort 'Sets the status to the new port
        If CurPort < ProgressScan.Max Then
            ProgressScan = CurPort
        Else
            Working = False 'cancel
        End If
        CurPort = CurPort + 1 'Increment the variable for the next time we ask for a new port
    End If
End Sub

Private Sub tmrStats_Timer()
    If tmrCheckStatus.Enabled Then
        If LastPort <> 0 Then
            Rate = 4 * (CurPort - LastPort) - 1
            Metrics.Visible = True
            Metrics.Left = Rate.Left + Rate.Width
            RemPort = ProgressScan.Max - ProgressScan
            Dim EstimatedTime As Long
            EstimatedTime = RemPort / Rate
            If (EstimatedTime / 60) > 1 Then
                Dim Secs As Long
                'EstimatedTime = EstimatedTime / 60
                Secs = (EstimatedTime / 60 - Int(EstimatedTime / 60)) * 60
                EstRem = Int(EstimatedTime / 60) & " min " & Secs & " secs"
            Else
                EstRem = Int(EstimatedTime) & " secs"
            End If
            Percent = Round(ProgressScan * 100 / ProgressScan.Max) & "%"
        End If
        LastPort = CurPort
    End If
End Sub

Private Sub chkTo_Click()
    'User enables Interval Search:
    If chkTo.Value = 0 Then
        txtLast.Enabled = False
    Else
        txtLast.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Not Working Then
        Working = True
        cmdAdd.Caption = "Stop"
        If IsAddress(txtIP) Then
            'Data is an IP adress:
            
            If chkTo.Value = 1 Then
                
                If Trim(txtLast) = "" Then 'Trim is used to "clean" adjacent spaces
                'User forgot to insert an address let's just ignore the interval:
                    If Not EntryExists(txtIP, lstAdr) Then
                        'Single address entry:
                        'There is no repeated entry, ok to proceed:
                        lstAdr.AddItem txtIP
                        'Check if Rem button is enabled
                        If cmdDel.Enabled = False Then cmdDel.Enabled = True
                    End If
                
                ElseIf IsAddress(txtLast) Then 'User has inserted some value
                    'Interval of Adresses search routine:
                    Dim FirstAddr As Double, LastAddr As Double, IpAddr As String, CurAddr As Double
                    'First and last address converted to long so they can be calculated later:
                    FirstAddr = AddrToLong(txtIP)
                    LastAddr = AddrToLong(txtLast)
                    'Obviously first addr must be smaller
                    If FirstAddr < LastAddr Then  'OK to proceed
                        For CurAddr = FirstAddr To LastAddr
                            'From first address to the last:
                            'Convert it to IP - string - so we can show it in list
                            IpAddr = LongToAddr(CurAddr)
                            If Not EntryExists(IpAddr, lstAdr) Then
                                'Assure there are no duplicates
                                lstAdr.AddItem IpAddr
                            End If
                            'we don't wan't to make our app to "freeze"
                            DoEvents
                            'check if user pressed cancel
                            If Not Working Then GoTo CleanUp
                        Next CurAddr
                    End If
                    'at least one entry was made, so let's check for Remove button:
                    If cmdDel.Enabled = False Then cmdDel.Enabled = True
                End If
            
            'No interval search (not checked)
            ElseIf Not EntryExists(txtIP, lstAdr) Then
                'Single address entry:
                'No repeated entry, ok to proceed...
                lstAdr.AddItem txtIP
                'Check if Rem button is enabled
                If cmdDel.Enabled = False Then cmdDel.Enabled = True
            End If
        End If
    Else
        Working = False
        Exit Sub
    End If
CleanUp:
    'Clean up process:
    txtIP = "": txtLast = ""
    cmdAdd.Caption = "Add"
    
End Sub

Private Sub cmdDel_Click()
    With lstAdr
        'Assure error free:
        If .ListCount = 0 Then cmdDel.Enabled = False: Exit Sub
        'Assure there is somethin selected
        If .ListIndex > -1 Then
            If .Selected(.ListIndex) Then
                .RemoveItem .ListIndex
            End If
        End If
        'If for some reason it's still enabled and
        'there's nothin to delete, disable button:
        If .ListCount = 0 Then cmdDel.Enabled = False
    End With
End Sub

Private Sub lstAdr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'To spare some resources let's do this only when it's needed:
    Static LastCount As Long
    'Check if value has changed so we can or not refresh tooltip text:
    If LastCount <> lstAdr.ListCount Then
        If lstAdr.ListCount = 1 Then
            'Only one address to ping:
            lstAdr.ToolTipText = lstAdr.ListCount & " address to ping"
        ElseIf lstAdr.ListCount = 0 Then
            'No addresses to ping:
            lstAdr.ToolTipText = "No addresses to ping"
        Else
            'More than one address to ping:
            lstAdr.ToolTipText = lstAdr.ListCount & " addresses to ping"
        End If
        'Refresh last value:
        LastCount = lstAdr.ListCount
    End If
End Sub

Private Sub lstAdr_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Enable drag and drop support:
    Dim Addr As String
    'Check the kind of incoming value
    Addr = data.GetData(vbCFText)
    'Check if is an address
    If IsAddress(Addr) Then
        'Check if it exists in list:
        If Not EntryExists(Addr, lstAdr) Then
            lstAdr.AddItem Addr
            If cmdDel.Enabled = False Then cmdDel.Enabled = True
        End If
    End If
End Sub

Private Sub txtIP_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Enable dragdrop in textbox:
    Dim Addr As String
    Addr = data.GetData(vbCFText)
    If IsAddress(Addr) Then
        txtIP = Addr
    End If
End Sub

Private Sub txtLast_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Enable dragdrop in textbox:
    Dim Addr As String
    Addr = data.GetData(vbCFText)
    If IsAddress(Addr) Then
        txtLast = Addr
    End If
End Sub

Function EntryExists(Entry As String, List As ListBox) As Boolean
    With List
        'Simple loop in matching each value of each list value with the disired
        Dim i As Integer, Last As Integer
        EntryExists = False
        Last = .ListCount - 1
        If Last < 0 Then Exit Function
        For i = 0 To Last
            If .List(i) = Entry Then
                EntryExists = True
                Exit Function
            End If
        Next i
    End With
End Function

Sub StartPing()
    If Working Then Working = False: Exit Sub 'User pressed then stop working
    If lstAdr.ListCount = 0 Then GoTo CleanUp
    'Start working:
    Working = True
    'Enable cancel button
    cmdStart.Caption = "Cancel"
    Dim i As Long, Elapsed As Long, Remaining As Long, ret As Long
    'Initialize progress bar
    ProgressPing.Min = 0
    ProgressPing.Max = lstAdr.ListCount - 1
    
    'Start ping loop:
    For i = 0 To ProgressPing.Max
        'Initialize clock:
        Elapsed = GetTickCount
        'Change progressbar:
        ProgressPing = i
        'Check if user pressed cancel:
        If Not Working Then
            txtLog = txtLog & "Operation aborted by user!" & vbNewLine
            GoTo CleanUp
        End If
        'Status
        lblPingStatus = "Pinging: " & lstAdr.List(i) & vbLf & "Hosts remaining: " & ProgressPing.Max - ProgressPing & vbLf & "Estimated time: " & Round((ProgressPing.Max - ProgressPing) * Remaining / 1000, 3) & " secs"
        'Refresh form, if ommited user can't cancel operation:
        DoEvents
        'ret->delay of ping
        'Ping host:
        ret = PingHostByAdress(lstAdr.List(i))
        If ret >= 0 Then
            txtLog = txtLog & lstAdr.List(i) & ": " & ret & " ms" & vbNewLine
            DoEvents
            If Not Working Then
                txtLog = txtLog & "Operation aborted by user!" & vbNewLine
                GoTo CleanUp
            End If
        End If
        
        'This sets the delay of last action that will be used
        'for calculating remaining time:
        Remaining = GetTickCount - Elapsed
Next i
CleanUp:
    'Safest way of cleaning progress bar:
    ProgressPing = ProgressPing.Min
    'Reset button caption:
    cmdStart.Caption = "Start"
    'Work is done
    Working = False
    'Clean status label:
    lblPingStatus = vbLf & vbLf & "Percentage Done"
End Sub

Sub StartScan(Optional Host As String, Optional Cancel As Boolean)
    If Cancel Or bolStop Then GoTo CleanUp
    Scanning = True
    If Host = "" Then Host = txtIP
    If Not IsAddress(Host) Then
            MsgBox "Invalid IP address to scan.", vbCritical, "Initialize Error"
            GoTo CleanUp
    End If
    
    If Not Working Then
        Working = True
        cmdStart.Caption = "Cancel"
        Dim ret As Long
        StatusBar.SimpleText = "Pinging " & Host & "..."
        ret = PingHostByAdress(Host)
        If ret < 0 Then
            GoTo CleanUp
        Else
            StatusBar.SimpleText = Host & " reply: " & ret & " ms"
        End If
        'Disable controls and change caption:
        fraPingHost.Enabled = False
        fraPortSettings.Enabled = False
        
        'Simple Mode:
        If chkPortAdv.Value = 0 Then
            'Check values:
            If IsNumeric(txtStartPort) And IsNumeric(txtEndPort) Then
                'Both values must be integer!
                If (Int(txtStartPort) - Val(txtStartPort)) <> 0 Or (Int(txtEndPort) - Val(txtEndPort)) <> 0 Then
                    MsgBox "One of the ports has a non integer value!", vbCritical, "Initialize Error"
                    GoTo CleanUp
                Else
                    StartPort = CInt(txtStartPort.Text)
                    EndPort = CInt(txtEndPort.Text)
                    Debug.Print StartPort; EndPort
                    If StartPort >= txtEndPort Then
                        If MsgBox("Starting Port is bigger than ending Port. Ending Port will be set to a bigger value", vbCritical Or vbOKCancel, "Initialize Error") = vbOK Then
                            EndPort = StartPort + 1
                        Else
                            GoTo CleanUp
                        End If
                    End If
                    'Check port in interval of values
                    If StartPort < 1 Then StartPort = 1
                    If StartPort >= 32767 Then StartPort = 32766
                    If EndPort > 32767 Then EndPort = 32767
                    If EndPort < 2 Then EndPort = 2
                End If
            Else
                MsgBox "On of the ports has a non numerical value!", vbCritical, "Initialize Error"
                GoTo CleanUp
            End If
            
            CurPort = StartPort 'Sets the global variable (set up at the top) to the first port the user wants scanned
            LastPort = 0 'used for stats only
            ProgressScan.Max = EndPort
            ProgressScan.Min = StartPort
            ProgressScan = StartPort
            'Actualize any changed data
            txtStartPort = StartPort: txtEndPort = EndPort
            LastOpenPort = 0
            
            sckScan.Close                       'Close the control
            sckScan.RemoteHost = Host           'This sets all the controls to want to connect to the target the user specified
            sckScan.RemotePort = CurPort        'Sets the port needed to be scanned
            sckScan.Connect                     'Try to get it to connect
            CurPort = CurPort + 1               'Makes curport get larget by one for the next control
            lblStatus.Caption = "Status - Scanning port:"
            Port = CurPort 'Sets the status message to display the port currently being scanned
                    
        'Advanced Mode:
        Else
            If lstPortInterval.ListCount = 0 Then
                MsgBox "No ports to scan!", vbCritical, "Initialize Error"
                GoTo CleanUp
            End If
            Dim i As Integer
            'Redimension array of selected ports:
            ReDim PortInt(lstPortInterval.ListCount - 1)
            For i = 0 To lstPortInterval.ListCount - 1
                'Set port
                PortInt(i) = lstPortInterval.List(i)
            Next i
            
            CurPort = 0 'Sets the global variable (set up at the top) to the first port the user wants scanned
            LastPort = 0 'used for stats porpose only
            ProgressScan.Max = lstPortInterval.ListCount - 1
            ProgressScan.Min = 0
            ProgressScan = 0
            'Actualize any changed data
            
            'Close the control
            sckScan.Close
            'This sets all the controls to want to connect to the target the user specified
            sckScan.RemoteHost = Host
            'Sets the port needed to be scanned
            sckScan.RemotePort = PortInt(CurPort)
            'Try to get it to connect
            sckScan.Connect
            'Makes curport get larget by one for the next control
            CurPort = CurPort + 1
            lblStatus.Caption = "Status - Scanning port:"
            'Sets the status message to say the port currently being scanned
            Port = CurPort
        
        End If
        
        txtLog = txtLog & "Scanning " & Host & "..." & vbNewLine
        tmrCheckStatus.Enabled = True   'The timer is sort of like the clean up person...it looks for controls that haven't connected and are just sitting there and assigns them a new port
        tmrStats.Enabled = True
    Else
CleanUp:
        Working = False
        cmdStart.Caption = "Start"      'Sets the caption back if the user wants to go again
        tmrCheckStatus.Enabled = False  'Stop the timer
        tmrStats.Enabled = False
        lblStatus.Caption = "Status - Idle" 'Sets the label status
        sckScan.Close
        If Action <> 3 Then txtStartPort = CurPort - 1
        CurPort = 0
        If CInt(txtStartPort) < 1 Then txtStartPort = 1
        Rate = ""
        fraPingHost.Enabled = True
        fraPortSettings.Enabled = True
        Port = ""
        RemPort = ""
        Rate = ""
        ProgressScan = ProgressScan.Min
        Metrics.Visible = False
        EstRem = ""
        Percent = ""
        'On Error Resume Next
        'Set flag for multiple hosts scan
        Scanning = False
    End If
End Sub

Sub SafeDeclare(SetVar As Variant, ValVar As Variant)
    'very usefull for when it comes to graphs
    'it makes the best use of resources and controls flickering
    If SetVar <> ValVar Then SetVar = ValVar
End Sub
