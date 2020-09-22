VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "Comct232.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Counter-Strike - Dedicated Server Manager"
   ClientHeight    =   4065
   ClientLeft      =   2460
   ClientTop       =   2685
   ClientWidth     =   8190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame1 
      BorderStyle     =   0  'None
      Height          =   3240
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   645
      Width           =   7890
      Begin VB.CommandButton cmdLstCont 
         Caption         =   "Re&load Maps"
         Height          =   375
         Index           =   2
         Left            =   6480
         TabIndex        =   7
         Top             =   1890
         Width           =   1125
      End
      Begin VB.CommandButton cmdLstCont 
         Caption         =   "&Run Server"
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   5
         Top             =   2820
         Width           =   1125
      End
      Begin VB.CommandButton cmdLstCont 
         Caption         =   "&Clear All"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   4
         Top             =   1020
         Width           =   1125
      End
      Begin VB.CommandButton cmdLstCont 
         Caption         =   "Select &All"
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   3
         Top             =   540
         Width           =   1125
      End
      Begin VB.ListBox lstSList 
         Columns         =   3
         Height          =   2760
         Left            =   45
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   465
         Width           =   6015
      End
      Begin VB.Label lblHLPath 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   2205
         TabIndex        =   25
         Top             =   180
         Width           =   3750
      End
      Begin VB.Label lblTotMap 
         Height          =   210
         Left            =   90
         TabIndex        =   24
         Top             =   180
         Width           =   2025
      End
   End
   Begin VB.Frame frame1 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   585
      Visible         =   0   'False
      Width           =   7920
      Begin VB.ComboBox cboMods 
         Height          =   315
         ItemData        =   "frmMain.frx":125A
         Left            =   1350
         List            =   "frmMain.frx":125C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   900
         Width           =   1470
      End
      Begin VB.CommandButton cmdBrow 
         Caption         =   "..."
         Height          =   375
         Left            =   6555
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox ChkClearMap 
         Caption         =   "Clear Map Selection on server start"
         Height          =   270
         Left            =   3570
         TabIndex        =   21
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2835
      End
      Begin VB.TextBox txtMPlayers 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5910
         TabIndex        =   20
         Top             =   1410
         Width           =   645
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   1
         Left            =   2910
         TabIndex        =   19
         Top             =   1425
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Local Network Only"
         Height          =   315
         Left            =   1350
         TabIndex        =   18
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   1350
         TabIndex        =   13
         Top             =   1425
         Width           =   915
      End
      Begin ComCtl2.UpDown UDown 
         Height          =   375
         Index           =   0
         Left            =   2265
         TabIndex        =   12
         Top             =   1425
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtPort(0)"
         BuddyDispid     =   196614
         BuddyIndex      =   0
         OrigLeft        =   2280
         OrigTop         =   1425
         OrigRight       =   2520
         OrigBottom      =   1800
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHLPath 
         Height          =   375
         Left            =   1365
         TabIndex        =   8
         Top             =   360
         Width           =   5190
      End
      Begin ComCtl2.UpDown UDown 
         Height          =   375
         Index           =   1
         Left            =   3825
         TabIndex        =   14
         Top             =   1425
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UDown 
         Height          =   375
         Index           =   2
         Left            =   6555
         TabIndex        =   17
         Top             =   1410
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Max Players:"
         Height          =   285
         Left            =   4965
         TabIndex        =   16
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "to"
         Height          =   330
         Left            =   2610
         TabIndex        =   15
         Top             =   1485
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Port to use:"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   1485
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "HL MOD Path:"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Half-Life Path:"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   420
         Width           =   1305
      End
   End
   Begin MSComctlLib.TabStrip ts1 
      Height          =   3735
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Servers"
            Key             =   "server"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuration"
            Key             =   "config"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intCurFrame As Integer
Dim strHLPath As String
Dim intPort As Integer
Dim SH As New Shell
Dim oFld As Folder

Private Sub Form_Load()
    
    intCurFrame = 1
    txtMPlayers.Text = "20"
    txtPort(0).Text = "27016"
    txtPort(1).Text = "27021"

    With UDown(0)
        .BuddyControl = "txtport(0)"
        .BuddyProperty = "text"
        .Max = txtPort(1).Text
        .Min = txtPort(0).Text
    End With

    With UDown(1)
        .BuddyControl = "txtport(1)"
        .BuddyProperty = "text"
        .Max = txtPort(1).Text
        .Min = txtPort(0).Text
        .Value = "27021"
    End With

    With UDown(2)
        .BuddyControl = "txtmplayers"
        .BuddyProperty = "text"
        .Max = "32"
        .Min = "2"
        .Value = "20"
    End With

    frmMain.Show

    strHLPath = QueryValue("HKCU\Software\Valve\Half-Life", "InstallPath")
    If strHLPath <> "" Then
        txtHLPath.Text = strHLPath
    Else
        MsgBox "Can't retrive the Half-Life path. " & _
            "Check the Configuration", vbOKOnly + vbInformation, "Warning"
        ts1.Tabs(2).Selected = True
        Exit Sub
    End If

    LoadMODs
    LoadMaps
    frmAbout.Show 1

    'ts1.Tabs(2).Selected = True

End Sub

Private Sub DispWarn()
    MsgBox "Running more than 5 dedicated server in a " & _
        vbCrLf & "single machine will make performance " & _
        "degradation.", vbOKOnly + vbCritical, "Caution"
End Sub

Private Sub cmdLstCont_Click(Index As Integer)
Dim intTotMaps As Integer
Dim nCount As Integer
Dim nPort1 As Integer
Dim strCurDir As String
Dim strCurDrv As String

    strCurDir = CurDir
    strCurDrv = Left(CurDir, 2)

    Select Case Index
        Case 0
            intTotMaps = lstSList.ListCount
            If intTotMaps > 5 Then DispWarn: intTotMaps = 5
            For nCount = 0 To intTotMaps - 1
                lstSList.Selected(nCount) = True
            Next
        Case 1
            intTotMaps = lstSList.ListCount
            For nCount = 0 To intTotMaps - 1
                lstSList.Selected(nCount) = False
            Next
        Case 2
            LoadMaps
        Case 3
            intTotMaps = lstSList.ListCount
            If lstSList.SelCount > 5 Then
                DispWarn
                lstSList.SetFocus
                Exit Sub
            End If
            nPort1 = Val(txtPort(0).Text) - 1
        
            ChDrive Left(strHLPath, 2)
            ChDir strHLPath

            For nCount = 0 To intTotMaps - 1
                If lstSList.Selected(nCount) = True Then
                    RunServer lstSList.List(nCount), Str(nPort1 + 1)
                    nPort1 = nPort1 + 1
                    If ChkClearMap.Value Then
                        lstSList.Selected(nCount) = False
                    End If
                End If
            Next
            ChDrive strCurDrv
            ChDir strCurDir
    End Select
End Sub

Private Sub RunServer(strMap As String, nPort As Integer)
Dim strCmd As String

strCmd = strHLPath & "\hlds.exe -game " & cboMods.Text
If Check1.Value Then
    strCmd = strCmd & " -nomaster -insecure +sv_lan 1 "
End If
strCmd = strCmd & " +maxplayers " & txtMPlayers.Text
strCmd = strCmd & " +map " & strMap
strCmd = strCmd & " +port " & nPort
Shell strCmd, vbMinimizedNoFocus

End Sub
Private Sub ShowLables()
    lblTotMap.Caption = Str(lstSList.ListCount) & " maps available."
    lblHLPath.Caption = strHLPath & "\" & cboMods.Text & "\maps"
End Sub
Private Sub LoadMaps()
Dim strMapFile As String

    If Not isFileExist(strHLPath & "\hl.exe") Then
        MsgBox "Invalid Half-Life path, hl.exe is not " & _
            "available in the specified path.", vbOKOnly + vbInformation, "Invalid Path"
        lstSList.Clear
        ShowLables
        Exit Sub
    End If
    
    lstSList.Clear
    strHLPath = txtHLPath.Text
    strMapFile = Dir(strHLPath & "\" & cboMods.Text & "\maps\*.bsp")
    If strMapFile = "" Then
        MsgBox "No map files available in the specified path.", _
            vbOKOnly + vbInformation, "Caution"
        Exit Sub
    Else
        Do
            lstSList.AddItem strMapFile
            strMapFile = Dir
        Loop Until strMapFile = ""
    End If
    ShowLables
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show 1

End Sub

Private Sub mnuToolsExit_Click()
Unload Me

End Sub

Private Sub ts1_Click()
    If ts1.SelectedItem.Index = intCurFrame Then Exit Sub
    frame1(ts1.SelectedItem.Index).Visible = True
    frame1(intCurFrame).Visible = False
    intCurFrame = ts1.SelectedItem.Index
    If ts1.SelectedItem.Index = 1 Then LoadMaps
End Sub

Private Sub LoadMODs()
Dim strFile As String

    cboMods.Clear
    If strHLPath = "" Then Exit Sub
    strFile = Dir(strHLPath & "\", vbDirectory)
    If Trim(strFile) = "" Then Exit Sub
    
    Do
        If Not InStr(1, strFile, ".") Then
        If isFileExist(strHLPath & "\" & strFile & _
            "\cl_dlls\client.dll") Then _
            cboMods.AddItem (strFile)
        End If
        strFile = Dir
    Loop Until strFile = ""
    If cboMods.ListCount >= 1 Then cboMods.ListIndex = 0
    
End Sub

Private Function isFileExist(cFile As String) As Boolean
    If Trim(cFile) = "" Then isFileExist = False: Exit Function
    On Error GoTo chkerr
    isFileExist = True
    Open cFile For Input As #1
    Close #1
    Exit Function
chkerr:
    isFileExist = False
    Resume Next
End Function

Private Sub txtHLPath_Change()
    strHLPath = txtHLPath.Text
    LoadMODs
End Sub

Private Sub cmdBrow_Click()
    On Error Resume Next
    Set oFld = SH.BrowseForFolder(hWnd, "Select Half-Life Folder", 0)
    strHLPath = oFld.Items.Item.Path
    txtHLPath.Text = strHLPath
    On Error GoTo 0
End Sub

