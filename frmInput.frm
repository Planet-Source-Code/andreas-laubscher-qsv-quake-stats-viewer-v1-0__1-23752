VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Games"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmmDlg 
      Left            =   4920
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Log File"
      Filter          =   "Log Files (*.log)|*.log|All Files|*.*"
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdInput 
         Caption         =   "..."
         Height          =   285
         Left            =   4980
         TabIndex        =   6
         Top             =   420
         Width           =   255
      End
      Begin VB.TextBox txtInputFile 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   4695
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Input File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go!"
         Height          =   375
         Left            =   4140
         TabIndex        =   9
         Top             =   3300
         Width           =   1155
      End
      Begin MSComctlLib.ListView lstGames 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
         View            =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Games"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   4380
      Width           =   5415
      Begin VB.CheckBox chkAutoHide 
         Caption         =   "Auto hide this form on game select"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================='
' Author            : Andreas Laubscher                                                         '
' Date Written      : 30 May 2001                                                               '
' Description       : The Input form. Fairly simple and straightforward                         '
'==============================================================================================='

Option Explicit

Private Sub cmdGo_Click()
Dim selGames As String

    If lstGames.ListItems.Count = 0 Then Exit Sub
    
    ' Parse and display the specified game(s)
    Screen.MousePointer = vbHourglass
    
    selGames = ""
    ' Get our games string. has to end with a semi-colon, else we'll miss the last game
    For i = 1 To lstGames.ListItems.Count
        If lstGames.ListItems(i).Selected Then
            selGames = selGames & lstGames.ListItems(i).Text & ";"
        End If
    Next
    
    GetStats selGames
    
    Screen.MousePointer = vbDefault
    
    ' If our main form is hidden then show it
    If Not frmMain.Visible Then frmMain.Visible = True
    
    ' If the user wants, hide this form
    If chkAutoHide.Value = vbChecked Then Me.Visible = False
    
End Sub

Private Sub cmdInput_Click()
    On Error GoTo InputError
    
    ' Set up and show the Open dialog
    cmmDlg.Flags = cdlOFNHideReadOnly + cdlOFNShareAware
    cmmDlg.ShowOpen
    
    If Len(Trim(cmmDlg.FileName)) = 0 Then
        Exit Sub
    End If
    
    txtInputFile.Text = cmmDlg.FileName
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    ' Input our cfg file to get the weapons data
    GetWeapons
    
    ' Input our log file, and extract the games. This will take a while for big files
    GetGames txtInputFile.Text
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
InputError:
    
    If Err.Number = 32755 Then      ' Cancel was pressed
        Err.Clear
    Else
        MsgBox "An error has occurred while trying to open the file." & Chr(13) & "Error Number " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    End If
    
End Sub

Private Sub Form_Load()

    cmmDlg.InitDir = GetSetting(App.Title, "Settings", "InitDir")
    chkAutoHide.Value = GetSetting(App.Title, "Settings", "AutoHide", vbChecked)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tmpform As Form

    If Not frmMain.Visible Then
        For Each tmpform In Forms
            Unload tmpform
        Next
    End If
    
    If cmmDlg.FileName <> "" Then SaveSetting App.Title, "Settings", "InitDir", cmmDlg.FileName
    SaveSetting App.Title, "Settings", "AutoHide", chkAutoHide.Value
    
End Sub

Private Sub lstGames_DblClick()
    
    If lstGames.ListItems.Count = 0 Then Exit Sub
    
    ' Parse and display the specified game
    Screen.MousePointer = vbHourglass
    GetStats lstGames.SelectedItem.Text & ";"
    Screen.MousePointer = vbDefault
    
    ' If our main form is hidden then show it
    If Not frmMain.Visible Then frmMain.Visible = True
    
    ' If the user wants, hide this form
    If chkAutoHide.Value = vbChecked Then Me.Visible = False
    
End Sub
