VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QSV"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7620
      TabIndex        =   149
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open..."
      Height          =   375
      Left            =   6360
      TabIndex        =   148
      Top             =   7680
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   13256
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Leaderboard"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Game"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Clans"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Players"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame9"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Weapons"
      TabPicture(4)   =   "frmMain.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame16"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame15"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame18"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Timeline"
      TabPicture(5)   =   "frmMain.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame19"
      Tab(5).Control(1)=   "Frame17"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame17 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   160
         Top             =   360
         Width           =   8595
         Begin VB.PictureBox picTimeline 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            FillStyle       =   0  'Solid
            Height          =   4515
            Left            =   180
            ScaleHeight     =   4455
            ScaleWidth      =   8235
            TabIndex        =   161
            Top             =   240
            Width           =   8295
         End
      End
      Begin VB.Frame Frame14 
         Height          =   7035
         Left            =   -74880
         TabIndex        =   126
         Top             =   360
         Width           =   2595
         Begin MSComctlLib.ListView lstWeapons 
            Height          =   6675
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   11774
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7035
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   2595
         Begin MSComctlLib.ListView lstClans 
            Height          =   6675
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   11774
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Height          =   7035
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   2595
         Begin MSComctlLib.ListView lstPlayers 
            Height          =   6675
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   11774
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   8595
         Begin VB.Label Label63 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Game Type:"
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
            Left            =   180
            TabIndex        =   155
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label62 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
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
            Left            =   3180
            TabIndex        =   154
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label lblGameType 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Game 999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   153
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label lblGameMap 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "24:60:60"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   152
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label lblGameDuration 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "24:60:60"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   45
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblGameNumber 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Game 999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   44
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label lblGameFile 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "C:\Games\Quake\Baseq3\Games.log"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   43
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label25 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Log File:"
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
            Left            =   180
            TabIndex        =   42
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblGameSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   41
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblGameFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   40
            Top             =   540
            Width           =   615
         End
         Begin VB.Label lblGamePlayers 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label21 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Suicides:"
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
            Left            =   6540
            TabIndex        =   38
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label20 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Frags:"
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
            Left            =   6540
            TabIndex        =   37
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label Label19 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Players:"
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
            Left            =   6540
            TabIndex        =   36
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Duration:"
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
            Left            =   3180
            TabIndex        =   35
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Game Name:"
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
            Left            =   180
            TabIndex        =   34
            Top             =   840
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8595
         Begin VB.Label Label58 
            Caption         =   "Winner"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   151
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblWinnerScore 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Score: 999"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6540
            TabIndex        =   4
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label lblWinnerClan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Clan: SomeClan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6540
            TabIndex        =   3
            Top             =   540
            Width           =   1875
         End
         Begin VB.Label lblWinnerName 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   180
            TabIndex        =   2
            Top             =   600
            Width           =   6435
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5955
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4695
         Begin MSComctlLib.ListView lstLeaderboard 
            Height          =   5415
            Left            =   120
            TabIndex        =   21
            Top             =   420
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   9551
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Rank"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Score"
               Object.Width           =   1270
            EndProperty
         End
         Begin VB.Label Label30 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Leaderboard:"
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
            Left            =   180
            TabIndex        =   49
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2895
         Left            =   4740
         TabIndex        =   6
         Top             =   1440
         Width           =   3975
         Begin VB.Label Label69 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective:"
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
            TabIndex        =   174
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label68 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Skill:"
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
            TabIndex        =   173
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lblPBSkill 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   172
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label lblPBEffective 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   171
            Top             =   1560
            Width           =   2715
         End
         Begin VB.Label lblPBSurvivor 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   28
            Top             =   2520
            Width           =   2715
         End
         Begin VB.Label lblPBMdwak 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   27
            Top             =   2280
            Width           =   2715
         End
         Begin VB.Label lblPBMkiol 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   26
            Top             =   2040
            Width           =   2715
         End
         Begin VB.Label lblPBNett 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   25
            Top             =   1800
            Width           =   2715
         End
         Begin VB.Label lblPBSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   24
            Top             =   1080
            Width           =   2715
         End
         Begin VB.Label lblPBDeaths 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   23
            Top             =   840
            Width           =   2715
         End
         Begin VB.Label lblPBFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   22
            Top             =   600
            Width           =   2715
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Nett:"
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
            TabIndex        =   14
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Survivor:"
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
            TabIndex        =   13
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mdwak:"
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
            TabIndex        =   12
            Top             =   2280
            Width           =   795
         End
         Begin VB.Label Label8 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mkiol:"
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
            TabIndex        =   11
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Players Billboard"
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
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   3675
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Suicides:"
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
            TabIndex        =   9
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Deaths:"
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
            TabIndex        =   8
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Frags:"
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
            TabIndex        =   7
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame Frame7 
         Height          =   5955
         Left            =   -74880
         TabIndex        =   46
         Top             =   1440
         Width           =   8595
         Begin MSComctlLib.ListView lstRawLog 
            Height          =   5415
            Left            =   120
            TabIndex        =   47
            Top             =   420
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   9551
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Killer"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Killee"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Weapon"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label29 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Raw Log:"
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
            Left            =   180
            TabIndex        =   48
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame9 
         Height          =   3375
         Left            =   -72360
         TabIndex        =   52
         Top             =   360
         Width           =   6075
         Begin VB.Label Label49 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Skill:"
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
            Left            =   3180
            TabIndex        =   165
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label46 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Effectivity:"
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
            Left            =   3180
            TabIndex        =   164
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label lblPEffectivity 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   163
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label lblPSkill 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   162
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label Label26 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Ping:"
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
            Left            =   3180
            TabIndex        =   117
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblPPing 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   116
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblPSurvived 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   115
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Survived:"
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
            Left            =   3180
            TabIndex        =   114
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Hit Rate:"
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
            TabIndex        =   113
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblPHitrate 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   112
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label22 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Nett:"
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
            Left            =   3180
            TabIndex        =   111
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label lblPNett 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   110
            ToolTipText     =   "Calculated as: Frags - Deaths - (2 * Suicides)"
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label lblPBestWeapon 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   77
            Top             =   3000
            Width           =   4515
         End
         Begin VB.Label lblPPickedOnBy 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   76
            Top             =   2700
            Width           =   4515
         End
         Begin VB.Label lblPPickedOn 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   75
            Top             =   2460
            Width           =   4515
         End
         Begin VB.Label lblPMdwak 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   74
            ToolTipText     =   "Most Deaths Without A Kill"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblPMkiol 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   73
            ToolTipText     =   "Most Kills In One Life"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblPSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   72
            Top             =   1860
            Width           =   1575
         End
         Begin VB.Label lblPDeaths 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   71
            Top             =   1860
            Width           =   1575
         End
         Begin VB.Label lblPFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   70
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label lblPScore 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   69
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label lblPClan 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   68
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblPRanking 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   67
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label lblPName 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
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
            Left            =   1440
            TabIndex        =   66
            Top             =   240
            Width           =   4515
         End
         Begin VB.Label Label31 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
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
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label32 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Clan:"
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
            TabIndex        =   63
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label33 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking:"
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
            TabIndex        =   62
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label34 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Score:"
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
            TabIndex        =   61
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label35 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Frags:"
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
            TabIndex        =   60
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label36 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Deaths:"
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
            TabIndex        =   59
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label Label37 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mkiol:"
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
            TabIndex        =   58
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label38 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Suicides:"
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
            Left            =   3180
            TabIndex        =   57
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label Label39 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mdwak:"
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
            Left            =   3180
            TabIndex        =   56
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label40 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Picked on:"
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
            TabIndex        =   55
            Top             =   2460
            Width           =   1095
         End
         Begin VB.Label Label41 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Picked on by:"
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
            TabIndex        =   54
            Top             =   2700
            Width           =   1095
         End
         Begin VB.Label Label42 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Best Weapon:"
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
            TabIndex        =   53
            Top             =   3000
            Width           =   1095
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3795
         Left            =   -72360
         TabIndex        =   65
         Top             =   3600
         Width           =   6075
         Begin TabDlg.SSTab SSTab2 
            Height          =   3435
            Left            =   180
            TabIndex        =   123
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   6059
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            ShowFocusRect   =   0   'False
            TabCaption(0)   =   "Frags Breakdown"
            TabPicture(0)   =   "frmMain.frx":03B2
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lstPKilled"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Weapon Usage"
            TabPicture(1)   =   "frmMain.frx":03CE
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lstPWeapons"
            Tab(1).ControlCount=   1
            Begin MSComctlLib.ListView lstPKilled 
               Height          =   2895
               Left            =   120
               TabIndex        =   124
               Top             =   420
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   5106
               View            =   3
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Name"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Frags"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Deaths"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Nett"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView lstPWeapons 
               Height          =   2895
               Left            =   -74880
               TabIndex        =   125
               Top             =   420
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   5106
               View            =   3
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Name"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Frags"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Deaths"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Nett"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3255
         Left            =   -72360
         TabIndex        =   80
         Top             =   360
         Width           =   6075
         Begin VB.Label lblCMembers 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   109
            Top             =   540
            Width           =   3195
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Members:"
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
            TabIndex        =   108
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblCPPDeaths 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   107
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblCPPSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   106
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblCPPInterfrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   105
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblCPPFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   104
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label56 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Interfrags:"
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
            TabIndex        =   103
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label lblCInterfrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   102
            Top             =   2400
            Width           =   1515
         End
         Begin VB.Label Label44 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Worst Player:"
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
            TabIndex        =   101
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label lblCWorstPlayer 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   100
            Top             =   1380
            Width           =   3195
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Best Player:"
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
            TabIndex        =   99
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label lblCBestPlayer 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   98
            Top             =   1140
            Width           =   3195
         End
         Begin VB.Label Label55 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mdwak:"
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
            TabIndex        =   94
            Top             =   2940
            Width           =   1095
         End
         Begin VB.Label Label54 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Suicides:"
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
            TabIndex        =   93
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label53 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Mkiol:"
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
            TabIndex        =   92
            Top             =   2700
            Width           =   1095
         End
         Begin VB.Label Label52 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Deaths:"
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
            TabIndex        =   91
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label51 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Frags:"
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
            TabIndex        =   90
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label50 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Score:"
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
            TabIndex        =   89
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label47 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
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
            TabIndex        =   88
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblCName 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   87
            Top             =   240
            Width           =   3195
         End
         Begin VB.Label lblCScore 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   86
            Top             =   840
            Width           =   3195
         End
         Begin VB.Label lblCFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   85
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label lblCDeaths 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   84
            Top             =   1920
            Width           =   1515
         End
         Begin VB.Label lblCSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   83
            Top             =   2160
            Width           =   1515
         End
         Begin VB.Label lblCMkiol 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   82
            Top             =   2700
            Width           =   3195
         End
         Begin VB.Label lblCMdwak 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   81
            Top             =   2940
            Width           =   3195
         End
      End
      Begin VB.Frame Frame12 
         Height          =   3915
         Left            =   -72360
         TabIndex        =   95
         Top             =   3480
         Width           =   6075
         Begin MSComctlLib.ListView lstCPlayers 
            Height          =   3375
            Left            =   180
            TabIndex        =   96
            Top             =   420
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   5953
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Frags"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label59 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Clan Members:"
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
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame16 
         Height          =   1215
         Left            =   -72360
         TabIndex        =   141
         Top             =   360
         Width           =   6075
         Begin VB.Label Label64 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Game Statistics:"
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
            TabIndex        =   157
            Top             =   240
            Width           =   5655
         End
         Begin VB.Label Label45 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Most Dangerous:"
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
            TabIndex        =   145
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblWDangerous 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   144
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label28 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Best Weapon:"
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
            TabIndex        =   143
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label lblWBestWeapon 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   142
            Top             =   540
            Width           =   4215
         End
      End
      Begin VB.Frame Frame15 
         Height          =   2535
         Left            =   -72360
         TabIndex        =   128
         Top             =   1440
         Width           =   6075
         Begin VB.Label lblWImmune 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   159
            Top             =   2160
            Width           =   4455
         End
         Begin VB.Label labelsomething 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Immune:"
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
            TabIndex        =   158
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label48 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Log Name:"
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
            TabIndex        =   147
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblWLogName 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   146
            Top             =   540
            Width           =   4455
         End
         Begin VB.Label Label81 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Best User:"
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
            TabIndex        =   140
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label80 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Picked on:"
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
            TabIndex        =   139
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label78 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Suicides:"
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
            TabIndex        =   138
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label75 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Frags:"
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
            TabIndex        =   137
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label74 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Score:"
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
            TabIndex        =   136
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label71 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
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
            TabIndex        =   135
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblWName 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   134
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label lblWScore 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   133
            Top             =   780
            Width           =   4455
         End
         Begin VB.Label lblWFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   132
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label lblWSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   131
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label lblWPickedOn 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   130
            Top             =   1920
            Width           =   4455
         End
         Begin VB.Label lblWBestUser 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   129
            Top             =   1620
            Width           =   4455
         End
      End
      Begin VB.Frame Frame18 
         Height          =   3555
         Left            =   -72360
         TabIndex        =   166
         Top             =   3840
         Width           =   6075
         Begin MSComctlLib.ListView lstWUsage 
            Height          =   3015
            Left            =   180
            TabIndex        =   167
            Top             =   420
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   5318
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Frags"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Deaths"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Nett"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label65 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Usage:"
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
            TabIndex        =   168
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   4740
         TabIndex        =   15
         Top             =   4200
         Width           =   3975
         Begin VB.Label lblCBNett 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   32
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label lblCBSuicides 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   31
            Top             =   1080
            Width           =   2715
         End
         Begin VB.Label lblCBDeaths 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   30
            Top             =   840
            Width           =   2715
         End
         Begin VB.Label lblCBFrags 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "AnExtremelyLongName"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1140
            TabIndex        =   29
            Top             =   600
            Width           =   2715
         End
         Begin VB.Label Label16 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Nett:"
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
            TabIndex        =   20
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Label15 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Suicides:"
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
            TabIndex        =   19
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Deaths:"
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
            TabIndex        =   18
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Frags:"
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
            TabIndex        =   17
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Clan Billboard"
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
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   3675
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1635
         Left            =   4740
         TabIndex        =   118
         Top             =   5760
         Width           =   3975
         Begin VB.OptionButton optSkill 
            Caption         =   "Skill"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   156
            Top             =   780
            Width           =   2415
         End
         Begin VB.OptionButton optRankCustom 
            Caption         =   "Custom Score"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   122
            Top             =   1260
            Width           =   2355
         End
         Begin VB.OptionButton optRankNett 
            Caption         =   "Nett"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   121
            Top             =   1020
            Width           =   2355
         End
         Begin VB.OptionButton optRankFrags 
            Caption         =   "Frags"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   120
            Top             =   540
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Scoring System"
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
            Left            =   180
            TabIndex        =   119
            Top             =   300
            Width           =   3675
         End
      End
      Begin VB.Frame Frame19 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   169
         Top             =   5100
         Width           =   8595
         Begin MSFlexGridLib.MSFlexGrid flxLegend 
            Height          =   1935
            Left            =   180
            TabIndex        =   170
            Top             =   240
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   0
            BackColorFixed  =   0
            BackColorBkg    =   0
            GridColor       =   4210752
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Label Label57 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Press 'F1' to view the help file"
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
      Left            =   180
      TabIndex        =   150
      Top             =   7800
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================='
' Author            : Andreas Laubscher                                                         '
' Date Written      : 30 May 2001                                                               '
' Description       : The stats viewer. Doesn't really do all that much processing. Mostly      '
'                   : queries the pre-created collections.                                      '
'==============================================================================================='
Option Explicit

Private Sub cmdOpen_Click()

    frmInput.Show
    
End Sub

Private Sub cmdQuit_Click()

    Unload Me
    
End Sub

Private Sub flxLegend_Click()

    flxLegend.CellFontStrikeThrough = Not flxLegend.CellFontStrikeThrough
    stsPlayers(flxLegend.Text).dspGraph = Not flxLegend.CellFontStrikeThrough
    GenerateTimeLine
    
End Sub

Private Sub Form_Load()
    
    ' Initialize all the controls
    InitializeStatsForm "All"
    
    ' I have decided not to upload the help file. People who want it can drop me a line at: qsv@mighty.co.za
    'App.HelpFile = App.Path & "\qsv.chm"
    
    frmBkg.Show
    frmInput.Show
    
    ' Initially we don't want to show the form
    Me.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tmpform                 As Form

    ' Close all the forms
    For Each tmpform In Forms
        Unload tmpform
    Next
    
End Sub


Private Sub lstClans_Click()
'==============================================================================================='
' Load the information to screen. Some entries are formatted to look better                     '
'==============================================================================================='

    If lstClans.ListItems.Count = 0 Then Exit Sub
    
    tmpString = lstClans.SelectedItem.Text
    i = stsClan(tmpString).lPlayers
    
    lblCName.Caption = stsClan(tmpString).sName
    lblCMembers.Caption = i
    lblCScore.Caption = stsClan(tmpString).lScore
    lblCBestPlayer.Caption = stsClan(tmpString).sBestPlayer
    lblCWorstPlayer.Caption = stsClan(tmpString).sWorstPlayer
    lblCFrags.Caption = stsClan(tmpString).lFrags
    lblCPPFrags.Caption = IIf(stsClan(tmpString).lFrags > 0, Format(stsClan(tmpString).lFrags / i, "0.0#"), 0)
    lblCDeaths.Caption = stsClan(tmpString).lDeaths
    lblCPPDeaths.Caption = IIf(stsClan(tmpString).lDeaths > 0, Format(stsClan(tmpString).lDeaths / i, "0.0#"), 0)
    lblCSuicides.Caption = stsClan(tmpString).lSuicides
    lblCPPSuicides.Caption = IIf(stsClan(tmpString).lSuicides > 0, Format(stsClan(tmpString).lSuicides / i, "0.0#"), 0)
    lblCInterfrags.Caption = stsClan(tmpString).lInterfrags
    lblCPPInterfrags.Caption = IIf(stsClan(tmpString).lInterfrags > 0, Format(stsClan(tmpString).lInterfrags / i, "0.0#"), 0)
    lblCMkiol.Caption = stsClan(tmpString).lMkiol
    lblCMdwak.Caption = stsClan(tmpString).lMdwak
    
    lstCPlayers.ListItems.Clear
    Tm1 = 32000
    Tm2 = -32000
    For i = 1 To stsClan(tmpString).colClanPlayers.Count
        With lstCPlayers.ListItems.Add(, , stsClan(tmpString).colClanPlayers(i).sName)
            .SubItems(1) = stsClan(tmpString).colClanPlayers(i).lFrags
        End With
        If stsClan(tmpString).colClanPlayers(i).lFrags < Tm1 Then
            Tm1 = stsClan(tmpString).colClanPlayers(i).lFrags
            lblCBestPlayer.Caption = stsClan(tmpString).colClanPlayers(i).sName
        End If
        If stsClan(tmpString).colClanPlayers(i).lFrags > Tm2 Then
            Tm2 = stsClan(tmpString).colClanPlayers(i).lFrags
            lblCWorstPlayer.Caption = stsClan(tmpString).colClanPlayers(i).sName
        End If
    Next
    
End Sub

Private Sub lstCPlayers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If (ColumnHeader.Position - 1) <> lstCPlayers.SortKey Then
        lstCPlayers.SortKey = ColumnHeader.Position - 1
        lstCPlayers.SortOrder = lvwAscending
    Else
        If lstCPlayers.SortOrder = lvwAscending Then
            lstCPlayers.SortOrder = lvwDescending
        Else
            lstCPlayers.SortOrder = lvwAscending
        End If
    End If
    
End Sub

Private Sub lstPKilled_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If (ColumnHeader.Position - 1) <> lstPKilled.SortKey Then
        lstPKilled.SortKey = ColumnHeader.Position - 1
        lstPKilled.SortOrder = lvwAscending
    Else
        If lstPKilled.SortOrder = lvwAscending Then
            lstPKilled.SortOrder = lvwDescending
        Else
            lstPKilled.SortOrder = lvwAscending
        End If
    End If
    
End Sub

Private Sub lstPlayers_Click()
'==============================================================================================='
' Load the information to screen. Some entries are formatted to look better                     '
'==============================================================================================='

    If lstPlayers.ListItems.Count = 0 Then Exit Sub
    
    tmpString = lstPlayers.SelectedItem.Text
    
    lblPName.Caption = stsPlayers(tmpString).sName
    lblPClan.Caption = stsPlayers(tmpString).sClan
    lblPEffectivity.Caption = IIf(stsPlayers(tmpString).lEffectivity <> 0, Format(stsPlayers(tmpString).lEffectivity, "0.0##"), "")
    lblPSkill.Caption = IIf(stsPlayers(tmpString).lSkill <> 0, Format(stsPlayers(tmpString).lSkill, "0.0##"), "")
    lblPRanking.Caption = IIf(stsPlayers(tmpString).lRanking = 1, "Winner", stsPlayers(tmpString).lRanking)
    lblPHitrate.Caption = Format((stsPlayers(tmpString).lFrags / (stsGame.Duration / 60)), "0.0#") & " Kpm"
    lblPPing = stsPlayers(tmpString).lPing
    lblPScore.Caption = stsPlayers(tmpString).lScore & " Points"
    lblPSurvived.Caption = IIf(stsPlayers(tmpString).lSurvived > 60, stsPlayers(tmpString).lSurvived \ 60 & ":" & Format(stsPlayers(tmpString).lSurvived Mod 60, "00") & " m", stsPlayers(tmpString).lSurvived & " s")
    lblPFrags.Caption = stsPlayers(tmpString).lFrags
    lblPNett.Caption = stsPlayers(tmpString).lNett
    lblPDeaths.Caption = stsPlayers(tmpString).lDeaths
    lblPSuicides.Caption = stsPlayers(tmpString).lSuicides
    lblPMkiol.Caption = stsPlayers(tmpString).lMkiol
    lblPMdwak.Caption = stsPlayers(tmpString).lMdwak
    
    If stsPlayers(tmpString).sBestWeapon <> "" Then
        lblPBestWeapon.Caption = stsWeapons(stsPlayers(tmpString).sBestWeapon).sDescription & " (" & stsPlayers(tmpString).colPWeapon(stsPlayers(tmpString).sBestWeapon).lFrags & " Frags)"
    Else
        lblPBestWeapon.Caption = "None"
    End If
    
    If stsPlayers(tmpString).sPickedOn = "" Then
        lblPPickedOn.Caption = "No-one... Lamer!"
    Else
        If stsPlayers(tmpString).sPickedOn = tmpString Then
            lblPPickedOn.Caption = "Himself... Hahaaa!"
        Else
            lblPPickedOn.Caption = stsPlayers(tmpString).sPickedOn
        End If
        lblPPickedOn.Caption = lblPPickedOn.Caption & " (" & stsPlayers(tmpString).colPPlayer(stsPlayers(tmpString).sPickedOn).lFrags & " Frags)"
    End If
    
    If stsPlayers(tmpString).sPickedOnBy = "" Then
        lblPPickedOnBy.Caption = "No-one... Impressive"
    Else
        lblPPickedOnBy.Caption = stsPlayers(tmpString).sPickedOnBy & " (" & stsPlayers(stsPlayers(tmpString).sPickedOnBy).colPPlayer(stsPlayers(tmpString).sName).lFrags & " Frags)"
    End If
    
    frmMain.lstPKilled.ListItems.Clear
    For i = 1 To stsPlayers(tmpString).colPPlayer.Count
        If stsPlayers(tmpString).colPPlayer(i).lFrags > 0 Then
            With lstPKilled.ListItems.Add(, , stsPlayers(tmpString).colPPlayer(i).sName)
                .SubItems(1) = stsPlayers(tmpString).colPPlayer(i).lFrags
                .SubItems(2) = stsPlayers(tmpString).colPPlayer(i).lDeaths
                .SubItems(3) = stsPlayers(tmpString).colPPlayer(i).lFrags - stsPlayers(tmpString).colPPlayer(i).lDeaths
            End With
        End If
    Next
    
    frmMain.lstPWeapons.ListItems.Clear
    For i = 1 To stsPlayers(tmpString).colPWeapon.Count
        If (stsPlayers(tmpString).colPWeapon(i).lFrags > 0) Or _
           (stsPlayers(tmpString).colPWeapon(i).lDeaths > 0) Or _
           (stsPlayers(tmpString).colPWeapon(i).lSuicides > 0) Then
            With lstPWeapons.ListItems.Add(, , stsWeapons(stsPlayers(tmpString).colPWeapon(i).sName).sDescription)
                .SubItems(1) = stsPlayers(tmpString).colPWeapon(i).lFrags
                .SubItems(2) = stsPlayers(tmpString).colPWeapon(i).lDeaths
                .SubItems(3) = stsPlayers(tmpString).colPWeapon(i).lFrags - stsPlayers(tmpString).colPWeapon(i).lDeaths
            End With
        End If
    Next
    
End Sub

Private Sub lstPWeapons_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If (ColumnHeader.Position - 1) <> lstPWeapons.SortKey Then
        lstPWeapons.SortKey = ColumnHeader.Position - 1
        lstPWeapons.SortOrder = lvwAscending
    Else
        If lstPWeapons.SortOrder = lvwAscending Then
            lstPWeapons.SortOrder = lvwDescending
        Else
            lstPWeapons.SortOrder = lvwAscending
        End If
    End If

End Sub

Private Sub lstRawLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If (ColumnHeader.Position - 1) <> lstRawLog.SortKey Then
        lstRawLog.SortKey = ColumnHeader.Position - 1
        lstRawLog.SortOrder = lvwAscending
    Else
        If lstRawLog.SortOrder = lvwAscending Then
            lstRawLog.SortOrder = lvwDescending
        Else
            lstRawLog.SortOrder = lvwAscending
        End If
    End If
    
End Sub

Private Sub lstWeapons_Click()
'==============================================================================================='
' Load the information to screen. Some entries are formatted to look better                     '
'==============================================================================================='

    If lstWeapons.ListItems.Count = 0 Then Exit Sub
    
    lstWUsage.ListItems.Clear
    tmpString = lstWeapons.SelectedItem.Key
    
    lblWName.Caption = stsWeapons(tmpString).sDescription
    lblWLogName.Caption = stsWeapons(tmpString).sName
    lblWScore.Caption = stsWeapons(tmpString).lScore
    lblWFrags.Caption = stsWeapons(tmpString).lFrags
    lblWSuicides.Caption = stsWeapons(tmpString).lSuicides
    lblWBestUser.Caption = IIf(stsWeapons(tmpString).sBestUser <> "", stsWeapons(tmpString).sBestUser & " (" & stsPlayers(stsWeapons(tmpString).sBestUser).colPWeapon(tmpString).lFrags & " Frags)", "No-one")
    lblWPickedOn.Caption = IIf(stsWeapons(tmpString).sPickedOn <> "", stsWeapons(tmpString).sPickedOn & " (" & stsWeapons(tmpString).colWPlayers(stsWeapons(tmpString).sPickedOn).lDeaths & " Deaths)", "No-one")
    lblWImmune.Caption = IIf(stsWeapons(tmpString).sImmune <> "", stsWeapons(tmpString).sImmune & " (" & stsWeapons(tmpString).colWPlayers(stsWeapons(tmpString).sImmune).lDeaths & " Deaths)", "No-one")
    
    For i = 1 To stsWeapons(tmpString).colWPlayers.Count
        If (stsWeapons(tmpString).colWPlayers(i).lDeaths <> 0) Or _
           (stsWeapons(tmpString).colWPlayers(i).lFrags <> 0) Then
           With lstWUsage.ListItems.Add(, , stsWeapons(tmpString).colWPlayers(i).sName)
                .SubItems(1) = stsWeapons(tmpString).colWPlayers(i).lFrags
                .SubItems(2) = stsWeapons(tmpString).colWPlayers(i).lDeaths
                .SubItems(3) = stsWeapons(tmpString).colWPlayers(i).lFrags - stsWeapons(tmpString).colWPlayers(i).lDeaths
           End With
        End If
    Next
    
End Sub

Private Sub lstWUsage_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If (ColumnHeader.Position - 1) <> lstWUsage.SortKey Then
        lstWUsage.SortKey = ColumnHeader.Position - 1
        lstWUsage.SortOrder = lvwAscending
    Else
        If lstWUsage.SortOrder = lvwAscending Then
            lstWUsage.SortOrder = lvwDescending
        Else
            lstWUsage.SortOrder = lvwAscending
        End If
    End If

End Sub

Private Sub optRankCustom_Click()

    modFileHandler.InitializeStatsForm "Some"
    CalculateRank woRankCustom
    
End Sub

Public Sub optRankFrags_Click()

    modFileHandler.InitializeStatsForm "Some"
    CalculateRank woRankFrags
    
End Sub

Private Sub optRankNett_Click()

    modFileHandler.InitializeStatsForm "Some"
    CalculateRank woRankNett
    
End Sub

Private Sub optSkill_Click()
    
    modFileHandler.InitializeStatsForm "Some"
    CalculateRank woRankSkill
    
End Sub

Public Sub picTimeline_Click()

    HasClicked = True
    GenerateTimeLine
    
End Sub
