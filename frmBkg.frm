VERSION 5.00
Begin VB.Form frmBkg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10530
   Enabled         =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9300
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   540
      Picture         =   "frmBkg.frx":0000
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   540
      Width           =   12000
   End
End
Attribute VB_Name = "frmBkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================='
' Author            : Andreas Laubscher                                                         '
' Date Written      : 30 May 2001                                                               '
' Description       : The background screen is there purely to look good, and jack up download  '
'                   : size of the project... hehehe                                             '
' Credits           : Credit should go to www.looroll.com for the brilliant Quake image. Check  '
'                   : out the website, it's quite good if you're looking for wallpapers.        '
'==============================================================================================='
Option Explicit

Private Sub Form_Load()

    ' Size the form to fit the screen
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width * Screen.TwipsPerPixelX
    Me.Height = Screen.Height * Screen.TwipsPerPixelY
    
    ' Center the image
    picMain.Left = (Me.Width / 2) - (picMain.Width / 2)
    picMain.Top = (Me.Height / 2) - (picMain.Height / 2)
    
End Sub
