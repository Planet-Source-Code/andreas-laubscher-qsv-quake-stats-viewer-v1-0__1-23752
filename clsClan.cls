VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colClanPlayers"
Option Explicit

Private msName              As String
Private mlRanking           As Long
Private mlScore             As Long
Private msBestPlayer        As String
Private msWorstPlayer       As String
Private mlPlayers           As Long
Private mlFrags             As Long
Private mlNett              As Long
Private mlDeaths            As Long
Private mlSuicides          As Long
Private mlInterfrags        As Long
Private mlMkiol             As Long
Private mlMdwak             As Long
Private mcolClanPlayers     As colClanPlayers

Public Property Get colClanPlayers() As colClanPlayers
    If mcolClanPlayers Is Nothing Then
        Set mcolClanPlayers = New colClanPlayers
    End If
    Set colClanPlayers = mcolClanPlayers
End Property

Public Property Set colClanPlayers(vData As colClanPlayers)
    Set mcolClanPlayers = vData
End Property
Private Sub Class_Terminate()
    Set mcolClanPlayers = Nothing
End Sub

Public Property Let lMdwak(ByVal vData As Long)
    mlMdwak = vData
End Property

Public Property Get lMdwak() As Long
    lMdwak = mlMdwak
End Property

Public Property Let lMkiol(ByVal vData As Long)
    mlMkiol = vData
End Property

Public Property Get lMkiol() As Long
    lMkiol = mlMkiol
End Property

Public Property Let lInterfrags(ByVal vData As Long)
    mlInterfrags = vData
End Property

Public Property Get lInterfrags() As Long
    lInterfrags = mlInterfrags
End Property

Public Property Let lNett(ByVal vData As Long)
    mlNett = vData
End Property

Public Property Get lNett() As Long
    lNett = mlNett
End Property

Public Property Let lSuicides(ByVal vData As Long)
    mlSuicides = vData
End Property

Public Property Get lSuicides() As Long
    lSuicides = mlSuicides
End Property

Public Property Let lDeaths(ByVal vData As Long)
    mlDeaths = vData
End Property

Public Property Get lDeaths() As Long
    lDeaths = mlDeaths
End Property

Public Property Let lFrags(ByVal vData As Long)
    mlFrags = vData
End Property

Public Property Get lFrags() As Long
    lFrags = mlFrags
End Property

Public Property Let sWorstPlayer(ByVal vData As String)
    msWorstPlayer = vData
End Property

Public Property Get sWorstPlayer() As String
    sWorstPlayer = msWorstPlayer
End Property

Public Property Let sBestPlayer(ByVal vData As String)
    msBestPlayer = vData
End Property

Public Property Get sBestPlayer() As String
    sBestPlayer = msBestPlayer
End Property

Public Property Let lScore(ByVal vData As Long)
    mlScore = vData
End Property

Public Property Get lScore() As Long
    lScore = mlScore
End Property

Public Property Let lRanking(ByVal vData As Long)
    mlRanking = vData
End Property

Public Property Get lRanking() As Long
    lRanking = mlRanking
End Property

Public Property Let sName(ByVal vData As String)
    msName = vData
End Property

Public Property Get sName() As String
    sName = msName
End Property

Public Property Let lPlayers(ByVal vData As String)
    mlPlayers = vData
End Property

Public Property Get lPlayers() As String
    lPlayers = mlPlayers
End Property
