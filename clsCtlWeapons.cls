VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlWeapons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colWPlayers"
Option Explicit

Private msName As String
Private msDescription As String
Private mlScore As Long
Private msBestUser As String
Private msPickedOn As String
Private msImmune As String
Private mlFrags As Long
Private mlSuicides As Long
Private mcolWPlayers As colWPlayers

Public Property Get colWPlayers() As colWPlayers
    If mcolWPlayers Is Nothing Then
        Set mcolWPlayers = New colWPlayers
    End If
    Set colWPlayers = mcolWPlayers
End Property

Public Property Set colWPlayers(vData As colWPlayers)
    Set mcolWPlayers = vData
End Property

Private Sub Class_Terminate()
    Set mcolWPlayers = Nothing
End Sub

Public Property Let lSuicides(ByVal vData As Long)
    mlSuicides = vData
End Property

Public Property Get lSuicides() As Long
    lSuicides = mlSuicides
End Property

Public Property Let lFrags(ByVal vData As Long)
    mlFrags = vData
End Property

Public Property Get lFrags() As Long
    lFrags = mlFrags
End Property

Public Property Let sBestUser(ByVal vData As String)
    msBestUser = vData
End Property

Public Property Get sBestUser() As String
    sBestUser = msBestUser
End Property

Public Property Let sPickedOn(ByVal vData As String)
    msPickedOn = vData
End Property

Public Property Get sPickedOn() As String
    sPickedOn = msPickedOn
End Property

Public Property Let sImmune(ByVal vData As String)
    msImmune = vData
End Property

Public Property Get sImmune() As String
    sImmune = msImmune
End Property


Public Property Let lScore(ByVal vData As Long)
    mlScore = vData
End Property

Public Property Get lScore() As Long
    lScore = mlScore
End Property

Public Property Let sDescription(ByVal vData As String)
    msDescription = vData
End Property

Public Property Get sDescription() As String
    sDescription = msDescription
End Property

Public Property Let sName(ByVal vData As String)
    msName = vData
End Property

Public Property Get sName() As String
    sName = msName
End Property
