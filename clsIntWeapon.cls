VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msName As String
Private mlFrags As Long
Private mlDeaths As Long
Private mlSuicides As Long

Public Property Let lFrags(ByVal vData As Long)
    mlFrags = vData
End Property

Public Property Get lFrags() As Long
    lFrags = mlFrags
End Property

Public Property Let sName(ByVal vData As String)
    msName = vData
End Property

Public Property Get sName() As String
    sName = msName
End Property

Public Property Let lDeaths(ByVal vData As Long)
    mlDeaths = vData
End Property

Public Property Get lDeaths() As Long
    lDeaths = mlDeaths
End Property

Public Property Let lSuicides(ByVal vData As Long)
    mlSuicides = vData
End Property

Public Property Get lSuicides() As Long
    lSuicides = mlSuicides
End Property

