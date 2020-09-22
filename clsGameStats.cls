VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGameStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mlFrags As Long
Private mlSuicides As Long
Private mlDuration As Long
Private msBestWeapon As String
Private msMostDangerousWeapon As String
Private msMapName As String
Private msGameType As String

Public Property Let sMostDangerousWeapon(ByVal vData As String)
    msMostDangerousWeapon = vData
End Property

Public Property Get sMostDangerousWeapon() As String
    sMostDangerousWeapon = msMostDangerousWeapon
End Property

Public Property Let sBestWeapon(ByVal vData As String)
    msBestWeapon = vData
End Property

Public Property Get sBestWeapon() As String
    sBestWeapon = msBestWeapon
End Property

Public Property Let sMapName(ByVal vData As String)
    msMapName = vData
End Property

Public Property Get sMapName() As String
    sMapName = msMapName
End Property

Public Property Let sGameType(ByVal vData As String)
    msGameType = vData
End Property

Public Property Get sGameType() As String
    sGameType = msGameType
End Property

Public Property Let Suicides(ByVal vData As Long)
    mlSuicides = vData
End Property

Public Property Get Suicides() As Long
    Suicides = mlSuicides
End Property

Public Property Let Frags(ByVal vData As Long)
    mlFrags = vData
End Property

Public Property Get Frags() As Long
    Frags = mlFrags
End Property

Public Property Let Duration(ByVal vData As Long)
    mlDuration = vData
End Property

Public Property Get Duration() As Long
    Duration = mlDuration
End Property
