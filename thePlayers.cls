VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "thePlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPlayer"
Attribute VB_Ext_KEY = "Member0" ,"clsPlayer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private tmpCls As clsPlayer

Public Function Add(sName As String, sClan As String, lRanking As Long, lScore As Long, lFrags As Long, lDeaths As Long, lSuicides As Long, lMkiol As Long, lMdwak As Long, sPickedOn As String, sPickedOnBy As String, sBestWeapon As String, Optional sKey As String) As clsPlayer
    
    'create a new object
    Dim objNewMember As clsPlayer
    Set objNewMember = New clsPlayer

    'set the properties passed into the method
    objNewMember.sName = sName
    objNewMember.sClan = sClan
    objNewMember.lRanking = lRanking
    objNewMember.lScore = lScore
    objNewMember.lFrags = lFrags
    objNewMember.lDeaths = lDeaths
    objNewMember.lSuicides = lSuicides
    objNewMember.sPickedOn = sPickedOn
    objNewMember.sPickedOnBy = sPickedOnBy
    objNewMember.sBestWeapon = sBestWeapon
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsPlayer
Attribute Item.VB_UserMemId = 0
    On Error GoTo GetError
    
    Set Item = mCol(vntIndexKey)

    Exit Property
    
GetError:
    
    Set tmpCls = New clsPlayer
    
    tmpCls.sName = vntIndexKey
    tmpCls.lFrags = 0
    tmpCls.lDeaths = 0
    tmpCls.lSuicides = 0
    tmpCls.lRanking = 0
    tmpCls.lScore = 0
    tmpCls.sPickedOn = ""
    tmpCls.sPickedOnBy = ""
    tmpCls.sBestWeapon = ""
    tmpCls.lSurvived = 0
    tmpCls.dspGraph = True
    tmpCls.lSkill = 1000
    
    Err.Clear
    
    ' This seems (rather obviously) the most sensible place to keep <world> out of our player collection.
    If (vntIndexKey <> cWorld) And (vntIndexKey <> "") Then
        mCol.Add tmpCls, vntIndexKey
        Set tmpCls = Nothing
        Resume
    Else
        Set Item = tmpCls
        Resume Next
    End If
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

