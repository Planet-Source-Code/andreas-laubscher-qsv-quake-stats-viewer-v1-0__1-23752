VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "theClans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsClan"
Attribute VB_Ext_KEY = "Member0" ,"clsClan"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private tmpCls As clsClan

Public Function Add(sName As String, lRanking As Long, lScore As Long, sBestPlayer As String, sWorstPlayer As String, lFrags As Long, lDeaths As Long, lSuicides As Long, lInterfrags As Long, lMkiol As Long, lMdwak As Long, Optional sKey As String) As clsClan
    
    'create a new object
    Dim objNewMember As clsClan
    Set objNewMember = New clsClan

    'set the properties passed into the method
    objNewMember.sName = sName
    objNewMember.lRanking = lRanking
    objNewMember.lScore = lScore
    objNewMember.sBestPlayer = sBestPlayer
    objNewMember.sWorstPlayer = sWorstPlayer
    objNewMember.lFrags = lFrags
    objNewMember.lDeaths = lDeaths
    objNewMember.lSuicides = lSuicides
    objNewMember.lInterfrags = lInterfrags
    objNewMember.lMkiol = lMkiol
    objNewMember.lMdwak = lMdwak
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsClan
Attribute Item.VB_UserMemId = 0
    On Error GoTo GetError
    
    Set Item = mCol(vntIndexKey)

    Exit Property
    
GetError:

    Set tmpCls = New clsClan
    
    tmpCls.sName = vntIndexKey
    
    mCol.Add tmpCls, vntIndexKey
    Set tmpCls = Nothing
    
    Err.Clear
    Resume
    
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
