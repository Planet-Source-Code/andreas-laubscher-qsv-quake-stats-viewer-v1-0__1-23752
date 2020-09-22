VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheWeapons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsCtlWeapons"
Attribute VB_Ext_KEY = "Member0" ,"clsCtlWeapons"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private tmpCls As clsCtlWeapons

Public Function Add(sName As String, sDescription As String, lScore As Long, Optional sKey As String) As clsCtlWeapons
    
    'create a new object
    Dim objNewMember As clsCtlWeapons
    Set objNewMember = New clsCtlWeapons

    'set the properties passed into the method
    objNewMember.sName = sName
    objNewMember.sDescription = sDescription
    objNewMember.lScore = lScore
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsCtlWeapons
Attribute Item.VB_UserMemId = 0
    On Error GoTo GetError
    
    Set Item = mCol(vntIndexKey)
    
    Exit Property
    
GetError:

    Set tmpCls = New clsCtlWeapons
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
