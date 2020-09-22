VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colPPlayer"
Attribute VB_Ext_KEY = "Member1" ,"colPWeapon"
Option Explicit

Private msName              As String
Private msClan              As String
Private mlPing              As Long
Private mlRanking           As Long
Private mlActive            As Double
Private mlScore             As Long
Private mlFrags             As Long
Private mlNett              As Long
Private mlDeaths            As Long
Private mlSuicides          As Long
Private mlMkiol             As Long
Private mlMdwak             As Long
Private msPickedOn          As String
Private msPickedOnBy        As String
Private msBestWeapon        As String
Private mlSurvived          As Long
Private IsKilling           As Boolean
Private IsDying             As Boolean
Private tmpKillSpree        As Long
Private tmpDieSpree         As Long
Private mcolPPlayer         As colPPlayer
Private mcolPWeapon         As colPWeapon
Private mlTLine             As Long
Private mltlX               As Single
Private mltlY               As Single
Private mClrR               As Integer
Private mClrG               As Integer
Private mClrB               As Integer
Private mDspGraph           As Boolean
Private mlSkill             As Single
Private mlEffectivity       As Single

Public Property Get colPWeapon() As colPWeapon
    If mcolPWeapon Is Nothing Then
        Set mcolPWeapon = New colPWeapon
    End If
    Set colPWeapon = mcolPWeapon
End Property

Public Property Set colPWeapon(vData As colPWeapon)
    Set mcolPWeapon = vData
End Property

Public Property Get colPPlayer() As colPPlayer
    If mcolPPlayer Is Nothing Then
        Set mcolPPlayer = New colPPlayer
    End If
    Set colPPlayer = mcolPPlayer
End Property

Public Property Set colPPlayer(vData As colPPlayer)
    Set mcolPPlayer = vData
End Property

Private Sub Class_Terminate()
    Set mcolPWeapon = Nothing
    Set mcolPPlayer = Nothing
End Sub

Public Property Let sBestWeapon(ByVal vData As String)
    msBestWeapon = vData
End Property

Public Property Get sBestWeapon() As String
    sBestWeapon = msBestWeapon
End Property

Public Property Let sPickedOnBy(ByVal vData As String)
    msPickedOnBy = vData
End Property

Public Property Get sPickedOnBy() As String
    sPickedOnBy = msPickedOnBy
End Property

Public Property Let sPickedOn(ByVal vData As String)
    msPickedOn = vData
End Property

Public Property Get sPickedOn() As String
    sPickedOn = msPickedOn
End Property

Public Property Get lMdwak() As Long
    lMdwak = mlMdwak
End Property

Public Property Get lMkiol() As Long
    lMkiol = mlMkiol
End Property

Public Property Get lPing() As Long
    lPing = mlPing
End Property

Public Property Let lPing(ByVal vData As Long)
    mlPing = vData
End Property

Public Property Get lSkill() As Single
    lSkill = mlSkill
End Property

Public Property Let lSkill(ByVal vData As Single)
    mlSkill = vData
End Property

Public Property Get lEffectivity() As Single
    lEffectivity = mlEffectivity
End Property

Public Property Let lEffectivity(ByVal vData As Single)
    mlEffectivity = vData
End Property

Public Property Get dspGraph() As Boolean
    dspGraph = mDspGraph
End Property

Public Property Let dspGraph(ByVal vData As Boolean)
    mDspGraph = vData
End Property

Public Property Get lTLine() As Long
    lTLine = mlTLine
End Property

Public Property Let lTLine(ByVal vData As Long)
    mlTLine = vData
End Property

Public Property Get ltlX() As Single
    ltlX = mltlX
End Property

Public Property Let ltlX(ByVal vData As Single)
    mltlX = vData
End Property

Public Property Get ltlY() As Single
    ltlY = mltlY
End Property

Public Property Let ltlY(ByVal vData As Single)
    mltlY = vData
End Property

Public Property Get clrR() As Integer
    clrR = mClrR
End Property

Public Property Let clrR(ByVal vData As Integer)
    mClrR = vData
End Property

Public Property Get clrG() As Integer
    clrG = mClrG
End Property

Public Property Let clrG(ByVal vData As Integer)
    mClrG = vData
End Property

Public Property Get clrB() As Integer
    clrB = mClrB
End Property

Public Property Let clrB(ByVal vData As Integer)
    mClrB = vData
End Property

Public Property Let lSuicides(ByVal vData As Long)
    
    mlSuicides = vData
    
    If tmpKillSpree > mlMkiol Then mlMkiol = tmpKillSpree
    IsKilling = False
    
    If IsDying Then
        tmpDieSpree = tmpDieSpree + 1
    Else
        IsDying = True
        tmpDieSpree = 1
    End If
    
End Property

Public Property Get lSuicides() As Long
    lSuicides = mlSuicides
End Property

Public Property Let lDeaths(ByVal vData As Long)
    
    mlDeaths = vData
    
    If tmpKillSpree > mlMkiol Then mlMkiol = tmpKillSpree
    IsKilling = False
    
    If IsDying Then
        tmpDieSpree = tmpDieSpree + 1
    Else
        IsDying = True
        tmpDieSpree = 1
    End If
    
End Property

Public Property Get lDeaths() As Long
    lDeaths = mlDeaths
End Property

Public Property Let lNett(ByVal vData As Long)
    mlNett = vData
End Property

Public Property Get lNett() As Long
    lNett = mlNett
End Property

Public Property Let lFrags(ByVal vData As Long)
    mlFrags = vData
    
    If tmpDieSpree > mlMdwak Then mlMdwak = tmpDieSpree
    IsDying = False
        
    If IsKilling Then
        tmpKillSpree = tmpKillSpree + 1
    Else
        IsKilling = True
        tmpKillSpree = 1
    End If
    
End Property

Public Property Get lFrags() As Long
    lFrags = mlFrags
End Property

Public Property Let lScore(ByVal vData As Long)
    mlScore = vData
End Property

Public Property Get lScore() As Long
    lScore = mlScore
End Property

Public Property Let lActive(ByVal vData As Double)
    mlActive = vData
End Property

Public Property Get lActive() As Double
    lActive = mlActive
End Property

Public Property Let lRanking(ByVal vData As Long)
    mlRanking = vData
End Property

Public Property Get lRanking() As Long
    lRanking = mlRanking
End Property

Public Property Let sClan(ByVal vData As String)
    msClan = vData
End Property

Public Property Get sClan() As String
    sClan = msClan
End Property

Public Property Let sName(ByVal vData As String)
    msName = vData
    msClan = GetClan(msName)
End Property

Public Property Get sName() As String
    sName = msName
End Property

Public Property Let lSurvived(ByVal vData As Long)
    mlSurvived = vData
End Property

Public Property Get lSurvived() As Long
    lSurvived = mlSurvived
End Property

Private Function GetClan(tmpString As String) As String
Dim t1 As Integer
Dim t2 As Integer

    t1 = InStr(1, tmpString, "[")
    
    If t1 > 0 Then
        t2 = InStr(t1, tmpString, "]")
    End If
    
    If (t1 > 0) And (t2 > 0) Then
        ' We don't want to add the square brackets
        t1 = t1 + 1
        GetClan = Mid(tmpString, t1, t2 - t1)
    Else
        GetClan = ""
    End If
    
End Function
