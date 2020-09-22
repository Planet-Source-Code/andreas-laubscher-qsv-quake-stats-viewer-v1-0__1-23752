Attribute VB_Name = "modFileHandler"
'==============================================================================================='
' Author            : Andreas Laubscher                                                         '
' Date Written      : 30 May 2001                                                               '
' Description       : This module does basically all the statistical processing required by the '
'                   : program. It is a rather large module, but it is broken down fairly well   '
'                   : into logical pieces. And it's well-commented. I have to give some credit  '
'                   : to Nick Philbin, who made the excellent clan logo for us.                 '
'                   : By the way, I have uploaded this software to PSC, so obviously you can    '
'                   : use and abuse it any way you see fit, just give credit where credit is    '
'                   : due.                                                                      '
'==============================================================================================='
Option Explicit
'==============================================================================================='
' All the declarations                                                                          '
'==============================================================================================='
' Enumerate the types of scoring systems available
Public Enum woRanking
    woRankFrags = 1
    woRankNett = 2
    woRankCustom = 3
    woRankSkill = 4
End Enum

' These constants define the strings checked against while parsing the file
Public Const cInitGame = "InitGame"
Public Const cShutDown = "ShutdownGame"
Public Const cKill = " Kill: "
Public Const cPing = " ping: "
Public Const cKilled = " killed "
Public Const cBy = " by "
Public Const cWorld = "<world>"
Public Const cClient = " client: "
Public Const cMapName = "\mapname\"
Public Const cGameType = "\g_gametype\"

Public theFile              As String
Public i                    As Long
Public i2                   As Long
Public i3                   As Long
Public Counter              As Integer
Dim Filen                   As Integer
Public tmpString            As String

Public Tm1                  As Long
Public Tm2                  As Long

Public stsGame              As clsGameStats
Public stsPlayers           As thePlayers
Public stsClan              As theClans
Public stsWeapons           As TheWeapons

Sub GetWeapons()
'==============================================================================================='
' Retrieves all available weapons information from our config file                              '
' The layout is as follows, delimited by semi-colons(;) :                                       '
' 1) Weapon name - as found in the log file (usually MOD_ something                             '
' 2) Weapon description - the display name that the user will see                               '
' 3) Score - The score value that will be used for custom scoring                               '
' I might be bothered to add a cfg editor to the program if I feel like it                      '
'==============================================================================================='
Dim wData()                 As String

    Set stsWeapons = New TheWeapons
    
    Filen = FreeFile
    Open App.Path & "\Weapons.cfg" For Input As Filen
    
    Do Until EOF(Filen)
    
        Line Input #Filen, tmpString
        wData = Split(tmpString, ";")
        
        stsWeapons(wData(0)).sDescription = wData(1)
        stsWeapons(wData(0)).lScore = CLng(wData(2))
        
    Loop
    
    Close Filen
    
End Sub

Sub GetGames(sFile As String)
'==============================================================================================='
' Gets all the available games from the selected log file. Also organizes the variable we'll    '
' be using later to rip out a specific game's data                                              '
'==============================================================================================='
    Filen = FreeFile
    Open sFile For Binary As Filen
    
    theFile = InputB(FileLen(sFile), Filen)
    theFile = StrConv(theFile, vbUnicode)
    
    i = 1
    Counter = 0
    Do Until i = 0
        i = InStr(i + 1, theFile, cInitGame)
        Counter = Counter + 1
    Loop
    
    frmInput.lstGames.ListItems.Clear
    For i = 1 To Counter
        frmInput.lstGames.ListItems.Add , , "Game No. " & i
    Next
    
    Close Filen
    
End Sub

Sub PopulateBillboards()
'==============================================================================================='
' Get the billboards information. These are the two lists on the bottom right hand side of the  '
' leaderboard tab                                                                               '
'==============================================================================================='
Dim tFrags                  As Integer
Dim tDeaths                 As Integer
Dim tSuicides               As Integer
Dim tNett                   As Integer
Dim tMkiol                  As Integer
Dim tMDwak                  As Integer
Dim tSurvivor               As Long
Dim tSkill                  As Double
Dim tEffective              As Single

Dim sFrags                  As String
Dim sDeaths                 As String
Dim sSuicides               As String
Dim sNett                   As String
Dim sMkiol                  As String
Dim sMDwak                  As String
Dim sSurvivor               As String
Dim sSkill                  As String
Dim sEffective              As String

    ' Nett score could be negative
    tNett = -32000
    
    For i = 1 To stsPlayers.Count
        
        If stsPlayers(i).lFrags > tFrags Then
            tFrags = stsPlayers(i).lFrags
            sFrags = stsPlayers(i).sName
        End If
        If stsPlayers(i).lDeaths > tDeaths Then
            tDeaths = stsPlayers(i).lDeaths
            sDeaths = stsPlayers(i).sName
        End If
        If stsPlayers(i).lSuicides > tSuicides Then
            tSuicides = stsPlayers(i).lSuicides
            sSuicides = stsPlayers(i).sName
        End If
        If stsPlayers(i).lNett > tNett Then
            tNett = stsPlayers(i).lNett
            sNett = stsPlayers(i).sName
        End If
        If stsPlayers(i).lMkiol > tMkiol Then
            tMkiol = stsPlayers(i).lMkiol
            sMkiol = stsPlayers(i).sName
        End If
        If stsPlayers(i).lMdwak > tMDwak Then
            tMDwak = stsPlayers(i).lMdwak
            sMDwak = stsPlayers(i).sName
        End If
        If stsPlayers(i).lSurvived > tSurvivor Then
            tSurvivor = stsPlayers(i).lSurvived
            sSurvivor = stsPlayers(i).sName
        End If
        If stsPlayers(i).lSkill > tSkill Then
            tSkill = stsPlayers(i).lSkill
            sSkill = stsPlayers(i).sName
        End If
        If stsPlayers(i).lEffectivity > tEffective Then
            tEffective = stsPlayers(i).lEffectivity
            sEffective = stsPlayers(i).sName
        End If
    Next
    
    ' Load them to screen. We load the tooltips too, in case the entry is too long to display
    frmMain.lblPBFrags.Caption = IIf(Len(sFrags) > 0, sFrags & "  (" & tFrags & " Frags)", "")
    frmMain.lblPBFrags.ToolTipText = frmMain.lblPBFrags.Caption
    frmMain.lblPBDeaths.Caption = IIf(Len(sDeaths) > 0, sDeaths & "  (" & tDeaths & " Deaths)", "")
    frmMain.lblPBDeaths.ToolTipText = frmMain.lblPBDeaths.Caption
    frmMain.lblPBSuicides.Caption = IIf(Len(sSuicides) > 0, sSuicides & "  (" & tSuicides & " Suicides)", "")
    frmMain.lblPBSuicides.ToolTipText = frmMain.lblPBSuicides.Caption
    frmMain.lblPBNett.Caption = IIf(Len(sNett) > 0, sNett & "  (" & tNett & " Nett)", "")
    frmMain.lblPBNett.ToolTipText = frmMain.lblPBNett.Caption
    frmMain.lblPBMkiol.Caption = IIf(Len(sMkiol) > 0, sMkiol & "  (" & tMkiol & " Frags)", "")
    frmMain.lblPBMkiol.ToolTipText = frmMain.lblPBMkiol.Caption
    frmMain.lblPBMdwak.Caption = IIf(Len(sMDwak) > 0, sMDwak & "  (" & tMDwak & " Deaths)", "")
    frmMain.lblPBMdwak.ToolTipText = frmMain.lblPBMdwak.Caption
    frmMain.lblPBSurvivor.Caption = IIf(Len(sSurvivor) > 0, sSurvivor & "  (" & _
            IIf(tSurvivor > 60, (tSurvivor \ 60) & ":" & Format(tSurvivor Mod 60, "00") & " minutes)", tSurvivor & " Seconds)"), "")
    frmMain.lblPBSurvivor.ToolTipText = frmMain.lblPBSurvivor.Caption
    frmMain.lblPBSkill.Caption = IIf(Len(sSkill) > 0, sSkill & "  (" & Format(tSkill, "0.0##") & " Points)", "")
    frmMain.lblPBEffective.Caption = IIf(Len(sEffective) > 0, sEffective & "  (" & Format(tEffective, "0.0#") & ")", "")
    
    tFrags = 0
    tDeaths = 0
    tSuicides = 0
    tNett = -10000
    
    sFrags = ""
    sDeaths = ""
    sSuicides = ""
    sNett = ""
    
    For i = 1 To stsClan.Count
    
        If stsClan(i).lFrags > tFrags Then
            tFrags = stsClan(i).lFrags
            sFrags = stsClan(i).sName
        End If
        
        If stsClan(i).lDeaths > tDeaths Then
            tDeaths = stsClan(i).lDeaths
            sDeaths = stsClan(i).sName
        End If
        
        If stsClan(i).lSuicides > tSuicides Then
            tSuicides = stsClan(i).lSuicides
            sSuicides = stsClan(i).sName
        End If
        
        If stsClan(i).lNett > tNett Then
            tNett = stsClan(i).lNett
            sNett = stsClan(i).sName
        End If
        
    Next
    
    frmMain.lblCBFrags.Caption = sFrags
    frmMain.lblCBDeaths.Caption = sDeaths
    frmMain.lblCBSuicides.Caption = sSuicides
    frmMain.lblCBNett.Caption = sNett
    
End Sub

Function ReEdit(tmpString As String) As String
'==============================================================================================='
' This exists to edit umm... editing information from the Players' names. This includes colours '
' and bolds and stuff like that. This information exists as a "^" followed by a number, for     '
' colours; and a "B" for bold, etc...                                                           '
'==============================================================================================='
Dim x                       As Integer

    x = InStr(1, tmpString, "^")
    Do Until x = 0
        If x > 1 Then
            tmpString = Mid(tmpString, 1, x - 1) & Mid(tmpString, x + 2)
        Else
            tmpString = Mid(tmpString, x + 2)
        End If
        
        x = InStr(x + 1, tmpString, "^")
    Loop
    
    ReEdit = tmpString
    
End Function

Sub InitializeStatsForm(sScope As String)
'==============================================================================================='
' Initialize all the controls on the stats form that carry variables. This is probably not the  '
' most efficient (code-wise) way of doing it, but it's the only way I could think of doing it   '
'==============================================================================================='
                                                                                                
    ' Leaderboard tab
    If sScope = "All" Then
        frmMain.lblWinnerName.Caption = ""
        frmMain.lblWinnerClan.Caption = ""
        frmMain.lblWinnerScore.Caption = ""
        
        frmMain.lstLeaderboard.ListItems.Clear
        frmMain.lstLeaderboard.ColumnHeaders(1).Width = frmMain.lstLeaderboard.Width / 5
        frmMain.lstLeaderboard.ColumnHeaders(2).Width = frmMain.lstLeaderboard.Width / 2
        frmMain.lstLeaderboard.ColumnHeaders(3).Width = frmMain.lstLeaderboard.Width / 5
        
        frmMain.lblPBFrags.Caption = ""
        frmMain.lblPBDeaths.Caption = ""
        frmMain.lblPBSuicides.Caption = ""
        frmMain.lblPBEffective.Caption = ""
        frmMain.lblPBSkill.Caption = ""
        frmMain.lblPBNett.Caption = ""
        frmMain.lblPBMkiol.Caption = ""
        frmMain.lblPBMdwak.Caption = ""
        frmMain.lblPBSurvivor.Caption = ""
        frmMain.lblCBFrags.Caption = ""
        frmMain.lblCBDeaths.Caption = ""
        frmMain.lblCBSuicides.Caption = ""
        frmMain.lblCBNett.Caption = ""
    End If
    
    ' Game tab
    If sScope = "All" Then
        frmMain.lstRawLog.ListItems.Clear
        frmMain.lstRawLog.ColumnHeaders(1).Width = frmMain.lstRawLog.Width / 5
        frmMain.lstRawLog.ColumnHeaders(2).Width = frmMain.lstRawLog.Width / 4
        frmMain.lstRawLog.ColumnHeaders(3).Width = frmMain.lstRawLog.Width / 4
        frmMain.lstRawLog.ColumnHeaders(4).Width = frmMain.lstRawLog.Width / 4
        
        frmMain.lblGameFile.Caption = ""
        frmMain.lblGameNumber.Caption = ""
        frmMain.lblGameType.Caption = ""
        frmMain.lblGameMap.Caption = ""
        frmMain.lblGameDuration.Caption = "00:00"
        frmMain.lblGamePlayers.Caption = 0
        frmMain.lblGameFrags.Caption = 0
        frmMain.lblGameSuicides.Caption = 0
        
    End If
    
    ' Clans tab
    If sScope = "All" Then
        frmMain.lstClans.ListItems.Clear
        frmMain.lstClans.ColumnHeaders(1).Width = frmMain.lstClans.Width * 9 / 10
    End If
    frmMain.lstCPlayers.ListItems.Clear
    frmMain.lstCPlayers.ColumnHeaders(1).Width = frmMain.lstCPlayers.Width * 2 / 4
    frmMain.lstCPlayers.ColumnHeaders(2).Width = frmMain.lstCPlayers.Width * 2 / 5
    
    frmMain.lblCName.Caption = ""
    frmMain.lblCMembers.Caption = ""
    frmMain.lblCScore.Caption = ""
    frmMain.lblCBestPlayer.Caption = ""
    frmMain.lblCWorstPlayer.Caption = ""
    frmMain.lblCFrags.Caption = ""
    frmMain.lblCPPFrags.Caption = ""
    frmMain.lblCDeaths.Caption = ""
    frmMain.lblCPPDeaths.Caption = ""
    frmMain.lblCSuicides.Caption = ""
    frmMain.lblCPPSuicides.Caption = ""
    frmMain.lblCInterfrags.Caption = ""
    frmMain.lblCPPInterfrags.Caption = ""
    frmMain.lblCMkiol.Caption = ""
    frmMain.lblCMdwak.Caption = ""
    
    'Players Tab
    If sScope = "All" Then
        frmMain.lstPlayers.ListItems.Clear
        frmMain.lstPlayers.ColumnHeaders(1).Width = frmMain.lstPlayers.Width * 9 / 10
    End If
    frmMain.lstPKilled.ListItems.Clear
    frmMain.lstPKilled.ColumnHeaders(1).Width = frmMain.lstPKilled.Width * 2 / 4
    frmMain.lstPKilled.ColumnHeaders(2).Width = frmMain.lstPKilled.Width * 2 / 5
    
    frmMain.lstPWeapons.ListItems.Clear
    frmMain.lstPWeapons.ColumnHeaders(1).Width = frmMain.lstPWeapons.Width / 2.5
    frmMain.lstPWeapons.ColumnHeaders(2).Width = frmMain.lstPWeapons.Width / 6
    frmMain.lstPWeapons.ColumnHeaders(3).Width = frmMain.lstPWeapons.Width / 6
    frmMain.lstPWeapons.ColumnHeaders(4).Width = frmMain.lstPWeapons.Width / 6
    
    frmMain.lstPKilled.ListItems.Clear
    frmMain.lstPKilled.ColumnHeaders(1).Width = frmMain.lstPKilled.Width / 2.5
    frmMain.lstPKilled.ColumnHeaders(2).Width = frmMain.lstPKilled.Width / 6
    frmMain.lstPKilled.ColumnHeaders(3).Width = frmMain.lstPKilled.Width / 6
    frmMain.lstPKilled.ColumnHeaders(4).Width = frmMain.lstPKilled.Width / 6
    
    frmMain.lblPName.Caption = ""
    frmMain.lblPClan.Caption = ""
    frmMain.lblPRanking.Caption = ""
    frmMain.lblPEffectivity.Caption = ""
    frmMain.lblPSkill.Caption = ""
    frmMain.lblPPing.Caption = ""
    frmMain.lblPScore.Caption = ""
    frmMain.lblPHitrate.Caption = ""
    frmMain.lblPSurvived.Caption = ""
    frmMain.lblPNett.Caption = ""
    frmMain.lblPFrags.Caption = ""
    frmMain.lblPDeaths.Caption = ""
    frmMain.lblPSuicides.Caption = ""
    frmMain.lblPMkiol.Caption = ""
    frmMain.lblPMdwak.Caption = ""
    frmMain.lblPPickedOn.Caption = ""
    frmMain.lblPPickedOnBy.Caption = ""
    frmMain.lblPBestWeapon.Caption = ""

    ' Weapons tab
    If sScope = "All" Then
        frmMain.lstWeapons.ListItems.Clear
        frmMain.lstWeapons.ColumnHeaders(1).Width = frmMain.lstWeapons.Width * 9 / 10
        frmMain.lblWBestWeapon.Caption = ""
        frmMain.lblWDangerous.Caption = ""
    End If
    
    frmMain.lstWUsage.ListItems.Clear
    frmMain.lstWUsage.ColumnHeaders(1).Width = frmMain.lstWUsage.Width / 2.5
    frmMain.lstWUsage.ColumnHeaders(2).Width = frmMain.lstWUsage.Width / 6
    frmMain.lstWUsage.ColumnHeaders(3).Width = frmMain.lstWUsage.Width / 6
    frmMain.lstWUsage.ColumnHeaders(4).Width = frmMain.lstWUsage.Width / 6
    
    frmMain.lblWName.Caption = ""
    frmMain.lblWLogName.Caption = ""
    frmMain.lblWScore.Caption = ""
    frmMain.lblWFrags.Caption = ""
    frmMain.lblWSuicides.Caption = ""
    frmMain.lblWBestUser.Caption = ""
    frmMain.lblWPickedOn.Caption = ""
    frmMain.lblWImmune.Caption = ""
    
    ' Timeline tab
    frmMain.picTimeline.Cls
    For i = 1 To frmMain.flxLegend.Rows - 1
        frmMain.flxLegend.RemoveItem 0
    Next
    frmMain.flxLegend.TextArray(0) = ""
    
    ' Give everything time to refresh
    DoEvents
    
End Sub
