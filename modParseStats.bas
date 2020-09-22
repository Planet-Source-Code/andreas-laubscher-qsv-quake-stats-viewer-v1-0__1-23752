Attribute VB_Name = "modParseStats"
Option Explicit

Dim tmpTime                 As String
Dim tmpKiller               As String
Dim tmpKillee               As String
Dim tmpWeapon               As String
Dim tmpPing                 As Long
Dim tmpStart                As Long
Dim tmpEnd                  As Long

Sub GetStats(sGame As String)

'==============================================================================================='
' The big one. Gets all the statistics (and I do mean ALL) for the selected game.               '
' You'll notice that there is no error handling here, even though we're referencing collections '
' Check out any one of the collections' Get event to see how I handle that.                     '
'==============================================================================================='
Dim tmpGameNo               As Integer
Dim theGame                 As String
Dim plc1                    As Integer
Dim plc2                    As Integer
Dim v1 As String

    ' Initialize our classes
    Set stsGame = New clsGameStats
    Set stsPlayers = New thePlayers
    Set stsClan = New theClans
    
    GetWeapons
    
    ' Clear the values currently on-screen
    InitializeStatsForm "All"
    
    ' Loop the following until we've done all the games
    '--------------------------------------------------
    plc1 = 1
    plc2 = InStr(plc1 + 1, sGame, ";")
    
    Do Until plc2 = 0
    
        
        tmpGameNo = Trim(Mid(Mid(sGame, plc1, plc2 - plc1), InStrRev(Mid(sGame, plc1, plc2 - plc1), ".") + 1))
        
        i = 0
        Counter = 1
    
        ' Step until we find the correct game number
        If tmpGameNo = 1 Then
            i = InStr(1, theFile, cInitGame)
        Else
            Do Until Counter = tmpGameNo
                i = InStr(i + 1, theFile, cInitGame)
                Counter = Counter + 1
            Loop
        End If
    
        ' Get the string containing all the data of the selected game
        theGame = Mid(theFile, InStrRev(theFile, Chr(10), i - 2) + 1, InStr(i, theFile, cShutDown) - i)
    
        ' Here we call the GetSingleGame sub
        GetSingleGame theGame
        
        ' Edit the game's duration
        stsGame.Duration = stsGame.Duration + (tmpEnd - tmpStart)
        
        ' OK, now we get all the stats that depend on total game information
        '-------------------------------------------------------------------
        FinaliseStats
    
        ' Find the next game's place holders
        If InStr(plc1 + 1, sGame, ";") <> 0 Then
            plc1 = plc2 + 1
            plc2 = InStr(plc1 + 1, sGame, ";")
        End If
        
    Loop
    
    If stsGame.Duration > 60 Then
        frmMain.lblGameDuration.Caption = (stsGame.Duration \ 60) & ":" & Format((stsGame.Duration Mod 60), "00") & " minutes"
    Else
        frmMain.lblGameDuration.Caption = stsGame.Duration & " seconds"
    End If
        
    ' Ranking time
    '-------------
    If frmMain.optRankFrags Then
        CalculateRank woRankFrags
    ElseIf frmMain.optRankNett Then
        CalculateRank woRankNett
    Else
        CalculateRank woRankCustom
    End If
    
    ' Do the billboards (on the leaderboard tab)
    '-------------------------------------------
    PopulateBillboards
    
    ' Now we add all the once-off entries to the main form
    '---------------------------------------------------------
    ' Gamefile
    If Len(frmInput.txtInputFile.Text) > 60 Then
        frmMain.lblGameFile.Caption = "..." & Mid(frmInput.txtInputFile.Text, InStr(Len(frmInput.txtInputFile.Text) - 35, frmInput.txtInputFile.Text, "\"))
    Else
        frmMain.lblGameFile.Caption = frmInput.txtInputFile.Text
    End If
    
    ' Game number. Edited to remove the required semicolon
    frmMain.lblGameNumber.Caption = Mid(sGame, 1, Len(sGame) - 1)
    
    ' Game Totals
    frmMain.lblGameFrags = stsGame.Frags
    frmMain.lblGameSuicides = stsGame.Suicides
    frmMain.lblGamePlayers = stsPlayers.Count
    frmMain.lblGameMap = stsGame.sMapName
    frmMain.lblGameType = stsGame.sGameType
    
    tmpString = stsGame.sBestWeapon
    frmMain.lblWBestWeapon.Caption = IIf(Len(stsWeapons(tmpString).sDescription) > 0, _
                                        stsWeapons(tmpString).sDescription & " (" & stsWeapons(tmpString).lFrags & " Frags)", "")
    
    tmpString = stsGame.sMostDangerousWeapon
    frmMain.lblWDangerous.Caption = IIf(Len(stsGame.sMostDangerousWeapon) > 0, stsWeapons(tmpString).sDescription & " (" & stsWeapons(tmpString).lSuicides & " Suicides)", "")
    
    ' And finally (phew) we add all the relevant lists to our stats form.
    For i = 1 To stsPlayers.Count
        frmMain.lstPlayers.ListItems.Add , , stsPlayers(i).sName
    Next
    For i = 1 To stsClan.Count
        frmMain.lstClans.ListItems.Add , , stsClan(i).sName
    Next
    For i = 1 To stsWeapons.Count
        If (stsWeapons(i).lFrags > 0) Or (stsWeapons(i).lSuicides > 0) Then
            frmMain.lstWeapons.ListItems.Add , stsWeapons(i).sName, stsWeapons(i).sDescription
        End If
    Next
    
    ' And load the timeline
    frmMain.picTimeline_Click
    
End Sub

Sub GetSingleGame(theGame As String)
    On Error GoTo StatsError
    
Dim theEntry                As String
Dim lRet                    As Long

    ' Rip out the time. We convert it to seconds to prevent some weird hassles further down
    tmpTime = Trim(Mid(theGame, 1, InStr(InStr(1, theGame, ":"), theGame, " ")))
    tmpStart = (60 * Mid(tmpTime, 1, InStr(1, tmpTime, ":") - 1)) + Mid(tmpTime, InStr(1, tmpTime, ":") + 1)
    
    ' This is done in case we encounter an empty game
    tmpEnd = tmpStart
    
    i = 1
    i2 = InStr(1, theGame, Chr(10))
    
    ' We step through our game until we're all out of line breaks (chr 10)
    Do Until i2 = 0
    
        ' Active line (or record, if you like)
        theEntry = Mid(theGame, i + 1, i2 - i - 1)
        i = i2
        
        ' We're checking for three types of file entries: Initgame, Kills and pings
        If InStr(1, theEntry, cInitGame) > 0 Then
            ' Initgame
            i2 = InStr(1, theEntry, cMapName) + 9
            stsGame.sMapName = Mid(theEntry, i2, InStr(i2, theEntry, "\") - i2)
            
            i2 = InStr(1, theEntry, cGameType) + 12
            stsGame.sGameType = Mid(theEntry, i2, InStr(i2, theEntry, "\") - i2)
            
        ElseIf InStr(1, theEntry, cKill) > 0 Then
            ' OK, so we've found a kill. Now let's do something with it.
        
            ' First, the time
            i2 = InStr(3, theEntry, " ")
            tmpTime = Trim(Mid(theEntry, 1, i2))
            
            ' Next, the killer
            i2 = InStr(InStr(1, theEntry, cKill) + 7, theEntry, ":") + 1
            i3 = InStr(i2, theEntry, cKilled)
            tmpKiller = ReEdit(Trim(Mid(theEntry, i2, i3 - i2)))
            
            ' Then, the killee
            i2 = i3 + 8
            i3 = InStr(i2, theEntry, cBy)
            tmpKillee = ReEdit(Mid(theEntry, i2, i3 - i2))
            
            ' Finally, the weapon
            i2 = i3 + 4
            tmpWeapon = Trim(Mid(theEntry, i2))
            
            ' Add the event to our raw log
            With frmMain.lstRawLog.ListItems.Add(, , tmpTime)
                .SubItems(1) = tmpKiller
                .SubItems(2) = tmpKillee
                .SubItems(3) = stsWeapons(tmpWeapon).sDescription
            End With
            
            ' Now that that's out of the way, we jack up our stats
            If (tmpKiller = cWorld) Or (tmpKiller = tmpKillee) Then
                ' Whoops, a suicide
                stsPlayers(tmpKillee).lSuicides = stsPlayers(tmpKillee).lSuicides + 1
                stsPlayers(tmpKillee).colPWeapon(tmpWeapon).lSuicides = stsPlayers(tmpKillee).colPWeapon(tmpWeapon).lSuicides + 1
                stsPlayers(tmpKillee).lSkill = stsPlayers(tmpKillee).lSkill - (stsPlayers(tmpKillee).lSkill / 1000)
                
                stsWeapons(tmpWeapon).lSuicides = stsWeapons(tmpWeapon).lSuicides + 1
                stsWeapons(tmpWeapon).colWPlayers(tmpKillee).lDeaths = stsWeapons(tmpWeapon).colWPlayers(tmpKillee).lDeaths + 1
                
                stsGame.Suicides = stsGame.Suicides + 1
            
            Else
                
                stsPlayers(tmpKiller).lFrags = stsPlayers(tmpKiller).lFrags + 1
                stsPlayers(tmpKiller).lScore = stsPlayers(tmpKiller).lScore + stsWeapons(tmpWeapon).lScore
                stsPlayers(tmpKiller).colPPlayer(tmpKillee).lFrags = stsPlayers(tmpKiller).colPPlayer(tmpKillee).lFrags + 1
                stsPlayers(tmpKiller).colPWeapon(tmpWeapon).lFrags = stsPlayers(tmpKiller).colPWeapon(tmpWeapon).lFrags + 1
                stsPlayers(tmpKiller).lSkill = stsPlayers(tmpKiller).lSkill + ((stsPlayers(tmpKillee).lSkill / stsPlayers(tmpKiller).lSkill) * 10)
                
                stsPlayers(tmpKillee).colPWeapon(tmpWeapon).lDeaths = stsPlayers(tmpKillee).colPWeapon(tmpWeapon).lDeaths + 1
                stsPlayers(tmpKillee).colPPlayer(tmpKiller).lDeaths = stsPlayers(tmpKillee).colPPlayer(tmpKiller).lDeaths + 1
                stsPlayers(tmpKillee).lDeaths = stsPlayers(tmpKillee).lDeaths + 1
                stsPlayers(tmpKillee).lSkill = stsPlayers(tmpKillee).lSkill + ((stsPlayers(tmpKiller).lSkill / stsPlayers(tmpKillee).lSkill) * 10)
                
                stsWeapons(tmpWeapon).lFrags = stsWeapons(tmpWeapon).lFrags + 1
                stsWeapons(tmpWeapon).colWPlayers(tmpKiller).lFrags = stsWeapons(tmpWeapon).colWPlayers(tmpKiller).lFrags + 1
                stsWeapons(tmpWeapon).colWPlayers(tmpKillee).lDeaths = stsWeapons(tmpWeapon).colWPlayers(tmpKillee).lDeaths + 1
                
                stsGame.Frags = stsGame.Frags + 1
                
            End If
            
            If tmpKiller <> cWorld Then
            
                ' An Interfrag is when a clan member kills one of his own clan
                If (stsPlayers(tmpKiller).sClan = stsPlayers(tmpKillee).sClan) And _
                   (stsPlayers(tmpKiller).sClan <> "") Then
                    stsClan(stsPlayers(tmpKiller).sClan).lInterfrags = stsClan(stsPlayers(tmpKiller).sClan).lInterfrags + 1
                End If
                
            End If
                        
        ElseIf InStr(1, theEntry, cPing) > 0 Then
            ' Aaaaah, a peeeng! 9+clent:
            
            tmpKiller = ReEdit(Trim(Mid(theEntry, InStr(InStr(1, theEntry, "client:") + 9, theEntry, " "))))
            
            i2 = InStr(1, theEntry, cPing) + 7
            i3 = InStr(i2 + 1, theEntry, " ")
            tmpPing = Mid(theEntry, i2, i3 - i2)
            
            If tmpKiller <> cWorld Then
                stsPlayers(tmpKiller).lPing = tmpPing
            End If
            
        End If
        
        i2 = InStr(i + 1, theGame, Chr(10))
        
    Loop
    
    ' A rather convoluted way of getting (and converting) the time, I know...
    tmpEnd = (60 * Mid(tmpTime, 1, InStr(1, tmpTime, ":") - 1)) + Mid(tmpTime, InStr(1, tmpTime, ":") + 1)

    Exit Sub
    
StatsError:
    lRet = MsgBox("An error occurred while attempting to parse the following log file entry: " & _
           vbCrLf & theEntry & vbCrLf & vbCrLf & _
           "This error is commonly caused by incorrect log file formatting. QSV might be able to recover from this error, but the statistics will probably be affected. It is suggested that you exit the application and correct the error in the log file." & vbCrLf & _
           "Would you like QSV to attempt to recover?", vbCritical + vbYesNo, "Error")
    If lRet = vbYes Then
        Resume Next
    Else
        MsgBox "The application will now exit.", vbCritical, "Goodbye"
        End
    End If

End Sub

Sub FinaliseStats()
'==============================================================================================='
' Here we organize all the information that depends on full game information                    '
'==============================================================================================='

    ' Finalize the players
    For i = 1 To stsPlayers.Count
        
        ' Set up the current player's nett score
        stsPlayers(i).lNett = stsPlayers(i).lFrags - stsPlayers(i).lDeaths - stsPlayers(i).lSuicides
        
        ' Get the longest survival
        tmpString = frmMain.lstRawLog.ListItems(1).Text
        Tm1 = (60 * Mid(tmpString, 1, InStr(1, tmpString, ":") - 1)) + Mid(tmpString, InStr(1, tmpString, ":") + 1)
        For i3 = 1 To frmMain.lstRawLog.ListItems.Count
            If frmMain.lstRawLog.ListItems(i3).SubItems(2) = stsPlayers(i).sName Then ' He died
            
                tmpString = frmMain.lstRawLog.ListItems(i3).Text
                Tm2 = (60 * Mid(tmpString, 1, InStr(1, tmpString, ":") - 1)) + Mid(tmpString, InStr(1, tmpString, ":") + 1)
                
                If (Tm2 - Tm1) > stsPlayers(i).lSurvived Then
                    stsPlayers(i).lSurvived = Tm2 - Tm1
                End If
                
                Tm1 = Tm2
                
            End If
        Next
        
        ' Picked on
        Tm1 = 0
        tmpString = ""
        For i3 = 1 To stsPlayers(i).colPPlayer.Count
            If stsPlayers(i).colPPlayer(i3).lFrags > Tm1 Then
                Tm1 = stsPlayers(i).colPPlayer(i3).lFrags
                tmpString = stsPlayers(i).colPPlayer(i3).sName
            End If
        Next
        stsPlayers(i).sPickedOn = tmpString
        
        ' Picked on by
        Tm1 = 0
        tmpString = ""
        For i3 = 1 To stsPlayers.Count
            If stsPlayers(i3).colPPlayer(stsPlayers(i).sName).lFrags > Tm1 Then
                Tm1 = stsPlayers(i3).colPPlayer(stsPlayers(i).sName).lFrags
                tmpString = stsPlayers(i3).sName
            End If
        Next
        stsPlayers(i).sPickedOnBy = tmpString
        
        ' Best Weapon
        Tm1 = 0
        tmpString = ""
        For i3 = 1 To stsPlayers(i).colPWeapon.Count
            If stsPlayers(i).colPWeapon(i3).lFrags > Tm1 Then
                Tm1 = stsPlayers(i).colPWeapon(i3).lFrags
                tmpString = stsPlayers(i).colPWeapon(i3).sName
            End If
        Next
        stsPlayers(i).sBestWeapon = tmpString
        
        ' Set up clan information with this player's information
        If stsPlayers(i).sClan <> "" Then
        
            tmpString = stsPlayers(i).sClan
            
            stsClan(tmpString).lFrags = stsClan(tmpString).lFrags + stsPlayers(i).lFrags
            stsClan(tmpString).lDeaths = stsClan(tmpString).lDeaths + stsPlayers(i).lDeaths
            stsClan(tmpString).lSuicides = stsClan(tmpString).lSuicides + stsPlayers(i).lSuicides
            stsClan(tmpString).lScore = stsClan(tmpString).lScore + stsPlayers(i).lScore
            ' We do not need to do the interfrags here, they are calculated with the kills
            If stsPlayers(i).lMkiol > stsClan(tmpString).lMkiol Then stsClan(tmpString).lMkiol = stsPlayers(i).lMkiol
            If stsPlayers(i).lMdwak > stsClan(tmpString).lMdwak Then stsClan(tmpString).lMdwak = stsPlayers(i).lMdwak
            
            stsClan(tmpString).colClanPlayers(stsPlayers(i).sName).lFrags = stsPlayers(i).lFrags
            
            stsClan(tmpString).lPlayers = stsClan(tmpString).lPlayers + 1
            
        End If
        
        If stsPlayers(i).lFrags <> stsPlayers(i).lSuicides Then
            stsPlayers(i).lEffectivity = ((stsPlayers(i).lFrags - stsPlayers(i).lSuicides) * 100 / (stsGame.Frags - stsGame.Suicides))
        Else
            stsPlayers(i).lEffectivity = 0
        End If
        
    Next
    
    ' Set up the clans' nett scores
    For i = 1 To stsClan.Count
        stsClan(i).lNett = (stsClan(i).lFrags - stsClan(i).lDeaths - (2 * stsClan(i).lSuicides)) / stsClan(i).lPlayers
    Next
    
    ' Weapons stats
    Tm1 = 0
    Tm2 = 0
    tmpString = ""
    tmpWeapon = ""
    For i = 1 To stsWeapons.Count
        ' Test for best weapon
        If stsWeapons(i).lFrags > Tm1 Then
            Tm1 = stsWeapons(i).lFrags
            tmpString = stsWeapons(i).sName
        End If
        
        ' Test for most dangerous weapon
        If stsWeapons(i).lSuicides > Tm2 Then
            Tm2 = stsWeapons(i).lSuicides
            tmpWeapon = stsWeapons(i).sName
        End If
        
        ' Check for best user
        i3 = 0
        For i2 = 1 To stsWeapons(i).colWPlayers.Count
            If stsWeapons(i).colWPlayers(i2).lFrags > i3 Then
                i3 = stsWeapons(i).colWPlayers(i2).lFrags
                stsWeapons(i).sBestUser = stsWeapons(i).colWPlayers(i2).sName
            End If
        Next
        
        ' Check for picked on
        i3 = 0
        For i2 = 1 To stsWeapons(i).colWPlayers.Count
            If stsWeapons(i).colWPlayers(i2).lDeaths > i3 Then
                i3 = stsWeapons(i).colWPlayers(i2).lDeaths
                stsWeapons(i).sPickedOn = stsWeapons(i).colWPlayers(i2).sName
            End If
        Next
        
        ' Check for immune
        i3 = stsGame.Frags
        For i2 = 1 To stsWeapons(i).colWPlayers.Count
            If stsWeapons(i).colWPlayers(i2).lDeaths < i3 Then
                i3 = stsWeapons(i).colWPlayers(i2).lDeaths
                stsWeapons(i).sImmune = stsWeapons(i).colWPlayers(i2).sName
            End If
        Next
        
    Next
    
    ' Final data added to the Game class
    stsGame.sBestWeapon = tmpString
    stsGame.sMostDangerousWeapon = tmpWeapon
    
End Sub

Sub CalculateRank(tmpType As woRanking)
'==============================================================================================='
' Get the rankings based on the scoring system selected                                         '
'==============================================================================================='
Dim tmpHighest              As Double
Dim AllHighest              As Double
Dim currRank                As Integer
Dim tmpRank                 As Integer
Dim ctlVal                  As Double
Dim tmpPos                  As Integer
Dim HasFoundWinner          As Boolean

    currRank = 1
    tmpHighest = -32000
    AllHighest = 32000
    frmMain.lstLeaderboard.ListItems.Clear
    
    ' Initialize the ranking, and set the Active value. The active value exists exclusively to
    ' prevent a span of Case statements further on down
    For i = 1 To stsPlayers.Count
        
        stsPlayers(i).lRanking = 0
        
        Select Case tmpType
        Case woRankFrags
            stsPlayers(i).lActive = stsPlayers(i).lFrags
        Case woRankNett
            stsPlayers(i).lActive = stsPlayers(i).lNett
        Case woRankCustom
            stsPlayers(i).lActive = stsPlayers(i).lScore
        Case woRankSkill
            stsPlayers(i).lActive = stsPlayers(i).lSkill
        End Select
        
    Next
    
    tmpRank = 1
    currRank = 1
    
    For i = 1 To stsPlayers.Count
        
        For i2 = 1 To stsPlayers.Count
            If (stsPlayers(i2).lActive > tmpHighest) And (stsPlayers(i2).lRanking = 0) Then
                tmpHighest = stsPlayers(i2).lActive
            End If
        Next
        
        For i2 = 1 To stsPlayers.Count
            If stsPlayers(i2).lActive = tmpHighest Then
                stsPlayers(i2).lRanking = currRank
                tmpRank = tmpRank + 1
            End If
        Next
        
        currRank = tmpRank
        tmpHighest = -32000
        
        ' Add a temporary place holder on the leaderboard
        frmMain.lstLeaderboard.ListItems.Add
        
    Next
    
    HasFoundWinner = False
    
    For i = 1 To stsPlayers.Count
    
        tmpPos = stsPlayers(i).lRanking
        
        ' It is handled like this in order to cater for multiple players with the same
        ' rank. For instance, if two guys were ranked no. 1, then there will be no rank
        ' position no. 2
        For i2 = 1 To stsPlayers.Count
        
            If stsPlayers(i2).lRanking = stsPlayers(i).lRanking Then
        
                With frmMain.lstLeaderboard.ListItems(tmpPos)
                    .Text = stsPlayers(i2).lRanking
                    .SubItems(1) = stsPlayers(i2).sName
                    .SubItems(2) = Round(stsPlayers(i2).lActive, 3)
                End With
            
                tmpPos = tmpPos + 1
                    
                If stsPlayers(i2).lRanking = 1 Then
                    If HasFoundWinner Then         ' a tie
                        frmMain.lblWinnerName.Caption = "...Tied Game..."
                        frmMain.lblWinnerClan.Caption = ""
                        frmMain.lblWinnerScore.Caption = ""
                    Else
                        frmMain.lblWinnerName.Caption = stsPlayers(i2).sName
                        frmMain.lblWinnerClan.Caption = IIf(stsPlayers(i2).sClan <> "", "Clan: " & stsPlayers(i2).sClan, "")
                        frmMain.lblWinnerScore.Caption = "Score: " & Round(stsPlayers(i2).lActive, 2)
                        HasFoundWinner = True
                    End If
                End If
                
            End If
        Next
    Next
    
End Sub
