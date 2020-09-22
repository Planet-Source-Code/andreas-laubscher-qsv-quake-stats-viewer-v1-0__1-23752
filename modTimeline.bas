Attribute VB_Name = "modTimeline"
Option Explicit

Public HasClicked As Boolean

Sub GenerateTimeLine()
Dim StepX As Single
Dim StepY As Single
Dim StepL As Single
Dim tmpPlayers() As Single
Dim currStep As Single
Dim currX As Single
Dim currY As Single
Dim lRet As Long
Dim HasValues As Boolean

    If Not frmMain.optRankFrags.Value Then
        lRet = MsgBox("The Timeline can only display the Frags Scoring System, but you have another Scoring System active. " & Chr(10) & "Would you like to revert to the Frags Scoring System?", vbQuestion + vbYesNo, "Illegal Scoring System")
        If lRet = vbYes Then
            frmMain.optRankFrags.Value = True
            frmMain.optRankFrags_Click
        Else
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    
    frmMain.picTimeline.Cls
    
    If frmMain.lstRawLog.ListItems.Count = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    ' Mmm... if the flexgrid contains values (more than one row), then we don't need to clear it.
    If frmMain.flxLegend.Rows > 1 Then
        HasValues = True
    Else
        HasValues = False
    End If
    
    ' Step sizes
    ' x = picture width / game duration
    If frmMain.lstRawLog.ListItems.Count = 0 Then
        MsgBox "The game contains no events! The timeline will not be drawn", vbExclamation, "What the...?"
        Exit Sub
    End If
    StepX = frmMain.picTimeline.Width / frmMain.lstRawLog.ListItems.Count
    
    ' y = height / winning score
    If stsPlayers(frmMain.lstLeaderboard.ListItems(1).SubItems(1)).lFrags = 0 Then
        MsgBox "The winner has no frags! The timeline will not be drawn", vbExclamation, "What the...?"
        Exit Sub
    End If
    StepY = frmMain.picTimeline.Height / stsPlayers(frmMain.lstLeaderboard.ListItems(1).SubItems(1)).lFrags
    
    ' Log step=picture width / game entries
    StepL = frmMain.picTimeline.Width / frmMain.lstRawLog.ListItems.Count
    
    ' Initialize the timeline variables
    
    For i = 1 To stsPlayers.Count
        stsPlayers(i).lTLine = 0
        stsPlayers(i).ltlX = 0
        stsPlayers(i).ltlY = frmMain.picTimeline.Height
        
        If (Not HasValues) Or HasClicked Then
            Randomize Timer
            stsPlayers(i).clrR = Rnd(100) * 255
            stsPlayers(i).clrG = Rnd(100) * 255
            stsPlayers(i).clrB = Rnd(100) * 255
        End If
        
    Next
    
    ' Loop through the entire raw log
    i = 1
    currStep = 0
    
    For i = 1 To frmMain.lstRawLog.ListItems.Count
    
        tmpString = frmMain.lstRawLog.ListItems(i).SubItems(1)
        
        If (tmpString = cWorld) Or (tmpString = frmMain.lstRawLog.ListItems(i).SubItems(2)) Then
            stsPlayers(tmpString).lTLine = stsPlayers(tmpString).lTLine - 1
        Else
            stsPlayers(tmpString).lTLine = stsPlayers(tmpString).lTLine + 1
        End If
        
        If stsPlayers(tmpString).dspGraph Then
            currX = i * StepX
            currY = frmMain.picTimeline.Height - (stsPlayers(tmpString).lTLine * StepY)
            If (stsPlayers(tmpString).ltlX <> currX) Or (stsPlayers(tmpString).ltlY <> currY) Then
                frmMain.picTimeline.Line (stsPlayers(tmpString).ltlX, stsPlayers(tmpString).ltlY)-(currX, currY), _
                                         RGB(stsPlayers(tmpString).clrR, stsPlayers(tmpString).clrG, stsPlayers(tmpString).clrB)
                stsPlayers(tmpString).ltlX = currX
                stsPlayers(tmpString).ltlY = currY
            End If
        End If
        
    Next
    
    If Not HasValues Then
        For i = 1 To frmMain.flxLegend.Rows - 1
            frmMain.flxLegend.RemoveItem 1
        Next
    End If
    
    frmMain.flxLegend.ColWidth(0) = frmMain.flxLegend.Width
    frmMain.flxLegend.ColAlignment(0) = 0
    
    ' Now set up the legend. We want to know which colour represents which player
    
    If Not HasValues Then
        For i = 1 To stsPlayers.Count
            frmMain.flxLegend.Rows = frmMain.flxLegend.Rows + 1
            frmMain.flxLegend.Row = frmMain.flxLegend.Rows - 1
            frmMain.flxLegend.CellForeColor = RGB(stsPlayers(i).clrR, stsPlayers(i).clrG, stsPlayers(i).clrB)
            frmMain.flxLegend.Text = stsPlayers(i).sName
        Next
    Else
        frmMain.flxLegend.Rows = frmMain.flxLegend.Rows + 1
        For i = 1 To stsPlayers.Count
            frmMain.flxLegend.Row = i - 1
            frmMain.flxLegend.CellForeColor = RGB(stsPlayers(i).clrR, stsPlayers(i).clrG, stsPlayers(i).clrB)
            frmMain.flxLegend.Text = stsPlayers(i).sName
        Next
    End If
    
    ' Initialize the strikethroughs.
    For i = 0 To frmMain.flxLegend.Rows - 1
        frmMain.flxLegend.Row = i
        frmMain.flxLegend.CellFontStrikeThrough = False
    Next
    
    ' Now remove the top and bottom rows, which are empty
    If frmMain.flxLegend.TextArray(0) = "" Then frmMain.flxLegend.RemoveItem 0
    If frmMain.flxLegend.TextArray(frmMain.flxLegend.Rows - 1) = "" Then frmMain.flxLegend.RemoveItem frmMain.flxLegend.Rows
    
    frmMain.flxLegend.Sort = flexSortGenericAscending
    
    ' Now that we've sorted and removed
    For i = 0 To frmMain.flxLegend.Rows - 1
        frmMain.flxLegend.Row = i
        frmMain.flxLegend.CellFontStrikeThrough = Not stsPlayers(frmMain.flxLegend.Text).dspGraph
    Next
    
    HasClicked = False
    Screen.MousePointer = vbDefault
    
End Sub
