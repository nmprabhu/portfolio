Attribute VB_Name = "ACC"
Option Explicit

Public Sub NamedRanges()

    'player position
    Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("A2"), Worksheets("ACC Statistics").range("A2").End(xlDown)).name = "PlayerPosition"
        
    'schools
    Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("B2"), Worksheets("ACC Statistics").range("B2").End(xlDown)).name = "Schools"
    
End Sub

Public Sub ParsePlayerInformation()

    'name ranges
    Call NamedRanges
    
    'defining variables
    Dim cell As range
    Dim comma As String
    Dim playerInfo As String
    Dim positionInfo As String
    
    'separating and recording player and position info
    For Each cell In Worksheets("ACC Statistics").range("PlayerPosition")
        
        comma = InStr(cell.Value, ",")
            
        'player info
        playerInfo = Left(cell, comma - 1)
        cell.Offset(0, 11).Value = playerInfo
        
        'position info
        positionInfo = Mid(cell, comma + 2)
        cell.Offset(0, 12).Value = positionInfo
            
    Next cell

    'naming player column
    Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("L2"), Worksheets("ACC Statistics").range("L2").End(xlDown)).name = "Players"
        
    'naming position column
    Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("M2"), Worksheets("ACC Statistics").range("M2").End(xlDown)).name = "Position"

End Sub

Public Sub FindPlayerInformation()

    'name ranges and create player and position
    Call ParsePlayerInformation
    
    'defining variables
    Dim COLPOS As Double
    Dim position As String
    Dim COLSCH As Double
    Dim school As String
    Dim COLTDS As Double
    Dim touchdowns As String
    Dim COLPTS As Double
    Dim points As String
    Dim playerName As String
    Dim cell As range
    Dim found As Double
    
    'constants for columns
    COLPOS = 1
    COLSCH = -10
    COLTDS = -6
    COLPTS = -2
    
    'input name of a player
    playerName = InputBox("Input a player name (i.e. Matt Dayes).", "Player Name")
    
    'loops though player names
    found = 0
    For Each cell In Worksheets("ACC Statistics").range("Players")
        If cell = playerName Then
            
            'player position
            position = cell.Offset(0, COLPOS).Value
                
            'player school
            school = cell.Offset(0, COLSCH).Value
                
            'player total points
            points = cell.Offset(0, COLPTS).Value
                
            'player touchdowns
            If cell.Offset(0, COLTDS).Value = 0 Then
                
                'no touchdowns
                Call MsgBox("ACC Player " & cell & " plays position " & position & " and attends " & school & ". " & "He has a total of " & points & " points.", , "Player Found")
                found = 1
            Else: touchdowns = cell.Offset(0, COLTDS).Value
                
                'some number of touchdowns
                Call MsgBox("ACC Player " & cell.Value & " plays position " & position & " and attends " & school & ". " & "He has scored " & touchdowns & " total touchdowns. " & "He has a total of " & points & " points.", , "Player Found")
                found = 1
            End If
        End If
    Next cell
        
    'error message
    If found <> 1 Then
        Call MsgBox("No ACC Player exists with the name " & playerName & ".", , "Player Not Found")
    End If
End Sub

Public Sub FindPlayerInformation2()

    'name ranges and create player and position
    Call ParsePlayerInformation
    
    'defining variables
    Dim COLPOS As Double
    Dim position As String
    Dim COLSCH As Double
    Dim school As String
    Dim COLTDS As Double
    Dim touchdowns As String
    Dim COLPTS As Double
    Dim points As String
    
    Dim playerName As String
    Dim cell As range
    Dim found As Double
    
    'constants for columns
    COLPOS = 1
    COLSCH = -10
    COLTDS = -6
    COLPTS = -2
    
    'input name of a player
    playerName = InputBox("Input a player name (i.e. Matt Dayes).", "Player Name")
    
    'loops though player names
    found = 0
    
    For Each cell In Worksheets("ACC Statistics").range("Players")
            
        If InStr(UCase(cell), UCase(playerName)) Then
            
            'player position
            position = cell.Offset(0, COLPOS).Value
                
            'player school
            school = cell.Offset(0, COLSCH).Value
                
            'player total points
            points = cell.Offset(0, COLPTS).Value
                
            'player touchdowns
            If cell.Offset(0, COLTDS).Value = 0 Then
                
                'no touchdowns
                Call MsgBox("ACC Player " & cell & " plays position " & position & " and attends " & school & ". " & "He has a total of " & points & " points.", , "Player Found")
                found = 1
                    
            Else: touchdowns = cell.Offset(0, COLTDS).Value
                
                'some number of touchdowns
                Call MsgBox("ACC Player " & cell.Value & " plays position " & position & " and attends " & school & ". " & "He has scored " & touchdowns & " total touchdowns. " & "He has a total of " & points & " points.", , "Player Found")
                    found = 1
                    
            End If
            
        End If
            
    Next cell
        
    'error message
    If found <> 1 Then
        Call MsgBox("No ACC Player exists with the name " & playerName & ".", , "Player Not Found")
    End If

End Sub

Public Sub FindMaxPointsPerGame()

    'call parse player info
    Call ParsePlayerInformation
    
    'defining variables
    Dim maxSoFar As Double
    Dim col As Double
    Dim cell As range
    Dim cellMax As range
    Dim row As range
    
    Dim COLPOS As Double
    Dim position As String
    Dim COLSCH As Double
    Dim school As String
    Dim COLNAME As Double
    Dim name As String
    Dim COLPTS As Double
    Dim points As String
    
    'columns
    COLPOS = 2
    COLSCH = -9
    COLNAME = 1
    COLPTS = -1
        
    'maximum loop
    maxSoFar = -1
    
    For Each cell In Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("K2"), Worksheets("ACC Statistics").range("K2").End(xlDown)).Cells
    
        If cell.Value > maxSoFar Then
        
            maxSoFar = cell.Value
            Set cellMax = cell
            
            'player position
            position = cell.Offset(0, COLPOS).Value
                
            'player school
            school = cell.Offset(0, COLSCH).Value
    
            'player name
            name = cell.Offset(0, COLNAME).Value
                
            'player total points
            points = cell.Offset(0, COLPTS).Value
            
        End If
        
    Next cell

    'message box
    Call MsgBox("The ACC player with the maximum points per game is " & name & " who plays the position " & position & ". He is from " & school & " with " & maxSoFar & " points per game and " & points & " total points.")

End Sub

Public Sub FindMaxPointsPerGamePosition()

    'defining variables
    Dim yesNo As String
    Dim position As String
    Dim cell As range
    Dim posCell As range
    Dim maxSoFar As Double
    Dim cellMax As range
    
    Dim COLSCH As Double
    Dim school As String
    Dim COLNAME As Double
    Dim name As String
    Dim COLPTS As Double
    Dim points As String
    
    'columns
    COLSCH = -9
    COLNAME = 1
    COLPTS = -1
    
    'ask user if they want to find max for certain position
    yesNo = MsgBox("Do you want to find the player with the maximum winnings for a particular position?", vbYesNo, "Position?")
    
    'if they say no
    If yesNo = vbNo Then
        Call FindMaxPointsPerGame
        Exit Sub
    End If
    
    'if they say yes
    If yesNo = vbYes Then
    
        'collect player position
        position = InputBox("Enter the position you would like to find the maximum for.", "Position for Max")
        
        'if they click cancel
        If position = "" Then
            Call MsgBox("Goodbye!", , "Goodbye")
            Exit Sub
        End If
        
        For Each cell In Worksheets("ACC Statistics").range(Worksheets("ACC Statistics").range("M2"), Worksheets("ACC Statistics").range("M2").End(xlDown))
        
            If UCase(cell.Value) = UCase(position) Then
            
                Set posCell = cell.Offset(0, -2)
    
                If posCell.Value > maxSoFar Then
        
                    maxSoFar = posCell.Value
                    Set cellMax = posCell
                
                    'player school
                    school = posCell.Offset(0, COLSCH).Value
    
                    'player name
                    name = posCell.Offset(0, COLNAME).Value
                
                    'player total points
                    points = posCell.Offset(0, COLPTS).Value
                    
                End If
                
            End If
        
        Next cell
 
    'message box
    Call MsgBox("The ACC player with the maximum points per game in position " & UCase(position) & " is " & name & ". He is from " & school & " with " & maxSoFar & " points per game and " & points & " total points.")
        
    End If
    
End Sub
