Option Explicit

Dim Grid(1 To 10, 1 To 10) As String
Dim PacManX As Integer, PacManY As Integer
Dim Ghosts() As Variant
Dim Score As Integer
Dim GameRunning As Boolean
Dim GhostSpeed As Double

' Start the game
Sub StartGame()
    Dim i As Integer, j As Integer

    ' Set up grid appearance
    SetupGrid

    ' Initialize grid
    For i = 1 To 10
        For j = 1 To 10
            Grid(i, j) = "." ' Dots
        Next j
    Next i

    ' Place Pac-Man
    PacManX = 5
    PacManY = 5
    Grid(PacManX, PacManY) = "P"

    ' Place Ghosts
    ReDim Ghosts(1 To 3)
    For i = 1 To UBound(Ghosts)
        Do
            Ghosts(i) = Array(Int((10 * Rnd) + 1), Int((10 * Rnd) + 1)) ' Random position
        Loop While Ghosts(i)(0) = PacManX And Ghosts(i)(1) = PacManY ' Ensure no overlap with Pac-Man
        Grid(Ghosts(i)(0), Ghosts(i)(1)) = "G"
    Next i

    ' Initialize game variables
    Score = 0
    GhostSpeed = 1 ' Initial ghost speed (in seconds)
    GameRunning = True

    ' Render the grid
    RenderGrid

    ' Setup controls
    SetupControls

    ' Start ghost movement
    Application.OnTime Now + GhostSpeed / 86400, "MoveGhosts"
End Sub

' Set up the grid appearance
Sub SetupGrid()
    Dim i As Integer, j As Integer

    ' Remove default gridlines
    Application.ActiveWindow.DisplayGridlines = False

    ' Format grid cells
    For i = 1 To 10
        For j = 1 To 10
            With Cells(i, j)
                .Interior.Color = RGB(255, 255, 255) ' White background
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 16
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        Next j
    Next i

    ' Set up score display
    With Range("L1:L2")
        .Merge
        .Value = "Score: 0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

' Render the grid
Sub RenderGrid()
    Dim i As Integer, j As Integer
    For i = 1 To 10
        For j = 1 To 10
            If Grid(i, j) = "P" Then
                Cells(i, j).Value = "??" ' Pac-Man
                Cells(i, j).Interior.Color = RGB(255, 255, 0) ' Yellow for Pac-Man
            ElseIf Grid(i, j) = "G" Then
                Cells(i, j).Value = "??" ' Ghost
                Cells(i, j).Interior.Color = RGB(255, 0, 0) ' Red for Ghost
            ElseIf Grid(i, j) = "." Then
                Cells(i, j).Value = "•" ' Dot
                Cells(i, j).Interior.Color = RGB(255, 255, 255) ' White for dots
            Else
                Cells(i, j).Value = ""
                Cells(i, j).Interior.Color = RGB(255, 255, 255) ' Clear cells
            End If
        Next j
    Next i
    ' Update score
    Range("L1").Value = "Score: " & Score
End Sub

' Move Pac-Man
Sub MovePacMan(Direction As String)
    If Not GameRunning Then Exit Sub

    ' Clear current position
    Grid(PacManX, PacManY) = ""

    ' Move Pac-Man based on direction
    Select Case Direction
        Case "UP"
            If PacManX > 1 Then PacManX = PacManX - 1
        Case "DOWN"
            If PacManX < 10 Then PacManX = PacManX + 1
        Case "LEFT"
            If PacManY > 1 Then PacManY = PacManY - 1
        Case "RIGHT"
            If PacManY < 10 Then PacManY = PacManY + 1
    End Select

    ' Check if Pac-Man eats a dot
    If Grid(PacManX, PacManY) = "." Then
        Score = Score + 10
        AdjustGhostSpeed
    End If

    ' Check if Pac-Man hits a ghost
    If Grid(PacManX, PacManY) = "G" Then
        GameOver
        Exit Sub
    End If

    ' Update position
    Grid(PacManX, PacManY) = "P"
    RenderGrid
End Sub

' Move Ghosts with smarter logic
Sub MoveGhosts()
    Dim i As Integer, GhostX As Integer, GhostY As Integer
    Dim NewGhostX As Integer, NewGhostY As Integer
    Randomize

    For i = 1 To UBound(Ghosts)
        GhostX = Ghosts(i)(0)
        GhostY = Ghosts(i)(1)

        ' Clear current position
        Grid(GhostX, GhostY) = ""

        ' Smarter movement: Move toward Pac-Man
        NewGhostX = GhostX
        NewGhostY = GhostY

        If Abs(GhostX - PacManX) > Abs(GhostY - PacManY) Then
            If GhostX > PacManX Then
                NewGhostX = GhostX - 1
            ElseIf GhostX < PacManX Then
                NewGhostX = GhostX + 1
            End If
        Else
            If GhostY > PacManY Then
                NewGhostY = GhostY - 1
            ElseIf GhostY < PacManY Then
                NewGhostY = GhostY + 1
            End If
        End If

        ' Check for collisions with Pac-Man
        If NewGhostX = PacManX And NewGhostY = PacManY Then
            GameOver
            Exit Sub
        End If

        ' Update ghost position
        Ghosts(i)(0) = NewGhostX
        Ghosts(i)(1) = NewGhostY
        Grid(NewGhostX, NewGhostY) = "G"
    Next i

    RenderGrid

    ' Continue ghost movement
    Application.OnTime Now + GhostSpeed / 86400, "MoveGhosts"
End Sub

' Adjust ghost speed based on score
Sub AdjustGhostSpeed()
    ' Increase speed only every 100 points, and decrease by smaller steps
    If Score Mod 100 = 0 And GhostSpeed > 0.5 Then
        GhostSpeed = GhostSpeed - 0.1 ' Decrease speed more gently
    End If
End Sub

' Setup controls
Sub SetupControls()
    Application.OnKey "{UP}", "MovePacManUp"
    Application.OnKey "{DOWN}", "MovePacManDown"
    Application.OnKey "{LEFT}", "MovePacManLeft"
    Application.OnKey "{RIGHT}", "MovePacManRight"
End Sub

' Control Pac-Man
Sub MovePacManUp()
    MovePacMan "UP"
End Sub

Sub MovePacManDown()
    MovePacMan "DOWN"
End Sub

Sub MovePacManLeft()
    MovePacMan "LEFT"
End Sub

Sub MovePacManRight()
    MovePacMan "RIGHT"
End Sub

' End the game
Sub GameOver()
    GameRunning = False
    MsgBox "Game Over! Final Score: " & Score, vbExclamation
End Sub


