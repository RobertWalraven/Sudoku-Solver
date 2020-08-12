Public Class CellButton

    ' Defines a single cell on the Sudoku game board

    Inherits Button

    ' Define possible key codes for both 9x9 and 16x16 games
    Enum CellKey
        KeyDelete = -1
        Key0 = 0
        Key1 = 1
        Key2 = 2
        Key3 = 3
        Key4 = 4
        Key5 = 5
        Key6 = 6
        Key7 = 7
        Key8 = 8
        Key9 = 9
        KeyA = 10
        KeyB = 11
        KeyC = 12
        KeyD = 13
        KeyE = 14
        KeyF = 15
        KeyUp = 16
        KeyDown = 17
        KeyLeft = 18
        KeyRight = 19
        KeySpace = 20
    End Enum

    Public ReadOnly x As Integer        ' Horizontal index for cell
    Public ReadOnly y As Integer        ' Vertical index for cell, top to bottom
    Public ReadOnly box As Integer      ' box index for the 9 3x3 or 16 4x4 boxes
    Public ReadOnly BoxSize As Integer  ' dimension of box: either 3 or 4
    Public ReadOnly BoxX As Integer     ' Horizontal position of the left side of the box this cell belongs to
    Public ReadOnly BoxY As Integer     ' Vertical position of the top of the box this cell belongs to
    Public Solution As Integer          ' Answer for this cell, or -1 if no answer yet
    Public OriginalSolution As Integer  ' Solution of cell when it receives focus from keyboard
    Public PendingSolution As Integer   ' Pending Solution before it is finalized or if a hint
    Public Guess As Boolean             ' The Solution is a guess
    Public forcedValue As Integer       ' Solution specified by user during keyboard entry of game
    Public Hint As Integer              ' Hint value to display
    Public Possibility(15) As Boolean   ' Remaining possibilities for cell
    Public PossibilityCount As Integer  ' Number of remaining possibilities

    Private SavedColor As Color         ' Original background color when color changed by getting focus
    ' The following are read from settings:
    Private ReadOnly mSize As Integer   ' dimension of game - either 9 or 16
    Private ReadOnly mStartValue As Integer     ' Starting index: 1 for 9x9 game, 0 for 16x16 game
    Private ReadOnly mEndValue As Integer       ' Ending index: 9 for 9x9 game, 15 for 16x16 game

    Public Event KeyPressed(ByVal sender As CellButton, ByVal key As CellKey)

    Sub New(ByVal x As Integer, ByVal y As Integer)
        ' Save the input values
        Me.x = x
        Me.y = y
        ' Get Settings data
        mSize = My.Settings.Size
        BoxSize = My.Settings.BoxSize
        mStartValue = My.Settings.StartValue
        mEndValue = My.Settings.EndValue
        box = (y \ BoxSize) * BoxSize + (x \ BoxSize)
        BoxX = (x \ BoxSize) * BoxSize
        BoxY = (y \ BoxSize) * BoxSize
        ' Initialize rest of data for a blank cell
        Solution = -1
        OriginalSolution = -1
        Guess = False
        forcedValue = -1
        PendingSolution = -1
        Hint = -1
        PossibilityCount = 0
        For i As Integer = mStartValue To mEndValue
            Possibility(i) = True
            PossibilityCount += 1
        Next
    End Sub

    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        ' Fire KeyPressed event for some particular keys.
        ' Everything else will be caught by ProcessDialogChar
        Select Case keyData
            Case Keys.Up
                RaiseEvent KeyPressed(Me, CellKey.KeyUp)
                Return True
            Case Keys.Down
                RaiseEvent KeyPressed(Me, CellKey.KeyDown)
                Return True
            Case Keys.Left, Keys.Back
                RaiseEvent KeyPressed(Me, CellKey.KeyLeft)
                Return True
            Case Keys.Right
                RaiseEvent KeyPressed(Me, CellKey.KeyRight)
                Return True
            Case Keys.Delete
                RaiseEvent KeyPressed(Me, CellKey.KeyDelete)
                Return True
            Case Else
                Return MyBase.ProcessDialogKey(keyData)
        End Select
    End Function

    Protected Overrides Function ProcessDialogChar(ByVal charCode As Char) As Boolean
        ' Allow possible value keys, convert period or space to KeyDelete,
        ' and ignore all other characters.
        Dim k As CellKey = CellKey.Key1
        Dim myCode As Char = UCase(charCode)
        Select Case myCode
            Case "0"c, "A"c, "B"c, "C"c, "D"c, "E"c, "F"c
                If mSize = 16 Then
                    Select Case myCode
                        Case "0"c : k = CellKey.Key0
                        Case "A"c : k = CellKey.KeyA
                        Case "B"c : k = CellKey.KeyB
                        Case "C"c : k = CellKey.KeyC
                        Case "D"c : k = CellKey.KeyD
                        Case "E"c : k = CellKey.KeyE
                        Case "F"c : k = CellKey.KeyF
                    End Select
                Else
                    k = CellKey.KeyDelete
                End If
            Case "1"c : k = CellKey.Key1
            Case "2"c : k = CellKey.Key2
            Case "3"c : k = CellKey.Key3
            Case "4"c : k = CellKey.Key4
            Case "5"c : k = CellKey.Key5
            Case "6"c : k = CellKey.Key6
            Case "7"c : k = CellKey.Key7
            Case "8"c : k = CellKey.Key8
            Case "9"c : k = CellKey.Key9
            Case "."c : k = CellKey.KeyDelete
            Case " "c : k = CellKey.KeySpace
            Case Else
                Return False
        End Select
        Guess = False
        ' If the key entered is not allowed, don't enter it and let user know
        If k < 0 Then Return False
        If k <= 15 Then
            If Possibility(k) = False AndAlso Solution = -1 Then
                MsgBox(k & " is not possible for this cell")
                Return False
            End If
        End If
        ' Otherwise raise the event
        RaiseEvent KeyPressed(Me, k)
        Return True
    End Function

    Public Sub SetForcedValue(ByVal v As Integer)
        ' Set a forced value in the data, which is value typed by the user.
        ForeColor = Color.Black
        Me.OriginalSolution = -1
        Me.PendingSolution = -1
        Me.Guess = False
        If v >= mStartValue And v <= mEndValue Then
            Me.Solution = v
            Me.forcedValue = v
            For i As Integer = mStartValue To mEndValue
                If i = v Then
                    Me.Possibility(i) = True
                Else
                    Me.Possibility(i) = False
                End If
            Next
            Me.PossibilityCount = 1
        Else
            ' Clear the forced value
            Me.Solution = -1
            Me.forcedValue = -1
            For i As Integer = mStartValue To mEndValue
                Me.Possibility(i) = True
            Next
        End If
        Me.Refresh()
    End Sub

    ' Define a big font for painting the cell
    Private Shared ReadOnly bigfont As New Font(FontFamily.GenericSansSerif, 20, FontStyle.Bold)

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        ' Fill in the cell contents

        MyBase.OnPaint(e)

        If Me.Solution <> -1 Then
            Dim myColor As Color
            If Guess Then
                myColor = Color.ForestGreen
            Else
                myColor = Me.ForeColor
            End If
            If mSize = 9 Then
                e.Graphics.DrawString(Hex(Me.Solution), bigfont, New SolidBrush(myColor), 9, 4)
            Else
                e.Graphics.DrawString(Hex(Me.Solution), bigfont, New SolidBrush(myColor), 14, 9)
            End If
            Exit Sub
        End If

        If Me.Hint <> -1 Then
            If mSize = 9 Then
                e.Graphics.DrawString(Hex(Me.Hint), bigfont, New SolidBrush(Color.Magenta), 5, 4)
            Else
                e.Graphics.DrawString(Hex(Me.Hint), bigfont, New SolidBrush(Color.Magenta), 5, 9)
            End If
        End If

        If Not My.Settings.ShowPossibilites Then Return

        Dim myBrush As New SolidBrush(Me.ForeColor)

        If mSize = 9 Then
            If Possibility(1) Then e.Graphics.DrawString("1", Me.Font, myBrush, 5, 3)
            If Possibility(2) Then e.Graphics.DrawString("2", Me.Font, myBrush, 15, 3)
            If Possibility(3) Then e.Graphics.DrawString("3", Me.Font, myBrush, 25, 3)
            If Possibility(4) Then e.Graphics.DrawString("4", Me.Font, myBrush, 5, 14)
            If Possibility(5) Then e.Graphics.DrawString("5", Me.Font, myBrush, 15, 14)
            If Possibility(6) Then e.Graphics.DrawString("6", Me.Font, myBrush, 25, 14)
            If Possibility(7) Then e.Graphics.DrawString("7", Me.Font, myBrush, 5, 25)
            If Possibility(8) Then e.Graphics.DrawString("8", Me.Font, myBrush, 15, 25)
            If Possibility(9) Then e.Graphics.DrawString("9", Me.Font, myBrush, 25, 25)
        Else
            If Possibility(0) Then e.Graphics.DrawString("0", Me.Font, myBrush, 5, 2)
            If Possibility(1) Then e.Graphics.DrawString("1", Me.Font, myBrush, 15, 2)
            If Possibility(2) Then e.Graphics.DrawString("2", Me.Font, myBrush, 25, 2)
            If Possibility(3) Then e.Graphics.DrawString("3", Me.Font, myBrush, 35, 2)
            If Possibility(4) Then e.Graphics.DrawString("4", Me.Font, myBrush, 5, 13)
            If Possibility(5) Then e.Graphics.DrawString("5", Me.Font, myBrush, 15, 13)
            If Possibility(6) Then e.Graphics.DrawString("6", Me.Font, myBrush, 25, 13)
            If Possibility(7) Then e.Graphics.DrawString("7", Me.Font, myBrush, 35, 13)
            If Possibility(8) Then e.Graphics.DrawString("8", Me.Font, myBrush, 5, 24)
            If Possibility(9) Then e.Graphics.DrawString("9", Me.Font, myBrush, 15, 24)
            If Possibility(10) Then e.Graphics.DrawString("A", Me.Font, myBrush, 25, 24)
            If Possibility(11) Then e.Graphics.DrawString("B", Me.Font, myBrush, 35, 24)
            If Possibility(12) Then e.Graphics.DrawString("C", Me.Font, myBrush, 5, 35)
            If Possibility(13) Then e.Graphics.DrawString("D", Me.Font, New SolidBrush(Me.ForeColor), 15, 35)
            If Possibility(14) Then e.Graphics.DrawString("E", Me.Font, New SolidBrush(Me.ForeColor), 25, 35)
            If Possibility(15) Then e.Graphics.DrawString("F", Me.Font, New SolidBrush(Me.ForeColor), 35, 35)
        End If

    End Sub

    Private Sub CellButton_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        SavedColor = Me.BackColor
        Me.BackColor = Color.AliceBlue
        Me.OriginalSolution = Me.Solution
    End Sub

    Private Sub CellButton_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        Me.BackColor = SavedColor
    End Sub

End Class
