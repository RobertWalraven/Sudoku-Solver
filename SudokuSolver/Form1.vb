Imports System.ComponentModel
Imports System.Drawing.Printing
Imports System.IO

Public Class Form1

#Region "Declarations"

    Private mSize As Integer            ' Dimension of board.  Either 9 or 16
    Private mBoxSize As Integer         ' Dimension of box.  Either 3 or 4
    Private mStartValue As Integer      ' Lowest possible value
    Private mEndValue As Integer        ' Largest possible value
    Private mShowHint As Boolean        ' Show a solution hint
    Private CellScreenSize As Integer   ' Horizontal and vertical screen size of cells
    Private Done As Boolean             ' Board has been solved

    Private ReadOnly ExtraSpace As Integer = 4   ' Screen spacing between boxes
    Private ReadOnly PrinterDoesColor As Boolean = False

    Friend Cell(15, 15) As CellButton       ' Allow enough cells for a 16x16 board
    Friend SavedCell(15, 15) As CellButton  ' Holds backup copy of Cell
    Friend RowLabels(15) As Label           ' Labels along left side of grid
    Friend ColumnLabels(15) As Label        ' Labels along top of grid
    Friend ColChr() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p"}
    Friend RowChr() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"}

#End Region
    '===============================================================================================
#Region "Initialization"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeBoard()
        Me.cbShowPossibilities.Checked = My.Settings.ShowPossibilites
        Me.cbSingleStep.Checked = My.Settings.Singlestep
        mShowHint = False
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub InitializeBoard()

        ' Get settings data
        mSize = My.Settings.Size

        'Define value range and box size
        If mSize = 9 Then
            mStartValue = 1
            mEndValue = 9
            mBoxSize = 3
            CellScreenSize = 41
        Else
            mStartValue = 0
            mEndValue = 15
            mBoxSize = 4
            CellScreenSize = 51
        End If
        My.Settings.StartValue = mStartValue
        My.Settings.EndValue = mEndValue
        My.Settings.BoxSize = mBoxSize
        My.Settings.Save()

        ' Define screen size
        If mSize = 9 Then
            Me.Width = 416 + 300
            Me.Height = 460
        Else
            Me.Width = 705 + 300 + 160
            Me.Height = 913
        End If

        ' Clear current cells from form
        For x As Integer = 0 To 15
            For y As Integer = 0 To 15
                If Cell(x, y) IsNot Nothing Then
                    Cell(x, y).Dispose()
                End If
            Next
        Next

        ' Clear current column/row labels
        For N As Integer = 0 To 15
            If RowLabels(N) IsNot Nothing Then
                RowLabels(N).Dispose()
            End If
            If ColumnLabels(N) IsNot Nothing Then
                ColumnLabels(N).Dispose()
            End If
        Next
        ' Add cells to display
        Dim c As CellButton
        For row As Integer = 0 To mSize - 1
            For column As Integer = 0 To mSize - 1
                c = New CellButton(column, row) With {
                    .Location = New Point(20 + Position(column), 20 + Position(row)),
                    .Size = New Size(CellScreenSize, CellScreenSize),
                    .BackColor = Color.WhiteSmoke,
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Font = Me.Font
                }
                Me.Controls.Add(c)
                Cell(column, row) = c
                If mSize = 9 Then
                    Select Case Cell(column, row).box
                        Case 0, 2, 4, 6, 8
                            c.BackColor = Color.White
                        Case Else
                            c.BackColor = Color.Gainsboro
                    End Select
                Else
                    Select Case Cell(column, row).box
                        Case 0, 2, 5, 7, 8, 10, 13, 15
                            c.BackColor = Color.White
                        Case Else
                            c.BackColor = Color.Gainsboro
                    End Select
                End If
                AddHandler c.KeyPressed, AddressOf CellsKeyPressed
            Next
        Next

        ' Add column and row names

        Dim L As Label
        For Column As Integer = 0 To mSize - 1
            L = New Label With {
                .Name = "C" & Chr(Asc("A") + Column),
                .Location = New Point(17 + Position(Column) + CellScreenSize \ 2, 5),
                .Width = 15,
                .Text = ColChr(Column)
            }
            Me.Controls.Add(L)
            ColumnLabels(Column) = L
        Next

        For row As Integer = 0 To mSize - 1
            L = New Label With {
                .Name = "R" & Chr(Asc("A") + row),
                .Location = New Point(5, 14 + Position(row) + CellScreenSize \ 2),
                .Width = 15,
                .Text = RowChr(row)
            }
            Me.Controls.Add(L)
            RowLabels(row) = L
        Next

        ' Finish up

        Label1.Text = ""
        Label2.Text = ""
        Done = False
        Me.btnSolve.Enabled = True
        Me.btnShowHint.Enabled = True
        Me.btnSearch.Visible = False
        Cell(0, 0).TabIndex = 0
        Cell(0, 0).Focus()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Function Position(ByVal d As Integer) As Integer
        ' Given the horizontal or vertical index of the cell, return its screen position
        Return d * CellScreenSize + (d \ mBoxSize) * ExtraSpace
    End Function

#End Region
    '===============================================================================================
#Region "Keystroke handler"

    Private Sub CellsKeyPressed(ByVal c As CellButton, ByVal k As CellButton.CellKey)
        Dim x As Integer = c.x
        Dim y As Integer = c.y
        Dim OriginalSolution As Integer = c.OriginalSolution
        Select Case k
            Case CellButton.CellKey.KeyLeft
                x -= 1
                If x < 0 Then x += mSize
            Case CellButton.CellKey.KeyRight
                x += 1
                If x >= mSize Then x -= mSize
            Case CellButton.CellKey.KeyUp
                y -= 1
                If y < 0 Then y += mSize
            Case CellButton.CellKey.KeyDown
                y += 1
                If y >= mSize Then y -= mSize
            Case CellButton.CellKey.Key0,
                 CellButton.CellKey.Key1 To CellButton.CellKey.Key9,
                 CellButton.CellKey.KeyA To CellButton.CellKey.KeyF
                If mSize = 9 Then
                    ' Some keys for size 16 are not allowed here
                    Select Case k
                        Case CellButton.CellKey.Key0
                            Return
                        Case CellButton.CellKey.KeyA To CellButton.CellKey.KeyF
                            Return
                    End Select
                End If
                If k <> c.Solution Then
                    If c.forcedValue <> -1 Then RestorePossibility(x, y, c.Solution)
                    c.SetForcedValue(k)
                    x += 1
                    If x >= mSize Then
                        x -= mSize
                        y += 1
                        If y >= mSize Then y = mSize - 1
                    End If
                End If
                If OriginalSolution <> c.Solution Then RestorePossibility(x, y, OriginalSolution)
                InitializePossibilities()
                UpdatePossibilities()
                Label1.Text = ""
                Label2.Text = ""
                If RemainingCellsToSolve() = 0 Then
                    Text = "The puzzle has been solved!"
                End If
                Me.Refresh()
            Case CellButton.CellKey.KeySpace
                x += 1
                If x >= mSize Then
                    x -= mSize
                    y += 1
                    If y >= mSize Then y = mSize - 1
                End If
            Case CellButton.CellKey.KeyDelete
                If c.forcedValue >= 0 Then
                    'RestorePossibility(x, y, c.Solution)
                    c.SetForcedValue(-1)
                    InitializePossibilities()
                    UpdatePossibilities()
                    Me.Refresh()
                End If
            Case Else
                Exit Sub
        End Select
        Cell(x, y).Focus()
    End Sub

#End Region
    '===============================================================================================
#Region "Control methods"

    Private Sub CbShowPossibilities_CheckedChanged(sender As Object, e As EventArgs) Handles cbShowPossibilities.CheckedChanged
        My.Settings.ShowPossibilites = cbShowPossibilities.Checked
        Me.Refresh()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        mShowHint = False
        'SearchForSolution()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub BtnShowHint_Click(sender As Object, e As EventArgs) Handles btnShowHint.Click

        mShowHint = True
        FindACellSolution()
        mShowHint = False
        Me.Refresh()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub BtnSolve_Click(sender As Object, e As EventArgs) Handles btnSolve.Click
        Dim myDone As Boolean = False
        Do Until myDone
            Label1.Text = ""
            Label2.Text = ""
            FindACellSolution()
            Me.Refresh()
            If Me.cbSingleStep.Checked Then myDone = True
            If Done Then myDone = True
        Loop
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        MsgBox("Sudoku Solver by Robert Walraven, 2017")
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuHelp_Click(sender As Object, e As EventArgs) Handles mnuHelp.Click
        MsgBox("Select a cell and press a number 1 to 9" & vbCrLf &
         "Press delete to clear a cell" & vbCrLf &
         "Use the arrows to navigate within the grid")
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuLoad_Click(sender As Object, e As EventArgs) Handles mnuLoad.Click
        ' Retrieve and display the previously saved game.
        Me.Cursor = Cursors.WaitCursor
        InitializeBoard()
        Dim IniForApp As String = GetIniForPath()
        Dim myLine As String
        Dim value As Integer
        If Not IO.File.Exists(IniForApp) Then
            If mSize = 9 Then
                MsgBox("No 9 x 9 game has been saved yet.")
            Else
                MsgBox("No 16 x 16 game has been saved yet.")
            End If
            Return
        End If
        Using sr As New StreamReader(IniForApp)
            myLine = sr.ReadLine
            myLine = sr.ReadLine
            Do While sr.Peek() >= 0
                myLine = sr.ReadLine
                Dim Column As Integer = Convert.ToInt32(myLine.Substring(4, 1), 16)
                Dim Row As Integer = Convert.ToInt32(myLine.Substring(5, 1), 16)
                If myLine.Length = 8 Then
                    Dim myChar As String = myLine.Substring(7, 1)
                    value = Convert.ToInt32(myChar, 16)
                    Cell(Column, Row).Guess = False
                    Cell(Column, Row).forcedValue = -1
                    Cell(Column, Row).PendingSolution = -1
                    If value <> -1 Then
                        Cell(Column, Row).Solution = value
                        For myValue As Integer = mStartValue To mEndValue
                            Cell(Column, Row).Possibility(myValue) = False
                        Next
                        Cell(Column, Row).Possibility(value) = True
                        Cell(Column, Row).PossibilityCount = 1
                        Cell(Column, Row).ForeColor = Color.Black
                    End If
                End If
            Loop
        End Using

        UpdatePossibilities()
        Me.Cursor = Cursors.Default

    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuNew9x9_Click(sender As Object, e As EventArgs) Handles mnuNew9x9.Click
        My.Settings.Size = 9
        My.Settings.Save()
        InitializeBoard()
        'SearchThread1.StartNew()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MenuNew16x16_Click(sender As Object, e As EventArgs) Handles MenuNew16x16.Click
        My.Settings.Size = 16
        My.Settings.Save()
        InitializeBoard()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuPrint_Click(sender As Object, e As EventArgs) Handles mnuPrint.Click
        PrintDocument1.Print()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim myImage As New Bitmap(Me.Width, Me.Height)
        Dim PrintSize As Size = e.MarginBounds.Size
        Dim scale As Double = 1D

        If Not PrinterDoesColor Then
            SwitchColors(Color.Blue, Color.DarkGray)
            SwitchColors(Color.Red, Color.Silver)
        End If

        Me.DrawToBitmap(myImage, New Rectangle(Point.Empty, Me.Size))
        myImage.RotateFlip(RotateFlipType.Rotate90FlipNone)

        PrintSize.Width = CInt(PrintSize.Width * 0.96) ' Convert to pixels
        PrintSize.Height = CInt(PrintSize.Height * 0.96)

        If myImage.Width > PrintSize.Width Then
            scale = PrintSize.Width / myImage.Width
            e.Graphics.ScaleTransform(CSng(scale), CSng(scale))
        End If

        If myImage.Height * scale > PrintSize.Height Then
            scale = PrintSize.Height / (myImage.Height * scale)
            e.Graphics.ScaleTransform(CSng(scale), CSng(scale))
        End If

        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        e.Graphics.DrawImage(myImage, e.MarginBounds.Location)

        myImage.Dispose()

        If Not PrinterDoesColor Then
            SwitchColors(Color.DarkGray, Color.Blue)
            SwitchColors(Color.Silver, Color.Red)
        End If

    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuQuit_Click(sender As Object, e As EventArgs) Handles mnuQuit.Click
        Close()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub MnuSave_Click(sender As Object, e As EventArgs) Handles mnuSave.Click
        ' Save the current board to a file.
        Me.Cursor = Cursors.WaitCursor
        Dim IniForApp As String = GetIniForPath()
        If IO.File.Exists(IniForApp) Then My.Computer.FileSystem.DeleteFile(IniForApp)
        Using outputFile As StreamWriter = File.CreateText(IniForApp)
            outputFile.WriteLine("SUDOKU, " & Application.ProductVersion)
            outputFile.WriteLine("")
            For row As Integer = 0 To mSize - 1
                For column As Integer = 0 To mSize - 1
                    outputFile.WriteLine("Cell" & Hex(column) & Hex(row) & "=" & Hex(Cell(column, row).Solution))
                Next
            Next
            outputFile.Close()
        End Using
        Me.Cursor = Cursors.Default
    End Sub

    Private Function GetIniForPath() As String
        'Specify location of saved game file.
        ' Each game size has its own saved game file for convenience.
        If mSize = 9 Then
            Return Application.CommonAppDataPath & "\SUDOKU9.INI"
        Else
            Return Application.CommonAppDataPath & "\SUDOKU16.INI"
        End If
    End Function

#End Region
    '===============================================================================================
#Region "Rules"
    Private Function OnlyChoice(ByRef Column As Integer,
                                ByRef Row As Integer,
                                ByRef Value As Integer,
                                ByRef Text As String) As Boolean
        ' When a row, column, or box contains all but one solved cells, then there ia only
        ' one choice remaining, so the remaining value must go int the empty cell.
        '
        ' If a cell is found that is an only choice, return true with the parameters of
        ' the column, row, and value and a text description of why this is the only
        ' choice.  Othewise return false.
        '
        ' Note:  arguments are meaningless if only-choice cell not found.

        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton

        Dim Type As String = ""

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1
                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                        Type = "row"
                    Case 2
                        GetCellsInColumn(i, myCells)
                        Type = "column"
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                        Type = "box"
                End Select

                ' If just one cell is empty, find it's index
                Dim Empty As Integer = 0
                Dim EmptyIndex As Integer = -1
                For j As Integer = 0 To mSize - 1
                    If myCells(j).Solution = -1 Then
                        EmptyIndex = j
                        Empty += 1
                        If Empty > 1 Then Exit For
                    End If
                Next

                If Empty = 1 Then
                    ' Remove illegal possibilities from empty cell
                    For j As Integer = 0 To mSize - 1
                        If j <> EmptyIndex Then
                            RemovePossibilityFromCell(myCells(EmptyIndex), myCells(j).Solution)
                        End If
                    Next
                    ' Locate the only value and return cell data
                    For myValue As Integer = mStartValue To mEndValue
                        If myCells(EmptyIndex).Possibility(myValue) Then
                            Column = myCells(EmptyIndex).x
                            Row = myCells(EmptyIndex).y
                            Value = myValue
                            If Not mShowHint Then Cell(Column, Row) = myCells(EmptyIndex)
                            Text = Hex(Value) & " at " & RowChr(Row) & ColChr(Column) &
                                " is the only value left in the " & Type
                            Return True
                        End If
                    Next
                End If
            Next
        Next

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function SinglePossibility(ByRef Column As Integer,
                                ByRef Row As Integer,
                                ByRef Value As Integer,
                                ByRef Text As String) As Boolean
        ' Find the next cell that has only one possibility and still has an empty solution.

        Dim myCells(mSize - 1, mSize - 1) As CellButton

        CopyCells(Cell, myCells)

        ' Search row for a cell with a single value

        For myRow As Integer = 0 To mSize - 1
            UpdatePossibilitiesForRow(myRow)
            For myColumn As Integer = 0 To mSize - 1
                With Cell(myColumn, myRow)
                    If (.Solution = -1) AndAlso (.PossibilityCount = 1) Then
                        For myValue As Integer = mStartValue To mEndValue
                            If .Possibility(Value) Then
                                Column = myColumn
                                Row = myRow
                                Value = myValue
                                Text = Hex(Value) & " is the only possibility for " &
                                    "cell " & RowChr(Row) & ColChr(Column)
                                If mShowHint Then CopyCells(myCells, Cell)
                                Return True
                            End If
                        Next
                    End If
                End With
            Next
        Next

        ' Search columns for a cell with a single value.

        For myColumn As Integer = 0 To mSize - 1
            UpdatePossibilitiesForColumn(myColumn)
            For myRow As Integer = 0 To mSize - 1
                With Cell(myColumn, myRow)
                    If (.Solution = -1) AndAlso (.PossibilityCount = 1) Then
                        For myValue As Integer = mStartValue To mEndValue
                            If .Possibility(myValue) Then
                                Column = myColumn
                                Row = myRow
                                Value = myValue
                                Text = Hex(Value) & " is the only possibility for " &
                                    "cell " & RowChr(Row) & ColChr(Column)
                                If mShowHint Then CopyCells(myCells, Cell)
                                Return True
                            End If
                        Next
                    End If
                End With
            Next
        Next

        ' Search box for a cell with a single value

        For BoxX As Integer = 0 To mBoxSize * (mBoxSize - 1) Step mBoxSize
            For BoxY As Integer = 0 To mBoxSize * (mBoxSize - 1) Step mBoxSize
                UpdatePossibilitiesForBox(BoxX, BoxY)
                For myColumn As Integer = BoxX To BoxX + mBoxSize - 1
                    For myRow As Integer = BoxY To BoxY + mBoxSize - 1
                        With Cell(myColumn, myRow)
                            If (.Solution = -1) AndAlso (.PossibilityCount = 1) Then
                                For myValue As Integer = mStartValue To mEndValue
                                    If .Possibility(Value) Then
                                        Column = myColumn
                                        Row = myRow
                                        Value = myValue
                                        Text = Hex(Value) & " is the only possibility for " &
                                            "cell " & RowChr(Row) & ColChr(Column)
                                        If mShowHint Then CopyCells(myCells, Cell)
                                        Return True
                                    End If
                                Next
                            End If
                        End With
                    Next
                Next
            Next
        Next


        Return False
    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function OnlySquare(ByRef Column As Integer,
                                ByRef Row As Integer,
                                ByRef Value As Integer,
                                ByRef Text As String) As Boolean
        ' Find a cell in a column, row, or box that is the only place a particular value can occur.
        ' The cell may contain other possibilities, but only the particular value is valid.
        '
        ' NOTE:  Before running this routine the routine SinglePossibility must be run repeatedly
        ' until it fails because that sets the valid probabilities for the cells in all columns,
        ' rows, and boxes.
        '
        ' Note:  arguments are meaningless if only-choice cell not found.

        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton

        Dim Type As String = ""

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1

                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                        Type = "row"
                    Case 2
                        GetCellsInColumn(i, myCells)
                        Type = "column"
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                        Type = "box"
                End Select
                Dim myIndex As Integer = -1
                For myValue As Integer = mStartValue To mEndValue
                    Dim Count As Integer = 0
                    For j As Integer = 0 To mSize - 1
                        If myCells(j).Possibility(myValue) AndAlso myCells(j).Solution = -1 Then
                            myIndex = j
                            Count += 1
                            If Count > 1 Then Exit For
                        End If
                    Next
                    If Count = 1 Then
                        Column = myCells(myIndex).x
                        Row = myCells(myIndex).y
                        Value = myValue
                        If Not mShowHint Then Cell(Column, Row) = myCells(myIndex)
                        Text = Hex(Value) & " at cell " & RowChr(Row) & ColChr(Column) &
                            " is the ony place this value can occur in the " & Type
                        Return True
                    End If
                Next

            Next

        Next

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function SubGroupExclusion(ByRef Text As String) As Boolean
        ' There are four conditions that constitute Subgroup Exclusion:
        ' 1. For a given row, if a particular value only occurs in the cells that are in one box
        '    then the possibility of the value an be removed from other rows in that box.
        ' 2. For a given column, if a particular value only occurs in the cells that are in one box
        '    then the possibility of the value can be removed from other columns of that box.
        ' 3. For a given box, if a particular value only occurs in a particular row of the box
        '    then the possibility of that value can be removed from thesame row in other boxes.
        ' 4. For a given box, if a particular value only occurs in a particular column of the box
        '    then the possibility of the value can be removed from the same column in other boxes.

        Dim myCells(mSize - 1) As CellButton

        Dim ExcludingIndex As Integer

        ' Look for condition 1
        For row As Integer = 0 To mSize - 1
            GetCellsInRow(row, myCells)
            For Value As Integer = mStartValue To mEndValue
                If SubGroupExclusion1(myCells, Value, ExcludingIndex) Then
                    Dim X As Integer = mBoxSize * ExcludingIndex
                    Dim N As Integer = Cell(X, row).box
                    GetCellsInBoxNumber(N, myCells)
                    Dim Count As Integer = 0
                    For i As Integer = 0 To mSize - 1
                        With myCells(i)
                            If .y <> row Then
                                If .Possibility(Value) And .Solution = -1 Then Count += 1
                                If Not mShowHint Then RemovePossibility(.x, .y, Value)
                            End If
                        End With
                    Next
                    If Count > 0 Then
                        Text = Hex(Value) & " only occurs in row " & RowChr(row) &
                            " in box at " & BoxChr(N) &
                            " so this value can be deleted from other cells of the box"
                        Return True
                    End If
                    GetCellsInRow(row, myCells)
                End If
            Next
        Next

        ' Look for condition 2
        For column As Integer = 0 To mSize - 1
            GetCellsInColumn(column, myCells)
            For Value As Integer = mStartValue To mEndValue
                If SubGroupExclusion1(myCells, Value, ExcludingIndex) Then
                    Dim Y As Integer = mBoxSize * ExcludingIndex
                    Dim N As Integer = Cell(column, Y).box
                    GetCellsInBoxNumber(N, myCells)
                    Dim Count As Integer = 0
                    For i As Integer = 0 To mSize - 1
                        With myCells(i)
                            If .x <> column Then
                                If .Possibility(Value) And .Solution = -1 Then Count += 1
                                If Not mShowHint Then RemovePossibility(.x, .y, Value)
                            End If
                        End With
                    Next
                    If Count > 0 Then
                        Text = Hex(Value) & " only occurs in column " & ColChr(column) &
                            " in box at " & BoxChr(N) &
                            " so this value can be deleted from other cells of the box"
                        Return True
                    End If
                    GetCellsInColumn(column, myCells)
                End If
            Next
        Next

        ' Look for condition 3
        For NBox As Integer = 0 To mSize - 1
            GetCellsInBoxNumber(NBox, myCells)
            For Value As Integer = mStartValue To mEndValue
                If SubGroupExclusion1(myCells, Value, ExcludingIndex) Then
                    Dim X As Integer = myCells(0).x
                    Dim Y As Integer = myCells(0).y + ExcludingIndex
                    GetCellsInRow(Y, myCells)
                    Dim Count As Integer = 0
                    For i As Integer = 0 To mSize - 1
                        With myCells(i)
                            If .box <> NBox Then
                                If .Possibility(Value) And .Solution = -1 Then Count += 1
                                If Not mShowHint Then RemovePossibility(.x, .y, Value)
                            End If
                        End With
                    Next
                    If Count > 0 Then
                        Text = Hex(Value) & " only occurs in row " & RowChr(Y) & " of box at " &
                            BoxChr(NBox) & " so this" &
                            " value can be deleted from cells in other boxes on " &
                            " this same row"
                        Return True
                    End If
                    GetCellsInBoxNumber(NBox, myCells)
                End If
            Next
        Next

        ' Look for condition 4
        Dim Flip As Boolean = True
        For NBox As Integer = 0 To mSize - 1
            GetCellsInBoxNumber(NBox, myCells, Flip)
            For Value As Integer = mStartValue To mEndValue
                If SubGroupExclusion1(myCells, Value, ExcludingIndex) Then
                    Dim X As Integer = myCells(0).BoxX + ExcludingIndex
                    Dim Y As Integer = myCells(0).y
                    GetCellsInColumn(X, myCells)
                    Dim Count As Integer = 0
                    For i As Integer = 0 To mSize - 1
                        With myCells(i)
                            If .box <> NBox Then
                                If .Possibility(Value) And .Solution = -1 Then Count += 1
                                If Not mShowHint Then RemovePossibility(.x, .y, Value)
                            End If
                        End With
                    Next
                    If Count > 0 Then
                        Text = Hex(Value) & " only occurs in column " & ColChr(X) & " of bax at " &
                            BoxChr(NBox) & " so this" &
                            " value can be deleted from cells in other boxes on" &
                            " this same column"
                        Return True
                    End If
                    GetCellsInBoxNumber(NBox, myCells, Flip)
                End If
            Next
        Next
        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function SubGroupExclusion1(ByVal CellSet() As CellButton,
                                        ByVal Value As Integer,
                                        ByRef ExcludingIndex As Integer) As Boolean
        ' Determine if SubGroupExclusion condition is present in a group of cells and
        ' a specific value. If True, return box number where excluding cells occur.

        Dim Count() As Integer = {0, 0, 0, 0}
        For i As Integer = 0 To mSize - 1
            If CellSet(i).Possibility(Value) And CellSet(i).Solution = -1 Then
                Dim j As Integer = CInt(i \ mBoxSize)
                Count(j) += 1
            End If
        Next

        Dim Filled As Integer = 0
        Dim myIndex As Integer = -1

        For i As Integer = 0 To mBoxSize - 1
            If Count(i) > 0 Then
                Filled += 1
                myIndex = i
                If Filled > 1 Then Return False
            End If
        Next
        If Filled = 0 Then Return False

        ExcludingIndex = myIndex
        Return True

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function NakedTwin(ByRef Text As String) As Boolean
        ' When two cells in a column, row, or box have only two possibilities and they are the
        ' same, then those two values can removed from the other cells in the corresponding
        ' column, row, or box.

        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1
                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                    Case 2
                        GetCellsInColumn(i, myCells)
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                End Select

                For j As Integer = 0 To mSize - 2
                    If myCells(j).PossibilityCount = 2 Then
                        Dim Value1 As Integer = -1
                        Dim Value2 As Integer = -1
                        GetOnlyTwoValues(myCells(j), Value1, Value2)
                        For k As Integer = j + 1 To mSize - 1
                            If myCells(k).PossibilityCount = 2 Then
                                Dim Value3 As Integer = -1
                                Dim Value4 As Integer = -1
                                GetOnlyTwoValues(myCells(k), Value3, Value4)
                                Dim Count As Integer = 0
                                If (Value1 = Value3) AndAlso (Value2 = Value4) _
                                AndAlso (Value2 <> -1) AndAlso (Value4 <> -1) Then
                                    For L As Integer = 0 To mSize - 1
                                        If (L <> j) _
                                        AndAlso (L <> k) _
                                        AndAlso myCells(L).Solution = -1 _
                                        AndAlso (myCells(L).Possibility(Value1) _
                                        OrElse myCells(L).Possibility(Value2)) Then
                                            Count += 1
                                            If Not mShowHint Then
                                                Dim X As Integer = myCells(L).x
                                                Dim Y As Integer = myCells(L).y
                                                RemovePossibility(X, Y, Value1)
                                                RemovePossibility(X, Y, Value2)
                                            End If
                                        End If
                                    Next L
                                    If Count > 0 Then
                                        Text = Hex(Value1) & " And " & Hex(Value2)
                                        Select Case Pass
                                            Case 1
                                                Text += " in row " & RowChr(i) &
                                                " only occur at " & ColChr(j) & " and " &
                                                ColChr(k) & " so these values can be " &
                                                "removed from other cells in the row"
                                            Case 2
                                                Text += " in column " & ColChr(i) &
                                                " only occur at " & RowChr(j) & " and " _
                                                & RowChr(k) & " so these values can be " &
                                                "removed from other cells in the column"
                                            Case 3
                                                Text += " in box at " & BoxChr(i) &
                                                " only occur at " &
                                                RowChr(myCells(j).y) & ColChr(myCells(j).x) &
                                                " and " &
                                                RowChr(myCells(k).y) & ColChr(myCells(k).x) &
                                                " so these values can be removed from other " &
                                                "cells in the box"
                                        End Select
                                        Return True
                                    End If
                                End If
                            End If
                        Next k
                    End If
                Next j
            Next i
        Next Pass

        Return False
    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function HiddenTwin(ByRef Text As String) As Boolean
        ' When two values in a column, row, or box only appear in two particular cells, then all
        ' possibilities for other values in the same two cells can be removed.

        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1
                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                    Case 2
                        GetCellsInColumn(i, myCells)
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                End Select

                For j As Integer = 0 To mSize - 2
                    If myCells(j).PossibilityCount = 2 Then
                        Dim Value1 As Integer = -1
                        Dim Value2 As Integer = -1
                        GetTWoValues(myCells(j), Value1, Value2)
                        For k As Integer = j + 1 To mSize - 1
                            If myCells(k).PossibilityCount = 2 Then
                                Dim Value3 As Integer = -1
                                Dim Value4 As Integer = -1
                                GetTWoValues(myCells(k), Value3, Value4)
                                Dim Count As Integer = 0
                                If (Value1 = Value3) AndAlso (Value2 = Value4) _
                                AndAlso (Value2 <> -1) AndAlso (Value4 <> -1) Then
                                    For Value As Integer = mStartValue To mEndValue
                                        If (Value <> Value1) AndAlso (Value <> Value2) Then
                                            If myCells(j).Possibility(Value) Then
                                                Count += 1
                                                If Not mShowHint Then
                                                    Dim X As Integer = myCells(j).x
                                                    Dim Y As Integer = myCells(j).y
                                                    RemovePossibility(X, Y, Value)
                                                End If
                                            End If
                                            If myCells(k).Possibility(Value) Then
                                                Count += 1
                                                If Not mShowHint Then
                                                    Dim X As Integer = myCells(k).x
                                                    Dim Y As Integer = myCells(k).y
                                                    RemovePossibility(X, Y, Value)
                                                End If
                                            End If
                                        End If
                                    Next
                                    If Count > 0 Then
                                        Text = Hex(Value1) & " and " & Hex(Value2)
                                        Select Case Pass
                                            Case 1
                                                Text += " in row " & RowChr(i) &
                                                " only occur at " & ColChr(j) & " and " &
                                                ColChr(k) & " so other values can be removed " &
                                                "from these two cells"
                                            Case 2
                                                Text += " in column " & ColChr(i) &
                                                " only occur at " & RowChr(j) & " and " &
                                                RowChr(k) & " so other values can be removed " &
                                                "from these two cells"
                                            Case 3
                                                Text += " in box at " & BoxChr(i) &
                                                "only occur at " &
                                                RowChr(myCells(j).y) & ColChr(myCells(j).x) &
                                                " and " &
                                                RowChr(myCells(k).y) & ColChr(myCells(k).x) &
                                                " so other values can be removed from " &
                                                "these two cells"
                                        End Select
                                        Return True
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        Next

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function NakedTriplet(ByRef Text As String) As Boolean
        ' When three cells in a column, row, or box are the only place where three particular
        ' possibilities exist and those cells have no other possibilities,  then those three
        ' values can removed from the other cells in the corresponding column, row, or box.
        ' Note: Some cells in the triplet can have only two possibilities.   If all three
        ' have just two possibilities then the pattern must be like "ab ac bc".
        '
        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton
        Dim Value31 As Integer = -1
        Dim Value32 As Integer = -1
        Dim Value33 As Integer = -1
        Dim Found3 As Boolean = False

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1

                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                    Case 2
                        GetCellsInColumn(i, myCells)
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                End Select

                ' Check all possible triplets
                ' For a particular triplet
                ' Check for a first possible candidate
                For index1 As Integer = 0 To mSize - 3
                    Found3 = False
                    Dim Value11 As Integer = -1
                    Dim Value12 As Integer = -1
                    Dim Value13 As Integer = -1
                    With myCells(index1)
                        ' It must be a cell with only 2 or 3 possibilities
                        If .PossibilityCount = 2 Or .PossibilityCount = 3 Then
                            ' A first candidate was found
                            GetOnlyTwoOrThreeValues(myCells(index1), Value11, Value12, Value13)
                        End If
                    End With
                    If Value12 <> -1 Then
                        ' Check for a second candidate
                        Dim FirstValue As Integer = Value11
                        Dim SecondValue As Integer = Value12
                        Dim ThirdValue As Integer = Value13
                        For index2 As Integer = index1 + 1 To mSize - 2
                            Dim Found2 As Boolean = False
                            Dim Value21 As Integer = -1
                            Dim Value22 As Integer = -1
                            Dim Value23 As Integer = -1
                            With myCells(index2)
                                If .PossibilityCount = 2 OrElse .PossibilityCount = 3 Then
                                    GetOnlyTwoOrThreeValues(myCells(index2), Value21, Value22, Value23)
                                    Found2 = TripletCheck(FirstValue, SecondValue, ThirdValue,
                                              Value21, Value22, Value23)
                                End If
                            End With
                            If Found2 Then
                                ' Check for third Candidate
                                Dim FirstValue3 As Integer = FirstValue
                                Dim SecondValue3 As Integer = SecondValue
                                Dim ThirdValue3 As Integer = ThirdValue
                                For index3 As Integer = index2 + 1 To mSize - 1
                                    Found3 = False
                                    FirstValue = FirstValue3
                                    SecondValue = SecondValue3
                                    ThirdValue = ThirdValue3
                                    With myCells(index3)
                                        If .PossibilityCount = 2 Or .PossibilityCount = 3 Then
                                            GetOnlyTwoOrThreeValues(myCells(index3), Value31, Value32, Value33)
                                            Found3 = TripletCheck(FirstValue, SecondValue, ThirdValue,
                                                                  Value31, Value32, Value33)
                                        End If
                                    End With
                                    If Found3 And ThirdValue <> -1 Then
                                        Dim Count As Integer = 0
                                        For N As Integer = 0 To mSize - 1
                                            If (N <> index1) AndAlso (N <> index2) AndAlso (N <> index3) Then
                                                With myCells(N)
                                                    If .Solution = -1 Then
                                                        If .Possibility(FirstValue) Then
                                                            Count += 1
                                                            If Not mShowHint Then RemovePossibility(.x, .y, FirstValue)
                                                        End If
                                                        If .Possibility(SecondValue) Then
                                                            Count += 1
                                                            If Not mShowHint Then RemovePossibility(.x, .y, SecondValue)
                                                        End If
                                                        If .Possibility(ThirdValue) Then
                                                            Count += 1
                                                            If Not mShowHint Then RemovePossibility(.x, .y, ThirdValue)
                                                        End If
                                                    End If
                                                End With
                                            End If
                                        Next
                                        If Count > 0 Then
                                            Text = "Triplet(" & Hex(FirstValue) & "," & Hex(SecondValue) & "," &
                                                    Hex(ThirdValue) & ") found in "
                                            Select Case Pass
                                                Case 1
                                                    Text += "row " & RowChr(i) & " so" &
                                                                " the values can be removed From other cells in the row"
                                                Case 2
                                                    Text += "column " & ColChr(i) & " so" &
                                                                " the values can be removed From other cells in the column"
                                                Case 3
                                                    Text += "box " & BoxChr(i) & " so" &
                                                                " the values can be removed From other cells in the box"
                                            End Select
                                            Return True
                                        End If
                                    End If
                                Next index3
                            End If
                        Next index2
                    End If
                Next index1

            Next i

        Next Pass

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function HiddenTriplet(ByRef Text As String) As Boolean
        ' When three values in a column, row, or box only appear in three particular cells,
        ' then all possibilities for other values in the same three cells can be removed.

        ' (If all three have just two possibilities then the pattern must be like "ab ac bc".)
        '

        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton
        Dim Value31 As Integer = -1
        Dim Value32 As Integer = -1
        Dim Value33 As Integer = -1
        Dim Found3 As Boolean = False

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1

                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                    Case 2
                        GetCellsInColumn(i, myCells)
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                End Select

                ' Check all possible triplets
                ' For a particular triplet
                ' Check for a first possible candidate
                For index1 As Integer = 0 To mSize - 3
                    Found3 = False
                    Dim Value11 As Integer = -1
                    Dim Value12 As Integer = -1
                    Dim Value13 As Integer = -1
                    With myCells(index1)
                        ' It must be a cell with only 2 or 3 possibilities
                        If .PossibilityCount = 2 Or .PossibilityCount = 3 Then
                            ' A first candidate was found
                            GetOnlyTwoOrThreeValues(myCells(index1), Value11, Value12, Value13)
                        End If
                    End With
                    If Value12 <> -1 Then
                        ' Check for a second candidate
                        Dim FirstValue As Integer = Value11
                        Dim SecondValue As Integer = Value12
                        Dim ThirdValue As Integer = Value13
                        For index2 As Integer = index1 + 1 To mSize - 2
                            Dim Found2 As Boolean = False
                            Dim Value21 As Integer = -1
                            Dim Value22 As Integer = -1
                            Dim Value23 As Integer = -1
                            With myCells(index2)
                                If .PossibilityCount = 2 OrElse .PossibilityCount = 3 Then
                                    GetOnlyTwoOrThreeValues(myCells(index2), Value21, Value22, Value23)
                                    Found2 = TripletCheck(FirstValue, SecondValue, ThirdValue,
                                              Value21, Value22, Value23)
                                End If
                            End With
                            If Found2 Then
                                ' Check for third Candidate
                                Dim FirstValue3 As Integer = FirstValue
                                Dim SecondValue3 As Integer = SecondValue
                                Dim ThirdValue3 As Integer = ThirdValue
                                For index3 As Integer = index2 + 1 To mSize - 1
                                    Found3 = False
                                    FirstValue = FirstValue3
                                    SecondValue = SecondValue3
                                    ThirdValue = ThirdValue3
                                    Found3 = False
                                    With myCells(index3)
                                        If .PossibilityCount = 2 Or .PossibilityCount = 3 Then
                                            GetOnlyTwoOrThreeValues(myCells(index3), Value31, Value32, Value33)
                                            Found3 = TripletCheck(FirstValue, SecondValue, ThirdValue,
                                                                  Value31, Value32, Value33)
                                        End If
                                    End With

                                    If Found3 And ThirdValue <> -1 Then
                                        Dim Count As Integer = 0
                                        For Value As Integer = mStartValue To mEndValue
                                            If (Value <> FirstValue) AndAlso (Value <> SecondValue) AndAlso (Value <> ThirdValue) Then
                                                If myCells(index1).Possibility(Value) Then
                                                    Count += 1
                                                    If Not mShowHint Then
                                                        Dim X As Integer = myCells(index1).x
                                                        Dim Y As Integer = myCells(index1).y
                                                        RemovePossibility(X, Y, Value)
                                                    End If
                                                End If
                                                If myCells(index2).Possibility(Value) Then
                                                    Count += 1
                                                    If Not mShowHint Then
                                                        Dim X As Integer = myCells(index2).x
                                                        Dim Y As Integer = myCells(index2).y
                                                        RemovePossibility(X, Y, Value)
                                                    End If
                                                End If
                                                If myCells(index3).Possibility(Value) Then
                                                    Count += 1
                                                    If Not mShowHint Then
                                                        Dim X As Integer = myCells(index3).x
                                                        Dim Y As Integer = myCells(index3).y
                                                        RemovePossibility(X, Y, Value)
                                                    End If
                                                End If
                                            End If
                                        Next
                                        If Count > 0 Then
                                            Text = "Triplet(" & Hex(FirstValue) & "," & Hex(SecondValue) & "," &
                                                        Hex(ThirdValue) & ") found in "
                                            Select Case Pass
                                                Case 1
                                                    Text += "row " & RowChr(i) & " so" &
                                                " other values can be removed from this cell"
                                                Case 2
                                                    Text += "column " & ColChr(i) & " so" &
                                                " other values can be removed from this cell"
                                                Case 3
                                                    Text += "box " & BoxChr(i) & " so" &
                                                " other values can be removed from this cell"
                                            End Select
                                            Return True
                                        End If
                                    End If
                                Next index3
                            End If
                        Next index2
                    End If
                Next index1

            Next i

        Next Pass

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function TripletCheck(ByRef FirstValue As Integer, ByRef SecondValue As Integer,
                                  ByRef ThirdValue As Integer, ByVal Value1 As Integer,
                                  ByVal Value2 As Integer, ByVal Value3 As Integer) As Boolean
        ' See if the triplet (Value1,Value2,Value3) is compatible with the triplet
        ' (FirstValue, SecondValue, ThirdValue) and if so, update the second triplet.
        ' Note: ThirdValue and Value3 may be -1, indicating not specified.


        If ThirdValue = -1 And Value3 = -1 Then
            If FirstValue = Value1 Then
                If SecondValue = Value2 Then
                    Return True
                ElseIf SecondValue < Value2 Then
                    ThirdValue = Value2
                    Return True
                Else
                    ThirdValue = SecondValue
                    SecondValue = Value2
                    Return True
                End If
            ElseIf FirstValue = Value2 Then
                ThirdValue = SecondValue
                SecondValue = FirstValue
                FirstValue = Value1
                Return True
            ElseIf SecondValue = Value1 Then
                ThirdValue = Value2
                Return True
            ElseIf SecondValue = Value2 Then
                If FirstValue > Value1 Then
                    ThirdValue = SecondValue
                    SecondValue = FirstValue
                    FirstValue = Value1
                    Return True
                Else
                    ThirdValue = SecondValue
                    SecondValue = Value1
                    Return True
                End If
            Else
                Return False
            End If
        ElseIf ThirdValue = -1 And Value3 <> -1 Then
            If FirstValue = Value1 Then
                If SecondValue = Value2 Then
                    ThirdValue = Value3
                    Return True
                ElseIf SecondValue = Value3 Then
                    ThirdValue = SecondValue
                    SecondValue = Value3
                    Return True
                Else
                    Return False
                End If
            Else
                If (FirstValue = Value2) AndAlso (SecondValue = Value3) Then
                    ThirdValue = SecondValue
                    SecondValue = FirstValue
                    FirstValue = Value1
                    Return True
                Else
                    Return False
                End If
            End If
        ElseIf ThirdValue <> -1 And Value3 = -1 Then
            If FirstValue = Value1 Then
                If (Value2 = SecondValue) Or (Value2 = ThirdValue) Then
                    Return True
                Else
                    Return False
                End If
            ElseIf (SecondValue = Value1) And (ThirdValue = Value2) Then
                Return True
            Else
                Return False
            End If
        Else
            If (FirstValue <> Value1) OrElse (SecondValue <> Value2) OrElse (ThirdValue <> Value3) Then
                Return False
            Else
                Return True
            End If
        End If

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function NakedQuad(ByRef Text As String) As Boolean
        ' When four cells in a column, row, or box are the only place where four particular
        ' possibilities exist and those cells have no other possibilities,  then those four
        ' values can removed from the other cells in the corresponding column, row, or box.
        ' Note: Some cells in the quad may have only two or three possibilities. If all four
        ' have just three possibilities, then the pattern must be like "abc abd acd bcd".
        '
        ' A common array is used to hold the cells for either a single row, column, or box,
        ' depending on which is being checked:
        Dim myCells(mSize - 1) As CellButton

        Dim index1 As Integer = -1
        Dim index2 As Integer = -1
        Dim index3 As Integer = -1
        Dim Index4 As Integer = -1
        Dim Found4 As Boolean = False

        For Pass As Integer = 1 To 3
            ' Pass 1     Do Rows
            ' Pass 2     Do Columns
            ' Pass 3     Do Boxs

            For i As Integer = 0 To mSize - 1

                Select Case Pass
                    Case 1
                        GetCellsInRow(i, myCells)
                    Case 2
                        GetCellsInColumn(i, myCells)
                    Case 3
                        GetCellsInBoxNumber(i, myCells)
                End Select

                ' Check all possible quads
                ' For a particular quad
                ' Check for a first possible candidate
                For index1 = 0 To mSize - 4
                    Dim Value11 As Integer = -1
                    Dim Value12 As Integer = -1
                    Dim Value13 As Integer = -1
                    Dim Value14 As Integer = -1
                    ' Get first candidate
                    GetOnlyTwoOrThreeOrFourValues(myCells(index1), Value11, Value12, Value13, Value14)
                    If Value12 <> -1 Then   ' Valid candidate must have at least two values
                        ' Get second candidate
                        For index2 = index1 + 1 To mSize - 3
                            Dim Found2 As Boolean = False
                            Dim FirstValue As Integer = Value11
                            Dim SecondValue As Integer = Value12
                            Dim ThirdValue As Integer = Value13
                            Dim FourthValue As Integer = Value14
                            Dim Value21 As Integer = -1
                            Dim Value22 As Integer = -1
                            Dim Value23 As Integer = -1
                            Dim Value24 As Integer = -1
                            With myCells(index2)
                                If .PossibilityCount >= 2 AndAlso .PossibilityCount <= 4 Then
                                    GetOnlyTwoOrThreeOrFourValues(myCells(index2), Value21, Value22, Value23, Value24)
                                    Found2 = QuadCheck(FirstValue, SecondValue, ThirdValue, FourthValue,
                                              Value21, Value22, Value23, Value24)
                                End If
                            End With
                            If Found2 Then
                                ' Get third candidate
                                Dim FirstValue3 As Integer = FirstValue
                                Dim SecondValue3 As Integer = SecondValue
                                Dim ThirdValue3 As Integer = ThirdValue
                                Dim FourthValue3 As Integer = FourthValue
                                For index3 = index2 + 1 To mSize - 2
                                    Dim Found3 As Boolean = False
                                    Dim Value31 As Integer = -1
                                    Dim Value32 As Integer = -1
                                    Dim Value33 As Integer = -1
                                    Dim Value34 As Integer = -1
                                    FirstValue = FirstValue3
                                    SecondValue = SecondValue3
                                    ThirdValue = ThirdValue3
                                    FourthValue = FourthValue3
                                    With myCells(index3)
                                        If .PossibilityCount >= 2 AndAlso .PossibilityCount <= 4 Then
                                            GetOnlyTwoOrThreeOrFourValues(myCells(index3), Value31, Value32, Value33, Value34)
                                            Found3 = QuadCheck(FirstValue, SecondValue, ThirdValue, FourthValue,
                                                           Value31, Value32, Value33, Value34)
                                        End If
                                    End With
                                    If Found3 Then
                                        ' Get fourth candidate
                                        Dim FirstValue4 As Integer = FirstValue
                                        Dim SecondValue4 As Integer = SecondValue
                                        Dim ThirdValue4 As Integer = ThirdValue
                                        Dim FourthValue4 As Integer = FourthValue
                                        For Index4 = index3 + 1 To mSize - 1
                                            Found4 = False
                                            Dim Value41 As Integer = -1
                                            Dim Value42 As Integer = -1
                                            Dim Value43 As Integer = -1
                                            Dim Value44 As Integer = -1
                                            FirstValue = FirstValue4
                                            SecondValue = SecondValue4
                                            ThirdValue = ThirdValue4
                                            FourthValue = FourthValue4
                                            With myCells(Index4)
                                                If .PossibilityCount >= 2 AndAlso .PossibilityCount <= 4 Then
                                                    GetOnlyTwoOrThreeOrFourValues(myCells(Index4), Value41, Value42, Value43, Value44)
                                                    Found4 = QuadCheck(FirstValue, SecondValue, ThirdValue, FourthValue,
                                                           Value41, Value42, Value43, Value44)
                                                End If
                                            End With
                                            If Found4 And FourthValue <> -1 Then
                                                Dim Count As Integer = 0
                                                For N As Integer = 0 To mSize - 1
                                                    If (N <> index1) AndAlso (N <> index2) AndAlso (N <> index3) AndAlso (N <> Index4) Then
                                                        With myCells(N)
                                                            If .Solution = -1 Then
                                                                If .Possibility(FirstValue) Then
                                                                    Count += 1
                                                                    If Not mShowHint Then RemovePossibility(.x, .y, FirstValue)
                                                                End If
                                                                If .Possibility(SecondValue) Then
                                                                    Count += 1
                                                                    If Not mShowHint Then RemovePossibility(.x, .y, SecondValue)
                                                                End If
                                                                If .Possibility(ThirdValue) Then
                                                                    Count += 1
                                                                    If Not mShowHint Then RemovePossibility(.x, .y, ThirdValue)
                                                                End If
                                                                If .Possibility(FourthValue) Then
                                                                    Count += 1
                                                                    If Not mShowHint Then RemovePossibility(.x, .y, FourthValue)
                                                                End If
                                                            End If
                                                        End With
                                                    End If
                                                Next
                                                If Count > 0 Then
                                                    Text = "Quad(" & Hex(FirstValue) & "," & Hex(SecondValue) & "," &
                                                        Hex(ThirdValue) & "," & Hex(FourthValue) & ") found in "
                                                    Select Case Pass
                                                        Case 1
                                                            Text += "row " & RowChr(i) & " so" &
                                                                    " the values can be removed From other cells in the row"
                                                        Case 2
                                                            Text += "column " & ColChr(i) & " so" &
                                                                    " the values can be removed From other cells in the column"
                                                        Case 3
                                                            Text += "box " & BoxChr(i) & " so" &
                                                                    " the values can be removed From other cells in the box"
                                                    End Select
                                                    Return True
                                                End If
                                            End If
                                        Next Index4
                                    End If
                                Next index3
                            End If
                        Next index2
                    End If
                Next index1

            Next i

        Next Pass

        Return False

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function QuadCheck(ByRef FirstValue As Integer, ByRef SecondValue As Integer,
                                  ByRef ThirdValue As Integer, ByRef FourthValue As Integer,
                                  ByVal Value1 As Integer, ByVal Value2 As Integer,
                                  ByVal Value3 As Integer, ByVal Value4 As Integer) As Boolean

        ' Definition: a defined value is one that is not -1.
        ' If Value1 is not found in the output quad (FirstValue, SecondValue, ThirdValue, FourthValue)
        ' and there are less than four defined values in the output quad, then add value1 to the
        ' output quad.  If there are already four defined values in the output quad, then value1
        ' cannot be added, so return false.  Repeat the operation For Value2, then value3, And then
        ' Value4. At this point the values in the output quad may be out of numeric order, so they
        ' need to be sorted into correct order, then return true.
        ' Note 1:  The output quad may contain 2, 3, or 4 defined values on exit.
        ' Note 2:  This routine uses a very different and simpler way of processing than TripletCheck.

        ' Define a working quad initially set to the output quad on entry
        Dim OutputValue1 As Integer = FirstValue
        Dim OutputValue2 As Integer = SecondValue
        Dim OutputValue3 As Integer = ThirdValue
        Dim OutputValue4 As Integer = FourthValue

        ' Determine how many defined values are in the output quad
        Dim NDefined As Integer = 0
        If FirstValue <> -1 Then NDefined += 1
        If SecondValue <> -1 Then NDefined += 1
        If ThirdValue <> -1 Then NDefined += 1
        If FourthValue <> -1 Then NDefined += 1

        ' Loop over all four values
        Dim Value As Integer = 0
        For i As Integer = 1 To 4
            Select Case i
                Case 1
                    Value = Value1
                Case 2
                    Value = Value2
                Case 3
                    Value = Value3
                Case 4
                    Value = Value4
            End Select

            ' Process the value.  If it is not defined, skip it
            If Value <> -1 Then
                If (Value <> OutputValue1) And (Value <> OutputValue2) And
                   (Value <> OutputValue2) And (Value <> OutputValue4) Then
                    If NDefined = 4 Then
                        Return False
                    Else
                        If OutputValue1 = -1 Then
                            OutputValue1 = Value
                        ElseIf OutputValue2 = -1 Then
                            OutputValue2 = Value
                        ElseIf OutputValue3 = -1 Then
                            OutputValue3 = Value
                        Else
                            OutputValue4 = Value
                        End If
                        NDefined += 1
                    End If
                End If
            End If
        Next

        ' sort the output using brute force
        Dim Temp As Integer
        If OutputValue1 > OutputValue2 Then
            Temp = OutputValue1
            OutputValue1 = OutputValue2
            OutputValue2 = Temp
        End If
        If OutputValue1 > OutputValue3 Then
            Temp = OutputValue1
            OutputValue1 = OutputValue3
            OutputValue3 = Temp
        End If
        If OutputValue1 > OutputValue4 Then
            Temp = OutputValue1
            OutputValue1 = OutputValue4
            OutputValue4 = Temp
        End If
        If OutputValue2 > OutputValue3 Then
            Temp = OutputValue2
            OutputValue2 = OutputValue3
            OutputValue3 = Temp
        End If
        If OutputValue2 > OutputValue4 Then
            Temp = OutputValue2
            OutputValue2 = OutputValue4
            OutputValue4 = Temp
        End If
        If OutputValue3 > OutputValue4 Then
            Temp = OutputValue3
            OutputValue3 = OutputValue4
            OutputValue4 = Temp
        End If

        ' Send the modified quad back to the caller
        FirstValue = OutputValue1
        SecondValue = OutputValue2
        ThirdValue = OutputValue3
        FourthValue = OutputValue4
        Return True

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Function XWing(ByRef Text As String) As Boolean
        ' If two rows have a value that occurs in only in two columns and the columns match in the
        ' two rows, then an X-Wing pattern is formed and that value must appear only on one of the
        ' diagonals of the X.   That means that value can be excluded from any cells in the two
        ' columns that are not on the X.
        ' The same applies to two columns that have a value that occurs only in two rows and the
        ' rows match, forming an X-wing pattern.   In this case the value can be removed from cells
        ' in the two rows that are not on the X.

        Dim Possibility(mSize - 1) As Boolean
        Dim myCells(mSize - 1) As CellButton

        For Pass As Integer = 1 To 2
            ' Pass 1    Rows
            ' Pass 2    Columns
            For value As Integer = mStartValue To mEndValue
                ' loop over all possibile values
                ' Initialize possibilites
                For Row As Integer = 0 To mSize - 1
                    Possibility(Row) = False
                Next
                ' Look for first row or column with a value in just two places.
                ' (First one found cannot be the very last one.)
                Dim Index1 As Integer = -1
                Dim Index2 As Integer = -1
                Dim N1 As Integer = -1
                Dim N2 As Integer = -1
                Dim Count As Integer
                For i As Integer = 0 To mSize - 2
                    Count = 0
                    Select Case Pass
                        Case 1
                            GetCellsInRow(i, myCells)
                        Case 2
                            GetCellsInColumn(i, myCells)
                    End Select
                    ' Keep track of the two locations
                    ' See if the value only occurs twice in the cells
                    For j As Integer = 0 To mSize - 1
                        If myCells(j).Possibility(value) And myCells(j).PossibilityCount > 1 Then
                            Select Case Count
                                Case 0
                                    N1 = j
                                Case 1
                                    N2 = j
                                    Index1 = i
                            End Select
                            Count += 1
                        End If
                        If Count > 2 Then Exit For
                    Next
                    If Count = 2 Then
                        ' If a first possible row or column found, look for a second,
                        ' which must be after the first.
                        For j As Integer = Index1 + 1 To mSize - 1
                            Count = 0
                            Select Case Pass
                                Case 1
                                    GetCellsInRow(j, myCells)
                                Case 2
                                    GetCellsInColumn(j, myCells)
                            End Select
                            For k As Integer = 0 To mSize - 1
                                If myCells(k).Possibility(value) And myCells(k).PossibilityCount > 1 Then
                                    Select Case Count
                                        Case 0
                                            If k <> N1 Then Exit For
                                        Case 1
                                            If k <> N2 Then Exit For
                                            Index2 = j
                                    End Select
                                    Count += 1
                                End If
                                If Count > 2 Then Exit For
                            Next
                            If Count = 2 Then
                                Dim Removed As Integer = 0
                                ' We have two rows or columns that match, so apply rule
                                For k As Integer = 1 To 2
                                    Dim N As Integer
                                    Select Case k
                                        Case 1
                                            N = N1
                                        Case 2
                                            N = N2
                                    End Select
                                    Select Case Pass
                                        Case 1
                                            GetCellsInColumn(N, myCells)
                                        Case 2
                                            GetCellsInRow(N, myCells)
                                    End Select
                                    For m As Integer = 0 To mSize - 1
                                        If m <> Index1 And m <> Index2 Then
                                            Dim X As Integer = myCells(m).x
                                            Dim Y As Integer = myCells(m).y
                                            If myCells(m).Possibility(value) = True Then Removed += 1
                                            If Not mShowHint Then RemovePossibility(X, Y, value)
                                        End If
                                    Next
                                Next
                                If Removed > 0 Then
                                    Text = "X-Wing pattern value " & Hex$(value) & " for "
                                    Select Case Pass
                                        Case 1
                                            Text += " rows " & RowChr(Index1) & " and " & RowChr(Index2) & " at columns " & ColChr(N1) & " and " & ColChr(N2) & "."
                                            Text += vbCrLf & "Values in columns not on X will be removed."
                                        Case 2
                                            Text += " columns " & ColChr(Index1) & " and " & ColChr(Index2) & " at rows " & RowChr(N1) & " and " & RowChr(N2) & "."
                                            Text += vbCrLf & "Values in rows not on X will be removed."
                                    End Select
                                    Return True
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        Next
        Return False
    End Function

#End Region
    '===============================================================================================
#Region "Cell Functions"
    '-----------------------------------------------------------------------------------------------
    Private Function BoxChr(ByVal j As Integer) As String
        Dim col As Integer = (j Mod mBoxSize) * mBoxSize
        Dim row As Integer = (j \ mBoxSize) * mBoxSize
        Return RowChr(row) & ColChr(col)
    End Function

    '-----------------------------------------------------------------------------------------------
    Private Sub CopyCells(ByVal SourceCells(,) As CellButton, ByRef DestinationCells(,) As CellButton)
        For column As Integer = 0 To mSize - 1
            For row As Integer = 0 To mSize - 1
                DestinationCells(column, row) = SourceCells(column, row)
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub GetCellsInBox(ByVal BoxX As Integer, ByVal BoxY As Integer, ByRef BoxCells() As CellButton)
        Dim i As Integer = 0
        For Row As Integer = BoxY To BoxY + mBoxSize - 1
            For column As Integer = BoxX To BoxX + mBoxSize - 1
                BoxCells(i) = Cell(column, Row)
                i += 1
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub GetCellsInBoxFlipped(ByVal BoxX As Integer, ByVal BoxY As Integer, ByRef BoxCells() As CellButton)
        ' Flip columns and rows in box
        Dim i As Integer = 0
        For column As Integer = BoxX To BoxX + mBoxSize - 1
            For Row As Integer = BoxY To BoxY + mBoxSize - 1
                BoxCells(i) = Cell(column, Row)
                i += 1
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub GetCellsInBoxNumber(ByVal BoxNumber As Integer, ByRef BoxCells() As CellButton, Optional Flip As Boolean = False)
        Dim BoxY As Integer = mBoxSize * CInt(BoxNumber \ mBoxSize)
        Dim BoxX As Integer = mBoxSize * (BoxNumber Mod mBoxSize)
        If Flip Then
            GetCellsInBoxFlipped(BoxX, BoxY, BoxCells)
        Else
            GetCellsInBox(BoxX, BoxY, BoxCells)
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub GetCellsInColumn(ByVal Column As Integer, ByRef ColumnCells() As CellButton)
        For row As Integer = 0 To mSize - 1
            ColumnCells(row) = Cell(Column, row)
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub GetCellsInRow(ByVal Row As Integer, ByRef RowCells() As CellButton)
        For Column As Integer = 0 To mSize - 1
            RowCells(Column) = Cell(Column, Row)
        Next
    End Sub

#End Region
    '===============================================================================================
#Region "Possibility Functions"

    '-----------------------------------------------------------------------------------------------
    Private Sub InitializePossibilities()
        For column As Integer = 0 To mSize - 1
            For row As Integer = 0 To mSize - 1
                ' For each cell
                Dim Count As Integer = 0
                With Cell(column, row)
                    For Value As Integer = mStartValue To mEndValue
                        ' for each value
                        If .Solution = -1 Then
                            ' If no solution yet, assume all values are possible
                            .Possibility(Value) = True
                            Count += 1
                        Else
                            ' Otherwise only solution is possible
                            If .Solution = Value Then
                                .Possibility(Value) = True
                                Count = 1
                            Else
                                .Possibility(Value) = False
                            End If
                        End If
                    Next
                    .PossibilityCount = Count
                    If .ForeColor = Color.Red Then .ForeColor = Color.Blue
                End With
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub RemovePossibility(ByVal Column As Integer,
                                  ByVal row As Integer,
                                  ByVal Value As Integer)
        With Cell(Column, row)
            If .Possibility(Value) = True AndAlso .PossibilityCount > 1 Then
                .Possibility(Value) = False
                .PossibilityCount -= 1
                If .PossibilityCount = 1 Then
                    For answer As Integer = mStartValue To mEndValue
                        If .Possibility(answer) Then
                            .PendingSolution = answer
                            Exit For
                        End If
                    Next
                    .OriginalSolution = -1
                End If
            End If
        End With
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub RemovePossibilityFromCell(ByRef Cell As CellButton, ByVal Value As Integer)
        With Cell
            If .Possibility(Value) = True AndAlso .PossibilityCount > 1 Then
                .Possibility(Value) = False
                .PossibilityCount -= 1
                If .PossibilityCount = 1 Then
                    For answer As Integer = mStartValue To mEndValue
                        If .Possibility(answer) Then
                            .PendingSolution = answer
                            Exit For
                        End If
                    Next
                    .OriginalSolution = -1
                End If
            End If
        End With
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub RestorePossibility(ByVal Column As Integer,
                                   ByVal Row As Integer,
                                   ByVal Value As Integer)
        If Value = -1 Then Return
        With Cell(Column, Row)
            If .Possibility(Value) = False Then
                .Possibility(Value) = True
                .PossibilityCount += 1
                .PendingSolution = -1
                .Solution = -1
                .Guess = False
            End If
            For myvalue As Integer = mStartValue To mEndValue
                .Possibility(myvalue) = True
            Next
            UpdatePossibilities()
        End With
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub UpdatePossibilitiesForColumn(ByVal Column As Integer)
        For row As Integer = 0 To mSize - 1
            Dim Solution As Integer = Cell(Column, row).Solution
            If Solution >= 0 Then
                For otherRow As Integer = 0 To mSize - 1
                    Dim otherSolution As Integer = Cell(Column, otherRow).Solution
                    If (otherRow <> row) AndAlso (otherSolution = -1) Then
                        RemovePossibility(Column, otherRow, Solution)
                    End If
                Next
            End If
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub UpdatePossibilitiesForRow(ByVal row As Integer)
        For column As Integer = 0 To mSize - 1
            Dim Solution As Integer = Cell(column, row).Solution
            If Solution >= 0 Then
                For otherColumn As Integer = 0 To mSize - 1
                    Dim otherSolution As Integer = Cell(otherColumn, row).Solution
                    If (otherColumn <> column) And (otherSolution = -1) Then
                        RemovePossibility(otherColumn, row, Solution)
                    End If
                Next
            End If
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub UpdatePossibilitiesForBox(ByVal boxX As Integer, ByVal boxY As Integer)
        For Column As Integer = boxX To boxX + mBoxSize - 1
            For row As Integer = boxY To boxY + mBoxSize - 1
                Dim Solution As Integer = Cell(Column, row).Solution
                If Solution >= 0 Then
                    For otherColumn As Integer = boxX To boxX + mBoxSize - 1
                        For otherRow As Integer = boxY To boxY + mBoxSize - 1
                            Dim otherSolution As Integer = Cell(otherColumn, otherRow).Solution
                            If ((otherColumn <> Column) OrElse (otherRow <> row)) AndAlso (otherSolution = -1) Then
                                RemovePossibility(otherColumn, otherRow, Solution)
                            End If
                        Next
                    Next
                End If
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub UpdatePossibilities()
        For column As Integer = 0 To mSize - 1
            For row As Integer = 0 To mSize - 1
                UpdatePossibilitiesForColumn(column)
                UpdatePossibilitiesForRow(row)
                UpdatePossibilitiesForBox(Cell(column, row).BoxX, Cell(column, row).BoxY)
            Next
        Next
    End Sub

#End Region
    '===============================================================================================
#Region "Support Routines"

    '-----------------------------------------------------------------------------------------------
    Private Function RemainingCellsToSolve() As Integer
        Dim Count As Integer = 0
        For column As Integer = 0 To mSize - 1
            For row As Integer = 0 To mSize - 1
                If Cell(column, row).Solution = -1 Then Count += 1
            Next
        Next
        Return Count
    End Function

    '-----------------------------------------------------------------------------------------------
    Private Sub FindACellSolution()

        Dim column As Integer
        Dim row As Integer
        Dim value As Integer
        Dim Text As String = ""
        Me.Label2.Text = ""
        ResetRedCellsToBlue()

        If OnlyChoice(column, row, value, Text) Then
            SetCellSolution(column, row, value)
        ElseIf SinglePossibility(column, row, value, Text) Then
            SetCellSolution(column, row, value)
        ElseIf OnlySquare(column, row, value, Text) Then
            SetCellSolution(column, row, value)
        ElseIf SubGroupExclusion(Text) Then
            ' This rule only removes possibilities
        ElseIf NakedTwin(Text) Then
            ' This rule only removes possibilities
        ElseIf HiddenTwin(Text) Then
            ' This rule only removes possibilities
        ElseIf NakedTriplet(Text) Then
            ' This rule only removes possibilities
        ElseIf HiddenTriplet(Text) Then
            ' This rule only removes possibilities
        ElseIf NakedQuad(Text) Then
            ' This rule only removes possibilities
        ElseIf XWing(Text) Then
            ' This rule only removes possibilities
        ElseIf RemainingCellsToSolve() = 0 Then
            Text = "The puzzle has been solved!"
            Done = True
        Else
            Text = "This is as far as we can go.  No more rules to apply"
            Done = True
        End If
        Me.Refresh()
        Me.Label1.Text = Text
        ValidateBoard()

    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub SetCellSolution(ByVal column As Integer, ByVal row As Integer, ByVal Value As Integer)
        With Cell(column, row)
            If mShowHint Then
                .Hint = Value
            Else
                .Solution = Value
                .ForeColor = Color.Red
                UpdatePossibilitiesForRow(row)
                UpdatePossibilitiesForColumn(column)
                UpdatePossibilitiesForBox(Cell(column, row).BoxX, Cell(column, row).BoxY)
            End If
        End With

        Me.Refresh()
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Function ValidateBoard() As Boolean
        ' Check columns, rows and boxes for valid values

        ' Check columns
        For column As Integer = 0 To mSize - 1
            ' Initially assume no values are present
            For value As Integer = mStartValue To mEndValue
                Dim FoundIt As Boolean = False
                For row As Integer = 0 To mSize - 1
                    If Cell(column, row).Solution = value Then
                        If FoundIt Then
                            Me.Label2.Text = "Column " & column & " has multiple entries of value " & value
                            Return False
                        Else
                            FoundIt = True
                        End If
                    Else
                        If Cell(column, row).Possibility(value) Then FoundIt = True
                    End If
                    If FoundIt Then Exit For
                Next
                If Not FoundIt Then
                    Me.Label2.Text = "Column " & column & " does not have the value " & value
                    Return False
                End If
            Next
        Next

        ' Check rows
        For row As Integer = 0 To mSize - 1
            ' Initially assume no values are present
            For value As Integer = mStartValue To mEndValue
                Dim FoundIt As Boolean = False
                For column As Integer = 0 To mSize - 1
                    If Cell(column, row).Solution = value Then
                        If FoundIt Then
                            Me.Label2.Text = "Row " & column & " has multiple entries of value " & value
                            Return False
                        Else
                            FoundIt = True
                        End If
                    Else
                        If Cell(column, row).Possibility(value) Then FoundIt = True
                    End If
                    If FoundIt Then Exit For
                Next
                If Not FoundIt Then
                    Me.Label2.Text = "Row " & row & " does not have the value " & value
                    Return False
                End If
            Next
        Next

        ' Check boxes
        For BoxX As Integer = 0 To mSize - 1 Step mBoxSize
            For BoxY As Integer = 0 To mSize - 1 Step mBoxSize
                ' Initially assume no values are present
                For value As Integer = mStartValue To mEndValue
                    Dim FoundIt As Boolean = False
                    For column As Integer = BoxX To BoxX + mBoxSize - 1
                        For row As Integer = BoxY To BoxY + mBoxSize - 1
                            If Cell(column, row).Solution = value Then
                                If FoundIt Then
                                    Me.Label2.Text = "Cell at (" & BoxX & "," & BoxY & ")" & " has multiple entries of value " & value
                                    Return False
                                Else
                                    FoundIt = True
                                End If
                            Else
                                If Cell(column, row).Possibility(value) Then FoundIt = True
                            End If
                            If FoundIt Then Exit For
                        Next
                        If FoundIt Then Exit For
                    Next
                    If Not FoundIt Then
                        Me.Label2.Text = "Cell at (" & BoxX & "," & BoxY & ")" & " does not have the value " & value
                        Return False
                    End If
                Next
            Next
        Next

        Return True

    End Function

    '-----------------------------------------------------------------------------------------------
    Private Sub GetOnlyTwoValues(ByVal myCell As CellButton, ByRef Value1 As Integer, ByRef Value2 As Integer)
        GetTWoValues(myCell, Value1, Value2)
        If myCell.PossibilityCount = 2 Then Return
        Value1 = -1
        Value2 = -1
    End Sub

    Private Sub GetTWoValues(ByVal myCells As CellButton, ByRef Value1 As Integer, ByRef Value2 As Integer)
        Value1 = -1
        Value2 = -1
        For value As Integer = mStartValue To mEndValue
            If myCells.Possibility(value) Then
                If Value1 = -1 Then
                    Value1 = value
                Else
                    Value2 = value
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub GetOnlyTwoOrThreeValues(ByVal myCell As CellButton, ByRef Value1 As Integer,
                                        ByRef Value2 As Integer, ByRef Value3 As Integer)
        GetTWoOrThreeValues(myCell, Value1, Value2, Value3)
        If myCell.PossibilityCount = 2 Or myCell.PossibilityCount = 3 Then Return
        Value1 = -1
        Value2 = -1
        Value3 = -1
    End Sub

    Private Sub GetTWoOrThreeValues(ByVal myCells As CellButton, ByRef Value1 As Integer,
                                    ByRef Value2 As Integer, ByRef Value3 As Integer)
        Value1 = -1
        Value2 = -1
        Value3 = -1
        For value As Integer = mStartValue To mEndValue
            If myCells.Possibility(value) Then
                If Value1 = -1 Then
                    Value1 = value
                ElseIf Value2 = -1 Then
                    Value2 = value
                Else
                    Value3 = value
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub GetOnlyTwoOrThreeOrFourValues(ByVal myCell As CellButton, ByRef Value1 As Integer,
                                        ByRef Value2 As Integer, ByRef Value3 As Integer,
                                         ByRef value4 As Integer)
        GetTwoOrThreeOrFourValues(myCell, Value1, Value2, Value3, value4)
        If myCell.PossibilityCount = 2 OrElse myCell.PossibilityCount = 3 OrElse myCell.PossibilityCount = 4 Then Return
        Value1 = -1
        Value2 = -1
        Value3 = -1
        value4 = -1
    End Sub

    Private Sub GetTwoOrThreeOrFourValues(ByVal myCell As CellButton, ByRef Value1 As Integer,
                                    ByRef Value2 As Integer, ByRef Value3 As Integer,
                                    ByRef Value4 As Integer)
        Value1 = -1
        Value2 = -1
        Value3 = -1
        Value4 = -1
        For value As Integer = mStartValue To mEndValue
            If myCell.Possibility(value) Then
                If Value1 = -1 Then
                    Value1 = value
                ElseIf Value2 = -1 Then
                    If value <> Value1 Then Value2 = value
                ElseIf Value3 = -1 Then
                    If value <> Value1 And value <> Value2 Then Value3 = value
                Else
                    If value <> Value1 And value <> Value2 And value <> Value3 Then
                        Value4 = value
                        Exit For
                    End If
                End If
            End If
        Next
    End Sub
    '-----------------------------------------------------------------------------------------------
    Private Sub ResetRedCellsToBlue()
        For column As Integer = 0 To mSize - 1
            For Row As Integer = 0 To mSize - 1
                If Cell(column, Row).ForeColor = Color.Red Then Cell(column, Row).ForeColor = Color.Blue
            Next
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------
    Private Sub SwitchColors(ByVal Color1 As Color, ByVal Color2 As Color)
        For column As Integer = 0 To mSize - 1
            For Row As Integer = 0 To mSize - 1
                With Cell(column, Row)
                    If .ForeColor = Color1 Then .ForeColor = Color2
                End With
            Next
        Next
    End Sub

#End Region

End Class
