<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.mnuNew9x9 = New System.Windows.Forms.MenuItem()
        Me.MenuNew16x16 = New System.Windows.Forms.MenuItem()
        Me.mnuLoad = New System.Windows.Forms.MenuItem()
        Me.mnuSave = New System.Windows.Forms.MenuItem()
        Me.mnuPrint = New System.Windows.Forms.MenuItem()
        Me.mnuQuit = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuAbout = New System.Windows.Forms.MenuItem()
        Me.cbShowPossibilities = New System.Windows.Forms.CheckBox()
        Me.btnSolve = New System.Windows.Forms.Button()
        Me.cbSingleStep = New System.Windows.Forms.CheckBox()
        Me.btnShowHint = New System.Windows.Forms.Button()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.Cornsilk
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(406, 266)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(293, 130)
        Me.Panel1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(8, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(282, 52)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(279, 58)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Label1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MenuItem1})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNew9x9, Me.MenuNew16x16, Me.mnuLoad, Me.mnuSave, Me.mnuPrint, Me.mnuQuit})
        Me.MenuItem4.Text = "File"
        '
        'mnuNew9x9
        '
        Me.mnuNew9x9.Index = 0
        Me.mnuNew9x9.Text = "New 9 x 9"
        '
        'MenuNew16x16
        '
        Me.MenuNew16x16.Index = 1
        Me.MenuNew16x16.Text = "New 16 x 16"
        '
        'mnuLoad
        '
        Me.mnuLoad.Index = 2
        Me.mnuLoad.Text = "Load"
        '
        'mnuSave
        '
        Me.mnuSave.Index = 3
        Me.mnuSave.Text = "Save"
        '
        'mnuPrint
        '
        Me.mnuPrint.Index = 4
        Me.mnuPrint.Text = "Print"
        '
        'mnuQuit
        '
        Me.mnuQuit.Index = 5
        Me.mnuQuit.Text = "Quit"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelp, Me.mnuAbout})
        Me.MenuItem1.Text = "Help"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 0
        Me.mnuHelp.Text = "How to use"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 1
        Me.mnuAbout.Text = "About"
        '
        'cbShowPossibilities
        '
        Me.cbShowPossibilities.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbShowPossibilities.AutoSize = True
        Me.cbShowPossibilities.Location = New System.Drawing.Point(442, 16)
        Me.cbShowPossibilities.Name = "cbShowPossibilities"
        Me.cbShowPossibilities.Size = New System.Drawing.Size(109, 17)
        Me.cbShowPossibilities.TabIndex = 2
        Me.cbShowPossibilities.Text = "Show Possibilities"
        Me.cbShowPossibilities.UseVisualStyleBackColor = True
        '
        'btnSolve
        '
        Me.btnSolve.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSolve.Location = New System.Drawing.Point(438, 48)
        Me.btnSolve.Name = "btnSolve"
        Me.btnSolve.Size = New System.Drawing.Size(128, 27)
        Me.btnSolve.TabIndex = 4
        Me.btnSolve.Text = "Solve"
        Me.btnSolve.UseVisualStyleBackColor = True
        '
        'cbSingleStep
        '
        Me.cbSingleStep.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbSingleStep.AutoSize = True
        Me.cbSingleStep.Location = New System.Drawing.Point(584, 52)
        Me.cbSingleStep.Name = "cbSingleStep"
        Me.cbSingleStep.Size = New System.Drawing.Size(80, 17)
        Me.cbSingleStep.TabIndex = 5
        Me.cbSingleStep.Text = "Single Step"
        Me.cbSingleStep.UseVisualStyleBackColor = True
        '
        'btnShowHint
        '
        Me.btnShowHint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowHint.Location = New System.Drawing.Point(438, 94)
        Me.btnShowHint.Name = "btnShowHint"
        Me.btnShowHint.Size = New System.Drawing.Size(128, 27)
        Me.btnShowHint.TabIndex = 6
        Me.btnShowHint.Text = "Show Hint"
        Me.btnShowHint.UseVisualStyleBackColor = True
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'PrintDocument1
        '
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearch.Location = New System.Drawing.Point(438, 142)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(128, 27)
        Me.btnSearch.TabIndex = 7
        Me.btnSearch.Text = "Search for Solution"
        Me.btnSearch.UseVisualStyleBackColor = True
        Me.btnSearch.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(700, 396)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.btnShowHint)
        Me.Controls.Add(Me.cbSingleStep)
        Me.Controls.Add(Me.btnSolve)
        Me.Controls.Add(Me.cbShowPossibilities)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sudoku Solver"
        Me.Panel1.ResumeLayout(False)
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents MainMenu1 As MainMenu
    Friend WithEvents MenuItem4 As MenuItem
    Friend WithEvents mnuNew9x9 As MenuItem
    Friend WithEvents mnuLoad As MenuItem
    Friend WithEvents mnuSave As MenuItem
    Friend WithEvents mnuQuit As MenuItem
    Friend WithEvents MenuItem1 As MenuItem
    Friend WithEvents mnuHelp As MenuItem
    Friend WithEvents mnuAbout As MenuItem
    Friend WithEvents MenuNew16x16 As MenuItem
    Friend WithEvents cbShowPossibilities As CheckBox
    Friend WithEvents btnSolve As Button
    Friend WithEvents cbSingleStep As CheckBox
    Friend WithEvents btnShowHint As Button
    Friend WithEvents FileSystemWatcher1 As IO.FileSystemWatcher
    Friend WithEvents mnuPrint As MenuItem
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents btnSearch As Button
End Class
