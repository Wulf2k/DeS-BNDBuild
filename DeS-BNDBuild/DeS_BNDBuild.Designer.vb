<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Des_BNDBuild
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtBNDfile = New System.Windows.Forms.TextBox()
        Me.lblGAFile = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.btnRebuild = New System.Windows.Forms.Button()
        Me.btnExtract = New System.Windows.Forms.Button()
        Me.txtInfo = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtBNDfile
        '
        Me.txtBNDfile.AllowDrop = True
        Me.txtBNDfile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBNDfile.Location = New System.Drawing.Point(44, 6)
        Me.txtBNDfile.Multiline = True
        Me.txtBNDfile.Name = "txtBNDfile"
        Me.txtBNDfile.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBNDfile.Size = New System.Drawing.Size(635, 166)
        Me.txtBNDfile.TabIndex = 26
        '
        'lblGAFile
        '
        Me.lblGAFile.AutoSize = True
        Me.lblGAFile.Location = New System.Drawing.Point(12, 9)
        Me.lblGAFile.Name = "lblGAFile"
        Me.lblGAFile.Size = New System.Drawing.Size(31, 13)
        Me.lblGAFile.TabIndex = 28
        Me.lblGAFile.Text = "Files:"
        '
        'btnBrowse
        '
        Me.btnBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowse.Location = New System.Drawing.Point(684, 6)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowse.TabIndex = 27
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'btnRebuild
        '
        Me.btnRebuild.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRebuild.Location = New System.Drawing.Point(684, 73)
        Me.btnRebuild.Name = "btnRebuild"
        Me.btnRebuild.Size = New System.Drawing.Size(75, 23)
        Me.btnRebuild.TabIndex = 30
        Me.btnRebuild.Text = "Rebuild"
        Me.btnRebuild.UseVisualStyleBackColor = True
        '
        'btnExtract
        '
        Me.btnExtract.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExtract.Location = New System.Drawing.Point(684, 44)
        Me.btnExtract.Name = "btnExtract"
        Me.btnExtract.Size = New System.Drawing.Size(75, 23)
        Me.btnExtract.TabIndex = 29
        Me.btnExtract.Text = "Extract"
        Me.btnExtract.UseVisualStyleBackColor = True
        '
        'txtInfo
        '
        Me.txtInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInfo.Location = New System.Drawing.Point(12, 182)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInfo.Size = New System.Drawing.Size(743, 246)
        Me.txtInfo.TabIndex = 31
        '
        'Des_BNDBuild
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(767, 440)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.btnRebuild)
        Me.Controls.Add(Me.btnExtract)
        Me.Controls.Add(Me.txtBNDfile)
        Me.Controls.Add(Me.lblGAFile)
        Me.Controls.Add(Me.btnBrowse)
        Me.DoubleBuffered = True
        Me.MinimumSize = New System.Drawing.Size(360, 179)
        Me.Name = "Des_BNDBuild"
        Me.Text = "Wulf's BND Rebuilder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtBNDfile As System.Windows.Forms.TextBox
    Friend WithEvents lblGAFile As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnRebuild As System.Windows.Forms.Button
    Friend WithEvents btnExtract As System.Windows.Forms.Button
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox

End Class
