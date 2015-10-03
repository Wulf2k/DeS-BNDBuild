<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DesBNDBuild
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
        Me.btnCompress = New System.Windows.Forms.Button()
        Me.btnDecompress = New System.Windows.Forms.Button()
        Me.btnBackup = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtBNDfile
        '
        Me.txtBNDfile.AllowDrop = True
        Me.txtBNDfile.Location = New System.Drawing.Point(44, 6)
        Me.txtBNDfile.Name = "txtBNDfile"
        Me.txtBNDfile.Size = New System.Drawing.Size(440, 20)
        Me.txtBNDfile.TabIndex = 26
        '
        'lblGAFile
        '
        Me.lblGAFile.AutoSize = True
        Me.lblGAFile.Location = New System.Drawing.Point(12, 9)
        Me.lblGAFile.Name = "lblGAFile"
        Me.lblGAFile.Size = New System.Drawing.Size(26, 13)
        Me.lblGAFile.TabIndex = 28
        Me.lblGAFile.Text = "File:"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(487, 4)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowse.TabIndex = 27
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'btnRebuild
        '
        Me.btnRebuild.Location = New System.Drawing.Point(487, 33)
        Me.btnRebuild.Name = "btnRebuild"
        Me.btnRebuild.Size = New System.Drawing.Size(75, 23)
        Me.btnRebuild.TabIndex = 30
        Me.btnRebuild.Text = "Rebuild"
        Me.btnRebuild.UseVisualStyleBackColor = True
        '
        'btnExtract
        '
        Me.btnExtract.Location = New System.Drawing.Point(409, 33)
        Me.btnExtract.Name = "btnExtract"
        Me.btnExtract.Size = New System.Drawing.Size(75, 23)
        Me.btnExtract.TabIndex = 29
        Me.btnExtract.Text = "Extract"
        Me.btnExtract.UseVisualStyleBackColor = True
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(12, 62)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInfo.Size = New System.Drawing.Size(548, 142)
        Me.txtInfo.TabIndex = 31
        '
        'btnCompress
        '
        Me.btnCompress.Enabled = False
        Me.btnCompress.Location = New System.Drawing.Point(92, 33)
        Me.btnCompress.Name = "btnCompress"
        Me.btnCompress.Size = New System.Drawing.Size(75, 23)
        Me.btnCompress.TabIndex = 33
        Me.btnCompress.Text = "Compress"
        Me.btnCompress.UseVisualStyleBackColor = True
        '
        'btnDecompress
        '
        Me.btnDecompress.Enabled = False
        Me.btnDecompress.Location = New System.Drawing.Point(14, 33)
        Me.btnDecompress.Name = "btnDecompress"
        Me.btnDecompress.Size = New System.Drawing.Size(75, 23)
        Me.btnDecompress.TabIndex = 32
        Me.btnDecompress.Text = "Decompress"
        Me.btnDecompress.UseVisualStyleBackColor = True
        '
        'btnBackup
        '
        Me.btnBackup.Location = New System.Drawing.Point(288, 33)
        Me.btnBackup.Name = "btnBackup"
        Me.btnBackup.Size = New System.Drawing.Size(75, 23)
        Me.btnBackup.TabIndex = 34
        Me.btnBackup.Text = "Backup"
        Me.btnBackup.UseVisualStyleBackColor = True
        '
        'DesBNDBuild
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(572, 216)
        Me.Controls.Add(Me.btnBackup)
        Me.Controls.Add(Me.btnCompress)
        Me.Controls.Add(Me.btnDecompress)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.btnRebuild)
        Me.Controls.Add(Me.btnExtract)
        Me.Controls.Add(Me.txtBNDfile)
        Me.Controls.Add(Me.lblGAFile)
        Me.Controls.Add(Me.btnBrowse)
        Me.Name = "DesBNDBuild"
        Me.Text = "Wulf's DeS BND Rebuilder 0.800"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtBNDfile As System.Windows.Forms.TextBox
    Friend WithEvents lblGAFile As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnRebuild As System.Windows.Forms.Button
    Friend WithEvents btnExtract As System.Windows.Forms.Button
    Friend WithEvents txtInfo As System.Windows.Forms.TextBox
    Friend WithEvents btnCompress As System.Windows.Forms.Button
    Friend WithEvents btnDecompress As System.Windows.Forms.Button
    Friend WithEvents btnBackup As System.Windows.Forms.Button

End Class
