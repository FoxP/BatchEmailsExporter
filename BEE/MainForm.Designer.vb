<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.cbStart = New System.Windows.Forms.Button()
        Me.cbStop = New System.Windows.Forms.Button()
        Me.progressBarCompletion = New System.Windows.Forms.ProgressBar()
        Me.cbAttachments = New System.Windows.Forms.CheckBox()
        Me.rbMsg = New System.Windows.Forms.RadioButton()
        Me.rbTxt = New System.Windows.Forms.RadioButton()
        Me.rbHtml = New System.Windows.Forms.RadioButton()
        Me.rbMht = New System.Windows.Forms.RadioButton()
        Me.cbAbout = New System.Windows.Forms.Button()
        Me.rbDoc = New System.Windows.Forms.RadioButton()
        Me.tbFolderPath = New System.Windows.Forms.TextBox()
        Me.cbFolderSelect = New System.Windows.Forms.Button()
        Me.rbRtf = New System.Windows.Forms.RadioButton()
        Me.ttHoverInfo = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'cbStart
        '
        Me.cbStart.Location = New System.Drawing.Point(9, 80)
        Me.cbStart.Name = "cbStart"
        Me.cbStart.Size = New System.Drawing.Size(200, 23)
        Me.cbStart.TabIndex = 10
        Me.cbStart.Text = "Start"
        Me.cbStart.UseVisualStyleBackColor = True
        '
        'cbStop
        '
        Me.cbStop.Enabled = False
        Me.cbStop.Location = New System.Drawing.Point(218, 80)
        Me.cbStop.Name = "cbStop"
        Me.cbStop.Size = New System.Drawing.Size(200, 23)
        Me.cbStop.TabIndex = 11
        Me.cbStop.Text = "Stop"
        Me.cbStop.UseVisualStyleBackColor = True
        '
        'progressBarCompletion
        '
        Me.progressBarCompletion.Location = New System.Drawing.Point(-1, 110)
        Me.progressBarCompletion.MarqueeAnimationSpeed = 10
        Me.progressBarCompletion.Name = "progressBarCompletion"
        Me.progressBarCompletion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.progressBarCompletion.Size = New System.Drawing.Size(462, 10)
        Me.progressBarCompletion.Step = 1
        Me.progressBarCompletion.TabIndex = 13
        '
        'cbAttachments
        '
        Me.cbAttachments.AutoSize = True
        Me.cbAttachments.Location = New System.Drawing.Point(341, 13)
        Me.cbAttachments.Name = "cbAttachments"
        Me.cbAttachments.Size = New System.Drawing.Size(115, 17)
        Me.cbAttachments.TabIndex = 7
        Me.cbAttachments.Text = " Save attachments"
        Me.cbAttachments.UseVisualStyleBackColor = True
        '
        'rbMsg
        '
        Me.rbMsg.AutoSize = True
        Me.rbMsg.Location = New System.Drawing.Point(10, 12)
        Me.rbMsg.Name = "rbMsg"
        Me.rbMsg.Size = New System.Drawing.Size(44, 17)
        Me.rbMsg.TabIndex = 1
        Me.rbMsg.TabStop = True
        Me.rbMsg.Text = "msg"
        Me.rbMsg.UseVisualStyleBackColor = True
        '
        'rbTxt
        '
        Me.rbTxt.AutoSize = True
        Me.rbTxt.Location = New System.Drawing.Point(69, 12)
        Me.rbTxt.Name = "rbTxt"
        Me.rbTxt.Size = New System.Drawing.Size(36, 17)
        Me.rbTxt.TabIndex = 2
        Me.rbTxt.TabStop = True
        Me.rbTxt.Text = "txt"
        Me.rbTxt.UseVisualStyleBackColor = True
        '
        'rbHtml
        '
        Me.rbHtml.AutoSize = True
        Me.rbHtml.Location = New System.Drawing.Point(119, 12)
        Me.rbHtml.Name = "rbHtml"
        Me.rbHtml.Size = New System.Drawing.Size(44, 17)
        Me.rbHtml.TabIndex = 3
        Me.rbHtml.TabStop = True
        Me.rbHtml.Text = "html"
        Me.rbHtml.UseVisualStyleBackColor = True
        '
        'rbMht
        '
        Me.rbMht.AutoSize = True
        Me.rbMht.Location = New System.Drawing.Point(177, 12)
        Me.rbMht.Name = "rbMht"
        Me.rbMht.Size = New System.Drawing.Size(42, 17)
        Me.rbMht.TabIndex = 4
        Me.rbMht.TabStop = True
        Me.rbMht.Text = "mht"
        Me.rbMht.UseVisualStyleBackColor = True
        '
        'cbAbout
        '
        Me.cbAbout.Location = New System.Drawing.Point(427, 80)
        Me.cbAbout.Name = "cbAbout"
        Me.cbAbout.Size = New System.Drawing.Size(24, 23)
        Me.cbAbout.TabIndex = 12
        Me.cbAbout.Text = "?"
        Me.cbAbout.UseVisualStyleBackColor = True
        '
        'rbDoc
        '
        Me.rbDoc.AutoSize = True
        Me.rbDoc.Location = New System.Drawing.Point(233, 12)
        Me.rbDoc.Name = "rbDoc"
        Me.rbDoc.Size = New System.Drawing.Size(43, 17)
        Me.rbDoc.TabIndex = 5
        Me.rbDoc.TabStop = True
        Me.rbDoc.Text = "doc"
        Me.rbDoc.UseVisualStyleBackColor = True
        '
        'tbFolderPath
        '
        Me.tbFolderPath.Location = New System.Drawing.Point(10, 44)
        Me.tbFolderPath.Name = "tbFolderPath"
        Me.tbFolderPath.ReadOnly = True
        Me.tbFolderPath.Size = New System.Drawing.Size(332, 20)
        Me.tbFolderPath.TabIndex = 8
        Me.tbFolderPath.TabStop = False
        '
        'cbFolderSelect
        '
        Me.cbFolderSelect.Location = New System.Drawing.Point(352, 43)
        Me.cbFolderSelect.Name = "cbFolderSelect"
        Me.cbFolderSelect.Size = New System.Drawing.Size(99, 22)
        Me.cbFolderSelect.TabIndex = 9
        Me.cbFolderSelect.Text = "Output folder"
        Me.cbFolderSelect.UseVisualStyleBackColor = True
        '
        'rbRtf
        '
        Me.rbRtf.AutoSize = True
        Me.rbRtf.Location = New System.Drawing.Point(290, 12)
        Me.rbRtf.Name = "rbRtf"
        Me.rbRtf.Size = New System.Drawing.Size(34, 17)
        Me.rbRtf.TabIndex = 6
        Me.rbRtf.TabStop = True
        Me.rbRtf.Text = "rtf"
        Me.rbRtf.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(460, 119)
        Me.Controls.Add(Me.rbRtf)
        Me.Controls.Add(Me.cbFolderSelect)
        Me.Controls.Add(Me.tbFolderPath)
        Me.Controls.Add(Me.rbDoc)
        Me.Controls.Add(Me.cbAbout)
        Me.Controls.Add(Me.progressBarCompletion)
        Me.Controls.Add(Me.rbMht)
        Me.Controls.Add(Me.rbHtml)
        Me.Controls.Add(Me.rbTxt)
        Me.Controls.Add(Me.rbMsg)
        Me.Controls.Add(Me.cbAttachments)
        Me.Controls.Add(Me.cbStop)
        Me.Controls.Add(Me.cbStart)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.Text = "BEE : BatchEmailsExporter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbStart As System.Windows.Forms.Button
    Friend WithEvents cbStop As System.Windows.Forms.Button
    Friend WithEvents progressBarCompletion As System.Windows.Forms.ProgressBar
    Friend WithEvents cbAttachments As System.Windows.Forms.CheckBox
    Friend WithEvents rbMsg As System.Windows.Forms.RadioButton
    Friend WithEvents rbTxt As System.Windows.Forms.RadioButton
    Friend WithEvents rbHtml As System.Windows.Forms.RadioButton
    Friend WithEvents rbMht As System.Windows.Forms.RadioButton
    Friend WithEvents cbAbout As System.Windows.Forms.Button
    Friend WithEvents rbDoc As System.Windows.Forms.RadioButton
    Friend WithEvents tbFolderPath As System.Windows.Forms.TextBox
    Friend WithEvents cbFolderSelect As System.Windows.Forms.Button
    Friend WithEvents rbRtf As System.Windows.Forms.RadioButton
    Friend WithEvents ttHoverInfo As System.Windows.Forms.ToolTip

End Class
