<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnMerge = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCircleLeaderReportLocation = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCIReportFileLocation = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtMergedFileName = New System.Windows.Forms.TextBox()
        Me.txtMergedFileLocation = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnMerge
        '
        Me.btnMerge.Location = New System.Drawing.Point(244, 219)
        Me.btnMerge.Name = "btnMerge"
        Me.btnMerge.Size = New System.Drawing.Size(75, 23)
        Me.btnMerge.TabIndex = 0
        Me.btnMerge.Text = "Merge Files"
        Me.btnMerge.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(48, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(161, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "CircleLeaderReport File Location"
        '
        'txtCircleLeaderReportLocation
        '
        Me.txtCircleLeaderReportLocation.Location = New System.Drawing.Point(215, 61)
        Me.txtCircleLeaderReportLocation.Name = "txtCircleLeaderReportLocation"
        Me.txtCircleLeaderReportLocation.Size = New System.Drawing.Size(325, 20)
        Me.txtCircleLeaderReportLocation.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(51, 105)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "CI ReportFile Location"
        '
        'txtCIReportFileLocation
        '
        Me.txtCIReportFileLocation.Location = New System.Drawing.Point(215, 97)
        Me.txtCIReportFileLocation.Name = "txtCIReportFileLocation"
        Me.txtCIReportFileLocation.Size = New System.Drawing.Size(325, 20)
        Me.txtCIReportFileLocation.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(54, 157)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "MergedFileName"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(57, 193)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(103, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Merged FileLocation"
        '
        'txtMergedFileName
        '
        Me.txtMergedFileName.Location = New System.Drawing.Point(215, 149)
        Me.txtMergedFileName.Name = "txtMergedFileName"
        Me.txtMergedFileName.Size = New System.Drawing.Size(166, 20)
        Me.txtMergedFileName.TabIndex = 7
        '
        'txtMergedFileLocation
        '
        Me.txtMergedFileLocation.Location = New System.Drawing.Point(215, 185)
        Me.txtMergedFileLocation.Name = "txtMergedFileLocation"
        Me.txtMergedFileLocation.Size = New System.Drawing.Size(325, 20)
        Me.txtMergedFileLocation.TabIndex = 8
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(583, 247)
        Me.Controls.Add(Me.txtMergedFileLocation)
        Me.Controls.Add(Me.txtMergedFileName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtCIReportFileLocation)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtCircleLeaderReportLocation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnMerge)
        Me.Name = "Form1"
        Me.Text = "EXCEL MERGE UTILITY"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnMerge As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCircleLeaderReportLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCIReportFileLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtMergedFileName As System.Windows.Forms.TextBox
    Friend WithEvents txtMergedFileLocation As System.Windows.Forms.TextBox

End Class
