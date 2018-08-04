<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReupdate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmReupdate))
        Me.Button4 = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.ChkLogo = New System.Windows.Forms.CheckBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(31, 91)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 159
        Me.Button4.Text = "Reupdate"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(128, 91)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(51, 23)
        Me.cmdClose.TabIndex = 160
        Me.cmdClose.Text = "E&xit"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'ChkLogo
        '
        Me.ChkLogo.AutoSize = True
        Me.ChkLogo.Location = New System.Drawing.Point(31, 26)
        Me.ChkLogo.Name = "ChkLogo"
        Me.ChkLogo.Size = New System.Drawing.Size(75, 17)
        Me.ChkLogo.TabIndex = 161
        Me.ChkLogo.Text = "With Logo"
        Me.ChkLogo.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(32, 48)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(273, 28)
        Me.ProgressBar1.TabIndex = 162
        '
        'FrmReupdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(319, 125)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.ChkLogo)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Button4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmReupdate"
        Me.Text = "Reupdate"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button4 As Button
    Friend WithEvents cmdClose As Button
    Friend WithEvents ChkLogo As CheckBox
    Friend WithEvents ProgressBar1 As ProgressBar
End Class
