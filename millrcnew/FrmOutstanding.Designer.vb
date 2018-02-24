<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmOutstanding
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmOutstanding))
        Me.RadioGodown = New System.Windows.Forms.RadioButton()
        Me.RadioTenant = New System.Windows.Forms.RadioButton()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'RadioGodown
        '
        Me.RadioGodown.AutoSize = True
        Me.RadioGodown.Location = New System.Drawing.Point(21, 30)
        Me.RadioGodown.Name = "RadioGodown"
        Me.RadioGodown.Size = New System.Drawing.Size(86, 17)
        Me.RadioGodown.TabIndex = 0
        Me.RadioGodown.TabStop = True
        Me.RadioGodown.Text = "Godownwise"
        Me.RadioGodown.UseVisualStyleBackColor = True
        '
        'RadioTenant
        '
        Me.RadioTenant.AutoSize = True
        Me.RadioTenant.Location = New System.Drawing.Point(122, 30)
        Me.RadioTenant.Name = "RadioTenant"
        Me.RadioTenant.Size = New System.Drawing.Size(80, 17)
        Me.RadioTenant.TabIndex = 2
        Me.RadioTenant.TabStop = True
        Me.RadioTenant.Text = "Tenantwise"
        Me.RadioTenant.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(162, 126)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(50, 22)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(93, 126)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(50, 22)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Print"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(21, 126)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(50, 22)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "View"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(139, 64)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(132, 20)
        Me.TextBox5.TabIndex = 151
        Me.TextBox5.Text = "Outstanding"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(22, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 41)
        Me.Label6.TabIndex = 152
        Me.Label6.Text = "Report File Name (Without Extn)"
        '
        'FrmOutstanding
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(383, 165)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.RadioTenant)
        Me.Controls.Add(Me.RadioGodown)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmOutstanding"
        Me.Text = "Outstanding Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents RadioGodown As RadioButton
    Friend WithEvents RadioTenant As RadioButton
    Friend WithEvents Button3 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents Label6 As Label
End Class
